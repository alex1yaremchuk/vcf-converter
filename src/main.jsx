import React, { useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import {
  contactsToRows,
  decodeFileText,
  getDefaultColumns,
  parseVcf,
} from "./vcf.js";
import "./styles.css";

function App() {
  const [fileName, setFileName] = useState("");
  const [contacts, setContacts] = useState([]);
  const [stats, setStats] = useState(null);
  const [columns, setColumns] = useState(getDefaultColumns);
  const [separator, setSeparator] = useState(" | ");
  const [error, setError] = useState("");
  const [dragIndex, setDragIndex] = useState(null);
  const fileInputRef = useRef(null);

  const enabledColumns = useMemo(
    () => columns.filter((column) => column.enabled),
    [columns],
  );

  const previewRows = useMemo(
    () => contactsToRows(contacts, columns, separator),
    [contacts, columns, separator],
  );

  async function handleFile(file) {
    if (!file) return;
    setError("");
    setFileName(file.name);

    try {
      const buffer = await file.arrayBuffer();
      const text = decodeFileText(buffer);
      const result = parseVcf(text);
      setContacts(result.contacts);
      setStats(result.stats);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Не удалось разобрать файл");
      setContacts([]);
      setStats(null);
    }
  }

  function toggleColumn(id) {
    setColumns((current) =>
      current.map((column) =>
        column.id === id ? { ...column, enabled: !column.enabled } : column,
      ),
    );
  }

  function renameColumn(id, label) {
    setColumns((current) =>
      current.map((column) =>
        column.id === id ? { ...column, label } : column,
      ),
    );
  }

  function moveColumn(from, to) {
    if (from === to || from == null || to == null) return;
    setColumns((current) => {
      const next = [...current];
      const [item] = next.splice(from, 1);
      next.splice(to, 0, item);
      return next;
    });
  }

  function resetColumns() {
    setColumns(getDefaultColumns());
  }

  function setAllColumns(enabled) {
    setColumns((current) => current.map((column) => ({ ...column, enabled })));
  }

  async function exportXlsx() {
    const rows = contactsToRows(contacts, columns, separator);
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Контакты", {
      views: [{ state: "frozen", ySplit: 1 }],
    });

    worksheet.columns = enabledColumns.map((column) => ({
      header: column.label,
      key: column.label,
      width: Math.min(Math.max(column.label.length + 8, 16), 40),
      style: { numFmt: "@" },
    }));

    worksheet.addRows(rows);
    worksheet.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: Math.max(rows.length + 1, 1), column: enabledColumns.length },
    };

    const header = worksheet.getRow(1);
    header.font = { bold: true, color: { argb: "FFFFFFFF" } };
    header.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF24394F" },
    };

    worksheet.eachRow((row) => {
      row.eachCell((cell) => {
        cell.alignment = { vertical: "top" };
        cell.border = {
          top: { style: "thin", color: { argb: "FFD8E0E8" } },
          left: { style: "thin", color: { argb: "FFD8E0E8" } },
          bottom: { style: "thin", color: { argb: "FFD8E0E8" } },
          right: { style: "thin", color: { argb: "FFD8E0E8" } },
        };
      });
    });

    const buffer = await workbook.xlsx.writeBuffer();
    downloadBlob(
      buffer,
      outputName("xlsx"),
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );
  }

  function exportCsv() {
    const rows = contactsToRows(contacts, columns, separator);
    const csv = rowsToCsv(rows, enabledColumns.map((column) => column.label));
    downloadBlob(`\ufeff${csv}`, outputName("csv"), "text/csv;charset=utf-8");
  }

  async function exportDocx(format) {
    const rows = contactsToRows(contacts, columns, separator);
    const blob = await createDocx({
      title: fileName.replace(/\.[^.]+$/, "") || "Контакты",
      format,
      rows,
      columns: enabledColumns.map((column) => column.label),
    });
    downloadBlob(
      blob,
      outputName(format === "list" ? "list.docx" : "table.docx"),
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    );
  }

  function outputName(ext) {
    const base = fileName.replace(/\.[^.]+$/, "") || "contacts";
    return `${base}.${ext}`;
  }

  const canExport = contacts.length > 0 && enabledColumns.length > 0;

  return (
    <main className="app">
      <header className="topbar">
        <div>
          <h1>VCF Converter</h1>
          <p>Локальная конвертация контактов в XLSX или CSV прямо в браузере.</p>
        </div>
        <div className="privacy">Данные не отправляются на сервер</div>
      </header>

      <section className="workspace">
        <aside className="side">
          <DropZone
            fileName={fileName}
            inputRef={fileInputRef}
            onFile={handleFile}
          />

          {error && <div className="error">{error}</div>}

          <Stats stats={stats} contacts={contacts} />

          <div className="panel">
            <div className="panelHeader">
              <h2>Экспорт</h2>
            </div>
            <label className="field">
              <span>Разделитель нескольких значений</span>
              <select value={separator} onChange={(event) => setSeparator(event.target.value)}>
                <option value=" | ">Вертикальная черта</option>
                <option value=", ">Запятая</option>
                <option value="\n">Новая строка</option>
              </select>
            </label>
            <div className="actions">
              <button type="button" onClick={exportXlsx} disabled={!canExport}>
                XLSX
              </button>
              <button type="button" onClick={exportCsv} disabled={!canExport}>
                CSV
              </button>
              <button type="button" onClick={() => exportDocx("list")} disabled={!canExport}>
                Word список
              </button>
              <button type="button" onClick={() => exportDocx("table")} disabled={!canExport}>
                Word таблица
              </button>
            </div>
          </div>
        </aside>

        <section className="mainPanel">
          <div className="panel columnsPanel">
            <div className="panelHeader">
              <h2>Колонки</h2>
              <div className="headerActions">
                <button type="button" className="ghost" onClick={() => setAllColumns(true)}>
                  Все
                </button>
                <button type="button" className="ghost" onClick={() => setAllColumns(false)}>
                  Ничего
                </button>
                <button type="button" className="ghost" onClick={resetColumns}>
                  Сбросить
                </button>
              </div>
            </div>
            <div className="columnList">
              {columns.map((column, index) => (
                <div
                  key={column.id}
                  className={`columnItem ${dragIndex === index ? "dragging" : ""}`}
                  draggable
                  onDragStart={() => setDragIndex(index)}
                  onDragOver={(event) => event.preventDefault()}
                  onDrop={() => {
                    moveColumn(dragIndex, index);
                    setDragIndex(null);
                  }}
                  onDragEnd={() => setDragIndex(null)}
                >
                  <span className="handle" aria-hidden="true">::</span>
                  <input
                    type="checkbox"
                    checked={column.enabled}
                    onChange={() => toggleColumn(column.id)}
                    aria-label={`Включить ${column.label}`}
                  />
                  <input
                    type="text"
                    value={column.label}
                    onChange={(event) => renameColumn(column.id, event.target.value)}
                  />
                </div>
              ))}
            </div>
          </div>

          <div className="panel previewPanel">
            <div className="panelHeader">
              <h2>Предпросмотр</h2>
              <span>
                {contacts.length
                  ? `${contacts.length} из ${contacts.length}`
                  : "файл не выбран"}
              </span>
            </div>
            <PreviewTable rows={previewRows} columns={enabledColumns} />
          </div>
        </section>
      </section>
    </main>
  );
}

function DropZone({ fileName, inputRef, onFile }) {
  const [active, setActive] = useState(false);

  return (
    <div
      className={`dropZone ${active ? "active" : ""}`}
      onDragOver={(event) => {
        event.preventDefault();
        setActive(true);
      }}
      onDragLeave={() => setActive(false)}
      onDrop={(event) => {
        event.preventDefault();
        setActive(false);
        onFile(event.dataTransfer.files[0]);
      }}
    >
      <input
        ref={inputRef}
        type="file"
        accept=".vcf,text/vcard,text/x-vcard"
        onChange={(event) => onFile(event.target.files[0])}
      />
      <button type="button" onClick={() => inputRef.current?.click()}>
        Выбрать VCF
      </button>
      <p>{fileName || "Перетащи файл контактов сюда"}</p>
    </div>
  );
}

function Stats({ stats, contacts }) {
  if (!stats) {
    return (
      <div className="panel stats empty">
        <h2>Сводка</h2>
        <p>После загрузки здесь появятся найденные контакты и поля.</p>
      </div>
    );
  }

  return (
    <div className="panel stats">
      <h2>Сводка</h2>
      <dl>
        <div>
          <dt>Контакты</dt>
          <dd>{contacts.length}</dd>
        </div>
        <div>
          <dt>Телефоны</dt>
          <dd>{stats.phones}</dd>
        </div>
        <div>
          <dt>Email</dt>
          <dd>{stats.emails}</dd>
        </div>
      </dl>
      <p className="small">
        Поля: {Object.keys(stats.fields).slice(0, 12).join(", ") || "нет"}
      </p>
    </div>
  );
}

function PreviewTable({ rows, columns }) {
  if (!rows.length) {
    return <div className="emptyPreview">Загрузи `.vcf`, чтобы увидеть таблицу.</div>;
  }

  return (
    <div className="tableWrap">
      <table>
        <thead>
          <tr>
            {columns.map((column) => (
              <th key={column.id}>{column.label}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, index) => (
            <tr key={index}>
              {columns.map((column) => (
                <td key={column.id}>{row[column.label]}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function downloadBlob(content, name, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = name;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function rowsToCsv(rows, headers) {
  const lines = [headers.map(escapeCsvCell).join(";")];
  for (const row of rows) {
    lines.push(headers.map((header) => escapeCsvCell(row[header] || "")).join(";"));
  }
  return lines.join("\r\n");
}

function escapeCsvCell(value) {
  const text = String(value ?? "");
  if (/[;"\r\n]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`;
  }
  return text;
}

async function createDocx({ title, format, rows, columns }) {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypesXml());
  zip.folder("_rels").file(".rels", rootRelsXml());
  zip.folder("word").file("document.xml", documentXml({ title, format, rows, columns }));
  zip.folder("word").folder("_rels").file("document.xml.rels", documentRelsXml());
  zip.folder("word").file("styles.xml", stylesXml());
  return zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

function documentXml({ title, format, rows, columns }) {
  const body =
    format === "table"
      ? docxTable(rows, columns)
      : rows.map((row, index) => docxContactBlock(row, columns, index)).join("");

  return xmlDoc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${paragraph(title, "Title")}
    ${paragraph(`Контактов: ${rows.length}`, "Subtitle")}
    ${body}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="850" w:header="708" w:footer="708" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`);
}

function docxContactBlock(row, columns, index) {
  const name = firstNonEmpty(row["Отображаемое имя"], row["Имя"], row["Фамилия"]) || `Контакт ${index + 1}`;
  const lines = columns
    .filter((column) => column !== "Отображаемое имя" && row[column])
    .map((column) => paragraph(`${column}: ${row[column]}`, "ContactLine"))
    .join("");
  return `${paragraph(name, "ContactName")}${lines}${paragraph("", "Spacer")}`;
}

function docxTable(rows, columns) {
  const headerCells = columns.map((column) => tableCell(column, true)).join("");
  const bodyRows = rows
    .map((row) => `<w:tr>${columns.map((column) => tableCell(row[column] || "", false)).join("")}</w:tr>`)
    .join("");
  return `<w:tbl>
    <w:tblPr>
      <w:tblStyle w:val="TableGrid"/>
      <w:tblW w:w="0" w:type="auto"/>
      <w:tblBorders>
        <w:top w:val="single" w:sz="4" w:space="0" w:color="D8E0E8"/>
        <w:left w:val="single" w:sz="4" w:space="0" w:color="D8E0E8"/>
        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="D8E0E8"/>
        <w:right w:val="single" w:sz="4" w:space="0" w:color="D8E0E8"/>
        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="D8E0E8"/>
        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="D8E0E8"/>
      </w:tblBorders>
    </w:tblPr>
    <w:tr>${headerCells}</w:tr>
    ${bodyRows}
  </w:tbl>`;
}

function tableCell(value, header) {
  return `<w:tc>
    <w:tcPr><w:tcW w:w="1800" w:type="dxa"/>${header ? '<w:shd w:fill="24394F"/>' : ""}</w:tcPr>
    ${paragraph(value, header ? "TableHeader" : "TableCell")}
  </w:tc>`;
}

function paragraph(text, style) {
  const lines = String(text ?? "").split("\n");
  const run = lines.map((line, index) => `${index ? "<w:br/>" : ""}<w:t xml:space="preserve">${escapeXml(line)}</w:t>`).join("");
  return `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr><w:r>${run || "<w:t></w:t>"}</w:r></w:p>`;
}

function firstNonEmpty(...values) {
  return values.find((value) => String(value || "").trim());
}

function escapeXml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function xmlDoc(value) {
  return value.trim();
}

function contentTypesXml() {
  return xmlDoc(`<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`);
}

function rootRelsXml() {
  return xmlDoc(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);
}

function documentRelsXml() {
  return xmlDoc(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`);
}

function stylesXml() {
  return xmlDoc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:after="120"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="36"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Subtitle">
    <w:name w:val="Subtitle"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:after="240"/></w:pPr>
    <w:rPr><w:color w:val="5D6978"/><w:sz w:val="22"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ContactName">
    <w:name w:val="ContactName"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="120" w:after="40"/></w:pPr>
    <w:rPr><w:b/><w:color w:val="1F4E78"/><w:sz w:val="24"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ContactLine">
    <w:name w:val="ContactLine"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:after="20"/></w:pPr>
    <w:rPr><w:sz w:val="20"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Spacer">
    <w:name w:val="Spacer"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:after="120"/></w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="TableHeader">
    <w:name w:val="TableHeader"/>
    <w:basedOn w:val="Normal"/>
    <w:rPr><w:b/><w:color w:val="FFFFFF"/><w:sz w:val="18"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="TableCell">
    <w:name w:val="TableCell"/>
    <w:basedOn w:val="Normal"/>
    <w:rPr><w:sz w:val="18"/></w:rPr>
  </w:style>
</w:styles>`);
}

createRoot(document.getElementById("root")).render(<App />);
