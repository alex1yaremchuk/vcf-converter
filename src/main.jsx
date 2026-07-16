import React, { useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import ExcelJS from "exceljs";
import {
  contactsToRows,
  decodeFileText,
  getDefaultColumns,
  parseVcf,
} from "./vcf.js";
import "./styles.css";

const PREVIEW_LIMIT = 100;

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
    () => contactsToRows(contacts.slice(0, PREVIEW_LIMIT), columns, separator),
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
            </div>
          </div>
        </aside>

        <section className="mainPanel">
          <div className="panel columnsPanel">
            <div className="panelHeader">
              <h2>Колонки</h2>
              <button type="button" className="ghost" onClick={resetColumns}>
                Сбросить
              </button>
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
                  ? `${Math.min(PREVIEW_LIMIT, contacts.length)} из ${contacts.length}`
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

createRoot(document.getElementById("root")).render(<App />);
