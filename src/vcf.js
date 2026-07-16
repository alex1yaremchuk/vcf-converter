const DEFAULT_COLUMNS = [
  { id: "displayName", label: "Отображаемое имя", enabled: true },
  { id: "firstName", label: "Имя", enabled: true },
  { id: "lastName", label: "Фамилия", enabled: true },
  { id: "nickname", label: "Никнейм", enabled: true },
  { id: "phones", label: "Телефоны", enabled: true },
  { id: "emails", label: "Email", enabled: true },
  { id: "organization", label: "Организация", enabled: true },
  { id: "title", label: "Должность", enabled: true },
  { id: "addresses", label: "Адреса", enabled: true },
  { id: "urls", label: "Сайт", enabled: true },
  { id: "birthday", label: "День рождения", enabled: true },
  { id: "notes", label: "Заметки", enabled: true },
  { id: "extra", label: "Дополнительно", enabled: true },
];

const FIELD_MAP = {
  FN: "displayName",
  N: "firstName/lastName",
  NICKNAME: "nickname",
  TEL: "phones",
  EMAIL: "emails",
  ORG: "organization",
  TITLE: "title",
  ADR: "addresses",
  URL: "urls",
  BDAY: "birthday",
  NOTE: "notes",
  "X-ANDROID-CUSTOM": "extra",
  "X-SAMSUNGADR": "addresses",
};

export function getDefaultColumns() {
  return DEFAULT_COLUMNS.map((column) => ({ ...column }));
}

export function decodeFileText(buffer) {
  const bytes = new Uint8Array(buffer);
  if (bytes[0] === 0xff && bytes[1] === 0xfe) {
    return new TextDecoder("utf-16le").decode(bytes);
  }
  if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    return new TextDecoder("utf-16be").decode(bytes);
  }
  return new TextDecoder("utf-8").decode(bytes);
}

export function parseVcf(text) {
  const normalized = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const lines = unfoldLines(normalized.split("\n"));
  const cards = collectCards(lines);
  const stats = {
    contacts: cards.length,
    phones: 0,
    emails: 0,
    fields: new Map(),
    unknownFields: new Map(),
  };

  const contacts = cards.map((card) => parseCard(card, stats));
  return {
    contacts,
    stats: {
      ...stats,
      fields: Object.fromEntries([...stats.fields.entries()].sort()),
      unknownFields: Object.fromEntries([...stats.unknownFields.entries()].sort()),
    },
  };
}

function unfoldLines(lines) {
  const logical = [];
  for (const line of lines) {
    if (logical.length && (line.startsWith(" ") || line.startsWith("\t"))) {
      logical[logical.length - 1] += line.slice(1);
      continue;
    }

    const previousHead = logical.length
      ? logical[logical.length - 1].split(":", 1)[0].toUpperCase()
      : "";

    if (
      logical.length &&
      previousHead.includes("ENCODING=QUOTED-PRINTABLE") &&
      logical[logical.length - 1].endsWith("=")
    ) {
      logical[logical.length - 1] = logical[logical.length - 1].slice(0, -1) + line;
      continue;
    }

    logical.push(line);
  }
  return logical;
}

function collectCards(lines) {
  const cards = [];
  let current = null;
  for (const line of lines) {
    const upper = line.toUpperCase();
    if (upper === "BEGIN:VCARD") {
      current = [];
    } else if (upper === "END:VCARD") {
      if (current) cards.push(current);
      current = null;
    } else if (current) {
      current.push(line);
    }
  }
  return cards;
}

function parseCard(lines, stats) {
  const data = {
    displayName: [],
    firstName: [],
    lastName: [],
    nickname: [],
    phones: [],
    emails: [],
    organization: [],
    title: [],
    addresses: [],
    urls: [],
    birthday: [],
    notes: [],
    extra: [],
  };

  for (const line of lines) {
    const parsed = parseLine(line);
    if (!parsed || parsed.name === "VERSION" || parsed.name === "PHOTO") continue;

    increment(stats.fields, parsed.name);
    if (!FIELD_MAP[parsed.name]) increment(stats.unknownFields, parsed.name);

    if (parsed.name === "FN") addUnique(data.displayName, parsed.value);
    else if (parsed.name === "N") parseName(parsed.value, data);
    else if (parsed.name === "NICKNAME") addUnique(data.nickname, parsed.value);
    else if (parsed.name === "TEL") addUnique(data.phones, parsed.value);
    else if (parsed.name === "EMAIL") addUnique(data.emails, parsed.value);
    else if (parsed.name === "ORG") addUnique(data.organization, joinStructured(parsed.value, " / "));
    else if (parsed.name === "TITLE") addUnique(data.title, parsed.value);
    else if (parsed.name === "ADR" || parsed.name === "X-SAMSUNGADR") {
      addUnique(data.addresses, joinStructured(parsed.value, ", "));
    } else if (parsed.name === "URL") addUnique(data.urls, parsed.value);
    else if (parsed.name === "BDAY") addUnique(data.birthday, parsed.value);
    else if (parsed.name === "NOTE") addUnique(data.notes, parsed.value);
    else if (parsed.name === "X-ANDROID-CUSTOM") parseAndroidCustom(parsed.value, data);
    else addUnique(data.extra, `${parsed.name}: ${parsed.value}`);
  }

  if (!data.displayName.length) {
    addUnique(data.displayName, data.nickname[0] || data.emails[0] || data.phones[0] || "");
  }

  stats.phones += data.phones.length;
  stats.emails += data.emails.length;

  return Object.fromEntries(
    Object.entries(data).map(([key, value]) => [key, value.filter(Boolean)]),
  );
}

function parseLine(line) {
  const separator = line.indexOf(":");
  if (separator === -1) return null;

  const left = line.slice(0, separator);
  const rawValue = line.slice(separator + 1);
  const chunks = left.split(";");
  const name = chunks[0].toUpperCase();
  const params = {};

  for (const chunk of chunks.slice(1)) {
    const equal = chunk.indexOf("=");
    if (equal === -1) {
      params.TYPE = params.TYPE ? `${params.TYPE},${chunk}` : chunk;
    } else {
      params[chunk.slice(0, equal).toUpperCase()] = chunk.slice(equal + 1);
    }
  }

  return { name, params, value: decodeValue(rawValue, params) };
}

function decodeValue(value, params) {
  const encoding = (params.ENCODING || "").toUpperCase();
  if (encoding === "QUOTED-PRINTABLE" || encoding === "QP") {
    const charset = params.CHARSET || "utf-8";
    value = decodeQuotedPrintable(value, charset);
  }

  return unescapeVcard(value);
}

function decodeQuotedPrintable(input, charset) {
  const bytes = [];
  for (let i = 0; i < input.length; i += 1) {
    if (input[i] === "=" && /^[0-9A-Fa-f]{2}$/.test(input.slice(i + 1, i + 3))) {
      bytes.push(parseInt(input.slice(i + 1, i + 3), 16));
      i += 2;
    } else {
      bytes.push(input.charCodeAt(i));
    }
  }

  try {
    return new TextDecoder(charset.toLowerCase()).decode(new Uint8Array(bytes));
  } catch {
    return new TextDecoder("utf-8").decode(new Uint8Array(bytes));
  }
}

function unescapeVcard(value) {
  return value
    .replace(/\\n/gi, " ")
    .replace(/\\;/g, ";")
    .replace(/\\,/g, ",")
    .replace(/\\\\/g, "\\")
    .trim();
}

function splitUnescaped(value, separator = ";") {
  const parts = [];
  let current = "";
  let escaped = false;

  for (const char of value) {
    if (escaped) {
      current += char;
      escaped = false;
    } else if (char === "\\") {
      escaped = true;
    } else if (char === separator) {
      parts.push(current);
      current = "";
    } else {
      current += char;
    }
  }

  parts.push(current);
  return parts;
}

function parseName(value, data) {
  const parts = splitUnescaped(value).map(unescapeVcard);
  addUnique(data.lastName, parts[0] || "");
  addUnique(data.firstName, parts[1] || "");
  addUnique(data.displayName, parts.filter(Boolean).join(" "));
}

function parseAndroidCustom(value, data) {
  const parts = splitUnescaped(value).map(unescapeVcard);
  if (parts[0] === "vnd.android.cursor.item/nickname" && parts[1]) {
    addUnique(data.nickname, parts[1]);
  } else {
    addUnique(data.extra, `X-ANDROID-CUSTOM: ${value}`);
  }
}

function joinStructured(value, separator) {
  const parts = splitUnescaped(value).map(unescapeVcard).filter(Boolean);
  return parts.length ? parts.join(separator) : value;
}

function addUnique(list, value) {
  const clean = String(value || "").trim();
  if (clean && !list.includes(clean)) list.push(clean);
}

function increment(map, key) {
  map.set(key, (map.get(key) || 0) + 1);
}

export function contactsToRows(contacts, columns, multiValueSeparator) {
  const enabled = columns.filter((column) => column.enabled);
  return contacts.map((contact) => {
    const row = {};
    for (const column of enabled) {
      row[column.label] = sanitizeForSpreadsheet(
        (contact[column.id] || []).join(multiValueSeparator),
      );
    }
    return row;
  });
}

export function sanitizeForSpreadsheet(value) {
  if (typeof value !== "string") return value;
  return /^[=+\-@]/.test(value) ? `'${value}` : value;
}
