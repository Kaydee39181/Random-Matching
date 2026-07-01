import * as XLSX from 'xlsx';

const EXCEL_MAX_SHEET_NAME_LENGTH = 31;
const DESCRIPTION_HEADER_KEY = 'description';
const EFFECTIVE_DATE_HEADER_KEY = 'effectivedate';
const RRN_HEADER = 'RRN';

function sanitizeSheetName(name, usedNames) {
  const cleanName = String(name)
    .replace(/[:\\/?*[\]]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, EXCEL_MAX_SHEET_NAME_LENGTH) || 'Sheet';

  let candidate = cleanName;
  let index = 2;

  while (usedNames.has(candidate)) {
    const suffix = ` ${index}`;
    candidate = `${cleanName.slice(0, EXCEL_MAX_SHEET_NAME_LENGTH - suffix.length)}${suffix}`;
    index += 1;
  }

  usedNames.add(candidate);
  return candidate;
}

function rowsToWorksheet(headers, rows) {
  const matrix = [
    headers,
    ...rows.map((row) => headers.map((header) => row[header] ?? '')),
  ];

  return XLSX.utils.aoa_to_sheet(matrix, {
    cellDates: true,
  });
}

function normalizeHeader(header) {
  return String(header ?? '')
    .trim()
    .toLowerCase()
    .replace(/[\s_-]+/g, '');
}

function findDescriptionColumn(headers) {
  return headers.find((header) => normalizeHeader(header) === DESCRIPTION_HEADER_KEY) ?? null;
}

function findEffectiveDateColumn(headers) {
  return headers.find((header) => normalizeHeader(header) === EFFECTIVE_DATE_HEADER_KEY) ?? null;
}

function parseEffectiveDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate()).getTime();
  }

  const text = String(value ?? '').trim();
  const match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/);

  if (!match) {
    return null;
  }

  const day = Number.parseInt(match[1], 10);
  const month = Number.parseInt(match[2], 10);
  const rawYear = Number.parseInt(match[3], 10);
  const year = rawYear < 100 ? 2000 + rawYear : rawYear;
  const date = new Date(year, month - 1, day);

  if (
    date.getFullYear() !== year ||
    date.getMonth() !== month - 1 ||
    date.getDate() !== day
  ) {
    return null;
  }

  return date.getTime();
}

function sortRowsByEffectiveDate(headers, rows) {
  const effectiveDateColumn = findEffectiveDateColumn(headers);

  if (!effectiveDateColumn) {
    return rows;
  }

  return rows
    .map((row, index) => ({
      row,
      index,
      timestamp: parseEffectiveDate(row[effectiveDateColumn]),
    }))
    .sort((left, right) => {
      if (left.timestamp === null && right.timestamp === null) {
        return left.index - right.index;
      }

      if (left.timestamp === null) {
        return 1;
      }

      if (right.timestamp === null) {
        return -1;
      }

      return left.timestamp - right.timestamp || left.index - right.index;
    })
    .map(({ row }) => row);
}

function extractRrn(description) {
  const text = String(description ?? '');
  const zenithInflowMatch = text.match(/\bISW(?:\s*ISW)?\s+INFLOW\s+CR\b[\s\S]*?\|\d{6}:(\d{11,12})\s+2MCS\w*\b/i);

  if (zenithInflowMatch) {
    return zenithInflowMatch[1];
  }

  const iswTerminalPairMatch = text.match(/\b2MCS\w*\s+(\d{11,12})\s+2MCS\w*\b/i);

  if (iswTerminalPairMatch) {
    return iswTerminalPairMatch[1];
  }

  const directIswMatch = text.match(/\bISW\s+(\d{11,12})\b/i);

  if (directIswMatch) {
    return directIswMatch[1];
  }

  const tokens = text.split(/\s+/).filter(Boolean);
  const terminalCodeIndex = tokens.findIndex(
    (token, index) => /^2MCS/i.test(token) && /^\d{11,12}$/.test(tokens[index - 1]),
  );
  const rrnBeforeTerminalCode = tokens[terminalCodeIndex - 1];

  if (/^\d{11,12}$/.test(rrnBeforeTerminalCode)) {
    return rrnBeforeTerminalCode;
  }

  return '';
}

function addRrnColumnAfterDescription(headers, rows) {
  const descriptionColumn = findDescriptionColumn(headers);

  if (!descriptionColumn) {
    return { headers, rows };
  }

  const nextHeaders = headers.includes(RRN_HEADER)
    ? headers
    : headers.flatMap((header) => (header === descriptionColumn ? [header, RRN_HEADER] : [header]));

  const nextRows = rows.map((row) => ({
    ...row,
    [RRN_HEADER]: extractRrn(row[descriptionColumn]),
  }));

  return { headers: nextHeaders, rows: nextRows };
}

export function exportValidTransactions(headers, rows) {
  const workbook = XLSX.utils.book_new();
  const validTransactions = addRrnColumnAfterDescription(headers, rows);
  const sortedRows = sortRowsByEffectiveDate(validTransactions.headers, validTransactions.rows);
  const worksheet = rowsToWorksheet(validTransactions.headers, sortedRows);
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Valid Transactions');
  XLSX.writeFile(workbook, 'Valid_Transactions.xlsx', {
    bookType: 'xlsx',
    cellDates: true,
  });
}

export function exportWrongRowsWorkbook(headers, groupedRows) {
  const workbook = XLSX.utils.book_new();
  const usedNames = new Set();

  for (const [groupName, rows] of Object.entries(groupedRows)) {
    if (rows.length === 0) {
      continue;
    }

    const worksheet = rowsToWorksheet(headers, sortRowsByEffectiveDate(headers, rows));
    XLSX.utils.book_append_sheet(workbook, worksheet, sanitizeSheetName(groupName, usedNames));
  }

  if (workbook.SheetNames.length === 0) {
    const worksheet = rowsToWorksheet(headers, []);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'No Wrong Rows');
  }

  XLSX.writeFile(workbook, 'Wrong_Rows.xlsx', {
    bookType: 'xlsx',
    cellDates: true,
  });
}
