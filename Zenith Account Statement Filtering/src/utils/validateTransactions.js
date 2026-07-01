const SOURCE_REFERENCE_HEADER_KEY = 'sourcereference';
const DEBIT_AMOUNT_HEADER_KEYS = new Set(['debitamount', 'debit']);
const CREDIT_AMOUNT_HEADER_KEYS = new Set(['creditamount', 'credit']);

export function normalizeHeader(header) {
  return String(header ?? '')
    .trim()
    .toLowerCase()
    .replace(/[\s_-]+/g, '');
}

export function findSourceReferenceColumn(headers) {
  return headers.find((header) => normalizeHeader(header) === SOURCE_REFERENCE_HEADER_KEY) ?? null;
}

export function findTransactionAmountColumns(headers) {
  const debitAmountColumn =
    headers.find((header) => DEBIT_AMOUNT_HEADER_KEYS.has(normalizeHeader(header))) ?? null;
  const creditAmountColumn =
    headers.find((header) => CREDIT_AMOUNT_HEADER_KEYS.has(normalizeHeader(header))) ?? null;

  return {
    debitAmountColumn,
    creditAmountColumn,
  };
}

export function normalizeCellText(value) {
  return String(value ?? '')
    .replace(/\s+/g, ' ')
    .trim();
}

function parseAmount(value) {
  const normalized = normalizeCellText(value);

  if (!normalized || normalized === '-') {
    return 0;
  }

  const isNegative = /^\(.*\)$/.test(normalized) || normalized.startsWith('-');
  const numericText = normalized.replace(/[(),\s]/g, '').replace(/[^\d.-]/g, '');
  const parsedAmount = Number.parseFloat(numericText);

  if (!Number.isFinite(parsedAmount)) {
    return null;
  }

  return isNegative ? -Math.abs(parsedAmount) : parsedAmount;
}

export function isValidTransaction(row, amountColumns) {
  const debitAmount = parseAmount(row[amountColumns.debitAmountColumn]);
  const creditAmount = parseAmount(row[amountColumns.creditAmountColumn]);

  return debitAmount === 0 && creditAmount > 0;
}

export function splitTransactions(rows, amountColumns) {
  const validRows = [];
  const rejectedRows = [];

  for (const row of rows) {
    if (isValidTransaction(row, amountColumns)) {
      validRows.push(row);
    } else {
      rejectedRows.push(row);
    }
  }

  return { validRows, rejectedRows };
}
