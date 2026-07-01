import * as XLSX from 'xlsx';
import Papa from 'papaparse';
import { findSourceReferenceColumn, findTransactionAmountColumns } from './validateTransactions.js';

const SUPPORTED_EXTENSIONS = new Set(['xlsx', 'xls', 'csv']);

function getFileExtension(fileName) {
  return String(fileName).split('.').pop()?.toLowerCase() ?? '';
}

function assertSupportedFile(file) {
  const extension = getFileExtension(file.name);

  if (!SUPPORTED_EXTENSIONS.has(extension)) {
    throw new Error('Unsupported file type. Please upload an .xlsx, .xls, or .csv file.');
  }

  return extension;
}

function dedupeHeaders(headers) {
  const counts = new Map();

  return headers.map((header, index) => {
    const normalizedHeader = String(header ?? '').trim() || `Column ${index + 1}`;
    const seen = counts.get(normalizedHeader) ?? 0;
    counts.set(normalizedHeader, seen + 1);
    return seen === 0 ? normalizedHeader : `${normalizedHeader} (${seen + 1})`;
  });
}

function buildRowsFromMatrix(matrix) {
  const hasContent = (row) => row.some((cell) => String(cell ?? '').trim() !== '');
  const firstNonEmptyRow = matrix.findIndex(hasContent);

  if (firstNonEmptyRow === -1) {
    throw new Error('The uploaded file is empty.');
  }

  const headerIndex = matrix.findIndex((row) => {
    if (!hasContent(row)) {
      return false;
    }

    const amountColumns = findTransactionAmountColumns(dedupeHeaders(row));
    return amountColumns.debitAmountColumn && amountColumns.creditAmountColumn;
  });

  if (headerIndex === -1) {
    throw new Error('Missing DebitAmount or CreditAmount column. The file must include both headers.');
  }

  const headers = dedupeHeaders(matrix[headerIndex]);
  const sourceReferenceColumn = findSourceReferenceColumn(headers);
  const amountColumns = findTransactionAmountColumns(headers);
  const dataRows = matrix.slice(headerIndex + 1);

  const rows = dataRows.map((row) =>
    Object.fromEntries(headers.map((header, index) => [header, row[index] ?? ''])),
  );

  return {
    headers,
    rows,
    sourceReferenceColumn,
    amountColumns,
    skippedRowsBeforeHeader: headerIndex,
  };
}

function parseCsv(file) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      skipEmptyLines: false,
      dynamicTyping: false,
      worker: true,
      complete: (result) => {
        if (result.errors?.length) {
          const firstError = result.errors[0];
          reject(new Error(`Could not parse CSV file: ${firstError.message}`));
          return;
        }

        try {
          resolve(buildRowsFromMatrix(result.data));
        } catch (error) {
          reject(error);
        }
      },
      error: (error) => reject(new Error(`Could not read CSV file: ${error.message}`)),
    });
  });
}

async function parseExcel(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, {
    type: 'array',
    cellDates: true,
    raw: true,
  });

  if (!workbook.SheetNames.length) {
    throw new Error('The uploaded workbook does not contain any worksheets.');
  }

  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const matrix = XLSX.utils.sheet_to_json(firstSheet, {
    header: 1,
    raw: false,
    defval: '',
    blankrows: true,
  });

  return buildRowsFromMatrix(matrix);
}

export async function parseFile(file) {
  if (!file) {
    throw new Error('Please choose a file to upload.');
  }

  const extension = assertSupportedFile(file);
  return extension === 'csv' ? parseCsv(file) : parseExcel(file);
}
