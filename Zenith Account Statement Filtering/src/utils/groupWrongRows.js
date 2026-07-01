import { normalizeCellText, normalizeHeader } from './validateTransactions.js';

export const WRONG_ROW_GROUPS = [
  {
    sheetName: 'POS Transactions',
    key: 'pos',
    keywords: ['POS', 'PURCHASE', 'TERMINAL'],
  },
  {
    sheetName: 'ATM Transactions',
    key: 'atm',
    keywords: ['ATM', 'CASH'],
  },
  {
    sheetName: 'Transfers',
    key: 'transfers',
    keywords: ['TRANSFER', 'NIP', 'TRF'],
  },
  {
    sheetName: 'Charges',
    key: 'charges',
    keywords: ['CHARGE', 'FEE', 'VAT'],
  },
];

export const EMPTY_DESCRIPTIONS_GROUP = {
  sheetName: 'Empty Descriptions',
  key: 'empty',
};

export const OTHERS_GROUP = {
  sheetName: 'Others',
  key: 'others',
};

function descriptionContainsAny(description, keywords) {
  const normalized = normalizeCellText(description).toUpperCase();
  return keywords.some((keyword) => normalized.includes(keyword));
}

function findDescriptionColumnFromRows(rows) {
  const firstRow = rows[0] ?? {};
  return Object.keys(firstRow).find((header) => normalizeHeader(header) === 'description') ?? null;
}

export function groupWrongRows(rows, groupingColumn) {
  const fallbackGroupingColumn = groupingColumn ?? findDescriptionColumnFromRows(rows);
  const groupedRows = {
    [EMPTY_DESCRIPTIONS_GROUP.sheetName]: [],
    ...Object.fromEntries(WRONG_ROW_GROUPS.map((group) => [group.sheetName, []])),
    [OTHERS_GROUP.sheetName]: [],
  };

  for (const row of rows) {
    const description = normalizeCellText(row[fallbackGroupingColumn]);

    if (!description) {
      groupedRows[EMPTY_DESCRIPTIONS_GROUP.sheetName].push(row);
      continue;
    }

    const matchingGroup = WRONG_ROW_GROUPS.find((group) =>
      descriptionContainsAny(description, group.keywords),
    );

    if (matchingGroup) {
      groupedRows[matchingGroup.sheetName].push(row);
    } else {
      groupedRows[OTHERS_GROUP.sheetName].push(row);
    }
  }

  return groupedRows;
}
