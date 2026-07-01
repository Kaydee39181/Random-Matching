# Local Transaction Filter

A fully local React + Vite application for filtering bank statement rows. The app extracts credit transactions whose `DebitAmount` is `0` and `CreditAmount` is greater than `0`, then exports rejected rows into a grouped workbook.

## Privacy

All parsing, filtering, grouping, and exporting happens in the browser on your machine. The app does not use databases, authentication, telemetry, cloud services, or external APIs.

## Supported Files

- `.xlsx`
- `.xls`
- `.csv`

## Run Locally

```bash
npm install
npm run dev
```

Open the local URL printed by Vite, usually `http://127.0.0.1:5173/`.

## Outputs

- `Valid_Transactions.xlsx`: rows whose `DebitAmount` is `0` and `CreditAmount` is greater than `0`.
- `Wrong_Rows.xlsx`: all rejected rows grouped into worksheets:
  - POS Transactions
  - ATM Transactions
  - Transfers
  - Charges
  - Empty Descriptions
  - Others

## Notes

- Headers are detected dynamically from the uploaded file.
- `DebitAmount` and `CreditAmount` headers are detected with whitespace and casing tolerance.
- Exported reports are sorted by `EffectiveDate`; dates like `29/04/2026` and `2/4/26` are supported.
- Original columns and column order are preserved in exported workbooks.
