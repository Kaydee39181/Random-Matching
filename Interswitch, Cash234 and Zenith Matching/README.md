# Three-Way Transaction Reconciliation App

Fully offline Streamlit app for comparing INTERSWITCH, CASH234, and ZENITH transaction files by normalized RRN.

## Features

- Upload `.xlsx`, `.xls`, and `.csv` files.
- Choose workbook sheets when multiple sheets exist.
- Normalize RRNs from:
  - INTERSWITCH: `Retrieval_Reference_Nr`
  - CASH234: `R R N`
  - ZENITH: `RRN`
- Produces all seven valid Venn result sets, pairwise intersection complement sheets, plus `INVALID_OR_BLANK_RRN`.
- Preserves all source columns, grouped by dataset with visible spacer columns in exported sheets.
- Removes duplicate valid RRNs per uploaded file before matching.
- Exports removed duplicate rows to a separate duplicates workbook with `INTERSWITCH`, `CASH234`, and `ZENITH` sheets.
- Interactive preview with search, filters, sorting, and pagination.
- Download the full Excel workbook, individual result sheets, or filtered previews.
- Includes summary/statistics sheets, charts, and a Venn-style visualization.

## Install

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Run

```powershell
streamlit run app.py
```

The app runs entirely on your local machine. Uploaded files are processed in memory and are not sent anywhere.

## Separate Tran ID vs Zenith Description Matching

The main reconciliation app now compares INTERSWITCH, CASH234, and ZENITH by RRN only.

To compare Interswitch `Tran_ID` values against useful token sections inside Zenith descriptions, run:

```powershell
.\run_tran_matcher.bat
```

This produces a separate `Tran_ID_vs_Zenith_Description.xlsx` report with matched, unmatched, and invalid-token sheets.
