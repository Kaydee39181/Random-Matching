# Interswitch Date Sorter

Local web app for sorting large `.xlsx` workbooks by the `Local_Date_Time` column.

## Run

```powershell
python app.py
```

Open `http://127.0.0.1:5000`, upload your workbook, choose oldest-first or newest-first, then download the sorted file.

The app also stores files locally:

- Uploaded originals: `uploads`
- Sorted workbooks: `sorted`

After sorting, the page shows a download link and a recent sorted files list. If a browser download does not start immediately, open the `sorted` folder or click the file in the recent list.

## Supported Date Formats

The sorter accepts normal Excel date cells plus strings like:

- `2026-05-22 22:37:58`
- `5/24/2026 9:08`

Rows with blank or unreadable `Local_Date_Time` values are kept at the bottom for oldest-first sorting.
