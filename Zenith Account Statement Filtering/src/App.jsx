import { useMemo, useState } from 'react';
import FileUpload from './components/FileUpload.jsx';
import SummaryCards from './components/SummaryCards.jsx';
import PreviewTable from './components/PreviewTable.jsx';
import DownloadButtons from './components/DownloadButtons.jsx';
import { parseFile } from './utils/parseFile.js';
import { splitTransactions } from './utils/validateTransactions.js';
import { groupWrongRows } from './utils/groupWrongRows.js';
import { exportValidTransactions, exportWrongRowsWorkbook } from './utils/exportWorkbook.js';

const initialState = {
  fileName: '',
  fileCount: 0,
  headers: [],
  rows: [],
  sourceReferenceColumn: '',
  validRows: [],
  rejectedRows: [],
  groupedWrongRows: {},
};

const SOURCE_FILE_HEADER = 'Source File';

function getSourceFileHeader(headers) {
  if (!headers.includes(SOURCE_FILE_HEADER)) {
    return SOURCE_FILE_HEADER;
  }

  let index = 2;
  let candidate = `${SOURCE_FILE_HEADER} (${index})`;

  while (headers.includes(candidate)) {
    index += 1;
    candidate = `${SOURCE_FILE_HEADER} (${index})`;
  }

  return candidate;
}

function addMissingHeaders(headers, nextHeaders) {
  const knownHeaders = new Set(headers);
  const mergedHeaders = [...headers];

  for (const header of nextHeaders) {
    if (!knownHeaders.has(header)) {
      knownHeaders.add(header);
      mergedHeaders.push(header);
    }
  }

  return mergedHeaders;
}

function getResultFileName(files) {
  if (files.length === 1) {
    return files[0].name;
  }

  return `${files.length} files: ${files.map((file) => file.name).join(', ')}`;
}

async function parseFiles(files) {
  const parsedFiles = await Promise.all(
    files.map(async (file) => {
      try {
        return {
          file,
          parsed: await parseFile(file),
        };
      } catch (error) {
        const message = error instanceof Error ? error.message : 'Could not process the file.';
        throw new Error(`${file.name}: ${message}`);
      }
    }),
  );

  const baseHeaders = parsedFiles.reduce(
    (headers, { parsed }) => addMissingHeaders(headers, parsed.headers),
    [],
  );
  const sourceFileHeader = files.length > 1 ? getSourceFileHeader(baseHeaders) : null;
  const headers = sourceFileHeader ? [...baseHeaders, sourceFileHeader] : baseHeaders;
  const rows = [];
  const validRows = [];
  const rejectedRows = [];

  for (const { file, parsed } of parsedFiles) {
    const normalizedRows = parsed.rows.map((row) => ({
      ...Object.fromEntries(headers.map((header) => [header, ''])),
      ...row,
      ...(sourceFileHeader ? { [sourceFileHeader]: file.name } : {}),
    }));
    const splitRows = splitTransactions(normalizedRows, parsed.amountColumns);

    rows.push(...normalizedRows);
    validRows.push(...splitRows.validRows);
    rejectedRows.push(...splitRows.rejectedRows);
  }

  return {
    fileName: getResultFileName(files),
    fileCount: files.length,
    headers,
    rows,
    sourceReferenceColumn: parsedFiles[0]?.parsed.sourceReferenceColumn ?? '',
    validRows,
    rejectedRows,
  };
}

export default function App() {
  const [result, setResult] = useState(initialState);
  const [error, setError] = useState('');
  const [isProcessing, setIsProcessing] = useState(false);

  const hasResults = result.headers.length > 0;

  const summary = useMemo(() => {
    if (!hasResults) {
      return null;
    }

    return {
      fileName: result.fileName,
      totalRows: result.rows.length,
      validRows: result.validRows.length,
      rejectedRows: result.rejectedRows.length,
    };
  }, [hasResults, result]);

  async function handleFileSelected(selectedFiles) {
    const files = Array.from(selectedFiles ?? []);

    if (files.length === 0) {
      return;
    }

    setError('');
    setIsProcessing(true);

    try {
      const parsed = await parseFiles(files);
      const descriptionColumn = parsed.headers.find(
        (header) => header.trim().toLowerCase().replace(/[\s_-]+/g, '') === 'description',
      );
      const groupedWrongRows = groupWrongRows(parsed.rejectedRows, descriptionColumn);

      setResult({
        fileName: parsed.fileName,
        fileCount: parsed.fileCount,
        headers: parsed.headers,
        rows: parsed.rows,
        sourceReferenceColumn: parsed.sourceReferenceColumn,
        validRows: parsed.validRows,
        rejectedRows: parsed.rejectedRows,
        groupedWrongRows,
      });
    } catch (caughtError) {
      setResult(initialState);
      setError(caughtError instanceof Error ? caughtError.message : 'Could not process the file.');
    } finally {
      setIsProcessing(false);
    }
  }

  function handleReset() {
    setResult(initialState);
    setError('');
  }

  return (
    <main className="app-shell">
      <header className="app-header">
        <div>
          <p className="eyebrow">Local reconciliation tool</p>
          <h1>ISW Transaction Filter</h1>
          <p>
            Extracts credit rows whose DebitAmount is zero and CreditAmount is greater than zero,
            then separates every other row into an auditable wrong-rows workbook.
          </p>
        </div>
      </header>

      <FileUpload onFileSelected={handleFileSelected} isProcessing={isProcessing} />

      {error ? <div className="error-banner" role="alert">{error}</div> : null}

      <SummaryCards summary={summary} />

      <DownloadButtons
        hasResults={hasResults}
        validCount={result.validRows.length}
        rejectedCount={result.rejectedRows.length}
        onDownloadValid={() => exportValidTransactions(result.headers, result.validRows)}
        onDownloadWrongRows={() => exportWrongRowsWorkbook(result.headers, result.groupedWrongRows)}
        onReset={handleReset}
      />

      {hasResults ? (
        <div className="preview-grid">
          <PreviewTable
            title="Valid Transactions Preview"
            headers={result.headers}
            rows={result.validRows}
            emptyMessage="No credit transactions found."
            badge={`${result.validRows.length} rows`}
          />

          <section className="wrong-preview">
            <div className="section-heading">
              <div>
                <h2>Wrong Rows Preview</h2>
                <p>Rejected rows grouped into workbook sheets.</p>
              </div>
              <span className="count-badge">{result.rejectedRows.length} rows</span>
            </div>

            <div className="group-stack">
              {Object.entries(result.groupedWrongRows).map(([groupName, rows]) => (
                <PreviewTable
                  key={groupName}
                  title={groupName}
                  headers={result.headers}
                  rows={rows}
                  emptyMessage="No rows in this group."
                  badge={`${rows.length} rows`}
                />
              ))}
            </div>
          </section>
        </div>
      ) : null}
    </main>
  );
}
