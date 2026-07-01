import { Download, RotateCcw } from 'lucide-react';

export default function DownloadButtons({
  hasResults,
  validCount,
  rejectedCount,
  onDownloadValid,
  onDownloadWrongRows,
  onReset,
}) {
  return (
    <div className="actions-bar">
      <button
        className="primary-button"
        type="button"
        disabled={!hasResults || validCount === 0}
        onClick={onDownloadValid}
      >
        <Download size={18} aria-hidden="true" />
        Download Valid Transactions
      </button>
      <button
        className="secondary-button"
        type="button"
        disabled={!hasResults || rejectedCount === 0}
        onClick={onDownloadWrongRows}
      >
        <Download size={18} aria-hidden="true" />
        Download Wrong Rows Workbook
      </button>
      <button className="ghost-button" type="button" disabled={!hasResults} onClick={onReset}>
        <RotateCcw size={18} aria-hidden="true" />
        Reset
      </button>
    </div>
  );
}
