import { useRef, useState } from 'react';
import { FileSpreadsheet, UploadCloud } from 'lucide-react';

export default function FileUpload({ onFileSelected, isProcessing }) {
  const [isDragging, setIsDragging] = useState(false);
  const inputRef = useRef(null);

  function handleFiles(files) {
    const selectedFiles = Array.from(files ?? []);
    if (selectedFiles.length) {
      onFileSelected(selectedFiles);
    }
  }

  return (
    <section
      className={`upload-zone ${isDragging ? 'is-dragging' : ''}`}
      onDragOver={(event) => {
        event.preventDefault();
        setIsDragging(true);
      }}
      onDragLeave={() => setIsDragging(false)}
      onDrop={(event) => {
        event.preventDefault();
        setIsDragging(false);
        handleFiles(event.dataTransfer.files);
      }}
    >
      <input
        ref={inputRef}
        className="visually-hidden"
        type="file"
        accept=".xlsx,.xls,.csv"
        multiple
        onChange={(event) => {
          handleFiles(event.target.files);
          event.target.value = '';
        }}
      />
      <div className="upload-icon">
        <UploadCloud size={30} aria-hidden="true" />
      </div>
      <div>
        <h2>Drop statement files here</h2>
        <p>.xlsx, .xls, and .csv files are processed locally in this browser.</p>
      </div>
      <button
        className="primary-button"
        type="button"
        disabled={isProcessing}
        onClick={() => inputRef.current?.click()}
      >
        <FileSpreadsheet size={18} aria-hidden="true" />
        {isProcessing ? 'Processing...' : 'Browse files'}
      </button>
    </section>
  );
}
