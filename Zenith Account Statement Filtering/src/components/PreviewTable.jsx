function formatCell(value) {
  if (value instanceof Date) {
    return value.toLocaleDateString();
  }

  if (value === null || value === undefined) {
    return '';
  }

  return String(value);
}

export default function PreviewTable({ title, headers, rows, emptyMessage, badge }) {
  const previewRows = rows.slice(0, 20);
  const visibleHeaders = headers.slice(0, 10);

  return (
    <section className="preview-section">
      <div className="section-heading">
        <div>
          <h2>{title}</h2>
          <p>{previewRows.length ? `Showing first ${previewRows.length} rows` : emptyMessage}</p>
        </div>
        {badge ? <span className="count-badge">{badge}</span> : null}
      </div>

      {previewRows.length ? (
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                {visibleHeaders.map((header) => (
                  <th key={header}>{header}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {previewRows.map((row, rowIndex) => (
                <tr key={`${title}-${rowIndex}`}>
                  {visibleHeaders.map((header) => (
                    <td key={header} title={formatCell(row[header])}>
                      {formatCell(row[header])}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : null}
    </section>
  );
}
