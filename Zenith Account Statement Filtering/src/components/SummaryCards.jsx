import { CheckCircle2, FileText, Rows3, XCircle } from 'lucide-react';

const cards = [
  { key: 'fileName', label: 'Uploaded files', icon: FileText },
  { key: 'totalRows', label: 'Total rows', icon: Rows3 },
  { key: 'validRows', label: 'Valid rows', icon: CheckCircle2 },
  { key: 'rejectedRows', label: 'Rejected rows', icon: XCircle },
];

export default function SummaryCards({ summary }) {
  if (!summary) {
    return null;
  }

  return (
    <section className="summary-grid" aria-label="Upload summary">
      {cards.map(({ key, label, icon: Icon }) => (
        <article className="summary-card" key={key}>
          <Icon size={20} aria-hidden="true" />
          <span>{label}</span>
          <strong title={String(summary[key])}>{summary[key]}</strong>
        </article>
      ))}
    </section>
  );
}
