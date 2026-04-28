export default function Badge({ value }) {
  if (!value) return <span className="text-xs text-brand-slate">unclassified</span>;
  const cls = {
    HIGH: 'badge-high',
    MEDIUM: 'badge-medium',
    UNCERTAIN: 'badge-uncertain',
  }[value] || 'badge-uncertain';
  return <span className={cls}>{value}</span>;
}
