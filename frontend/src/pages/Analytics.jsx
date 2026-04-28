import { useEffect, useState } from 'react';
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell, Legend } from 'recharts';
import { api } from '../api';

const COLORS = {
  HIGH: '#10b981',
  MEDIUM: '#f59e0b',
  UNCERTAIN: '#ef4444',
  CRITICAL: '#ef4444',
  ERROR: '#f97316',
  WARNING: '#fbbf24',
  INFO: '#7E8AA2',
};

function toBarData(obj) {
  return Object.entries(obj || {}).map(([name, value]) => ({ name, value }));
}

export default function Analytics() {
  const [data, setData] = useState(null);
  const [error, setError] = useState(null);

  useEffect(() => {
    api.analytics().then(setData).catch((e) => setError(e.message));
  }, []);

  if (error) return <div className="bg-rose-50 border border-rose-200 text-rose-800 px-3 py-2 rounded">{error}</div>;
  if (!data) return <div className="text-brand-slate">Loading…</div>;

  const bySource = toBarData(data.by_source);
  const bySeverity = toBarData(data.by_severity);
  const byBadge = toBarData(data.by_badge);
  const byModule = toBarData(data.by_module);

  return (
    <div className="space-y-6">
      <h1 className="text-2xl font-semibold text-brand-navy">Analytics</h1>

      <div className="grid grid-cols-3 gap-4">
        <Stat label="Total errors" value={data.total_errors} />
        <Stat label="Classified" value={data.total_classified} />
        <Stat label="Retry rate" value={`${(data.retry_rate * 100).toFixed(1)}%`} />
      </div>

      <div className="grid grid-cols-2 gap-4">
        <Card title="Errors by source">
          <ResponsiveContainer width="100%" height={240}>
            <BarChart data={bySource}>
              <XAxis dataKey="name" stroke="#7E8AA2" fontSize={12} />
              <YAxis stroke="#7E8AA2" fontSize={12} />
              <Tooltip />
              <Bar dataKey="value" fill="#263248" />
            </BarChart>
          </ResponsiveContainer>
        </Card>
        <Card title="Errors by severity">
          <ResponsiveContainer width="100%" height={240}>
            <BarChart data={bySeverity}>
              <XAxis dataKey="name" stroke="#7E8AA2" fontSize={12} />
              <YAxis stroke="#7E8AA2" fontSize={12} />
              <Tooltip />
              <Bar dataKey="value">
                {bySeverity.map((d, i) => <Cell key={i} fill={COLORS[d.name] || '#7E8AA2'} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </Card>
        <Card title="Confidence badge mix">
          <ResponsiveContainer width="100%" height={240}>
            <PieChart>
              <Pie data={byBadge} dataKey="value" nameKey="name" outerRadius={80} label>
                {byBadge.map((d, i) => <Cell key={i} fill={COLORS[d.name] || '#7E8AA2'} />)}
              </Pie>
              <Legend />
              <Tooltip />
            </PieChart>
          </ResponsiveContainer>
        </Card>
        <Card title="Errors by module">
          <ResponsiveContainer width="100%" height={240}>
            <BarChart data={byModule}>
              <XAxis dataKey="name" stroke="#7E8AA2" fontSize={12} />
              <YAxis stroke="#7E8AA2" fontSize={12} />
              <Tooltip />
              <Bar dataKey="value" fill="#FF9800" />
            </BarChart>
          </ResponsiveContainer>
        </Card>
      </div>
    </div>
  );
}

function Stat({ label, value }) {
  return (
    <div className="bg-white border border-brand-silver rounded-md p-4">
      <div className="text-xs text-brand-slate uppercase tracking-wide">{label}</div>
      <div className="text-2xl font-semibold text-brand-navy mt-1">{value}</div>
    </div>
  );
}

function Card({ title, children }) {
  return (
    <div className="bg-white border border-brand-silver rounded-md p-4">
      <div className="text-sm font-semibold text-brand-navy mb-2">{title}</div>
      {children}
    </div>
  );
}
