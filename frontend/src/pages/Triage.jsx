import { useEffect, useState } from 'react';
import { Link } from 'react-router-dom';
import { RefreshCw, Filter, Play } from 'lucide-react';
import { api } from '../api';
import Badge from '../components/Badge.jsx';

const SOURCES = ['', 'ST22', 'SM21', 'SLG1', 'SM13', 'RZ20', 'SM37'];
const SEVERITIES = ['', 'CRITICAL', 'ERROR', 'WARNING'];
const BADGES = ['', 'HIGH', 'MEDIUM', 'UNCERTAIN'];

export default function Triage() {
  const [errors, setErrors] = useState([]);
  const [loading, setLoading] = useState(false);
  const [sweeping, setSweeping] = useState(false);
  const [error, setError] = useState(null);
  const [filters, setFilters] = useState({ source: '', severity: '', badge: '' });

  async function load() {
    setLoading(true);
    setError(null);
    try {
      const data = await api.listErrors({
        limit: 100,
        source: filters.source || undefined,
        severity: filters.severity || undefined,
        badge: filters.badge || undefined,
      });
      setErrors(data);
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }

  async function sweep() {
    setSweeping(true);
    try {
      await api.sweep(50);
      await load();
    } catch (e) {
      setError(e.message);
    } finally {
      setSweeping(false);
    }
  }

  useEffect(() => {
    load();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [filters.source, filters.severity, filters.badge]);

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-semibold text-brand-navy">Triage</h1>
          <p className="text-sm text-brand-slate">Live SAP errors ranked by confidence × severity</p>
        </div>
        <div className="flex gap-2">
          <button
            onClick={sweep}
            disabled={sweeping}
            className="flex items-center gap-2 px-3 py-2 rounded-md bg-brand-amber text-brand-navy font-medium text-sm disabled:opacity-50"
          >
            <Play className="w-4 h-4" />
            {sweeping ? 'Classifying…' : 'Classify unclassified'}
          </button>
          <button
            onClick={load}
            disabled={loading}
            className="flex items-center gap-2 px-3 py-2 rounded-md bg-white border border-brand-silver text-brand-navy text-sm disabled:opacity-50"
          >
            <RefreshCw className={`w-4 h-4 ${loading ? 'animate-spin' : ''}`} />
            Refresh
          </button>
        </div>
      </div>

      <div className="bg-white rounded-md border border-brand-silver p-3 flex items-center gap-3">
        <Filter className="w-4 h-4 text-brand-slate" />
        <select
          value={filters.source}
          onChange={(e) => setFilters({ ...filters, source: e.target.value })}
          className="text-sm border border-brand-silver rounded px-2 py-1"
        >
          {SOURCES.map((s) => <option key={s} value={s}>{s || 'All sources'}</option>)}
        </select>
        <select
          value={filters.severity}
          onChange={(e) => setFilters({ ...filters, severity: e.target.value })}
          className="text-sm border border-brand-silver rounded px-2 py-1"
        >
          {SEVERITIES.map((s) => <option key={s} value={s}>{s || 'All severities'}</option>)}
        </select>
        <select
          value={filters.badge}
          onChange={(e) => setFilters({ ...filters, badge: e.target.value })}
          className="text-sm border border-brand-silver rounded px-2 py-1"
        >
          {BADGES.map((b) => <option key={b} value={b}>{b || 'All badges'}</option>)}
        </select>
        <span className="text-sm text-brand-slate ml-auto">{errors.length} rows</span>
      </div>

      {error && (
        <div className="bg-rose-50 border border-rose-200 text-rose-800 px-3 py-2 rounded text-sm">{error}</div>
      )}

      <div className="bg-white rounded-md border border-brand-silver overflow-hidden">
        <table className="w-full text-sm">
          <thead className="bg-brand-navy text-white">
            <tr>
              <th className="px-3 py-2 text-left">Source</th>
              <th className="px-3 py-2 text-left">Severity</th>
              <th className="px-3 py-2 text-left">Occurred</th>
              <th className="px-3 py-2 text-left">Short text</th>
              <th className="px-3 py-2 text-left">Top proposal</th>
              <th className="px-3 py-2 text-left">Confidence</th>
              <th className="px-3 py-2 text-left">Badge</th>
            </tr>
          </thead>
          <tbody>
            {errors.map((e) => (
              <tr key={e.hash_key} className="border-t border-brand-silver hover:bg-gray-50">
                <td className="px-3 py-2 font-mono text-xs">{e.source}</td>
                <td className={`px-3 py-2 severity-${(e.severity || '').toLowerCase()}`}>{e.severity}</td>
                <td className="px-3 py-2 text-brand-slate text-xs">{new Date(e.occurred_at).toLocaleString()}</td>
                <td className="px-3 py-2 max-w-md truncate">
                  <Link to={`/errors/${e.hash_key}`} className="text-brand-navy hover:underline">
                    {e.short_text}
                  </Link>
                </td>
                <td className="px-3 py-2 max-w-xs truncate text-brand-slate">{e.top_proposal_title || '—'}</td>
                <td className="px-3 py-2">{e.composite_confidence ? `${(e.composite_confidence * 100).toFixed(0)}%` : '—'}</td>
                <td className="px-3 py-2"><Badge value={e.badge} /></td>
              </tr>
            ))}
            {!loading && errors.length === 0 && (
              <tr><td colSpan="7" className="px-3 py-8 text-center text-brand-slate">No errors found.</td></tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}
