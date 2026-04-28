import { useEffect, useState } from 'react';
import { useParams, Link } from 'react-router-dom';
import { ArrowLeft, RefreshCw, ThumbsUp, ThumbsDown, ExternalLink } from 'lucide-react';
import { api } from '../api';
import Badge from '../components/Badge.jsx';

export default function ErrorDetail() {
  const { hashKey } = useParams();
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [classifying, setClassifying] = useState(false);
  const [feedbackOpen, setFeedbackOpen] = useState(false);
  const [error, setError] = useState(null);

  async function load() {
    setLoading(true);
    setError(null);
    try {
      setData(await api.getError(hashKey));
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }

  async function reclassify() {
    setClassifying(true);
    try {
      await api.classify(hashKey);
      await load();
    } catch (e) {
      setError(e.message);
    } finally {
      setClassifying(false);
    }
  }

  useEffect(() => { load(); }, [hashKey]);

  if (loading) return <div className="text-brand-slate">Loading…</div>;
  if (error) return <div className="bg-rose-50 border border-rose-200 text-rose-800 px-3 py-2 rounded">{error}</div>;
  if (!data) return null;

  const c = data.classification;
  const proposals = c?.proposals || [];

  return (
    <div className="space-y-6">
      <div className="flex items-center justify-between">
        <Link to="/triage" className="flex items-center gap-1 text-brand-slate hover:text-brand-navy text-sm">
          <ArrowLeft className="w-4 h-4" /> Back to triage
        </Link>
        <div className="flex gap-2">
          <button
            onClick={reclassify}
            disabled={classifying}
            className="flex items-center gap-2 px-3 py-2 rounded-md bg-brand-amber text-brand-navy text-sm font-medium disabled:opacity-50"
          >
            <RefreshCw className={`w-4 h-4 ${classifying ? 'animate-spin' : ''}`} />
            {classifying ? 'Classifying…' : 'Re-classify'}
          </button>
        </div>
      </div>

      {/* Error header */}
      <div className="bg-white border border-brand-silver rounded-md p-4">
        <div className="flex items-start justify-between gap-4">
          <div>
            <div className="flex items-center gap-2 text-xs text-brand-slate">
              <span className="font-mono">{data.source}</span>
              <span>·</span>
              <span className={`severity-${(data.severity || '').toLowerCase()}`}>{data.severity}</span>
              <span>·</span>
              <span>{new Date(data.occurred_at).toLocaleString()}</span>
              <span>·</span>
              <span className="font-mono">{data.system_id}</span>
            </div>
            <div className="mt-2 text-lg font-semibold text-brand-navy">{data.short_text}</div>
            <div className="mt-1 text-sm text-brand-slate">
              {data.transaction && <span className="mr-3">tcode <span className="font-mono text-brand-navy">{data.transaction}</span></span>}
              {data.program && <span>program <span className="font-mono text-brand-navy">{data.program}</span></span>}
            </div>
          </div>
          {c && (
            <div className="text-right">
              <Badge value={c.badge} />
              <div className="text-xs text-brand-slate mt-1">
                {c.composite_confidence != null ? `${(c.composite_confidence * 100).toFixed(0)}% confidence` : ''}
              </div>
              {c.retried && <div className="text-xs text-amber-600 mt-1">retried</div>}
            </div>
          )}
        </div>
        {data.long_text && (
          <pre className="mt-4 bg-gray-50 border border-brand-silver rounded p-3 text-xs font-mono whitespace-pre-wrap max-h-64 overflow-auto">
            {data.long_text}
          </pre>
        )}
      </div>

      {/* Proposals */}
      {!c && (
        <div className="bg-amber-50 border border-amber-200 text-amber-800 px-3 py-2 rounded text-sm">
          This error has not been classified yet. Click <strong>Re-classify</strong> above.
        </div>
      )}
      {proposals.length > 0 && (
        <div className="space-y-3">
          <h2 className="font-semibold text-brand-navy">Proposed solutions</h2>
          {proposals.map((p, i) => (
            <div key={i} className="bg-white border border-brand-silver rounded-md p-4">
              <div className="flex items-start justify-between gap-4">
                <div>
                  <div className="text-xs text-brand-slate">Rank {p.rank}</div>
                  <div className="font-semibold text-brand-navy mt-0.5">{p.title}</div>
                </div>
                <div className="text-right">
                  <Badge value={p.badge} />
                  <div className="text-xs text-brand-slate mt-1">
                    composite {(p.composite_confidence * 100).toFixed(0)}%
                  </div>
                </div>
              </div>

              <ol className="list-decimal list-inside mt-3 space-y-1 text-sm">
                {(p.steps || []).map((s, j) => <li key={j}>{s}</li>)}
              </ol>

              {p.risks?.length > 0 && (
                <div className="mt-3 text-xs">
                  <div className="font-semibold text-rose-700">Risks</div>
                  <ul className="list-disc list-inside text-rose-700">
                    {p.risks.map((r, k) => <li key={k}>{r}</li>)}
                  </ul>
                </div>
              )}

              <details className="mt-3 text-xs">
                <summary className="cursor-pointer text-brand-slate">Confidence breakdown</summary>
                <div className="mt-1 grid grid-cols-3 gap-2">
                  <div>vector sim: <span className="font-mono">{(p.vector_similarity * 100).toFixed(0)}%</span></div>
                  <div>model self-conf: <span className="font-mono">{(p.model_self_confidence * 100).toFixed(0)}%</span></div>
                  <div>corpus: <span className="font-mono">{(p.corpus_corroboration * 100).toFixed(0)}%</span></div>
                </div>
              </details>
            </div>
          ))}
        </div>
      )}

      {/* Critique notes */}
      {c?.critique_notes && (
        <div className="bg-white border border-brand-silver rounded-md p-4">
          <h2 className="font-semibold text-brand-navy mb-2">Critique notes</h2>
          <div className="text-sm text-brand-slate">{c.critique_notes}</div>
        </div>
      )}

      {/* Citations */}
      {c?.citations?.length > 0 && (
        <div className="bg-white border border-brand-silver rounded-md p-4">
          <h2 className="font-semibold text-brand-navy mb-2">Corpus citations</h2>
          <ul className="space-y-1 text-sm">
            {c.citations.slice(0, 12).map((url, i) => (
              <li key={i}>
                <a href={url} target="_blank" rel="noreferrer" className="text-brand-navy hover:underline inline-flex items-center gap-1">
                  <ExternalLink className="w-3 h-3" /> {url}
                </a>
              </li>
            ))}
          </ul>
        </div>
      )}

      {/* Feedback */}
      {c && (
        <FeedbackBlock hashKey={hashKey} open={feedbackOpen} setOpen={setFeedbackOpen} />
      )}
    </div>
  );
}

function FeedbackBlock({ hashKey, open, setOpen }) {
  const [reviewer, setReviewer] = useState('');
  const [comments, setComments] = useState('');
  const [accepted, setAccepted] = useState(true);
  const [submitted, setSubmitted] = useState(false);
  const [error, setError] = useState(null);

  async function submit() {
    setError(null);
    try {
      await api.feedback(hashKey, { reviewer, accepted, comments });
      setSubmitted(true);
    } catch (e) {
      setError(e.message);
    }
  }

  if (submitted) {
    return <div className="bg-emerald-50 border border-emerald-200 text-emerald-800 px-3 py-2 rounded text-sm">Feedback recorded. Thanks.</div>;
  }

  return (
    <div className="bg-white border border-brand-silver rounded-md p-4">
      <button
        onClick={() => setOpen(!open)}
        className="text-brand-navy font-semibold text-sm"
      >
        {open ? 'Hide feedback' : 'Submit feedback'}
      </button>
      {open && (
        <div className="mt-3 space-y-2">
          <div className="flex gap-2">
            <button
              onClick={() => setAccepted(true)}
              className={`flex items-center gap-1 px-3 py-1 rounded text-sm ${accepted ? 'bg-emerald-100 text-emerald-800' : 'bg-gray-100 text-brand-slate'}`}
            >
              <ThumbsUp className="w-4 h-4" /> Accept
            </button>
            <button
              onClick={() => setAccepted(false)}
              className={`flex items-center gap-1 px-3 py-1 rounded text-sm ${!accepted ? 'bg-rose-100 text-rose-800' : 'bg-gray-100 text-brand-slate'}`}
            >
              <ThumbsDown className="w-4 h-4" /> Reject
            </button>
          </div>
          <input
            value={reviewer}
            onChange={(e) => setReviewer(e.target.value)}
            placeholder="Your name / SAP user"
            className="w-full border border-brand-silver rounded px-2 py-1 text-sm"
          />
          <textarea
            value={comments}
            onChange={(e) => setComments(e.target.value)}
            placeholder="Optional comments"
            rows={3}
            className="w-full border border-brand-silver rounded px-2 py-1 text-sm"
          />
          {error && <div className="text-xs text-rose-700">{error}</div>}
          <button
            onClick={submit}
            disabled={!reviewer}
            className="px-3 py-1 rounded bg-brand-amber text-brand-navy text-sm font-medium disabled:opacity-50"
          >
            Submit
          </button>
        </div>
      )}
    </div>
  );
}
