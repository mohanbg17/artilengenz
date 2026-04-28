import React, { useEffect, useMemo, useState } from 'react'
import ReactMarkdown from 'react-markdown'
import remarkGfm from 'remark-gfm'
import { api } from './api'

// =====================================================================
// utils
// =====================================================================
function fmtDate(s) {
  if (!s) return '—'
  try {
    const d = new Date(s)
    return d.toLocaleString('en-US', {
      year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false,
    })
  } catch { return s }
}

function fmtNum(n) {
  if (n == null) return '0'
  return Number(n).toLocaleString()
}

function shortHash(h) { return h ? h.slice(0, 12) : '—' }

const BADGE_COLORS = {
  HIGH_CONFIDENCE:   { fg: '#3DDC84', label: 'HIGH'    },
  MEDIUM_CONFIDENCE: { fg: '#FFAA33', label: 'MEDIUM'  },
  LOW_CONFIDENCE:    { fg: '#FF9800', label: 'LOW'     },
  UNCERTAIN:         { fg: '#9eaabe', label: 'UNCERTAIN' },
  FAILED:            { fg: '#FF4444', label: 'FAILED'  },
  null:              { fg: '#4a566c', label: '—'       },
}

const SEVERITY_COLORS = {
  CRITICAL: '#FF4444',
  ERROR:    '#FF9800',
  WARNING:  '#FFAA33',
  INFO:     '#3DDC84',
}

// =====================================================================
// header / hud
// =====================================================================
function Hud({ stats }) {
  const t = stats?.tokens || {}
  const cacheRate = t.sonnet_input
    ? Math.round((100 * (t.sonnet_cache_reads || 0)) / Math.max(1, t.sonnet_input))
    : 0
  return (
    <header className="border-b border-cable bg-carbon/60 backdrop-blur sticky top-0 z-30">
      <div className="max-w-[1600px] mx-auto px-6 py-4 flex items-center gap-8">
        <div className="flex items-center gap-3">
          <div className="w-9 h-9 border border-signal/60 flex items-center justify-center relative overflow-hidden">
            <div className="w-3.5 h-3.5 rounded-full bg-signal animate-pulse_slow"/>
            <div className="absolute inset-0 -translate-x-full animate-scan bg-gradient-to-r from-transparent via-signal/30 to-transparent"/>
          </div>
          <div>
            <div className="font-display text-[15px] tracking-wider2 uppercase text-silver">
              Artilegenz
            </div>
            <div className="font-mono text-[10px] tracking-widest text-muted -mt-0.5">
              ERROR_INTELLIGENCE_CONSOLE_v1.0
            </div>
          </div>
        </div>

        <div className="hidden md:flex flex-1 items-center gap-6">
          <Tile label="ERRORS"      val={fmtNum(stats?.total_errors)}     accent />
          <Tile label="CLASSIFIED"  val={fmtNum(stats?.unique_classified)} />
          <Tile label="CORPUS"      val={fmtNum(stats?.corpus_embedded)}   />
          <Tile label="CACHE_HIT%"  val={`${cacheRate}%`}                  />
          <Tile label="LAST_EVENT"  val={fmtDate(stats?.latest_error_at)?.split(' ')[1] || '—'} mono />
        </div>

        <div className="font-mono text-[10px] tracking-wider text-muted ml-auto">
          <span className="inline-block w-1.5 h-1.5 rounded-full bg-ok mr-2 animate-pulse_slow"/>
          UPLINK_OK
        </div>
      </div>
    </header>
  )
}

function Tile({ label, val, accent, mono }) {
  return (
    <div className="flex flex-col">
      <div className="font-mono text-[9px] tracking-widest text-muted uppercase">{label}</div>
      <div className={[
        'font-display text-lg leading-none',
        mono ? 'font-mono text-base' : '',
        accent ? 'text-signal' : 'text-silver',
      ].join(' ')}>{val}</div>
    </div>
  )
}

// =====================================================================
// list row
// =====================================================================
function ErrorRow({ row, selected, onClick }) {
  const badge = BADGE_COLORS[row.badge] || BADGE_COLORS.null
  const sevColor = SEVERITY_COLORS[row.severity] || '#7E8AA2'
  const conf = row.composite_confidence != null ? Number(row.composite_confidence) : null

  return (
    <button
      onClick={onClick}
      className={[
        'w-full text-left grid items-start gap-4 px-4 py-3 border-l-2',
        'hover:bg-carbon/70 transition-colors',
        selected
          ? 'bg-carbon border-l-signal'
          : 'border-l-transparent border-b border-b-cable/50',
      ].join(' ')}
      style={{
        gridTemplateColumns: '90px 90px 70px 1fr 100px 90px',
      }}
    >
      <div className="font-mono text-[11px] text-muted truncate">
        {fmtDate(row.occurred_at).split(' ')[1] || ''}
      </div>
      <div className="font-mono text-[11px] text-slate-50">{row.source || '—'}</div>
      <div className="font-mono text-[11px]" style={{color: sevColor}}>
        {row.severity || '—'}
      </div>
      <div className="font-body text-[12px] text-silver line-clamp-2 leading-snug">
        {row.short_text_preview || '(no preview)'}
      </div>
      <div className="font-mono text-[11px] flex items-center gap-1.5">
        <span
          className="inline-block w-1.5 h-1.5 rounded-full"
          style={{background: badge.fg}}
        />
        <span style={{color: badge.fg}}>{badge.label}</span>
      </div>
      <div className="font-mono text-[11px] text-right">
        {conf != null
          ? <span className="text-silver">{conf.toFixed(2)}</span>
          : <span className="text-muted">—</span>}
      </div>
    </button>
  )
}

function ListHeader() {
  return (
    <div
      className="grid items-center gap-4 px-4 py-2 border-b border-cable bg-carbon/70 sticky top-0 z-10"
      style={{ gridTemplateColumns: '90px 90px 70px 1fr 100px 90px' }}
    >
      {['TIME','SOURCE','SEVERITY','SHORT_TEXT','BADGE','CONF'].map(h => (
        <div key={h} className="font-mono text-[9px] tracking-widest uppercase text-muted">{h}</div>
      ))}
    </div>
  )
}

// =====================================================================
// detail panel
// =====================================================================
function FeedbackBar({ classification, existing, onSubmitted }) {
  const [busy, setBusy] = useState(false)
  const [done, setDone] = useState(existing?.accepted ?? null)

  async function submit(accepted) {
    setBusy(true)
    try {
      await api.feedback({
        classification_id: classification.classification_id,
        accepted,
        reviewer: 'mohan',
      })
      setDone(accepted)
      onSubmitted && onSubmitted()
    } catch (e) {
      console.error(e)
      alert('Feedback failed: ' + e.message)
    } finally {
      setBusy(false)
    }
  }

  if (!classification) return null

  return (
    <div className="flex items-center gap-3 px-4 py-3 border-t border-cable bg-graphite/40">
      <span className="font-mono text-[10px] uppercase tracking-widest text-muted">
        Diagnosis review:
      </span>
      <button
        disabled={busy}
        onClick={() => submit(true)}
        className={[
          'font-mono text-[11px] uppercase tracking-wider px-3 py-1.5 border transition',
          done === true
            ? 'bg-ok/20 border-ok text-ok'
            : 'border-cable text-slate-50 hover:border-ok hover:text-ok',
        ].join(' ')}
      >
        ▲ Accept
      </button>
      <button
        disabled={busy}
        onClick={() => submit(false)}
        className={[
          'font-mono text-[11px] uppercase tracking-wider px-3 py-1.5 border transition',
          done === false
            ? 'bg-critical/20 border-critical text-critical'
            : 'border-cable text-slate-50 hover:border-critical hover:text-critical',
        ].join(' ')}
      >
        ▼ Reject
      </button>
      {done !== null && (
        <span className="font-mono text-[10px] text-muted ml-auto">
          recorded
        </span>
      )}
    </div>
  )
}

function DetailPanel({ hash, onRefresh }) {
  const [data, setData]   = useState(null)
  const [loading, setLoading] = useState(true)
  const [err, setErr]     = useState(null)
  const [classifying, setClassifying] = useState(false)

  async function load() {
    if (!hash) return
    setLoading(true); setErr(null)
    try {
      const d = await api.errorDetail(hash)
      setData(d)
    } catch (e) {
      setErr(e.message)
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => { load() }, [hash])

  async function classifyNow() {
    setClassifying(true)
    try {
      await api.classifyNow(hash)
      await load()
      onRefresh && onRefresh()
    } catch (e) {
      alert('Classify failed: ' + e.message)
    } finally {
      setClassifying(false)
    }
  }

  if (!hash) {
    return (
      <div className="flex flex-col items-center justify-center h-full text-muted font-mono text-[11px] tracking-widest uppercase">
        <div className="border border-cable px-6 py-8 text-center">
          <div className="text-2xl mb-3 text-signal/40">◉</div>
          Select an error from the queue
        </div>
      </div>
    )
  }

  if (loading) {
    return (
      <div className="p-6 font-mono text-[11px] text-muted uppercase tracking-wider">
        loading_record [{shortHash(hash)}]...
      </div>
    )
  }
  if (err) {
    return (
      <div className="p-6 font-mono text-[11px] text-critical">
        ERROR: {err}
      </div>
    )
  }
  if (!data) return null

  const e = data.error
  const c = data.classification
  const fb = data.feedback

  return (
    <div className="flex flex-col h-full">
      {/* sticky header strip */}
      <div className="px-5 py-4 border-b border-cable bg-graphite/30 sticky top-0 z-10">
        <div className="flex items-center justify-between gap-4">
          <div className="flex-1 min-w-0">
            <div className="font-mono text-[10px] uppercase tracking-widest text-muted mb-1">
              record_id
            </div>
            <div className="font-mono text-[11px] text-silver break-all">
              {e.hash_key}
            </div>
          </div>
          {!c && (
            <button
              onClick={classifyNow}
              disabled={classifying}
              className="font-mono text-[10px] uppercase tracking-wider px-3 py-2 border border-signal text-signal hover:bg-signal hover:text-ink transition disabled:opacity-50 disabled:cursor-wait"
            >
              {classifying ? 'analyzing...' : '▶ Run Classifier'}
            </button>
          )}
        </div>
      </div>

      <div className="flex-1 overflow-y-auto">
        {/* Metadata grid */}
        <div className="grid grid-cols-2 gap-x-6 gap-y-3 px-5 py-4 border-b border-cable">
          <Field label="SOURCE"   v={e.source} />
          <Field label="SEVERITY" v={e.severity} color={SEVERITY_COLORS[e.severity]} />
          <Field label="OCCURRED" v={fmtDate(e.occurred_at)} />
          <Field label="SYSTEM"   v={e.system_id} />
          <Field label="USER"     v={e.user_name} />
          <Field label="ERROR_ID" v={e.error_id} />
          <Field label="TXN"      v={e.transaction} />
          <Field label="PROGRAM"  v={e.program} />
        </div>

        {/* Diagnosis */}
        {c ? (
          <>
            <div className="px-5 py-4 border-b border-cable bg-carbon/30">
              <div className="flex items-center justify-between">
                <div>
                  <div className="font-mono text-[10px] uppercase tracking-widest text-muted">
                    diagnosis
                  </div>
                  <div className="font-display text-amber-50 mt-1 text-base leading-tight">
                    {c.top_proposal_title}
                  </div>
                </div>
                <BadgePill badge={c.badge} confidence={c.composite_confidence} />
              </div>
            </div>

            <div className="px-5 py-4 md-render">
              <ReactMarkdown remarkPlugins={[remarkGfm]}>
                {c.summary_md || '_(no summary written)_'}
              </ReactMarkdown>
            </div>

            <FeedbackBar
              classification={c}
              existing={fb}
              onSubmitted={load}
            />
          </>
        ) : (
          <div className="px-5 py-12 text-center font-mono text-[11px] text-muted uppercase tracking-wider">
            <div className="text-3xl text-cable mb-3">◌</div>
            unclassified
          </div>
        )}

        {/* Raw payload */}
        <details className="border-t border-cable">
          <summary className="px-5 py-3 font-mono text-[10px] uppercase tracking-widest text-muted cursor-pointer hover:text-silver">
            ► raw_short_text
          </summary>
          <pre className="px-5 pb-4 text-[11px] font-mono text-slate-50 whitespace-pre-wrap break-all">
{e.short_text || '(empty)'}
          </pre>
        </details>

        {e.long_text && (
          <details className="border-t border-cable">
            <summary className="px-5 py-3 font-mono text-[10px] uppercase tracking-widest text-muted cursor-pointer hover:text-silver">
              ► raw_long_text [{e.long_text.length} chars]
            </summary>
            <pre className="px-5 pb-4 text-[11px] font-mono text-slate-50 whitespace-pre-wrap break-all max-h-96 overflow-y-auto">
{e.long_text}
            </pre>
          </details>
        )}
      </div>
    </div>
  )
}

function Field({ label, v, color }) {
  return (
    <div>
      <div className="font-mono text-[9px] uppercase tracking-widest text-muted">{label}</div>
      <div
        className="font-mono text-[12px] mt-0.5 break-all"
        style={{ color: color || '#D9D9D9' }}
      >
        {v || '—'}
      </div>
    </div>
  )
}

function BadgePill({ badge, confidence }) {
  const b = BADGE_COLORS[badge] || BADGE_COLORS.null
  const conf = confidence != null ? Number(confidence).toFixed(2) : '—'
  return (
    <div className="flex items-center gap-3 ml-4 shrink-0">
      <div
        className="font-mono text-[10px] tracking-widest px-2.5 py-1 border"
        style={{ color: b.fg, borderColor: b.fg }}
      >
        {b.label}
      </div>
      <div className="font-mono text-base text-silver">{conf}</div>
    </div>
  )
}

// =====================================================================
// app
// =====================================================================
export default function App() {
  const [stats, setStats]     = useState(null)
  const [errors, setErrors]   = useState([])
  const [total, setTotal]     = useState(0)
  const [selected, setSelected] = useState(null)
  const [loading, setLoading] = useState(true)
  const [filterClassified, setFilterClassified] = useState(false)
  const [filterBadge, setFilterBadge] = useState('')

  async function loadAll() {
    setLoading(true)
    try {
      const [s, e] = await Promise.all([
        api.stats(),
        api.listErrors({
          limit: 100,
          classified_only: filterClassified || undefined,
          badge: filterBadge || undefined,
        }),
      ])
      setStats(s)
      setErrors(e.errors)
      setTotal(e.total)
      // Auto-select the first classified error if nothing's selected
      if (!selected && e.errors.length) {
        const firstClassified = e.errors.find(r => r.classification_id) || e.errors[0]
        setSelected(firstClassified.hash_key)
      }
    } catch (err) {
      console.error(err)
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => { loadAll() }, [filterClassified, filterBadge])

  return (
    <div className="min-h-screen flex flex-col">
      <Hud stats={stats} />

      {/* Filter bar */}
      <div className="border-b border-cable bg-carbon/30 px-6 py-2 flex items-center gap-3">
        <span className="font-mono text-[10px] tracking-widest uppercase text-muted">filter:</span>
        <button
          onClick={() => setFilterClassified(v => !v)}
          className={[
            'font-mono text-[10px] tracking-wider px-3 py-1 border uppercase transition',
            filterClassified
              ? 'border-signal text-signal'
              : 'border-cable text-slate-50 hover:border-signal',
          ].join(' ')}
        >
          {filterClassified ? '▣' : '▢'} classified_only
        </button>
        <select
          value={filterBadge}
          onChange={(e) => setFilterBadge(e.target.value)}
          className="font-mono text-[10px] tracking-wider uppercase bg-carbon border border-cable text-slate-50 px-3 py-1 hover:border-signal focus:border-signal focus:outline-none"
        >
          <option value="">all_badges</option>
          <option value="HIGH_CONFIDENCE">high_confidence</option>
          <option value="MEDIUM_CONFIDENCE">medium_confidence</option>
          <option value="LOW_CONFIDENCE">low_confidence</option>
          <option value="UNCERTAIN">uncertain</option>
        </select>
        <span className="font-mono text-[10px] tracking-widest uppercase text-muted ml-auto">
          showing {errors.length} of {total}
        </span>
      </div>

      {/* Two-column layout */}
      <div className="flex-1 grid" style={{ gridTemplateColumns: 'minmax(0,1fr) minmax(540px,640px)' }}>
        {/* List */}
        <div className="border-r border-cable bg-ink overflow-y-auto" style={{maxHeight: 'calc(100vh - 110px)'}}>
          <ListHeader />
          {loading && (
            <div className="p-4 font-mono text-[11px] text-muted">loading...</div>
          )}
          {!loading && errors.length === 0 && (
            <div className="p-6 text-center font-mono text-[11px] text-muted uppercase tracking-wider">
              ◌ no records match
            </div>
          )}
          {errors.map((r) => (
            <ErrorRow
              key={r.hash_key}
              row={r}
              selected={selected === r.hash_key}
              onClick={() => setSelected(r.hash_key)}
            />
          ))}
        </div>

        {/* Detail */}
        <div className="bg-graphite/20 overflow-hidden" style={{maxHeight: 'calc(100vh - 110px)'}}>
          <DetailPanel hash={selected} onRefresh={loadAll} />
        </div>
      </div>
    </div>
  )
}
