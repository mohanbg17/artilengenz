// Vite proxies /api -> http://localhost:8000 (see vite.config.js)
const BASE = '/api'

async function jget(path) {
  const r = await fetch(`${BASE}${path}`)
  if (!r.ok) {
    const text = await r.text().catch(() => '')
    throw new Error(`GET ${path} failed: ${r.status} ${text.slice(0, 200)}`)
  }
  return r.json()
}

async function jpost(path, body) {
  const r = await fetch(`${BASE}${path}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: body == null ? '{}' : JSON.stringify(body),
  })
  if (!r.ok) {
    const text = await r.text().catch(() => '')
    throw new Error(`POST ${path} failed: ${r.status} ${text.slice(0, 200)}`)
  }
  return r.json()
}

export const api = {
  stats:        ()                => jget('/stats'),
  listErrors:   (params = {})     => {
    const q = new URLSearchParams(
      Object.entries(params).filter(([_, v]) => v != null && v !== '')
    ).toString()
    return jget(`/errors${q ? `?${q}` : ''}`)
  },
  errorDetail:  (hash)            => jget(`/errors/${hash}`),
  feedback:     (payload)         => jpost('/feedback', payload),
  classifyNow:  (hash)            => jpost(`/classify-now/${hash}`),
}
