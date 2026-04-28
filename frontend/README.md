# Artilegenz Error Intelligence — MVP Frontend

Operator-console dashboard for the Artilegenz SAP Error Intelligence platform.

## What's in this drop

| File | Goes to | Purpose |
|---|---|---|
| `dashboard.py` | `api/dashboard.py` | New FastAPI router with read/feedback/classify-now endpoints |
| `app_dashboard.py` | `api/app_dashboard.py` | Combined entrypoint (mounts ingest + dashboard) |
| `package.json` | `frontend/package.json` | Vite + React deps |
| `vite.config.js` | `frontend/vite.config.js` | Dev server + `/api` proxy → :8000 |
| `tailwind.config.js` | `frontend/tailwind.config.js` | Brand palette + operator-console tokens |
| `postcss.config.js` | `frontend/postcss.config.js` | PostCSS chain |
| `index.html` | `frontend/index.html` | HTML entry + Google Fonts |
| `src/main.jsx` | `frontend/src/main.jsx` | React entry |
| `src/index.css` | `frontend/src/index.css` | Tailwind + CRT scanline + markdown styles |
| `src/api.js` | `frontend/src/api.js` | API client |
| `src/App.jsx` | `frontend/src/App.jsx` | Main UI (HUD + list + detail + feedback) |

## Steps to run locally

### 1. Drop API files

```powershell
# From the deploy script that should accompany this README, or by hand:
Copy-Item dashboard.py     C:\sap-error-ai\sap-error-intelligence\api\dashboard.py     -Force
Copy-Item app_dashboard.py C:\sap-error-ai\sap-error-intelligence\api\app_dashboard.py -Force
```

### 2. Start the new FastAPI server

If `uvicorn api.app_ingest:app` is currently running, stop it (Ctrl+C in that window). Then:

```powershell
cd C:\sap-error-ai\sap-error-intelligence
.\.venv\Scripts\Activate.ps1

# Make sure env vars are set (PG creds + Anthropic/Voyage/Pinecone for classify-now endpoint)
$env:PG_DATABASE        = "sap_errors"
$env:PG_USER            = "artilegenz"
$env:PG_PASSWORD        = "artilegenz_local_dev"
$env:ANTHROPIC_API_KEY  = "sk-ant-..."
$env:VOYAGE_API_KEY     = "pa-..."
$env:PINECONE_API_KEY   = "pcsk_..."
$env:ARTILEGENZ_INGEST_TOKEN = "<your-ingest-token>"   # so SAP push still works

# Start combined app
uvicorn api.app_dashboard:app --host 0.0.0.0 --port 8000
```

Smoke test the new endpoints:

```powershell
curl http://localhost:8000/api/stats
curl "http://localhost:8000/api/errors?limit=5"
```

### 3. Drop frontend files

```powershell
$frontDir = "C:\sap-error-ai\sap-error-intelligence\frontend"
New-Item -ItemType Directory -Force -Path "$frontDir\src" | Out-Null

Copy-Item package.json       "$frontDir\package.json"       -Force
Copy-Item vite.config.js     "$frontDir\vite.config.js"     -Force
Copy-Item tailwind.config.js "$frontDir\tailwind.config.js" -Force
Copy-Item postcss.config.js  "$frontDir\postcss.config.js"  -Force
Copy-Item index.html         "$frontDir\index.html"         -Force
Copy-Item src\main.jsx       "$frontDir\src\main.jsx"       -Force
Copy-Item src\index.css      "$frontDir\src\index.css"      -Force
Copy-Item src\api.js         "$frontDir\src\api.js"         -Force
Copy-Item src\App.jsx        "$frontDir\src\App.jsx"        -Force
```

### 4. Install + run frontend

In a NEW PowerShell window:

```powershell
cd C:\sap-error-ai\sap-error-intelligence\frontend
npm install
npm run dev
```

Vite starts on **http://localhost:5173**. Open it in a browser.

### 5. What you should see

- HUD across top: ARTILEGENZ logo, ERRORS / CLASSIFIED / CORPUS / CACHE_HIT% / LAST_EVENT
- Filter bar: classified_only toggle, badge dropdown
- Left column: scrollable error list with TIME / SOURCE / SEVERITY / SHORT_TEXT / BADGE / CONF
- Right column: detail panel — metadata grid, diagnosis title with badge pill, markdown summary, ▲ Accept / ▼ Reject feedback bar
- Subtle CRT scanline overlay (the operator-console aesthetic)
- Click any row → detail loads
- For unclassified errors, a `▶ Run Classifier` button appears. Click it → ~30-60 sec wait → fresh diagnosis renders.

### 6. Submit feedback

Click ▲ Accept or ▼ Reject on a classification. POSTs to `/api/feedback`, writes to `intel.human_feedback`. The button highlights to show it was recorded.

Verify in DB:
```powershell
docker exec -it artilegenz-postgres psql -U artilegenz -d sap_errors -c "SELECT feedback_id, classification_id, accepted, reviewer, reviewed_at FROM intel.human_feedback ORDER BY reviewed_at DESC LIMIT 10;"
```

## Troubleshooting

| Symptom | Fix |
|---|---|
| `npm install` fails with EACCES / network | Check firewall; npm needs registry.npmjs.org |
| Frontend loads but stats are all zero | API isn't running on :8000 — start uvicorn |
| Frontend shows "Failed to fetch" | CORS issue. Should be auto-handled by Vite proxy. Check vite.config.js proxy points to 8000. |
| Classify button errors | Anthropic/Voyage/Pinecone env vars not in uvicorn's session |
| Markdown renders raw `**bold**` text | react-markdown didn't install. Re-run `npm install`. |

## What this is and isn't

**Is**: a working demo dashboard you can show in interviews, with real classifier output, real feedback writing to a real DB.

**Isn't**: production-grade. No auth. CORS lax. SQL injection safe (parameterized) but no rate limiting. Localhost only — don't expose this externally without adding security.

## Next steps after MVP

1. Process the remaining 56 unclassified records
2. NSSM-install both workers (embedding + classifier) so they auto-start
3. Improve ABAP extractor for richer FLIST data
4. Add SLG1 + SM21 extractors
5. Use accepted feedback as training signal for future corpus enrichment
