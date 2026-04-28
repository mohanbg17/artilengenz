# Classifier Runbook

Sonnet first pass + Opus critique for SAP error diagnosis with adaptive retrieval.

## Files

| File | Purpose |
|---|---|
| `03_migration_classifier.sql` | Adds tracking columns + views to `intel.classifications` |
| `retrieval.py` | Adaptive top-K retrieval across both Pinecone namespaces |
| `prompts.py` | Sonnet + Opus prompt builders with cache_control |
| `classifier_core.py` | Main pipeline: retrieval → Sonnet → Opus → Postgres |
| `classify.py` | CLI: classify a single error |
| `classify_worker.py` | Batch worker: continuous or `--once` mode |
| `run_classifier.bat` | Windows launcher for NSSM |
| `install_classifier_service.ps1` | NSSM service installer |

## Cost expectations (per error)

| Phase | Tokens (avg) | Cost (cached) | Cost (first-of-batch) |
|---|---|---|---|
| Sonnet 4.6 input + retrieval | ~10K cache + 3K live | ~$0.012 | $0.04 |
| Sonnet 4.6 output | ~1.5K | $0.023 | $0.023 |
| Opus 4.7 input | ~3K cache + 5K live | ~$0.027 | $0.04 |
| Opus 4.7 output | ~2K | $0.05 | $0.05 |
| **Total per error** | ~21K tokens | **~$0.11** | **~$0.16** |

So at ~$0.11 cached:
- 61 errors backlog: **~$7**
- 1000 errors/day: **~$110/day = ~$3.3K/month**

5-minute cache window means batches keep cache hot. If errors arrive sparsely (every 10 min), you mostly pay first-of-batch rates.

## Steps

### 1. Apply migration

```powershell
Get-Content C:\sap-error-ai\sap-error-intelligence\db\migrations\03_migration_classifier.sql | docker exec -i artilegenz-postgres psql -U artilegenz -d sap_errors
```

Expected output ends with `Migration 03 complete`.

Verify:
```powershell
docker exec -it artilegenz-postgres psql -U artilegenz -d sap_errors -c "SELECT * FROM intel.v_classification_status;"
```
Should show `classified=0 pending=61 failed=0 total=61`.

### 2. Get Anthropic API key

If you don't have one yet:
1. Go to https://console.anthropic.com/
2. Sign up / log in
3. **API Keys** → **Create Key**
4. Copy (starts with `sk-ant-...`)

### 3. Install dependency

```powershell
cd C:\sap-error-ai\sap-error-intelligence
.\.venv\Scripts\Activate.ps1
pip install anthropic
```

### 4. Smoke test the CLI on one error

```powershell
cd C:\sap-error-ai\sap-error-intelligence\classifier

# Set env vars
$env:ANTHROPIC_API_KEY = "sk-ant-..."        # your real key
$env:VOYAGE_API_KEY    = "pa-..."            # already set from earlier session
$env:PINECONE_API_KEY  = "pcsk_..."          # already set
$env:PG_DATABASE       = "sap_errors"
$env:PG_USER           = "artilegenz"
$env:PG_PASSWORD       = "artilegenz_local_dev"

# Pick latest error and classify it
python classify.py --latest
```

Expected output: ~30-60 seconds total, then a pretty diagnosis with:
- Root cause sentence
- Severity
- Verdict from Opus (CONFIRM/REFINE/REJECT)
- Numbered remediation steps
- Markdown summary
- Token counts (Sonnet + Opus)
- Adaptive retrieval split

If it looks reasonable -> classifier works end-to-end.

### 5. Backlog: classify all 61 records

```powershell
python classify_worker.py --once
```

Polls `v_pending_classifications` view in batches of 5, processes each through the full pipeline, sleeps briefly between calls. Total time ~10-15 min for 61 records.

Watch logs:
```
{"event": "worker_starting", ...}
{"event": "startup_status", "classified": 0, "pending": 61, "failed": 0, "total": 61}
{"event": "batch_starting", "size": 5}
{"event": "classifying", "hash_key": "abc123..."}
{"event": "classified_ok", "confidence": 0.85, "badge": "HIGH_CONFIDENCE", "cache_savings_sonnet": 8200, ...}
...
{"event": "no_pending_records"}
{"event": "once_mode_complete_exiting", "processed": 61}
```

`cache_savings_sonnet` shows how many tokens were read from cache (paying 10% rate). Usually 0 on first batch, then ~8000 on every following call within the 5-min window.

### 6. Inspect results

```powershell
# Status
docker exec -it artilegenz-postgres psql -U artilegenz -d sap_errors -c "SELECT * FROM intel.v_classification_status;"

# Top diagnoses by confidence
docker exec -it artilegenz-postgres psql -U artilegenz -d sap_errors -c "
SELECT badge, COUNT(*), AVG(composite_confidence)::numeric(4,3)
FROM intel.classifications WHERE status = 'classified'
GROUP BY badge ORDER BY 2 DESC;
"

# Verdicts
docker exec -it artilegenz-postgres psql -U artilegenz -d sap_errors -c "
SELECT opus_critique->>'verdict' AS verdict, COUNT(*)
FROM intel.classifications WHERE status = 'classified'
GROUP BY 1 ORDER BY 2 DESC;
"

# Sample one diagnosis
docker exec -it artilegenz-postgres psql -U artilegenz -d sap_errors -c "
SELECT top_proposal_title, composite_confidence, badge,
       LEFT(summary_md, 500) AS summary_preview
FROM intel.classifications WHERE status = 'classified'
ORDER BY created_at DESC LIMIT 1;
"
```

### 7. Install as Windows service (production)

After backlog is processed and quality looks good:

```powershell
# Run as Administrator
cd C:\sap-error-ai\sap-error-intelligence\classifier
.\install_classifier_service.ps1
```

Prompts for the three API keys (Anthropic, Voyage, Pinecone), then installs `ArtilegenzClassifierWorker` as an auto-starting Windows service.

Useful service commands:
```powershell
nssm status   ArtilegenzClassifierWorker
nssm stop     ArtilegenzClassifierWorker
nssm start    ArtilegenzClassifierWorker
nssm restart  ArtilegenzClassifierWorker
nssm edit     ArtilegenzClassifierWorker     # GUI

Get-Content C:\sap-error-ai\logs\classifier_stdout.log -Wait -Tail 50
```

## Operations cheat sheet

```powershell
# Manual reclassify (e.g. after model upgrade or prompt change)
docker exec -it artilegenz-postgres psql -U artilegenz -d sap_errors -c "
DELETE FROM intel.classifications WHERE status = 'classified';
"
nssm restart ArtilegenzClassifierWorker

# Find failed classifications
docker exec -it artilegenz-postgres psql -U artilegenz -d sap_errors -c "
SELECT classification_id, error_hash_key, error_message, retry_count
FROM intel.classifications WHERE status = 'failed'
ORDER BY created_at DESC LIMIT 20;
"

# Reset failed records for retry
docker exec -it artilegenz-postgres psql -U artilegenz -d sap_errors -c "
DELETE FROM intel.classifications WHERE status = 'failed' AND retry_count < 3;
"

# Cost so far (rough) -- count tokens from intel.classifications
docker exec -it artilegenz-postgres psql -U artilegenz -d sap_errors -c "
SELECT
  SUM((sonnet_tokens->>'input_tokens')::int) AS sonnet_input,
  SUM((sonnet_tokens->>'cache_read_input_tokens')::int) AS sonnet_cached,
  SUM((sonnet_tokens->>'output_tokens')::int) AS sonnet_output,
  SUM((opus_tokens->>'input_tokens')::int) AS opus_input,
  SUM((opus_tokens->>'output_tokens')::int) AS opus_output
FROM intel.classifications WHERE status = 'classified';
"
```

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| `anthropic.AuthenticationError` | Bad/missing key | Verify `ANTHROPIC_API_KEY` env var |
| `json.JSONDecodeError` from Sonnet/Opus | Model returned non-JSON | Check `sonnet_response` column in DB, may be a model glitch -- re-run; if persistent, prompt may need reinforcement |
| All confidence < 0.5 | Retrieval is poor (no relevant corpus) | Expand corpus (run scraper for more records) or check Pinecone has both namespaces |
| Service crashes on startup | Missing env vars | Edit service: `nssm edit ArtilegenzClassifierWorker` |
| `cache_read_input_tokens=0` always | Cache TTL expiring (slow/sparse traffic) | Reduce `CLASSIFY_POLL_INTERVAL_SEC` so calls stay within 5-min window, OR use `ttl: "1h"` extended cache (more expensive write but longer hit window) |
| Duplicate classifications | Worker raced | View `intel.v_latest_classifications` -- it dedups via DISTINCT ON. The raw table keeps history; latest view shows current. |

## What's next

After classifier runs cleanly:
1. **Frontend** (Vite+React+Tailwind dashboard reading from `intel.v_latest_classifications` + Postgres + Pinecone)
2. **Human feedback loop** -- accepted/corrected diagnoses feed back into corpus for retrieval improvement
3. **Snowflake replication** -- mirror Postgres → Snowflake for analytics + Tableau dashboards
4. **More extractors** -- SLG1, SM21 broaden raw_errors coverage
