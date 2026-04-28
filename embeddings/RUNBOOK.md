# Embedding Worker Deployment Runbook

This guide takes you from "Voyage and Pinecone signed up" to "embeddings flowing automatically as new SAP errors arrive."

## Prerequisites checklist

- [x] Postgres running in Docker (`artilegenz-postgres` container)
- [x] FastAPI uvicorn ingestion endpoint running
- [x] ngrok tunnel running with reserved domain
- [x] ABAP push pipeline tested end-to-end (59 records in `raw.raw_errors`)
- [ ] Voyage AI account + API key (starts with `pa-`)
- [ ] Pinecone account + API key (starts with `pcsk_`)
- [ ] Pinecone index `artilegenz-sap-errors` created (1024 dim, cosine, AWS us-east-1)
- [ ] NSSM downloaded from https://nssm.cc/download (extract to `C:\nssm\`)

---

## Step 1: Apply the SQL migration

Creates `intel.raw_embeddings_log` table + helper views.

```powershell
# Copy migration to a place Docker can read it
docker cp C:\path\to\01_migration_raw_embeddings.sql artilegenz-postgres:/tmp/migration.sql

# Apply
docker exec -it artilegenz-postgres psql -U artilegenz -d sap_errors -f /tmp/migration.sql
```

Expected output ends with:
```
 status
----------------------
 Migration complete
(1 row)
```

Verify:
```powershell
docker exec -it artilegenz-postgres psql -U artilegenz -d sap_errors -c "SELECT * FROM intel.v_embedding_status;"
```

Should show `pending=59 embedded=0 failed=0 total=59`.

---

## Step 2: Copy worker files to project

```powershell
# Worker files into embeddings/ folder of the existing project
$src = "C:\path\to\downloaded\embedding-worker"
$dst = "C:\sap-error-ai\sap-error-intelligence\embeddings"

Copy-Item "$src\raw_worker.py"          "$dst\"
Copy-Item "$src\raw_embeddings_db.py"   "$dst\"
Copy-Item "$src\query_similar.py"       "$dst\"
Copy-Item "$src\run_worker.bat"         "$dst\"
Copy-Item "$src\.env.template"          "$dst\.env.template"
```

---

## Step 3: Smoke test from a PowerShell window (no service yet)

Best to validate the worker runs cleanly before installing as a service.

```powershell
cd C:\sap-error-ai\sap-error-intelligence
.\.venv\Scripts\Activate.ps1

# Set env vars for this session
$env:VOYAGE_API_KEY    = "pa-PASTE-YOUR-KEY-HERE"
$env:PINECONE_API_KEY  = "pcsk_PASTE-YOUR-KEY-HERE"
$env:VOYAGE_MODEL      = "voyage-3-large"
$env:PINECONE_INDEX    = "artilegenz-sap-errors"
$env:PINECONE_NAMESPACE = "raw_errors_v1"
$env:PG_DATABASE       = "sap_errors"
$env:PG_USER           = "artilegenz"
$env:PG_PASSWORD       = "artilegenz_local_dev"

# Dry run first -- no API calls, just verifies DB query and config
cd embeddings
python raw_worker.py --once --dry-run
```

Expected: logs show `pending_batch_found count=32`, then `dry_run_skip_api_calls`, then `once_mode_complete_exiting` (after one or two passes through the 59 records).

Now real run (1 batch of 32):
```powershell
python raw_worker.py --once
```

Expected output (JSON logs):
```
{"event": "worker_starting", "voyage_model": "voyage-3-large", ...}
{"event": "pinecone_create_index", "name": "artilegenz-sap-errors", ...}   # only on first run
{"event": "startup_status", "embedded": 0, "pending": 59, ...}
{"event": "pending_batch_found", "count": 32}
{"event": "batch_prepared", "size": 32, ...}
{"event": "batch_complete", "embedded": 32}
{"event": "batch_outcome", "success": 32, "failure": 0}
{"event": "pending_batch_found", "count": 27}
...
{"event": "no_pending_records", ...}
{"event": "once_mode_complete_exiting"}
```

Verify in Postgres:
```powershell
docker exec -it artilegenz-postgres psql -U artilegenz -d sap_errors -c "SELECT * FROM intel.v_embedding_status;"
```

Should show `embedded=59 pending=0 failed=0 total=59`.

---

## Step 4: Try a similarity search

```powershell
python query_similar.py --text "ABAP zerodivide error" --top-k 5
```

Expected: top 5 matches from your 59 records, ranked by cosine similarity. The `FC022COMPUTE_INT_ZERODIVIDE` records should rank highest.

Try filter:
```powershell
python query_similar.py --text "authorization failure" --source ST22 --top-k 3
```

Try "find similar to existing":
```powershell
# Pick any hash_key from raw_errors
docker exec -it artilegenz-postgres psql -U artilegenz -d sap_errors -c "SELECT hash_key, short_text FROM raw.raw_errors LIMIT 1;"
# Copy the hash_key, then:
python query_similar.py --hash-key <paste-hash-here> --top-k 5
```

If similarity search returns sensible results -> embeddings work end-to-end.

---

## Step 5: Install as Windows service via NSSM

Now that smoke test passed, register as a service.

```powershell
# Run as Administrator!
cd C:\sap-error-ai\sap-error-intelligence\embeddings
# (Use the install script copied earlier, or wherever you saved it)
.\install_nssm_service.ps1
```

Script will prompt for Voyage and Pinecone keys (so they don't end up in any committed file).

Verify:
```powershell
nssm status ArtilegenzEmbeddingWorker
# Expected: SERVICE_RUNNING

# Watch logs
Get-Content C:\sap-error-ai\logs\embedding_worker_stdout.log -Wait -Tail 50
```

---

## Step 6: End-to-end test -- new error -> embedded automatically

Now push fresh ST22 records from SAP:

1. **SAP** -> SE38 -> `Z_ARTILEGENZ_ERROR_EXPORT` -> F8
2. Within 30 seconds (poll interval), worker picks up new records
3. Watch service logs:
   ```powershell
   Get-Content C:\sap-error-ai\logs\embedding_worker_stdout.log -Wait -Tail 50
   ```

You should see new `pending_batch_found` -> `batch_complete` cycles.

Verify in Postgres:
```powershell
docker exec -it artilegenz-postgres psql -U artilegenz -d sap_errors -c "
  SELECT
    DATE_TRUNC('minute', l.embedded_at) AS embedded_minute,
    COUNT(*) AS records
  FROM intel.raw_embeddings_log l
  GROUP BY 1
  ORDER BY 1 DESC
  LIMIT 10;
"
```

---

## Operations cheat sheet

```powershell
# Service control
nssm status   ArtilegenzEmbeddingWorker
nssm start    ArtilegenzEmbeddingWorker
nssm stop     ArtilegenzEmbeddingWorker
nssm restart  ArtilegenzEmbeddingWorker
nssm edit     ArtilegenzEmbeddingWorker     # GUI to edit config
nssm remove   ArtilegenzEmbeddingWorker confirm

# Logs
Get-Content C:\sap-error-ai\logs\embedding_worker_stdout.log -Wait -Tail 50
Get-Content C:\sap-error-ai\logs\embedding_worker_stderr.log -Tail 100

# DB status checks
docker exec -it artilegenz-postgres psql -U artilegenz -d sap_errors -c "SELECT * FROM intel.v_embedding_status;"
docker exec -it artilegenz-postgres psql -U artilegenz -d sap_errors -c "SELECT * FROM intel.raw_embeddings_log WHERE embed_error IS NOT NULL LIMIT 10;"

# Force re-embed everything (e.g. after upgrading model)
docker exec -it artilegenz-postgres psql -U artilegenz -d sap_errors -c "DELETE FROM intel.raw_embeddings_log WHERE embedding_version != 'voyage-3-large';"
nssm restart ArtilegenzEmbeddingWorker
```

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| `config_missing_env_vars` on start | Env vars not passed to service | Re-run `install_nssm_service.ps1` or `nssm edit` to set them |
| `pinecone.exceptions.PineconeApiException: 404` | Index name mismatch | Check Pinecone dashboard; ensure `artilegenz-sap-errors` exists with 1024 dim |
| `voyageai.error.AuthenticationError` | Bad/expired key | Regenerate at https://dashboard.voyageai.com/ |
| Worker runs but `pending` count never drops | Records have `embed_error` set, retry maxed | `SELECT embed_error FROM intel.raw_embeddings_log WHERE embed_error IS NOT NULL;` to see why |
| Service won't start | Path / permission issue | Check `C:\sap-error-ai\logs\embedding_worker_stderr.log` |
| `psycopg2.OperationalError: connection refused` | Postgres container stopped | `docker ps`; `docker start artilegenz-postgres` |

---

## What's next

After this is running stable for a day or two, the next layer is the **Claude classifier** (option 3) -- pull a raw error, retrieve top-k similar past errors via Pinecone, prompt Claude to propose root cause + remediation with self-critique. That writes to `intel.classifications` and feeds the eventual frontend.
