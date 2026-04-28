# Corpus Scraper Runbook

Stack Overflow → Postgres → Pinecone, ~500 SAP error/solution records.

## Files

| File | Purpose |
|---|---|
| `02_migration_corpus_extend.sql` | Adds tracking columns + dedup unique index to `intel.corpus` |
| `fetch_so.py` | Hits Stack Exchange API for SAP-tagged questions w/ accepted answers |
| `load_corpus.py` | Inserts JSON output into `intel.corpus` (idempotent, ON CONFLICT DO UPDATE) |
| `embed_corpus.py` | Embeds corpus rows → Pinecone `corpus_v1` namespace |
| `query_corpus.py` | Verify corpus retrieval works |
| `run_corpus_pipeline.bat` | One-shot orchestrator |

## Steps

### 1. Apply migration

```powershell
Get-Content C:\sap-error-ai\sap-error-intelligence\db\migrations\02_migration_corpus_extend.sql | docker exec -i artilegenz-postgres psql -U artilegenz -d sap_errors
```

Expected: `Migration 02 complete`. Adds columns to `intel.corpus`, doesn't drop existing data.

Verify:
```powershell
docker exec -it artilegenz-postgres psql -U artilegenz -d sap_errors -c "SELECT * FROM intel.v_corpus_status;"
```

### 2. Smoke test the SO scraper (20 records, ~1 min)

```powershell
cd C:\sap-error-ai\sap-error-intelligence\scraper
..\.venv\Scripts\Activate.ps1

# (env vars from earlier session should still be set; if not, re-set them)

python fetch_so.py --target 20 --output smoke.json
```

Expected: ~1 min runtime, JSON file with 20 records. Inspect:
```powershell
python -c "import json; d=json.load(open('smoke.json')); print(f'records: {len(d)}'); print('first title:', d[0]['title']); print('first module:', d[0]['module'])"
```

### 3. Run full pipeline (target: 500)

```powershell
.\run_corpus_pipeline.bat 500
```

Three stages, ~15-25 min total:

**Stage 1** — `fetch_so.py --target 500`:
- Iterates SAP_TAGS (`sap`, `abap`, `sap-erp`, `sap-basis`, `sap-fiori`, `sap-hana`, `sap-bw`)
- 1.0s delay between API calls
- Filters: accepted answer, score≥3, error keywords match, last 5 years, body length ≥100 chars
- Stops when 500 collected or all tags exhausted
- Writes `scraped_so.json`

**Stage 2** — `load_corpus.py`:
- Reads JSON, normalizes, generates `corpus_id` = sha256(`source_platform:source_id`)
- Strips HTML, extracts error signatures (CX_*, COMPUTE_*, MESSAGE patterns)
- UPSERTs into `intel.corpus` — running twice is safe, no duplicates
- ~30s for 500 records

**Stage 3** — `embed_corpus.py`:
- Reads unembedded rows ordered by upvote_score DESC (best first)
- Voyage `voyage-3-large`, 32 records/batch
- Upserts to Pinecone `corpus_v1` namespace
- Marks `embedding_version = 'voyage-3-large'` in Postgres
- ~3-5 min for 500 records (depends on Voyage rate limits)

### 4. Verify

**Postgres:**
```powershell
docker exec -it artilegenz-postgres psql -U artilegenz -d sap_errors -c "SELECT * FROM intel.v_corpus_status;"
docker exec -it artilegenz-postgres psql -U artilegenz -d sap_errors -c "SELECT system_module, COUNT(*) FROM intel.corpus GROUP BY system_module ORDER BY 2 DESC;"
```

**Pinecone:**
```powershell
python -c "import os; from pinecone import Pinecone; pc=Pinecone(api_key=os.environ['PINECONE_API_KEY']); print(pc.Index('artilegenz-sap-errors').describe_index_stats())"
```
Expect both namespaces: `raw_errors_v1` (~61 vectors) and `corpus_v1` (~500 vectors).

**Retrieval:**
```powershell
python query_corpus.py --text "ABAP zero divide computation error" --top-k 5
python query_corpus.py --text "authorization check failed" --top-k 5
python query_corpus.py --text "transport error" --module BASIS --top-k 5
```

### 5. Re-run incrementally (anytime)

The pipeline is idempotent. Re-running:
- `fetch_so.py` re-fetches but `load_corpus.py` UPSERTs (no duplicates)
- `embed_corpus.py` only embeds rows where `embedding_version` is NULL or stale
- Increase `--target` to gradually grow the corpus

```powershell
# Extend corpus to 1000 records later
.\run_corpus_pipeline.bat 1000
```

## Quality tuning

If results aren't useful:

| Symptom | Fix |
|---|---|
| Too few results returned | Lower `score >= 3` filter in `fetch_so.py` `passes_quality_filter()` |
| Too much off-topic content | Tighten `ERROR_KEYWORDS` regex in `fetch_so.py` |
| Module classification wrong | Edit `MODULE_TAG_MAP` in `fetch_so.py` |
| Duplicate-feeling results | Already deduped by `(source_platform, source_id)` unique index |

## What's next

After corpus is loaded + embedded, the **classifier** will:
1. Take a raw error (`raw.raw_errors` row)
2. Embed it as a query
3. Retrieve top-K from BOTH `raw_errors_v1` (similar past errors) AND `corpus_v1` (real solutions)
4. Hand to Claude Sonnet with retrieved context for diagnosis
5. Self-critique with Claude Opus on low-confidence cases
6. Store result in `intel.classifications` (JSON + markdown)

That's the next session's build.

## Troubleshooting

**`requests.exceptions.HTTPError: 502 Bad Gateway`**
SO API hiccup. Wait 30s, re-run — fetch is incremental, won't re-fetch what already loaded.

**`fetch_so.py` returns far fewer than 500**
Quality filters too strict for available SO questions. Loosen or run on more tags. Check log lines like `kept=2 total=137` to see filter pass rate.

**`embed_corpus.py` says `voyage_failed`**
Check Voyage rate limits (free tier: 3 RPM / 10K TPM until billing added). Reduce `--batch-size` or add card.

**Pinecone `Resource not found: index`**
Index name mismatch. Confirm `PINECONE_INDEX=artilegenz-sap-errors` and that the index exists in Pinecone dashboard.
