# IDoc Extractor — MVP Runbook

Pulls failed SAP IDocs from S/4HANA via the standard OData service and posts
them into the platform's existing `/ingest/errors` endpoint as
`source='IDOC'` rows. From the platform's perspective, an IDoc failure is
just another `raw.raw_errors` record — no new tables, no new endpoint.

## What's new

| File | Purpose |
|---|---|
| `extractors/idoc/models.py` | `IDoc`, `IDocStatus`, `IDocSegment` dataclasses |
| `extractors/idoc/severity.py` | Status-code → severity mapping; default failed-status sets |
| `extractors/idoc/connector.py` | `IDocConnector` ABC |
| `extractors/idoc/odata_client.py` | OData v2 client for `API_IDOC_SRV` |
| `extractors/idoc/mock_client.py` | JSON-fixture connector for tests / dry-runs |
| `extractors/idoc/transformer.py` | IDoc → `AbapErrorRecord` for `/ingest/errors` |
| `extractors/idoc/watermark.py` | DB-backed and in-memory watermark stores |
| `extractors/idoc/poller.py` | Watermark-driven pull-and-post loop |
| `extractors/idoc/cli.py` | `python -m extractors.idoc.cli --once / --watch` |
| `classifier/idoc_prompts.py` | IDoc calibration block for the Sonnet/Opus prompts |
| `tests/extractors/idoc/` | 73 unit + integration tests, no live SAP needed |

## How it fits the existing platform

```
┌───────────────────────┐     OData v2     ┌──────────────────────┐
│  S/4HANA              │ ───────────────► │  IDocPoller          │
│  API_IDOC_SRV         │                  │  (Windows service)   │
│  (no SAP code change) │                  └──────────┬───────────┘
└───────────────────────┘                             │ HTTPS POST
                                                      │ Bearer token
                                                      ▼
                                          ┌─────────────────────────┐
                                          │  /ingest/errors         │
                                          │  (existing FastAPI)     │
                                          └──────────┬──────────────┘
                                                     │
                                                     ▼
                                          ┌─────────────────────────┐
                                          │  raw.raw_errors         │
                                          │  source='IDOC'          │
                                          │  raw JSONB = full EDIDC │
                                          │   + EDIDS + EDID4       │
                                          └──────────┬──────────────┘
                                                     │
                                          (existing classifier picks it up;
                                           idoc_prompts adds calibration
                                           when source='IDOC')
```

Zero schema changes. The existing classifier, dashboard, and frontend
operate on IDoc rows the same way they operate on ST22/SM21 rows.

## Field mapping

| `raw.raw_errors` column | IDoc source                                            |
|---|---|
| `source`        | literal `'IDOC'`                                            |
| `error_id`      | EDIDC-DOCNUM (16-char IDoc number)                          |
| `occurred_at`   | EDIDC creation timestamp (ABAP `YYYYMMDDTHHMMSS` format)    |
| `severity`      | derived from STATUS+direction (see `severity.py`)           |
| `short_text`    | most-recent EDIDS-STATXT                                    |
| `long_text`     | full EDIDS history + selected EDID4 segments (rendered)     |
| `transaction`   | `WE19` (inbound) / `BD87` (outbound)                        |
| `object`        | EDIDC-MESTYP (message type — ORDERS, MATMAS, …)             |
| `sub_object`    | EDIDC-IDOCTP (basic type — ORDERS05, MATMAS05, …)           |
| `raw` (JSONB)   | full EDIDC + EDIDS list + EDID4 list                        |

## Setup on the laptop

### 1. Configure environment

Append to `.env`:

```env
# ─── IDoc OData extractor ─────────────────────────────────────────
IDOC_ODATA_BASE_URL=https://your-s4-host:44300/sap/opu/odata/sap/API_IDOC_SRV
IDOC_SAP_USERNAME=ALE_REPORT_USER
IDOC_SAP_PASSWORD=...                # use a service user with S_IDOC_ALL
IDOC_SAP_HOST=S4D                     # short host id; goes into system_id
IDOC_SAP_SYSNR=00
IDOC_SAP_CLIENT=100
IDOC_INGEST_URL=http://localhost:8000/ingest/errors
# ARTILEGENZ_INGEST_TOKEN already exists — re-used.
```

If your S/4 service exposes different entity-set or field names, override:
`IDOC_ENTITY_HEADER`, `IDOC_ENTITY_STATUS`, `IDOC_ENTITY_SEGMENT`,
`IDOC_FIELD_DOCNUM`, `IDOC_FIELD_STATUS`, etc. (see `odata_client.py`).

### 2. Validate against fixtures (no SAP, no DB writes)

```powershell
$env:ARTILEGENZ_INGEST_TOKEN = "smoke-token"
python -m extractors.idoc.cli --once `
  --mock tests\extractors\idoc\fixtures `
  --in-memory-watermark `
  --host S4D
```

The CLI logs `idoc_cli_done fetched=4 inserted=4` (the 5th fixture is a
status-69 informational marker that's intentionally excluded from the
default failed-status set).

### 3. Smoke against the running ingest endpoint

```powershell
# Make sure uvicorn api.app_ingest:app is running on :8000
python -m extractors.idoc.cli --once `
  --mock tests\extractors\idoc\fixtures `
  --in-memory-watermark `
  --host S4D
```

Verify rows landed:

```powershell
docker exec -it artilegenz-postgres psql -U artilegenz -d sap_errors -c `
  "SELECT hash_key, source, error_id, severity, short_text FROM raw.raw_errors WHERE source='IDOC';"
```

### 4. Wire to live SAP

```powershell
python -m extractors.idoc.cli --once
```

Or run continuously every 60 seconds:

```powershell
python -m extractors.idoc.cli --watch --interval 60
```

For Windows-service mode, copy `embeddings/install_nssm_service.ps1` and
adjust the script path — the patterns are identical.

### 5. Wire IDoc calibration into the classifier (one-line change)

In `classifier/prompts.py`, adjust `build_sonnet_messages`:

```python
from classifier.idoc_prompts import maybe_augment_user_text
# ...inside build_sonnet_messages, after constructing user_text:
user_text = maybe_augment_user_text(raw_error, user_text)
```

This adds the IDoc diagnostic block (status-code semantics, EDIDC/EDIDS/EDID4
roles, transaction-code remediation style) to the *user* turn — uncached, so
the cached system prompt stays unchanged for non-IDoc traffic and the
cache-hit ratio doesn't degrade.

## Status-code coverage

Default poll picks up these failure statuses:

**Inbound:** 51 (app document not posted), 56 (IDoc with errors added),
63 (error passing to application), 65 (ALE service error), 68 (no further
processing).

**Outbound:** 02 (error passing to port), 04 (control info error), 05
(translation error), 26 (syntax check error), 29 (ALE service error).

Override with `--statuses 51 56` for a narrower poll, or
`--statuses 51 56 60 61 63 65 68 02 04 05 26 29 30 32` for everything that
isn't a clean success.

Status 69 ("IDoc was edited") and 32 ("IDoc was edited" outbound) are
informational markers — excluded from the default poll. Use `--statuses 69`
to surface them when you want to track manual repairs.

## Cost / volume expectations

Same as the existing classifier — ~$0.11/error cached, ~$0.16 first-of-batch
on Sonnet 4.6 + Opus 4.7. At 50 IDoc failures/day → ~$5.50/day. The IDoc
calibration block adds ~600 tokens to the user turn (uncached), so each IDoc
classification costs a hair more than an ST22 dump.

## Future work

- **Extractor**: add a CDS-view-based connector for releases without
  `API_IDOC_SRV` exposed. Same `IDocConnector` interface.
- **Dashboard**: add a `source` filter chip to `frontend/src/pages/Triage.jsx`
  so operators can scope to IDocs.
- **Corpus enrichment**: scrape SAP Help / SCN IDoc threads into
  `intel.corpus` (existing scraper) so retrieval pulls IDoc-specific
  reference solutions.
- **Auto-resubmit**: when classifier confidence > 0.9 and root cause is
  data-only (e.g. status 51 with a known fix), surface a "WE19 + reprocess"
  action in the operator console.

## Tests

73 tests covering severity mapping, OData parsing (v2 ticks + ISO 8601 + v4
`value` arrays), mock connector filtering, transformer round-trip through
the `AbapErrorRecord` Pydantic model, poller end-to-end with mocked HTTP
transport, and prompt calibration.

```powershell
python -m pytest tests/extractors/idoc -v
```
