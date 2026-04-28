# Handoff prompt for laptop Claude Code

Paste **everything below the `--- BEGIN ---` line** as your first message to
the local Claude Code instance after launching it inside
`C:\sap-error-ai\sap-error-intelligence` on branch `claude/idoc-mvp`.

It runs the bootstrap, smokes the IDoc pipeline, and then **stops** at a
checkpoint so you can decide whether to wire live SAP. It will not paste
credentials anywhere or push to git without you saying so.

---
--- BEGIN ---

You are continuing IDoc MVP work for the Artilegenz SAP Error Intelligence
Platform. A sandboxed Claude session already built the code on branch
`claude/idoc-mvp` (origin: `mohanbg17/artilengenz`) and proved it end-to-end
in a Linux sandbox. Your job is to do the same on this Windows laptop, then
help wire it to the live S/4 system.

## Context you need

- 73 unit/integration tests pass in the sandbox (`tests/extractors/idoc/`)
- A real Postgres + FastAPI + IDoc CLI run in the sandbox already validated
  the full pipeline (cold start: 4 inserted, warm watermark: 0 fetched,
  forced re-fetch: 4 dedup, status-69 explicit poll: 1 inserted)
- A bug in `db/migrations/03_migration_classifier.sql` was caught and
  committed (commit `f05766d`)
- The classifier prompt extension `classifier/idoc_prompts.py` is wired but
  needs one line in `classifier/prompts.py` to activate (see step 7)

## What I want you to do — execute in order, stop at each CHECKPOINT

### 1. Verify branch and update

```powershell
git status
git fetch origin
git checkout claude/idoc-mvp
git pull
```

If branch doesn't exist locally yet, run:
```powershell
git fetch origin claude/idoc-mvp:claude/idoc-mvp
git checkout claude/idoc-mvp
```

### 2. Run the bootstrap script

```powershell
.\scripts\bootstrap.ps1
```

This sets up everything deterministic: .venv, pip install, .env (with
auto-generated `ARTILEGENZ_INGEST_TOKEN` if absent), Postgres via Docker
Compose, all migrations, runs the test suite, starts uvicorn in a new
Windows Terminal tab on `:8000`, runs the smoke test, installs frontend deps.

If anything fails, **diagnose it, fix it, re-run.** Common issues:
- ExecutionPolicy: re-invoke as `powershell.exe -ExecutionPolicy Bypass -File .\scripts\bootstrap.ps1`
- Docker Desktop not running: tell me to start it, then re-run
- Port 8000 already taken: find what's on it (`Get-NetTCPConnection -LocalPort 8000`) and ask whether to stop it
- Python not on PATH: try `py -3` instead of `python`

### 3. Re-run the smoke test independently

```powershell
.\scripts\smoke-idoc.ps1
```

Expect output ending with `Smoke test PASSED`. The 8 stages must all show
`ok:`. If any stage fails, stop and report the exact stage and error.

### 4. CHECKPOINT — stop and report

Tell me:
- Did the bootstrap and smoke pass?
- Are the IDoc rows visible in Postgres? Show me:
  ```powershell
  docker exec -it artilegenz-postgres psql -U artilegenz -d sap_errors -c `
    "SELECT source, error_id, severity, sub_object FROM raw.raw_errors WHERE source='IDOC' ORDER BY occurred_at;"
  ```
- What's running? `Get-NetTCPConnection -LocalPort 8000,5173 -State Listen -ErrorAction SilentlyContinue`

**Wait for me to say "go" before continuing.** I may want to inspect the
console output myself.

### 5. Optional: start the React operator console

When I say go:

```powershell
.\scripts\start-services.ps1 -Frontend
```

Open `http://127.0.0.1:5173` in a browser and verify the dashboard loads
showing the IDoc rows. Tell me if the UI shows them with `source=IDOC`.

### 6. Optional: ngrok tunnel for ABAP push from real S/4

When I say go (only if I want to also accept ABAP-pushed errors from a
real SAP system on this same laptop):

```powershell
.\scripts\start-services.ps1 -Tunnel
```

Capture the ngrok HTTPS public URL from the new tab and tell me. Don't
share it widely — anyone with the URL + token can post errors.

### 7. Wire the IDoc classifier prompt extension

This is a one-line edit so the classifier adds the IDoc calibration block
when it sees `source='IDOC'` rows. Show me the diff before applying:

In `classifier/prompts.py`, find `def build_sonnet_messages(` and inside
that function, after `user_text = ...` is constructed but before the
`user_messages = [...]` line, insert:

```python
from classifier.idoc_prompts import maybe_augment_user_text
user_text = maybe_augment_user_text(raw_error, user_text)
```

Or — cleaner — add the import at the top of `prompts.py` and the call
in-line. Show me the diff, wait for approval, then commit:

```powershell
git add classifier/prompts.py
git commit -m "Wire idoc_prompts.maybe_augment_user_text into Sonnet pass"
git push origin claude/idoc-mvp
```

### 8. CHECKPOINT — live SAP wiring decision

Ask me which path I want:

**Path A — start with mock fixtures only.** Stop here. The platform is
fully working with synthetic IDocs.

**Path B — wire the live S/4 OData service.** I will give you:
- The OData base URL (probably
  `https://<host>:<port>/sap/opu/odata/sap/API_IDOC_SRV`)
- A service user + password (BASIC auth)
- Whether the service uses non-standard entity / field names

You will:
1. Add those values to `.env` under the `IDOC_*` keys (see `.env.example`
   for the full set)
2. Run a one-shot live poll:
   ```powershell
   $env:IDOC_INGEST_URL = 'http://127.0.0.1:8000/ingest/errors'
   .venv\Scripts\python.exe -m extractors.idoc.cli --once `
     --in-memory-watermark `
     --statuses 51 56
   ```
3. If it fails (4xx, 5xx, JSON parse, certificate), diagnose and report.
   Common issues:
   - 401: wrong creds — ask me to re-check
   - 403: user lacks `S_IDOC_*` authorization — ask me to grant or pick a
     different user
   - 404: service path wrong — try variants
     (`/sap/opu/odata/sap/IDOC_DISPLAY_SRV`, custom Z-service)
   - 500 with "Filter cannot be applied": the field name for status or
     timestamp differs on this system. Set `IDOC_FIELD_STATUS` /
     `IDOC_FIELD_CREATED` env overrides to whatever the metadata
     document at `<base>/$metadata` advertises
   - SSL: if the system uses a self-signed cert, add `--insecure` to the
     CLI for now and tell me — we'll fix at the cert level later
4. When live poll succeeds, switch to DB-backed watermark (drop
   `--in-memory-watermark`) and run continuously:
   ```powershell
   .venv\Scripts\python.exe -m extractors.idoc.cli --watch --interval 60
   ```
   in its own Windows Terminal tab.

### 9. Optional: classifier sweep over the new IDoc rows

If `ANTHROPIC_API_KEY`, `VOYAGE_API_KEY`, and `PINECONE_API_KEY` are
present in `.env`:

```powershell
.\scripts\start-services.ps1 -Classifier
```

This kicks off `classifier\classify_worker.py` over `intel.v_pending_classifications`.
It costs ~$0.11/error cached (~$0.16 first-of-batch). Watch the tab for
each classification result and a final summary.

### 10. CHECKPOINT — done

Report:
- How many IDocs in `raw.raw_errors` with `source='IDOC'`
- How many classified (`SELECT * FROM intel.v_classification_status;`)
- The Sonnet+Opus diagnoses for the highest-severity one
  (`SELECT summary_md FROM intel.v_latest_classifications ORDER BY composite_confidence DESC LIMIT 1;`)

## Hard rules

- **Don't paste credentials in chat.** Always write to `.env` directly.
- **Don't `git push --force`** — ever, on any branch.
- **Don't drop tables / volumes** without me saying so explicitly.
- **Don't create a PR** unless I ask. The branch is on origin; reviewing
  via the GitHub web UI is fine.
- **Stop at every CHECKPOINT.** Wait for "go" before continuing.
- **Show diffs before committing** any code change.
- **Never `--no-verify` a commit.** If a hook fails, fix the underlying
  issue.

## When you're stuck

If a step fails and you can't figure out why in 2-3 attempts, **stop and
ask me.** Don't keep trying random fixes. Show me the exact command you
ran and the exact output you got.

--- END ---
