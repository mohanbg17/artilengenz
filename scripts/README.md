# scripts/

Automation for the Artilegenz SAP Error Intelligence Platform on Windows.

| File | Purpose |
|---|---|
| `bootstrap.ps1` | One-shot setup: prereqs, .venv, deps, .env, Postgres, migrations, tests, smoke. Idempotent. |
| `start-services.ps1` | Bring up long-running services (Postgres, uvicorn, Vite, ngrok, classifier) in Windows Terminal tabs. |
| `stop-services.ps1` | Graceful shutdown. `-StopDb` also stops Postgres; `-RemoveDb` wipes data (destructive). |
| `smoke-idoc.ps1` | Re-runnable IDoc end-to-end smoke. Resets state, runs 4 polls, verifies DB rows + dedup. |
| `handoff-claude-code.md` | Prompt to paste into local Claude Code so it drives the whole pipeline. |

## Daily workflow

```powershell
# First time on this laptop
.\scripts\bootstrap.ps1

# Day to day
.\scripts\start-services.ps1                      # API + Postgres
.\scripts\start-services.ps1 -Frontend            # also React UI
.\scripts\start-services.ps1 -Frontend -Tunnel    # also ngrok for ABAP push
.\scripts\smoke-idoc.ps1                          # validate IDoc pipeline
.\scripts\stop-services.ps1                       # done for the day
```

## Notes

- `.env` is **never overwritten**. Bootstrap auto-fills `ARTILEGENZ_INGEST_TOKEN`
  if absent but leaves everything else alone.
- All scripts assume `.venv\Scripts\python.exe` exists. Run `bootstrap.ps1`
  first if you haven't.
- Tabs require Windows Terminal (`wt`). Without it, services launch in
  detached windows.
- If `powershell.exe` complains about execution policy, run as:
  ```powershell
  powershell.exe -ExecutionPolicy Bypass -File .\scripts\bootstrap.ps1
  ```
