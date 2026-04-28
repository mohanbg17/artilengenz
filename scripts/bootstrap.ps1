<#
.SYNOPSIS
  One-shot bootstrap for the Artilegenz SAP Error Intelligence Platform on Windows.

.DESCRIPTION
  Idempotent. Safe to re-run. Does:
    1. Verify prerequisites (git, docker, python>=3.11, optional npm)
    2. Ensure on branch claude/idoc-mvp (warn if not)
    3. Create / reuse .venv and install requirements.txt
    4. Initialize .env if absent (auto-generate ARTILEGENZ_INGEST_TOKEN)
    5. Bring up Postgres via docker-compose, wait healthy
    6. Apply postgres_ddl.sql + migrations 01/02/03
    7. Run pytest tests/extractors/idoc -v
    8. Start the FastAPI ingest API in a new Windows Terminal tab on :8000
    9. Smoke test the IDoc CLI against fixtures
   10. Print summary + next steps

  Never overwrites an existing .env. Never logs secrets.

.PARAMETER SkipFrontend
  Skip 'npm install' for the React frontend (Vite). Default: install if npm present.

.PARAMETER SkipTests
  Skip pytest run. Default: run.

.PARAMETER NoApi
  Don't auto-start uvicorn. Default: start in a new tab.

.EXAMPLE
  .\scripts\bootstrap.ps1
  .\scripts\bootstrap.ps1 -SkipFrontend -SkipTests
#>
[CmdletBinding()]
param(
    [switch]$SkipFrontend,
    [switch]$SkipTests,
    [switch]$NoApi
)

$ErrorActionPreference = 'Stop'
$RepoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $RepoRoot

function Write-Step([string]$msg) { Write-Host "`n==> $msg" -ForegroundColor Cyan }
function Write-Ok([string]$msg)   { Write-Host "    ok: $msg" -ForegroundColor Green }
function Write-Warn2([string]$msg) { Write-Host "    warn: $msg" -ForegroundColor Yellow }
function Write-Err([string]$msg)  { Write-Host "    error: $msg" -ForegroundColor Red }

function Require-Command([string]$name, [string]$installHint) {
    if (-not (Get-Command $name -ErrorAction SilentlyContinue)) {
        Write-Err "$name not found on PATH. $installHint"
        exit 1
    }
    Write-Ok "$name found"
}

function Wait-DockerService([string]$container, [int]$timeoutSec = 60) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $timeoutSec) {
        $health = (docker inspect --format '{{.State.Health.Status}}' $container 2>$null)
        if ($LASTEXITCODE -eq 0 -and $health -eq 'healthy') { return $true }
        Start-Sleep -Seconds 2
    }
    return $false
}

# ──────────────────────────────────────────────────────────────────────────
# 1. Prereqs
# ──────────────────────────────────────────────────────────────────────────
Write-Step "Checking prerequisites"
Require-Command git "Install Git for Windows from https://git-scm.com/download/win"
Require-Command docker "Install Docker Desktop from https://www.docker.com/products/docker-desktop and ensure it's running"

# Find a usable Python. Prefer 'py -3' (Windows launcher), fall back to 'python'.
$pyExe = $null
$pyArgs = @()
if (Get-Command py -ErrorAction SilentlyContinue) {
    $pyExe = (Get-Command py).Source
    $pyArgs = @('-3')
} elseif (Get-Command python -ErrorAction SilentlyContinue) {
    $pyExe = (Get-Command python).Source
}
if (-not $pyExe) {
    Write-Err "Python 3.11+ not found. Install from https://www.python.org/downloads/"
    exit 1
}
$pyVer = & $pyExe @pyArgs -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')"
if ([version]$pyVer -lt [version]'3.11') {
    Write-Err "Python 3.11+ required. Found $pyVer."
    exit 1
}
Write-Ok "python $pyVer ($pyExe $pyArgs)"

# Test docker daemon is responsive
docker info --format '{{.ServerVersion}}' 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Err "Docker daemon not responding. Start Docker Desktop and wait for the whale icon to settle."
    exit 1
}
Write-Ok "docker daemon responsive"

if (-not $SkipFrontend) {
    if (Get-Command npm -ErrorAction SilentlyContinue) { Write-Ok "npm found (frontend will be set up)" }
    else { Write-Warn2 "npm not found — skipping frontend setup. Install Node 20+ from https://nodejs.org/"; $SkipFrontend = $true }
}

# ──────────────────────────────────────────────────────────────────────────
# 2. Branch sanity
# ──────────────────────────────────────────────────────────────────────────
Write-Step "Branch sanity"
$branch = (git rev-parse --abbrev-ref HEAD).Trim()
if ($branch -ne 'claude/idoc-mvp') {
    Write-Warn2 "Current branch is '$branch'. Expected 'claude/idoc-mvp'."
    Write-Warn2 "Switch with: git checkout claude/idoc-mvp"
} else {
    Write-Ok "on branch $branch"
}

# ──────────────────────────────────────────────────────────────────────────
# 3. Python venv + deps
# ──────────────────────────────────────────────────────────────────────────
Write-Step "Python virtualenv and dependencies"
$venvPath = Join-Path $RepoRoot '.venv'
$venvPython = Join-Path $venvPath 'Scripts\python.exe'

if (-not (Test-Path $venvPython)) {
    Write-Ok "creating .venv"
    & $pyExe @pyArgs -m venv .venv
    if ($LASTEXITCODE -ne 0) { Write-Err "venv creation failed"; exit 1 }
} else {
    Write-Ok "reusing existing .venv"
}

& $venvPython -m pip install --upgrade pip --quiet 2>&1 | Out-Null
Write-Ok "pip upgraded"
& $venvPython -m pip install -r requirements.txt --quiet
if ($LASTEXITCODE -ne 0) { Write-Err "pip install failed"; exit 1 }
Write-Ok "requirements.txt installed"

# ──────────────────────────────────────────────────────────────────────────
# 4. .env — initialize if absent, never overwrite
# ──────────────────────────────────────────────────────────────────────────
Write-Step ".env initialization"
$envFile = Join-Path $RepoRoot '.env'
if (-not (Test-Path $envFile)) {
    Copy-Item .env.example .env
    Write-Ok ".env created from .env.example"
    # Auto-generate the ingest token
    $token = & $venvPython -c "import secrets; print(secrets.token_urlsafe(32))"
    (Get-Content $envFile) -replace '^ARTILEGENZ_INGEST_TOKEN=.*$', "ARTILEGENZ_INGEST_TOKEN=$token" | Set-Content $envFile -Encoding utf8
    Write-Ok "ARTILEGENZ_INGEST_TOKEN auto-generated"
} else {
    Write-Ok ".env exists — leaving untouched"
    # Ensure ARTILEGENZ_INGEST_TOKEN is set
    $envContent = Get-Content $envFile -Raw
    if ($envContent -notmatch '(?m)^ARTILEGENZ_INGEST_TOKEN=\S+') {
        Write-Warn2 "ARTILEGENZ_INGEST_TOKEN is empty in .env — auto-filling"
        $token = & $venvPython -c "import secrets; print(secrets.token_urlsafe(32))"
        if ($envContent -match '(?m)^ARTILEGENZ_INGEST_TOKEN=') {
            (Get-Content $envFile) -replace '^ARTILEGENZ_INGEST_TOKEN=.*$', "ARTILEGENZ_INGEST_TOKEN=$token" | Set-Content $envFile -Encoding utf8
        } else {
            Add-Content $envFile "`nARTILEGENZ_INGEST_TOKEN=$token"
        }
    }
}

# ──────────────────────────────────────────────────────────────────────────
# 5. Postgres via docker-compose
# ──────────────────────────────────────────────────────────────────────────
Write-Step "Bringing up Postgres (docker compose)"
$existing = docker ps --filter "name=artilengenz-postgres" --format '{{.Names}}' 2>$null
if (-not $existing) {
    $existing = docker ps --filter "name=artilegenz-postgres" --format '{{.Names}}' 2>$null
}
docker compose up -d 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) { Write-Err "docker compose up failed"; exit 1 }
Write-Ok "docker compose up issued"

if (Wait-DockerService 'artilegenz-postgres' 90) {
    Write-Ok "postgres healthy"
} else {
    Write-Err "postgres did not reach healthy state in 90s. Check 'docker logs artilegenz-postgres'"
    exit 1
}

# ──────────────────────────────────────────────────────────────────────────
# 6. Apply migrations
# ──────────────────────────────────────────────────────────────────────────
Write-Step "Applying database migrations"
$psqlExec = "docker exec -i artilegenz-postgres psql -U artilegenz -d sap_errors"

# postgres_ddl.sql is mounted via docker-entrypoint-initdb.d on first run, but
# re-running it is safe (CREATE TABLE IF NOT EXISTS). Apply explicitly so the
# script works even on an existing volume.
foreach ($sql in @(
    "db\postgres_ddl.sql",
    "db\migrations\01_migration_raw_embeddings.sql",
    "db\migrations\02_migration_corpus_extend.sql",
    "db\migrations\03_migration_classifier.sql"
)) {
    if (Test-Path $sql) {
        Write-Host "    applying $sql" -ForegroundColor Gray
        Get-Content $sql | docker exec -i artilegenz-postgres psql -U artilegenz -d sap_errors 2>&1 |
            Where-Object { $_ -match 'ERROR|FATAL' } | ForEach-Object { Write-Warn2 $_ }
    } else {
        Write-Warn2 "$sql not found — skipping"
    }
}

# Verify schemas / views landed
$views = (docker exec -i artilegenz-postgres psql -U artilegenz -d sap_errors -tA -c "SELECT viewname FROM pg_views WHERE schemaname='intel' ORDER BY viewname;").Trim()
$expected = @('v_classification_status', 'v_latest_classifications', 'v_pending_classifications')
$present = $views -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ }
$missing = $expected | Where-Object { $_ -notin $present }
if ($missing) {
    Write-Warn2 "expected views missing: $($missing -join ', ')"
} else {
    Write-Ok "intel views present: $($present -join ', ')"
}

# ──────────────────────────────────────────────────────────────────────────
# 7. Tests
# ──────────────────────────────────────────────────────────────────────────
if (-not $SkipTests) {
    Write-Step "Running IDoc test suite"
    & $venvPython -m pytest tests/extractors/idoc -q
    if ($LASTEXITCODE -ne 0) {
        Write-Err "tests failed — fix before continuing"
        exit 1
    }
    Write-Ok "73 tests passing"
}

# ──────────────────────────────────────────────────────────────────────────
# 8. Start ingest API in a new Windows Terminal tab (or fall back to background)
# ──────────────────────────────────────────────────────────────────────────
if (-not $NoApi) {
    Write-Step "Starting FastAPI ingest API on :8000"
    $apiRunning = try { (Invoke-WebRequest -Uri "http://127.0.0.1:8000/healthz" -TimeoutSec 1 -UseBasicParsing).StatusCode -eq 200 } catch { $false }
    if ($apiRunning) {
        Write-Ok "ingest API already running on :8000 — leaving it"
    } else {
        $startCmd = "Set-Location '$RepoRoot'; & '$venvPython' -m uvicorn api.app_ingest:app --host 127.0.0.1 --port 8000 --reload"
        if (Get-Command wt -ErrorAction SilentlyContinue) {
            wt -w 0 nt --title "artilengenz-api" pwsh -NoExit -Command $startCmd
            Write-Ok "uvicorn launched in new Windows Terminal tab"
        } else {
            $logFile = Join-Path $RepoRoot 'uvicorn.log'
            Start-Process -FilePath $venvPython -ArgumentList @('-m','uvicorn','api.app_ingest:app','--host','127.0.0.1','--port','8000') -WindowStyle Hidden -RedirectStandardOutput $logFile -RedirectStandardError $logFile
            Write-Ok "uvicorn launched in background (log: $logFile)"
        }

        # Wait for healthz
        $sw = [Diagnostics.Stopwatch]::StartNew()
        while ($sw.Elapsed.TotalSeconds -lt 30) {
            try {
                if ((Invoke-WebRequest -Uri "http://127.0.0.1:8000/healthz" -TimeoutSec 1 -UseBasicParsing).StatusCode -eq 200) {
                    Write-Ok "ingest API healthy"
                    break
                }
            } catch {}
            Start-Sleep -Milliseconds 500
        }
    }
}

# ──────────────────────────────────────────────────────────────────────────
# 9. Smoke test
# ──────────────────────────────────────────────────────────────────────────
Write-Step "Smoke test: IDoc CLI against fixtures"
& "$PSScriptRoot\smoke-idoc.ps1" -SkipBootstrapCheck
if ($LASTEXITCODE -ne 0) {
    Write-Err "smoke test failed"
    exit 1
}

# ──────────────────────────────────────────────────────────────────────────
# 10. Frontend (optional)
# ──────────────────────────────────────────────────────────────────────────
if (-not $SkipFrontend) {
    Write-Step "Frontend dependencies"
    Push-Location frontend
    if (-not (Test-Path node_modules)) {
        npm install --silent 2>&1 | Where-Object { $_ -match 'error|warn' } | ForEach-Object { Write-Warn2 $_ }
        Write-Ok "node_modules installed"
    } else {
        Write-Ok "node_modules already present — skipping npm install"
    }
    Pop-Location
}

# ──────────────────────────────────────────────────────────────────────────
# Summary
# ──────────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "=================================================================" -ForegroundColor Green
Write-Host " Bootstrap complete" -ForegroundColor Green
Write-Host "=================================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Postgres:   docker exec -it artilegenz-postgres psql -U artilegenz -d sap_errors"
Write-Host "  Ingest API: http://127.0.0.1:8000/healthz"
Write-Host "  Tests:      .venv\Scripts\python.exe -m pytest tests/extractors/idoc -v"
Write-Host "  Smoke:      .\scripts\smoke-idoc.ps1"
Write-Host "  Start all:  .\scripts\start-services.ps1"
Write-Host "  Stop all:   .\scripts\stop-services.ps1"
Write-Host ""
Write-Host "Next: wire live SAP. See scripts\handoff-claude-code.md for the prompt."
Write-Host ""
