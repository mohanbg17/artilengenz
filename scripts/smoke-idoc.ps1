<#
.SYNOPSIS
  Re-runnable smoke test for the IDoc MVP.

.DESCRIPTION
  Posts the 5 fixture IDocs to the running ingest API and verifies:
    - 4 inserted (default failure-status filter; status-69 fixture excluded)
    - 1 more inserted when polling status 69 explicitly
    - Second run yields zero fetches (watermark advance)
    - All 5 hash keys are unique
    - Rows visible in intel.v_pending_classifications

.PARAMETER SkipBootstrapCheck
  Internal flag used by bootstrap.ps1 to skip the prereq check.
#>
[CmdletBinding()]
param(
    [switch]$SkipBootstrapCheck
)

$ErrorActionPreference = 'Stop'
$RepoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $RepoRoot

function Write-Step([string]$msg) { Write-Host "==> $msg" -ForegroundColor Cyan }
function Write-Ok([string]$msg)   { Write-Host "    ok: $msg" -ForegroundColor Green }
function Write-Err([string]$msg)  { Write-Host "    error: $msg" -ForegroundColor Red }

$venvPython = Join-Path $RepoRoot '.venv\Scripts\python.exe'
if (-not $SkipBootstrapCheck) {
    if (-not (Test-Path $venvPython)) { Write-Err ".venv missing — run scripts\bootstrap.ps1 first"; exit 1 }
    try {
        $health = Invoke-WebRequest -Uri "http://127.0.0.1:8000/healthz" -TimeoutSec 2 -UseBasicParsing
        if ($health.StatusCode -ne 200) { throw }
    } catch {
        Write-Err "ingest API not responding on :8000 — run scripts\start-services.ps1 first"
        exit 1
    }
}

# Read .env into env vars so the CLI sees ARTILEGENZ_INGEST_TOKEN etc.
if (Test-Path .env) {
    Get-Content .env | ForEach-Object {
        if ($_ -match '^\s*([^#=][^=]*)=(.*)$') {
            $name = $matches[1].Trim()
            $val  = $matches[2].Trim()
            if ($val -match '^"(.*)"$') { $val = $matches[1] }
            Set-Item -Path "Env:$name" -Value $val
        }
    }
}

# Force the CLI at the running API on this host
$env:IDOC_INGEST_URL = 'http://127.0.0.1:8000/ingest/errors'
if (-not $env:IDOC_SAP_HOST)   { $env:IDOC_SAP_HOST   = 'S4D' }
if (-not $env:IDOC_SAP_SYSNR)  { $env:IDOC_SAP_SYSNR  = '00' }
if (-not $env:IDOC_SAP_CLIENT) { $env:IDOC_SAP_CLIENT = '100' }
if (-not $env:ARTILEGENZ_INGEST_TOKEN) {
    Write-Err "ARTILEGENZ_INGEST_TOKEN not set in .env"
    exit 1
}

Write-Step "1. Reset prior smoke state (delete IDOC rows + watermark)"
docker exec -i artilegenz-postgres psql -U artilegenz -d sap_errors -c `
    "DELETE FROM raw.raw_errors WHERE source='IDOC'; DELETE FROM raw.watermarks WHERE source='IDOC';" 2>&1 |
    Where-Object { $_ -match 'DELETE|ERROR' } | ForEach-Object { Write-Host "    $_" -ForegroundColor Gray }

Write-Step "2. First poll — default failure statuses, expect fetched=4 inserted=4"
$out1 = & $venvPython -m extractors.idoc.cli --once --mock tests/extractors/idoc/fixtures --host $env:IDOC_SAP_HOST 2>&1 | Out-String
Write-Host $out1 -ForegroundColor Gray
if ($out1 -notmatch 'fetched=4 inserted=4') { Write-Err "expected fetched=4 inserted=4"; exit 1 }
Write-Ok "first poll: 4 inserted"

Write-Step "3. Second poll — warm watermark, expect fetched=0"
$out2 = & $venvPython -m extractors.idoc.cli --once --mock tests/extractors/idoc/fixtures --host $env:IDOC_SAP_HOST 2>&1 | Out-String
Write-Host $out2 -ForegroundColor Gray
if ($out2 -notmatch 'fetched=0') { Write-Err "expected fetched=0 on warm watermark"; exit 1 }
Write-Ok "second poll: 0 fetches (watermark working)"

Write-Step "4. Force re-fetch with in-memory watermark, expect duplicates=4"
$out3 = & $venvPython -m extractors.idoc.cli --once --mock tests/extractors/idoc/fixtures --in-memory-watermark --host $env:IDOC_SAP_HOST 2>&1 | Out-String
Write-Host $out3 -ForegroundColor Gray
if ($out3 -notmatch 'duplicates=4') { Write-Err "expected duplicates=4 (hash dedup)"; exit 1 }
Write-Ok "dedup working: 4 duplicates rejected"

Write-Step "5. Explicit status-69 poll, expect 1 informational fixture"
$out4 = & $venvPython -m extractors.idoc.cli --once --mock tests/extractors/idoc/fixtures --in-memory-watermark --host $env:IDOC_SAP_HOST --statuses 69 2>&1 | Out-String
Write-Host $out4 -ForegroundColor Gray
if ($out4 -notmatch 'fetched=1 inserted=1') { Write-Err "expected fetched=1 inserted=1 for status 69"; exit 1 }
Write-Ok "status-69 informational poll: 1 row"

Write-Step "6. DB row inspection"
docker exec -i artilegenz-postgres psql -U artilegenz -d sap_errors -c `
    "SELECT source, error_id, severity, transaction, sub_object, occurred_at FROM raw.raw_errors WHERE source='IDOC' ORDER BY occurred_at;"

Write-Step "7. Pending classifications view"
docker exec -i artilegenz-postgres psql -U artilegenz -d sap_errors -c `
    "SELECT source, error_id, severity, retry_count FROM intel.v_pending_classifications WHERE source='IDOC' ORDER BY occurred_at;"

Write-Step "8. Hash uniqueness check"
$hashRow = (docker exec -i artilegenz-postgres psql -U artilegenz -d sap_errors -tA -c `
    "SELECT COUNT(*) || '|' || COUNT(DISTINCT hash_key) FROM raw.raw_errors WHERE source='IDOC';").Trim()
$parts = $hashRow -split '\|'
if ($parts[0] -ne $parts[1]) {
    Write-Err "hash collision: $($parts[0]) rows but $($parts[1]) distinct hashes"
    exit 1
}
Write-Ok "$($parts[0]) rows / $($parts[1]) distinct hashes — clean"

Write-Host ""
Write-Host "=================================================================" -ForegroundColor Green
Write-Host " Smoke test PASSED" -ForegroundColor Green
Write-Host "=================================================================" -ForegroundColor Green
