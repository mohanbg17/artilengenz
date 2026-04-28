<#
.SYNOPSIS
  Start the long-running services (Postgres, ingest API, frontend, optional ngrok).

.DESCRIPTION
  Starts each service in its own Windows Terminal tab so logs stay visible.
  Idempotent — already-running services are left alone.

  Service map:
    Postgres        docker compose up -d
    Ingest API      uvicorn api.app_ingest:app --port 8000  (tab: artilengenz-api)
    Frontend        npm run dev                              (tab: artilengenz-ui)
    ngrok           ngrok http 8000                          (tab: artilengenz-ngrok, with -Tunnel)
    Classifier      python classifier/classify_worker.py    (tab: artilengenz-classifier, with -Classifier)

.PARAMETER Frontend
  Start the React + Vite dev server on :5173.

.PARAMETER Tunnel
  Start an ngrok tunnel that exposes :8000. Requires ngrok on PATH.

.PARAMETER Classifier
  Start the classifier worker (requires ANTHROPIC_API_KEY, VOYAGE_API_KEY,
  PINECONE_API_KEY in .env).

.EXAMPLE
  .\scripts\start-services.ps1
  .\scripts\start-services.ps1 -Frontend -Tunnel
#>
[CmdletBinding()]
param(
    [switch]$Frontend,
    [switch]$Tunnel,
    [switch]$Classifier
)

$ErrorActionPreference = 'Stop'
$RepoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $RepoRoot

function Write-Step([string]$msg) { Write-Host "==> $msg" -ForegroundColor Cyan }
function Write-Ok([string]$msg)   { Write-Host "    ok: $msg" -ForegroundColor Green }
function Write-Skip([string]$msg) { Write-Host "    skip: $msg" -ForegroundColor Gray }
function Write-Err([string]$msg)  { Write-Host "    error: $msg" -ForegroundColor Red }

$venvPython = Join-Path $RepoRoot '.venv\Scripts\python.exe'
if (-not (Test-Path $venvPython)) { Write-Err ".venv missing — run scripts\bootstrap.ps1 first"; exit 1 }
$useWt = (Get-Command wt -ErrorAction SilentlyContinue) -ne $null

function Test-PortListening([int]$port) {
    try {
        $tcp = Test-NetConnection -ComputerName 127.0.0.1 -Port $port -WarningAction SilentlyContinue
        return $tcp.TcpTestSucceeded
    } catch { return $false }
}

function Launch-InTab([string]$title, [string]$command) {
    if ($useWt) {
        wt -w 0 nt --title $title pwsh -NoExit -Command $command
        Write-Ok "$title launched in new tab"
    } else {
        $logFile = Join-Path $RepoRoot "$title.log"
        Start-Process pwsh -ArgumentList '-NoExit','-Command',$command -WindowStyle Normal
        Write-Ok "$title launched in new window (Windows Terminal not found — install for tabs)"
    }
}

# ──────────────────────────────────────────────────────────────────────────
Write-Step "Postgres"
docker compose up -d 2>&1 | Out-Null
$health = (docker inspect --format '{{.State.Health.Status}}' artilegenz-postgres 2>$null)
if ($health -eq 'healthy') { Write-Ok "postgres healthy" } else { Write-Err "postgres not healthy ($health)"; exit 1 }

# ──────────────────────────────────────────────────────────────────────────
Write-Step "Ingest API on :8000"
if (Test-PortListening 8000) {
    Write-Skip "something is already listening on :8000"
} else {
    Launch-InTab 'artilengenz-api' "Set-Location '$RepoRoot'; & '$venvPython' -m uvicorn api.app_ingest:app --host 127.0.0.1 --port 8000 --reload"
    # Wait for ready
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt 30) {
        try {
            if ((Invoke-WebRequest -Uri "http://127.0.0.1:8000/healthz" -TimeoutSec 1 -UseBasicParsing).StatusCode -eq 200) {
                Write-Ok "API healthy"; break
            }
        } catch {}
        Start-Sleep -Milliseconds 500
    }
}

# ──────────────────────────────────────────────────────────────────────────
if ($Frontend) {
    Write-Step "Frontend (Vite) on :5173"
    if (Test-PortListening 5173) {
        Write-Skip "something is already listening on :5173"
    } elseif (-not (Test-Path frontend\package.json)) {
        Write-Err "frontend\package.json missing"
    } elseif (-not (Test-Path frontend\node_modules)) {
        Write-Err "frontend\node_modules missing — run scripts\bootstrap.ps1 (without -SkipFrontend)"
    } else {
        Launch-InTab 'artilengenz-ui' "Set-Location '$RepoRoot\frontend'; npm run dev"
    }
}

# ──────────────────────────────────────────────────────────────────────────
if ($Tunnel) {
    Write-Step "ngrok tunnel for :8000"
    if (-not (Get-Command ngrok -ErrorAction SilentlyContinue)) {
        Write-Err "ngrok not on PATH. Install from https://ngrok.com/download then run ngrok config add-authtoken <your-token>"
    } else {
        Launch-InTab 'artilengenz-ngrok' "ngrok http 8000"
        Start-Sleep -Seconds 3
        try {
            $tunnels = Invoke-RestMethod -Uri "http://127.0.0.1:4040/api/tunnels" -TimeoutSec 5
            $publicUrl = $tunnels.tunnels | Where-Object { $_.proto -eq 'https' } | Select-Object -First 1 -ExpandProperty public_url
            if ($publicUrl) { Write-Ok "ngrok public URL: $publicUrl" }
        } catch { Write-Skip "ngrok dashboard not reachable yet — check the tab" }
    }
}

# ──────────────────────────────────────────────────────────────────────────
if ($Classifier) {
    Write-Step "Classifier worker"
    $envContent = if (Test-Path .env) { Get-Content .env -Raw } else { '' }
    $missing = @()
    if ($envContent -notmatch '(?m)^ANTHROPIC_API_KEY=\S+') { $missing += 'ANTHROPIC_API_KEY' }
    if ($envContent -notmatch '(?m)^VOYAGE_API_KEY=\S+')    { $missing += 'VOYAGE_API_KEY' }
    if ($envContent -notmatch '(?m)^PINECONE_API_KEY=\S+')  { $missing += 'PINECONE_API_KEY' }
    if ($missing) {
        Write-Err "missing in .env: $($missing -join ', ')"
    } else {
        Launch-InTab 'artilengenz-classifier' "Set-Location '$RepoRoot'; & '$venvPython' classifier\classify_worker.py"
    }
}

Write-Host ""
Write-Host "Services up. Tabs in Windows Terminal show logs." -ForegroundColor Green
Write-Host "Tear down with: .\scripts\stop-services.ps1" -ForegroundColor Gray
