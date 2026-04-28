<#
.SYNOPSIS
  Graceful shutdown of all platform services.

.DESCRIPTION
  Stops each service that bootstrap / start-services started:
    - kills python processes serving uvicorn on :8000
    - kills node processes serving Vite on :5173
    - kills ngrok if running
    - kills classifier worker if running
    - optionally stops Postgres (with -StopDb)

.PARAMETER StopDb
  Also bring down Postgres with 'docker compose down'. Default: leave running.

.PARAMETER RemoveDb
  Bring down Postgres AND delete the volume (DESTRUCTIVE — wipes data).
#>
[CmdletBinding()]
param(
    [switch]$StopDb,
    [switch]$RemoveDb
)

$ErrorActionPreference = 'Continue'
$RepoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $RepoRoot

function Write-Step([string]$msg) { Write-Host "==> $msg" -ForegroundColor Cyan }
function Write-Ok([string]$msg)   { Write-Host "    ok: $msg" -ForegroundColor Green }
function Write-Skip([string]$msg) { Write-Host "    skip: $msg" -ForegroundColor Gray }

function Kill-PortListener([int]$port, [string]$label) {
    $pids = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty OwningProcess -Unique
    if (-not $pids) { Write-Skip "$label not running on :$port"; return }
    foreach ($pid in $pids) {
        try {
            $proc = Get-Process -Id $pid -ErrorAction Stop
            $proc | Stop-Process -Force
            Write-Ok "$label stopped (pid $pid, name $($proc.ProcessName))"
        } catch { Write-Skip "could not stop pid $pid" }
    }
}

Write-Step "Stopping ingest API on :8000"
Kill-PortListener 8000 'uvicorn'

Write-Step "Stopping frontend on :5173"
Kill-PortListener 5173 'vite'

Write-Step "Stopping ngrok"
$ng = Get-Process -Name ngrok -ErrorAction SilentlyContinue
if ($ng) { $ng | Stop-Process -Force; Write-Ok "ngrok stopped" } else { Write-Skip "ngrok not running" }

Write-Step "Stopping classifier worker (any python running classify_worker.py)"
$workers = Get-CimInstance Win32_Process -Filter "Name='python.exe'" -ErrorAction SilentlyContinue |
           Where-Object { $_.CommandLine -match 'classify_worker' }
if ($workers) {
    foreach ($w in $workers) { Stop-Process -Id $w.ProcessId -Force; Write-Ok "classifier pid $($w.ProcessId) stopped" }
} else { Write-Skip "no classifier worker running" }

if ($RemoveDb) {
    Write-Step "Bringing down Postgres AND deleting volume (destructive)"
    docker compose down -v 2>&1 | Out-Null
    Write-Ok "postgres and volume removed"
} elseif ($StopDb) {
    Write-Step "Bringing down Postgres (data preserved on volume)"
    docker compose down 2>&1 | Out-Null
    Write-Ok "postgres container stopped"
} else {
    Write-Step "Postgres"
    Write-Skip "leaving Postgres running. Use -StopDb to stop, -RemoveDb to wipe data."
}

Write-Host ""
Write-Host "Shutdown complete." -ForegroundColor Green
