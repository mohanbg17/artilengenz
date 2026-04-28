# ============================================================
#  NSSM Service Installation -- Artilegenz Embedding Worker
# ============================================================
#  Run this PowerShell script ONCE as Administrator to install
#  the worker as a Windows service.
#
#  Prerequisites:
#    1. Download NSSM from https://nssm.cc/download
#       Extract to C:\nssm\
#       (or set NSSM_PATH below to wherever you put nssm.exe)
#
#    2. Voyage + Pinecone API keys ready
#    3. embedding-worker/ files copied to C:\sap-error-ai\sap-error-intelligence\embeddings\
#    4. SQL migration applied (01_migration_raw_embeddings.sql)
#    5. Postgres + ngrok + uvicorn already running (we check below)
#
#  Stop / uninstall later:
#    nssm stop ArtilegenzEmbeddingWorker
#    nssm remove ArtilegenzEmbeddingWorker confirm
# ============================================================

#Requires -RunAsAdministrator

# ---- Settings (edit if needed) ----
$ServiceName     = "ArtilegenzEmbeddingWorker"
$ServiceDisplay  = "Artilegenz Embedding Worker"
$ServiceDesc     = "Polls raw.raw_errors and pushes embeddings to Pinecone via Voyage AI"
$NSSM_PATH       = "C:\nssm\nssm.exe"           # adjust if NSSM lives elsewhere
$LauncherBat     = "C:\sap-error-ai\sap-error-intelligence\embeddings\run_worker.bat"
$LogDir          = "C:\sap-error-ai\logs"

# ---- API keys -- PROMPT (don't paste in script files committed to git!) ----
Write-Host ""
Write-Host "=== Service install: $ServiceDisplay ===" -ForegroundColor Cyan
Write-Host ""
$VoyageKey   = Read-Host "Voyage API key (starts with pa-)"
$PineconeKey = Read-Host "Pinecone API key (starts with pcsk_)"

if ([string]::IsNullOrWhiteSpace($VoyageKey) -or [string]::IsNullOrWhiteSpace($PineconeKey)) {
    Write-Error "Both keys required. Aborting."
    exit 1
}

# ---- Sanity: NSSM exists ----
if (-not (Test-Path $NSSM_PATH)) {
    Write-Error "NSSM not found at $NSSM_PATH. Download from https://nssm.cc/download"
    exit 1
}

# ---- Sanity: launcher .bat exists ----
if (-not (Test-Path $LauncherBat)) {
    Write-Error "Launcher batch not found at $LauncherBat"
    exit 1
}

# ---- Logs dir ----
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
}

# ---- Remove existing service if present ----
$existing = & sc.exe query $ServiceName 2>&1
if ($LASTEXITCODE -eq 0) {
    Write-Host "Removing existing service first..." -ForegroundColor Yellow
    & $NSSM_PATH stop $ServiceName 2>&1 | Out-Null
    & $NSSM_PATH remove $ServiceName confirm 2>&1 | Out-Null
    Start-Sleep -Seconds 2
}

# ---- Install service ----
Write-Host "Installing service..." -ForegroundColor Cyan
& $NSSM_PATH install $ServiceName $LauncherBat
& $NSSM_PATH set $ServiceName DisplayName $ServiceDisplay
& $NSSM_PATH set $ServiceName Description $ServiceDesc
& $NSSM_PATH set $ServiceName Start SERVICE_AUTO_START
& $NSSM_PATH set $ServiceName AppStdout "$LogDir\embedding_worker_stdout.log"
& $NSSM_PATH set $ServiceName AppStderr "$LogDir\embedding_worker_stderr.log"
& $NSSM_PATH set $ServiceName AppRotateFiles 1
& $NSSM_PATH set $ServiceName AppRotateOnline 1
& $NSSM_PATH set $ServiceName AppRotateBytes 10485760  # rotate at 10MB

# ---- Environment variables ----
Write-Host "Setting environment variables..." -ForegroundColor Cyan
$envBlock = @(
    "VOYAGE_API_KEY=$VoyageKey"
    "VOYAGE_MODEL=voyage-3-large"
    "VOYAGE_DIM=1024"
    "PINECONE_API_KEY=$PineconeKey"
    "PINECONE_INDEX=artilegenz-sap-errors"
    "PINECONE_NAMESPACE=raw_errors_v1"
    "PINECONE_CLOUD=aws"
    "PINECONE_REGION=us-east-1"
    "PG_HOST=localhost"
    "PG_PORT=5432"
    "PG_USER=artilegenz"
    "PG_PASSWORD=artilegenz_local_dev"
    "PG_DATABASE=sap_errors"
    "POLL_INTERVAL_SEC=30"
    "BATCH_SIZE=32"
    "MAX_TEXT_CHARS=100000"
    "MAX_RETRIES=3"
    "LOG_LEVEL=INFO"
) -join "`r`n"

& $NSSM_PATH set $ServiceName AppEnvironmentExtra $envBlock

# ---- Stop / restart behavior ----
& $NSSM_PATH set $ServiceName AppStopMethodSkip 0
& $NSSM_PATH set $ServiceName AppStopMethodConsole 30000   # 30s graceful shutdown
& $NSSM_PATH set $ServiceName AppRestartDelay 5000          # 5s restart delay
& $NSSM_PATH set $ServiceName AppExit Default Restart

# ---- Start ----
Write-Host "Starting service..." -ForegroundColor Cyan
& $NSSM_PATH start $ServiceName

Start-Sleep -Seconds 3
$status = & $NSSM_PATH status $ServiceName
Write-Host ""
Write-Host "Service status: $status" -ForegroundColor Green
Write-Host ""
Write-Host "=== Useful commands ===" -ForegroundColor Cyan
Write-Host "  Status:   nssm status $ServiceName"
Write-Host "  Stop:     nssm stop $ServiceName"
Write-Host "  Start:    nssm start $ServiceName"
Write-Host "  Restart:  nssm restart $ServiceName"
Write-Host "  Logs:     Get-Content '$LogDir\embedding_worker_stdout.log' -Wait -Tail 50"
Write-Host "  Edit:     nssm edit $ServiceName"
Write-Host "  Remove:   nssm remove $ServiceName confirm"
Write-Host ""
Write-Host "  Watch logs in real time:" -ForegroundColor Yellow
Write-Host "    Get-Content '$LogDir\embedding_worker_stdout.log' -Wait -Tail 50"
Write-Host ""
