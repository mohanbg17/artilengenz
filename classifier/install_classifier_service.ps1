# ============================================================
#  NSSM Service Installation -- Artilegenz Classifier Worker
# ============================================================
#  Run as Administrator. Prompts for API keys, registers the
#  classifier as a Windows service.
# ============================================================

#Requires -RunAsAdministrator

$ServiceName     = "ArtilegenzClassifierWorker"
$ServiceDisplay  = "Artilegenz Classifier Worker"
$ServiceDesc     = "Classifies SAP errors using Sonnet+Opus with retrieval-augmented generation"
$NSSM_PATH       = "C:\nssm\nssm.exe"
$LauncherBat     = "C:\sap-error-ai\sap-error-intelligence\classifier\run_classifier.bat"
$LogDir          = "C:\sap-error-ai\logs"

Write-Host ""
Write-Host "=== Service install: $ServiceDisplay ===" -ForegroundColor Cyan
Write-Host ""
$AnthropicKey = Read-Host "Anthropic API key (starts with sk-ant-)"
$VoyageKey    = Read-Host "Voyage API key (starts with pa-)"
$PineconeKey  = Read-Host "Pinecone API key (starts with pcsk_)"

if ([string]::IsNullOrWhiteSpace($AnthropicKey) -or
    [string]::IsNullOrWhiteSpace($VoyageKey) -or
    [string]::IsNullOrWhiteSpace($PineconeKey)) {
    Write-Error "All three keys required. Aborting."
    exit 1
}

if (-not (Test-Path $NSSM_PATH)) {
    Write-Error "NSSM not found at $NSSM_PATH. Download from https://nssm.cc/download"
    exit 1
}
if (-not (Test-Path $LauncherBat)) {
    Write-Error "Launcher batch not found at $LauncherBat"
    exit 1
}
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
}

# Remove existing
$existing = & sc.exe query $ServiceName 2>&1
if ($LASTEXITCODE -eq 0) {
    Write-Host "Removing existing service..." -ForegroundColor Yellow
    & $NSSM_PATH stop $ServiceName 2>&1 | Out-Null
    & $NSSM_PATH remove $ServiceName confirm 2>&1 | Out-Null
    Start-Sleep -Seconds 2
}

Write-Host "Installing service..." -ForegroundColor Cyan
& $NSSM_PATH install $ServiceName $LauncherBat
& $NSSM_PATH set $ServiceName DisplayName $ServiceDisplay
& $NSSM_PATH set $ServiceName Description $ServiceDesc
& $NSSM_PATH set $ServiceName Start SERVICE_AUTO_START
& $NSSM_PATH set $ServiceName AppStdout "$LogDir\classifier_stdout.log"
& $NSSM_PATH set $ServiceName AppStderr "$LogDir\classifier_stderr.log"
& $NSSM_PATH set $ServiceName AppRotateFiles 1
& $NSSM_PATH set $ServiceName AppRotateOnline 1
& $NSSM_PATH set $ServiceName AppRotateBytes 10485760

$envBlock = @(
    "ANTHROPIC_API_KEY=$AnthropicKey"
    "SONNET_MODEL=claude-sonnet-4-6"
    "OPUS_MODEL=claude-opus-4-7"
    "SONNET_MAX_TOKENS=2000"
    "OPUS_MAX_TOKENS=2500"
    "VOYAGE_API_KEY=$VoyageKey"
    "VOYAGE_MODEL=voyage-3-large"
    "VOYAGE_DIM=1024"
    "PINECONE_API_KEY=$PineconeKey"
    "PINECONE_INDEX=artilegenz-sap-errors"
    "PINECONE_NAMESPACE=raw_errors_v1"
    "PINECONE_CORPUS_NAMESPACE=corpus_v1"
    "PINECONE_CLOUD=aws"
    "PINECONE_REGION=us-east-1"
    "PG_HOST=localhost"
    "PG_PORT=5432"
    "PG_USER=artilegenz"
    "PG_PASSWORD=artilegenz_local_dev"
    "PG_DATABASE=sap_errors"
    "CLASSIFY_POLL_INTERVAL_SEC=60"
    "CLASSIFY_BATCH_SIZE=5"
    "CLASSIFY_INTER_CALL_DELAY=1.0"
    "RETRIEVAL_TOTAL_K=10"
    "LOG_LEVEL=INFO"
) -join "`r`n"

& $NSSM_PATH set $ServiceName AppEnvironmentExtra $envBlock

& $NSSM_PATH set $ServiceName AppStopMethodSkip 0
& $NSSM_PATH set $ServiceName AppStopMethodConsole 30000
& $NSSM_PATH set $ServiceName AppRestartDelay 5000
& $NSSM_PATH set $ServiceName AppExit Default Restart

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
Write-Host "  Logs:     Get-Content '$LogDir\classifier_stdout.log' -Wait -Tail 50"
Write-Host "  Edit:     nssm edit $ServiceName"
Write-Host "  Remove:   nssm remove $ServiceName confirm"
Write-Host ""
