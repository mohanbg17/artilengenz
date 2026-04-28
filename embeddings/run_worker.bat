@echo off
REM ============================================================
REM  Artilegenz Embedding Worker Launcher
REM  Used by NSSM service. Activates venv, sets env, runs worker.
REM ============================================================

REM ---- Project paths
set "PROJECT_DIR=C:\sap-error-ai\sap-error-intelligence"
set "VENV_PYTHON=%PROJECT_DIR%\.venv\Scripts\python.exe"
set "WORKER_DIR=%PROJECT_DIR%\embeddings"
set "WORKER_SCRIPT=%WORKER_DIR%\raw_worker.py"

REM ---- Logs
set "LOG_DIR=C:\sap-error-ai\logs"
if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"

REM ---- Sanity
if not exist "%VENV_PYTHON%" (
    echo [ERROR] Python venv not found at %VENV_PYTHON%
    exit /b 1
)

if not exist "%WORKER_SCRIPT%" (
    echo [ERROR] Worker script not found at %WORKER_SCRIPT%
    exit /b 1
)

REM ---- Environment vars are inherited from NSSM service config
REM       (set via: nssm set ArtilegenzEmbeddingWorker AppEnvironmentExtra ...)
REM       For manual runs from PowerShell, set these before calling this .bat:
REM         $env:VOYAGE_API_KEY = "pa-..."
REM         $env:PINECONE_API_KEY = "pcsk_..."
REM         $env:PG_HOST = "localhost"
REM         $env:PG_DATABASE = "sap_errors"
REM         etc.

echo [%date% %time%] Starting Artilegenz embedding worker...
cd /d "%WORKER_DIR%"
"%VENV_PYTHON%" "%WORKER_SCRIPT%"
exit /b %errorlevel%
