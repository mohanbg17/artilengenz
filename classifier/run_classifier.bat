@echo off
REM ============================================================
REM   Artilegenz Classifier Worker Launcher (NSSM)
REM ============================================================

setlocal

set "PROJECT_DIR=C:\sap-error-ai\sap-error-intelligence"
set "VENV_PYTHON=%PROJECT_DIR%\.venv\Scripts\python.exe"
set "WORKER_DIR=%PROJECT_DIR%\classifier"
set "WORKER_SCRIPT=%WORKER_DIR%\classify_worker.py"

set "LOG_DIR=C:\sap-error-ai\logs"
if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"

if not exist "%VENV_PYTHON%" (
    echo [ERROR] Python venv not found at %VENV_PYTHON%
    exit /b 1
)
if not exist "%WORKER_SCRIPT%" (
    echo [ERROR] Worker script not found at %WORKER_SCRIPT%
    exit /b 1
)

echo [%date% %time%] Starting Artilegenz classifier worker...
cd /d "%WORKER_DIR%"
"%VENV_PYTHON%" "%WORKER_SCRIPT%"
exit /b %errorlevel%
