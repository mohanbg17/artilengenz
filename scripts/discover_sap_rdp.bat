@echo off
:: ============================================================================
:: SAP S/4HANA Connection Discovery Script
:: Run this INSIDE the Remote Desktop session (no admin rights needed).
:: It prints everything needed to configure the NLP-SAP engine.
:: ============================================================================

setlocal enabledelayedexpansion
echo.
echo ============================================================
echo  SAP S/4HANA Connection Discovery
echo  Run this inside the Remote Desktop session
echo ============================================================
echo.

:: ── 1. Machine hostname and IP ──────────────────────────────────────────────
echo [1] Machine Info
echo     Hostname : %COMPUTERNAME%
for /f "tokens=2 delims=:" %%i in ('ipconfig ^| findstr /i "IPv4"') do (
    set ip=%%i
    set ip=!ip: =!
    echo     IP       : !ip!
)
echo.

:: ── 2. SAP system info from registry ────────────────────────────────────────
echo [2] SAP Systems (from registry / SAP GUI config)
set "SAPKEY=HKCU\SOFTWARE\SAP\SAPGUI Front\SAP Frontend Server\Connections"
reg query "%SAPKEY%" 2>nul && (
    for /f "tokens=*" %%k in ('reg query "%SAPKEY%" 2^>nul') do (
        echo     SAP Entry: %%k
        reg query "%%k" /v Description 2>nul | findstr /i "Description"
        reg query "%%k" /v Server     2>nul | findstr /i "Server"
        reg query "%%k" /v SystemID   2>nul | findstr /i "SystemID"
        reg query "%%k" /v SNCName    2>nul | findstr /i "SNCName"
        reg query "%%k" /v Client     2>nul | findstr /i "Client"
        echo.
    )
) || echo     (SAP GUI registry entries not found - try SAPUILandscape.xml)
echo.

:: ── 3. SAPUILandscape.xml (another location for SAP system config) ──────────
echo [3] SAPUILandscape.xml entries
set "LANDSCAPE=%APPDATA%\SAP\Common\SAPUILandscape.xml"
if exist "%LANDSCAPE%" (
    echo     Found: %LANDSCAPE%
    findstr /i "server systemid description client" "%LANDSCAPE%"
) else (
    echo     Not found at: %LANDSCAPE%
    echo     Searching...
    for /r "%USERPROFILE%" %%f in (SAPUILandscape.xml) do echo     Found: %%f
)
echo.

:: ── 4. Common SAP OData ports reachable from this machine ───────────────────
echo [4] SAP ICM Port Probe (common ports)
echo     Testing TCP connectivity to SAP OData/HTTP ports...
echo.

:: Get SAP server hostnames from landscape or known hosts
set "TEST_HOSTS="
if exist "%LANDSCAPE%" (
    for /f "tokens=2 delims== " %%h in ('findstr /i "server=" "%LANDSCAPE%"') do (
        set host=%%h
        set host=!host:server=!
        set host=!host:"=!
        if not "!host!"=="" (
            set "TEST_HOSTS=!TEST_HOSTS! !host!"
        )
    )
)

:: Also test localhost and the machine's own IP
set "TEST_HOSTS=%COMPUTERNAME% localhost %TEST_HOSTS%"

for %%H in (%TEST_HOSTS%) do (
    echo     Host: %%H
    for %%P in (443 8443 44300 44301 8000 8001 8080 50000 50001) do (
        powershell -command "
            try {
                $t = New-Object Net.Sockets.TcpClient
                $t.Connect('%%H', %%P)
                if ($t.Connected) { Write-Host '      PORT %%P : OPEN  [SAP ICM candidate]' }
                $t.Close()
            } catch { Write-Host '      PORT %%P : closed' }
        "
    )
    echo.
)

:: ── 5. SAP Message Server / hosts file ──────────────────────────────────────
echo [5] SAP entries in hosts file
findstr /i "sap\|hana\|s4\|erp\|abap" %WINDIR%\System32\drivers\etc\hosts 2>nul || echo     (none found)
echo.

:: ── 6. Running SAP processes ────────────────────────────────────────────────
echo [6] Running SAP-related processes
tasklist 2>nul | findstr /i "sap\|disp\|gwrd\|msg_server\|icman\|sapstartsrv" || echo     (none found - SAP may be remote)
echo.

:: ── 7. Browser test URL for OData ───────────────────────────────────────────
echo [7] OData Test URLs to try in a browser inside this RDP session:
echo     (Replace SAPSID with your 3-char system ID, CLIENT with 3-digit client)
echo.
echo     http://^<SAP-HOST^>:8000/sap/opu/odata/IWFND/CATALOGSERVICE?$format=json
echo     https://^<SAP-HOST^>:44300/sap/opu/odata/IWFND/CATALOGSERVICE?$format=json
echo     http://localhost:8000/sap/opu/odata/IWFND/CATALOGSERVICE?$format=json
echo.
echo     If ANY of these returns JSON (even a login prompt), OData is available.
echo.

:: ── 8. Summary of what to put in .env ───────────────────────────────────────
echo [8] What to fill in your .env file:
echo.
echo     SAP_HOST=^<hostname or IP from above^>
echo     SAP_HTTP_PORT=^<open port from above - typically 44300 or 8000^>
echo     SAP_HTTPS=true   ^(if port 44300 / 443 / 8443^)
echo     SAP_HTTPS=false  ^(if port 8000 / 8001 / 8080^)
echo     SAP_CLIENT=^<3-digit client from SAP logon screen^>
echo     SAP_USERNAME=^<your SAP username^>
echo     SAP_PASSWORD=^<your SAP password^>
echo     SAP_SYSTEM_TYPE=S4HANA
echo     MOCK_SAP=false
echo.

echo ============================================================
echo  Discovery complete. Share the output above.
echo ============================================================
pause
