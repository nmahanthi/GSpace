@echo off
setlocal enabledelayedexpansion
set SCRIPT_DIR=%~dp0
set OUTDIR=%SCRIPT_DIR%output

REM ---------------------------------------------------------------------------
REM Option B fix for Step 4A 403 errors:
REM The Sites API v1 requires the caller to have explicit Drive-level Viewer
REM access to each site - being a Workspace super admin is NOT enough. This
REM script uses GAM's elevated (domain-wide-delegated) access to grant the
REM account that will call the Sites API (%SITES_ADMIN_EMAIL%) Reader access
REM to every site in the inventory, before 03a_get_published_urls.js runs.
REM ---------------------------------------------------------------------------

REM Resolve GAM executable (same 3-strategy resolution as 01_run_gam_exports.cmd)
if defined GAM_PATH goto :verify_gam

set GAM_CFG=%SCRIPT_DIR%gam.cfg
if exist "%GAM_CFG%" (
    for /f "usebackq tokens=1,* delims==" %%A in ("%GAM_CFG%") do (
        if /i "%%A"=="GAM_PATH" set "GAM_PATH=%%B"
    )
)
if defined GAM_PATH goto :verify_gam

for %%X in (gam.exe gam) do (
    set "_found=%%~$PATH:X"
    if defined _found (
        set "GAM_PATH=%%~$PATH:X"
        goto :verify_gam
    )
)

echo ERROR: GAM executable not found.
echo   Set the GAM_PATH environment variable, add gam.exe to your PATH,
echo   or create a gam.cfg file in the script directory with:
echo     GAM_PATH=^<full path to gam.exe^>
exit /b 1

:verify_gam
if not exist "%GAM_PATH%" (
    echo ERROR: GAM not found at "%GAM_PATH%"
    echo   Check your GAM_PATH environment variable or gam.cfg configuration.
    exit /b 1
)
echo [GAM] Using GAM at: %GAM_PATH%

if not defined SITES_ADMIN_EMAIL (
    echo ERROR: SITES_ADMIN_EMAIL environment variable not set.
    echo   This must be the email of the account that will call the Sites API
    echo   ^(the account you run "gcloud auth login" with^).
    exit /b 1
)

set INVENTORY=%OUTDIR%\GSites_Inventory_Detailed.csv
if not exist "%INVENTORY%" (
    echo ERROR: Inventory file not found: %INVENTORY%
    echo   Run GAM export first ^(Step 1^).
    exit /b 1
)

echo [Grant Access] Granting %SITES_ADMIN_EMAIL% Reader access to all sites in the inventory...
echo [Grant Access] This is required because the Sites API v1 requires the caller
echo [Grant Access] to have explicit Drive-level access to each site, even for domain admins.

REM GAM dynamically omits the driveId column from the CSV entirely when none of the
REM exported sites live in a Shared Drive. Detect whether the column exists before
REM using it in a matchfield filter, otherwise GAM errors with "field not found".
set HEADER_LINE=
for /f "usebackq delims=" %%L in ("%INVENTORY%") do (
    if not defined HEADER_LINE set "HEADER_LINE=%%L"
)
set HAS_DRIVEID=0
echo !HEADER_LINE! | findstr /I "driveId" >nul 2>&1
if not errorlevel 1 set HAS_DRIVEID=1

echo [1/2] Granting access to My Drive sites (impersonating each site's owner)...
if "!HAS_DRIVEID!"=="1" (
    "%GAM_PATH%" config num_threads 30 csv "%INVENTORY%" matchfield driveId "^$" gam user "~Owner" add drivefileacl "~id" user "%SITES_ADMIN_EMAIL%" role reader
) else (
    "%GAM_PATH%" config num_threads 30 csv "%INVENTORY%" gam user "~Owner" add drivefileacl "~id" user "%SITES_ADMIN_EMAIL%" role reader
)

echo [2/2] Granting access to Shared Drive sites (using domain admin access)...
if "!HAS_DRIVEID!"=="1" (
    "%GAM_PATH%" config num_threads 30 csv "%INVENTORY%" matchfield driveId ".+" gam add drivefileacl "~id" user "%SITES_ADMIN_EMAIL%" role reader adminaccess
) else (
    echo [INFO] No driveId column found in inventory - no Shared Drive sites to process, skipping.
)

echo.
echo Access grant completed. %SITES_ADMIN_EMAIL% should now have Reader access to all sites.
echo Note: Some individual grants may show errors above (e.g. if the account already
echo had access) - these are not fatal, Step 4A will still proceed.
exit /b 0
