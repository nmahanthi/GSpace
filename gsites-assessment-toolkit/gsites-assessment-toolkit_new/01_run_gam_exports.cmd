@echo off
setlocal
set SCRIPT_DIR=%~dp0
set OUTDIR=%SCRIPT_DIR%output
if not exist "%OUTDIR%" mkdir "%OUTDIR%"

REM ---------------------------------------------------------------------------
REM Resolve GAM executable - three strategies, in priority order:
REM   1. GAM_PATH environment variable (set externally or in gam.cfg)
REM   2. gam.cfg file next to this script  (key=value: GAM_PATH=<path>)
REM   3. gam.exe / gam found on the system PATH
REM
REM To configure permanently, create a file called gam.cfg in the same
REM folder as this script with one line, e.g.:
REM   GAM_PATH=C:\tools\gam\gam.exe
REM ---------------------------------------------------------------------------

REM Strategy 1: honour an already-set GAM_PATH environment variable
if defined GAM_PATH goto :verify_gam

REM Strategy 2: read gam.cfg if it exists beside this script
set GAM_CFG=%SCRIPT_DIR%gam.cfg
if exist "%GAM_CFG%" (
    for /f "usebackq tokens=1,* delims==" %%A in ("%GAM_CFG%") do (
        if /i "%%A"=="GAM_PATH" set "GAM_PATH=%%B"
    )
)
if defined GAM_PATH goto :verify_gam

REM Strategy 3: search for gam.exe (or gam) on the system PATH
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

REM Build Sites query if a name filter was provided by the orchestrator
if defined GAM_SITES_FILTER (
    set "SITES_QUERY=mimeType='application/vnd.google-apps.site' and trashed=false and (%GAM_SITES_FILTER%)"
    echo [INFO] Restricting Sites scan to selected site names
) else (
    set "SITES_QUERY=mimeType='application/vnd.google-apps.site' and trashed=false"
)

if defined GAM_TARGET_FILE (
    set GAM_USER_TARGET=csv "%GAM_TARGET_FILE%" gam user "~Email"
    echo [INFO] Restricting GAM to scan ONLY specific user drives provided in the CSV.
) else (
    set GAM_USER_TARGET=all users
    echo [INFO] Note: Because GAM must search every user's Drive to find these sites,
    echo [INFO] it will first fetch the list of all users. This may take a few minutes
    echo [INFO] before you see any progress on the screen. Please be patient!
)

echo [1/3] Minimal Google Sites sanity export...
if defined GAM_SITES_FILTER (
    "%GAM_PATH%" config auto_batch_min 1 num_threads 30 redirect csv "%OUTDIR%\GSites_Inventory_Min.csv" multiprocess %GAM_USER_TARGET% print filelist corpora alldrives query "%SITES_QUERY%" fields id,name,mimetype
) else (
    "%GAM_PATH%" config auto_batch_min 1 num_threads 30 redirect csv "%OUTDIR%\GSites_Inventory_Min.csv" multiprocess redirect stderr - multiprocess %GAM_USER_TARGET% print filelist corpora alldrives fields id,name,mimetype filepath showmimetype gsite
)
if errorlevel 1 goto :fail

echo [2/3] Detailed Google Sites inventory...
"%GAM_PATH%" config auto_batch_min 1 num_threads 30 redirect csv "%OUTDIR%\GSites_Inventory_Detailed.csv" multiprocess %GAM_USER_TARGET% print filelist corpora alldrives query "%SITES_QUERY%" fields id,name,mimetype,webviewlink,createdtime,modifiedtime,owners,shared,parents,driveid,drivename,size,quotabytesused,version,viewedbymetime,copyrequireswriterpermission,viewerscancopycontent,writerscanshare,inheritedpermissionsdisabled,starred,modifiedbyme,modifiedbymetime,viewedbyme,explicitlytrashed,spaces,thumbnaillink,thumbnailversion,hasthumbnail,exportlinks
if errorlevel 1 goto :fail

echo [3/3] Google Sites permissions and security...
"%GAM_PATH%" config auto_batch_min 1 num_threads 30 redirect csv "%OUTDIR%\GSites_Permissions.csv" multiprocess %GAM_USER_TARGET% print filelist corpora alldrives query "%SITES_QUERY%" fields id,name,webviewlink,owners,basicpermissions,shared,copyrequireswriterpermission,viewerscancopycontent,writerscanshare,inheritedpermissionsdisabled oneitemperrow

if errorlevel 1 goto :fail

echo.
echo GAM exports completed successfully.
echo Output folder: %OUTDIR%

REM Clean up environment variable so it does not affect future runs
set GAM_SITES_FILTER=
set SITES_QUERY=

echo Cleaning CSV headers...
powershell -NoProfile -Command "Get-ChildItem '%OUTDIR%\*.csv' | ForEach-Object { $lines = @(Get-Content $_.FullName); if ($lines.Count -gt 0) { $lines[0] = $lines[0] -replace '\.[0-9]+\.', '.'; $lines | Set-Content $_.FullName } }"
exit /b 0

:fail
echo.
echo GAM export failed. Review the command above and GAM error output.
exit /b 1
