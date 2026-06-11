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

echo [1/6] Minimal Google Sites sanity export...
"%GAM_PATH%" config auto_batch_min 1 num_threads 10 redirect csv "%OUTDIR%\01_GSites_Inventory_Min.csv" multiprocess redirect stderr - multiprocess all users print filelist fields id,name,mimetype filepath showmimetype gsite
if errorlevel 1 goto :fail

echo [2/6] Detailed Google Sites inventory...
"%GAM_PATH%" config auto_batch_min 1 num_threads 10 redirect csv "%OUTDIR%\02_GSites_Inventory_Detailed.csv" multiprocess all users print filelist query "mimeType='application/vnd.google-apps.site' and trashed=false" fields id,name,mimetype,webviewlink,createdtime,modifiedtime,owners,shared,parents,driveid,size,quotabytesused,version,viewedbymetime,copyrequireswriterpermission,viewerscancopycontent,writerscanshare,inheritedpermissionsdisabled,starred,modifiedbyme,modifiedbymetime,viewedbyme,explicitlytrashed,spaces,thumbnaillink,thumbnailversion,hasthumbnail,exportlinks
if errorlevel 1 goto :fail

echo [3/6] Google Sites permissions and security...
"%GAM_PATH%" config auto_batch_min 1 num_threads 10 redirect csv "%OUTDIR%\03_GSites_Permissions.csv" multiprocess all users print filelist query "mimeType='application/vnd.google-apps.site' and trashed=false" fields id,name,webviewlink,owners,basicpermissions,shared,copyrequireswriterpermission,viewerscancopycontent,writerscanshare,inheritedpermissionsdisabled oneitemperrow
if errorlevel 1 goto :fail

echo [4/6] Broad candidate Google Sheets inventory...
"%GAM_PATH%" config auto_batch_min 1 num_threads 10 redirect csv "%OUTDIR%\04_Candidate_Sheets.csv" multiprocess all users print filelist query "mimeType='application/vnd.google-apps.spreadsheet' and trashed=false" fields id,name,mimetype,webviewlink,createdtime,modifiedtime,owners,shared,parents,driveid
if errorlevel 1 goto :fail

echo [5/6] Broad candidate Google Forms inventory...
"%GAM_PATH%" config auto_batch_min 1 num_threads 10 redirect csv "%OUTDIR%\05_Candidate_Forms.csv" multiprocess all users print filelist query "mimeType='application/vnd.google-apps.form' and trashed=false" fields id,name,mimetype,webviewlink,createdtime,modifiedtime,owners,shared,parents,driveid
if errorlevel 1 goto :fail

echo [6/6] Broad candidate Apps Script inventory...
"%GAM_PATH%" config auto_batch_min 1 num_threads 10 redirect csv "%OUTDIR%\06_Candidate_Scripts.csv" multiprocess all users print filelist query "mimeType='application/vnd.google-apps.script' and trashed=false" fields id,name,mimetype,webviewlink,createdtime,modifiedtime,owners,shared,parents,driveid
if errorlevel 1 goto :fail

echo.
echo GAM exports completed successfully.
echo Output folder: %OUTDIR%
exit /b 0

:fail
echo.
echo GAM export failed. Review the command above and GAM error output.
exit /b 1
