@echo off
setlocal enabledelayedexpansion
set SCRIPT_DIR=%~dp0
set OUTDIR=%SCRIPT_DIR%output
if not exist "%OUTDIR%" mkdir "%OUTDIR%"

REM ===========================================================================
REM Standalone export: Google Sites hosted on SPECIFIC Shared Drives.
REM
REM This is a separate utility from 01_run_gam_exports.cmd (which now scans
REM My Drive only, per customer requirement). Use this script only when you
REM need to inventory Sites living on a known, hand-picked list of Shared
REM Drives.
REM
REM Usage:
REM   01d_run_gam_exports_shareddrive.cmd <SharedDriveIDs.csv> <ScanningUserEmail>
REM
REM   <SharedDriveIDs.csv>    Required. CSV file with a header row containing
REM                           a "driveId" column, one Shared Drive ID per row.
REM                           Example:
REM                             driveId
REM                             0AbCDeFGhIJKLmnUK9PVA
REM                             0XyzTeamDriveIdHere123
REM
REM   <ScanningUserEmail>     Required. The Workspace user GAM will impersonate
REM                           to query these Shared Drives. Must either be a
REM                           member of every Shared Drive listed, or a Super
REM                           Admin (Super Admins can see all Shared Drives in
REM                           the domain regardless of membership). Can also be
REM                           set via the GAM_ADMIN_USER environment variable
REM                           instead of passing it as an argument.
REM
REM Outputs (written to output/):
REM   GSites_SharedDrive_Inventory.csv    One row per Site found in the drives
REM   GSites_SharedDrive_Permissions.csv  One row per (Site, grantee) permission,
REM                                       flagged with whether access is direct
REM                                       or inherited from Shared Drive membership
REM ===========================================================================

if "%~1"=="" (
    echo ERROR: Missing required argument: path to Shared Drive IDs CSV.
    echo.
    echo Usage: %~nx0 ^<SharedDriveIDs.csv^> ^<ScanningUserEmail^>
    exit /b 1
)
set SHAREDDRIVE_CSV=%~1
if not exist "%SHAREDDRIVE_CSV%" (
    echo ERROR: Shared Drive IDs CSV not found: %SHAREDDRIVE_CSV%
    exit /b 1
)

if not "%~2"=="" set GAM_ADMIN_USER=%~2
if not defined GAM_ADMIN_USER (
    echo ERROR: No scanning user specified.
    echo   Pass it as the 2nd argument, or set the GAM_ADMIN_USER environment
    echo   variable to a Workspace user/admin email that has access to the
    echo   Shared Drives listed in %SHAREDDRIVE_CSV%.
    exit /b 1
)

REM --- GAM path resolution: GAM_PATH env var -^> gam.cfg -^> system PATH ---
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
echo [INFO] Scanning user (impersonated for Shared Drive access): %GAM_ADMIN_USER%
echo [INFO] Shared Drive IDs source: %SHAREDDRIVE_CSV%

REM Number of parallel GAM worker processes, one per Shared Drive ID row.
if not defined GAM_NUM_THREADS set GAM_NUM_THREADS=5
echo [INFO] Using num_threads=%GAM_NUM_THREADS% for GAM multiprocess operations

set "SITES_QUERY=mimeType='application/vnd.google-apps.site' and trashed=false"

echo [1/2] Google Sites inventory across selected Shared Drives...
"%GAM_PATH%" config auto_batch_min 1 num_threads %GAM_NUM_THREADS% redirect csv "%OUTDIR%\GSites_SharedDrive_Inventory.csv" multiprocess csv "%SHAREDDRIVE_CSV%" gam user %GAM_ADMIN_USER% print filelist select teamdriveid "~driveId" query "%SITES_QUERY%" fields id,name,mimetype,description,webviewlink,createdtime,modifiedtime,owners,lastmodifyinguser,shared,driveid,size,hasaugmentedpermissions,capabilities.canshare,capabilities.canedit,capabilities.candelete,capabilities.candownload,capabilities.cancopy,capabilities.canremovechildren,spaces,thumbnaillink showdrivename
if errorlevel 1 goto :fail

echo [2/2] Google Sites permissions across selected Shared Drives (direct vs inherited)...
"%GAM_PATH%" config auto_batch_min 1 num_threads %GAM_NUM_THREADS% redirect csv "%OUTDIR%\GSites_SharedDrive_Permissions.csv" multiprocess csv "%SHAREDDRIVE_CSV%" gam user %GAM_ADMIN_USER% print filelist select teamdriveid "~driveId" query "%SITES_QUERY%" fields id,name,webviewlink,owners,lastmodifyinguser,basicpermissions,permissiondetails,shared,driveid,copyrequireswriterpermission,viewerscancopycontent,writerscanshare,inheritedpermissionsdisabled oneitemperrow showshareddrivepermissions
if errorlevel 1 goto :fail

echo.
echo Shared Drive GSites export completed successfully.
echo Output folder: %OUTDIR%

set GAM_NUM_THREADS=
set SITES_QUERY=

echo Cleaning CSV headers...
powershell -NoProfile -Command "Get-ChildItem '%OUTDIR%\GSites_SharedDrive_*.csv' | ForEach-Object { $lines = @(Get-Content $_.FullName); if ($lines.Count -gt 0) { $lines[0] = $lines[0] -replace '\.[0-9]+\.', '.'; $lines | Set-Content $_.FullName } }"
exit /b 0

:fail
echo.
echo Shared Drive GSites export failed. Review the command above and GAM error output.
exit /b 1
