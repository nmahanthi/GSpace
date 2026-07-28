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
REM   01d_run_gam_exports_shareddrive.cmd <SharedDriveIDs.csv> [DefaultScanningUserEmail]
REM
REM   <SharedDriveIDs.csv>    Required. CSV file with a header row containing
REM                           "driveId" AND "account" columns - one row per
REM                           Shared Drive, each paired with the Workspace
REM                           user account that has access to THAT drive
REM                           (its owner/organizer/member, or a Super Admin).
REM                           This supports drives governed by different
REM                           owners, so no single admin needs access to
REM                           every drive. Example:
REM                             driveId,account
REM                             0AbCDeFGhIJKLmnUK9PVA,owner1@rocheua.com
REM                             0XyzTeamDriveIdHere123,owner2@rocheua.com
REM
REM   [DefaultScanningUserEmail]  Optional. Fallback account used for any row
REM                           whose "account" cell is left blank. Can also be
REM                           set via the GAM_ADMIN_USER environment variable.
REM                           If a row has no account and no default is given,
REM                           the script fails fast with a clear error.
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
    echo Usage: %~nx0 ^<SharedDriveIDs.csv^> [DefaultScanningUserEmail]
    exit /b 1
)
set SHAREDDRIVE_CSV=%~1
if not exist "%SHAREDDRIVE_CSV%" (
    echo ERROR: Shared Drive IDs CSV not found: %SHAREDDRIVE_CSV%
    exit /b 1
)

set "DEFAULT_ACCOUNT=%~2"
if not defined DEFAULT_ACCOUNT if defined GAM_ADMIN_USER set "DEFAULT_ACCOUNT=%GAM_ADMIN_USER%"

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
echo [INFO] Shared Drive IDs source: %SHAREDDRIVE_CSV% (per-row "account" column selects the scanning user for that drive)

REM Number of parallel GAM worker processes, one per Shared Drive ID row.
if not defined GAM_NUM_THREADS set GAM_NUM_THREADS=5
echo [INFO] Using num_threads=%GAM_NUM_THREADS% for GAM multiprocess operations

REM Validate the input CSV has "driveId" and "account" columns, fill any
REM blank "account" cells with DEFAULT_ACCOUNT (if provided), and write the
REM result to a temp CSV used to drive GAM's multiprocess substitution below.
REM This is what lets each Shared Drive be scanned by its own owner/
REM organizer instead of requiring one Super Admin with access to all drives.
set "EFFECTIVE_CSV=%OUTDIR%\_SharedDriveIDs_effective.csv"
powershell -NoProfile -Command ^
  "$rows = @(Import-Csv '%SHAREDDRIVE_CSV%'); if ($rows.Count -eq 0) { Write-Host '[ERROR] CSV contains no data rows: %SHAREDDRIVE_CSV%'; exit 1 }; if (-not ($rows[0].PSObject.Properties.Name -contains 'driveId')) { Write-Host '[ERROR] CSV is missing required column: driveId'; exit 1 }; if (-not ($rows[0].PSObject.Properties.Name -contains 'account')) { Write-Host '[ERROR] CSV is missing required column: account. Each row must specify driveId,account - the Workspace user with access to that specific Shared Drive.'; exit 1 }; $default = '%DEFAULT_ACCOUNT%'; $missing = @($rows | Where-Object { [string]::IsNullOrWhiteSpace($_.account) }); if ($missing.Count -gt 0 -and [string]::IsNullOrWhiteSpace($default)) { Write-Host ('[ERROR] Missing account value for driveId: ' + (($missing | ForEach-Object { $_.driveId }) -join ', ') + '. Provide an account per row, or pass a DefaultScanningUserEmail as the 2nd argument.'); exit 1 }; foreach ($r in $rows) { if ([string]::IsNullOrWhiteSpace($r.account)) { $r.account = $default } }; $rows | Export-Csv -NoTypeInformation -Path '%EFFECTIVE_CSV%'"
if errorlevel 1 goto :fail

set "SITES_QUERY=mimeType='application/vnd.google-apps.site' and trashed=false"

echo [1/2] Google Sites inventory across selected Shared Drives...
"%GAM_PATH%" config auto_batch_min 1 num_threads %GAM_NUM_THREADS% redirect csv "%OUTDIR%\GSites_SharedDrive_Inventory.csv" multiprocess csv "%EFFECTIVE_CSV%" gam user "~account" print filelist select teamdriveid "~driveId" query "%SITES_QUERY%" fields id,name,mimetype,description,webviewlink,createdtime,modifiedtime,owners,lastmodifyinguser,shared,driveid,size,hasaugmentedpermissions,capabilities.canshare,capabilities.canedit,capabilities.candelete,capabilities.candownload,capabilities.cancopy,capabilities.canremovechildren,spaces,thumbnaillink showdrivename
if errorlevel 1 goto :fail

echo [2/2] Google Sites permissions across selected Shared Drives (direct vs inherited)...
"%GAM_PATH%" config auto_batch_min 1 num_threads %GAM_NUM_THREADS% redirect csv "%OUTDIR%\GSites_SharedDrive_Permissions.csv" multiprocess csv "%EFFECTIVE_CSV%" gam user "~account" print filelist select teamdriveid "~driveId" query "%SITES_QUERY%" fields id,name,webviewlink,owners,lastmodifyinguser,basicpermissions,permissiondetails,shared,driveid,copyrequireswriterpermission,viewerscancopycontent,writerscanshare,inheritedpermissionsdisabled oneitemperrow showshareddrivepermissions
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
