@echo off
setlocal enabledelayedexpansion
set SCRIPT_DIR=%~dp0
set OUTDIR=%SCRIPT_DIR%output
if not exist "%OUTDIR%" mkdir "%OUTDIR%"

REM ===========================================================================
REM Standalone export: Shared Drive-level metadata (restrictions + organizers)
REM for the same Shared Drives targeted by 01d_run_gam_exports_shareddrive.cmd.
REM
REM This is about the DRIVE itself (sharing policy, who administers it), not
REM the Sites hosted on it - use 01d for the Sites inventory/permissions.
REM
REM Usage:
REM   01e_list_shareddrive_metadata.cmd <SharedDriveIDs.csv> <ScanningUserEmail>
REM
REM   <SharedDriveIDs.csv>    Required. Same input file used by 01d - a CSV
REM                           with a header row containing a "driveId" column.
REM   <ScanningUserEmail>     Required. Must either be a member (Manager
REM                           /Organizer for full restriction visibility) of
REM                           every listed Shared Drive, or a Super Admin.
REM                           Can also be set via GAM_ADMIN_USER env var.
REM
REM Outputs (written to output/):
REM   GSites_SharedDrive_Settings.csv     One row per drive: sharing
REM                                       restrictions (domainUsersOnly,
REM                                       driveMembersOnly, etc.)
REM   GSites_SharedDrive_Organizers.csv   One row per (drive, organizer)
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

if not defined GAM_NUM_THREADS set GAM_NUM_THREADS=5
echo [INFO] Using num_threads=%GAM_NUM_THREADS% for GAM multiprocess operations

set "SETTINGS_JSONL=%OUTDIR%\GSites_SharedDrive_Settings.jsonl"
if exist "%SETTINGS_JSONL%" del /q "%SETTINGS_JSONL%"

REM NOTE: num_threads forced to 1 here (unlike the other exports) because the
REM output is one JSON object per drive; running this single-threaded avoids
REM any risk of interleaved output from concurrent workers corrupting the
REM JSON before it is parsed into CSV below.
echo [1/2] Shared Drive restrictions/settings...
"%GAM_PATH%" config num_threads 1 redirect stdout "%SETTINGS_JSONL%" multiprocess csv "%SHAREDDRIVE_CSV%" gam user %GAM_ADMIN_USER% info shareddrive "~driveId" fields id,name,createdtime,restrictions formatjson
if errorlevel 1 goto :fail

echo [2/2] Shared Drive organizers (drive-level Manager access)...
"%GAM_PATH%" config auto_batch_min 1 num_threads %GAM_NUM_THREADS% redirect csv "%OUTDIR%\GSites_SharedDrive_Organizers.csv" multiprocess csv "%SHAREDDRIVE_CSV%" gam user %GAM_ADMIN_USER% print shareddriveorganizers shareddrives "~driveId" includetypes user,group oneorganizer false
if errorlevel 1 goto :fail

echo.
echo Converting Shared Drive settings JSON output to CSV...
REM GAM's formatjson may pretty-print each drive's JSON object across multiple
REM lines, so objects are concatenated back-to-back rather than one per line.
REM Insert a separator between adjacent "}{" boundaries and wrap as a JSON
REM array so ConvertFrom-Json can parse all drives regardless of formatting.
powershell -NoProfile -Command ^
  "$raw = ''; if (Test-Path '%SETTINGS_JSONL%') { $raw = (Get-Content '%SETTINGS_JSONL%' -Raw) }; $raw = $raw.Trim(); if ([string]::IsNullOrWhiteSpace($raw)) { Write-Host '[WARN] No Shared Drive settings captured - check %SETTINGS_JSONL% for errors'; exit }; $joined = '[' + ($raw -replace '}\s*{', '},{') + ']'; try { $objs = $joined | ConvertFrom-Json } catch { Write-Host \"[WARN] Failed to parse Shared Drive settings JSON: $($_.Exception.Message)\"; exit }; $out = @($objs | ForEach-Object { [pscustomobject]@{ id=$_.id; name=$_.name; createdTime=$_.createdTime; adminManagedRestrictions=$_.restrictions.adminManagedRestrictions; domainUsersOnly=$_.restrictions.domainUsersOnly; driveMembersOnly=$_.restrictions.driveMembersOnly; copyRequiresWriterPermission=$_.restrictions.copyRequiresWriterPermission; sharingFoldersRequiresOrganizerPermission=$_.restrictions.sharingFoldersRequiresOrganizerPermission } }); if ($out.Count -gt 0) { $out | Export-Csv -NoTypeInformation -Path '%OUTDIR%\GSites_SharedDrive_Settings.csv' } else { Write-Host '[WARN] No Shared Drive settings parsed - check %SETTINGS_JSONL% for errors' }"

echo.
echo Shared Drive metadata export completed successfully.
echo Output folder: %OUTDIR%

set GAM_NUM_THREADS=

echo Cleaning CSV headers...
powershell -NoProfile -Command "Get-ChildItem '%OUTDIR%\GSites_SharedDrive_Organizers.csv' -ErrorAction SilentlyContinue | ForEach-Object { $lines = @(Get-Content $_.FullName); if ($lines.Count -gt 0) { $lines[0] = $lines[0] -replace '\.[0-9]+\.', '.'; $lines | Set-Content $_.FullName } }"
exit /b 0

:fail
echo.
echo Shared Drive metadata export failed. Review the command above and GAM error output.
exit /b 1
