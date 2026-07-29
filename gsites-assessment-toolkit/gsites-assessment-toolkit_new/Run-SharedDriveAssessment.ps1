<#
.SYNOPSIS
    Shared Drive GSites Assessment Orchestrator.

.DESCRIPTION
    Runs the full standalone Shared Drive assessment pipeline in one command:
      1. 01d_run_gam_exports_shareddrive.cmd  - Site inventory + permissions
      2. 01e_list_shareddrive_metadata.cmd    - Drive-level settings/organizers
      3. 05b_score_shareddrive_sites.ps1      - Dedup + complexity/security scoring

.PARAMETER SharedDriveIdsCsv
    Path to the input CSV with header "driveId,account" - one row per Shared
    Drive paired with the Workspace user account that has access to that
    specific drive (its owner/organizer/member, or a Super Admin).

.PARAMETER PrimaryDomain
    Your primary domain (e.g. "rocheua.com"), used to distinguish internal
    vs. external grantees in the permissions-based security score.

.PARAMETER DefaultScanningUserEmail
    Optional fallback account used for any row in SharedDriveIdsCsv whose
    "account" cell is left blank. Can also be set via GAM_ADMIN_USER env var.

.PARAMETER GamThreads
    Number of parallel GAM worker processes for the export steps (default: 5).

.PARAMETER SkipExport
    Skip step 1 (01d) - reuse existing GSites_SharedDrive_Inventory/Permissions.csv.

.PARAMETER SkipMetadata
    Skip step 2 (01e) - reuse existing GSites_SharedDrive_Settings/Organizers.csv.

.EXAMPLE
    .\Run-SharedDriveAssessment.ps1 -SharedDriveIdsCsv "SharedDriveIDs.csv" -PrimaryDomain "rocheua.com"
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$SharedDriveIdsCsv,

    [Parameter(Mandatory = $true)]
    [string]$PrimaryDomain,

    [string]$DefaultScanningUserEmail,

    [int]$GamThreads = 5,

    [switch]$SkipExport,
    [switch]$SkipMetadata
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ScriptDir = $PSScriptRoot
$OutputDir = Join-Path $ScriptDir 'output'
$LogsDir = Join-Path $ScriptDir 'logs'

function Write-Step { param([string]$Message) Write-Host "`n========================================" -ForegroundColor Cyan; Write-Host "  $Message" -ForegroundColor Cyan; Write-Host "========================================`n" -ForegroundColor Cyan }
function Write-Success { param([string]$Message) Write-Host "[OK] $Message" -ForegroundColor Green }
function Write-Info { param([string]$Message) Write-Host "[INFO] $Message" -ForegroundColor Yellow }
function Write-Error-Custom { param([string]$Message) Write-Host "[ERROR] $Message" -ForegroundColor Red }

function Write-LogTail {
    param([string]$Path, [string]$Label, [ConsoleColor]$Color = [ConsoleColor]::Gray, [int]$Tail = 20)
    if (-not (Test-Path $Path)) { return }
    Write-Info $Label
    foreach ($line in @(Get-Content $Path -Tail $Tail)) { Write-Host "    $line" -ForegroundColor $Color }
}

function Invoke-LoggedCmd {
    param([Parameter(Mandatory = $true)][string]$CmdScript, [string[]]$ScriptArgs = @(), [Parameter(Mandatory = $true)][string]$LogPrefix)

    if (-not (Test-Path $LogsDir)) { New-Item -ItemType Directory -Path $LogsDir -Force | Out-Null }
    $stdoutLog = Join-Path $LogsDir "${LogPrefix}_stdout.log"
    $stderrLog = Join-Path $LogsDir "${LogPrefix}_stderr.log"
    if (Test-Path $stdoutLog) { Remove-Item $stdoutLog -Force }
    if (Test-Path $stderrLog) { Remove-Item $stderrLog -Force }

    $argList = @('/c', "`"$CmdScript`"") + $ScriptArgs
    $proc = Start-Process -FilePath 'cmd.exe' -ArgumentList $argList -WorkingDirectory $ScriptDir -NoNewWindow -Wait -PassThru `
        -RedirectStandardOutput $stdoutLog -RedirectStandardError $stderrLog

    return [pscustomobject]@{ ExitCode = $proc.ExitCode; StdOutLog = $stdoutLog; StdErrLog = $stderrLog }
}

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "  Shared Drive GSites Assessment - Orchestrator" -ForegroundColor Cyan
Write-Host "============================================================`n" -ForegroundColor Cyan
Write-Info "Shared Drive IDs CSV: $SharedDriveIdsCsv"
Write-Info "Primary Domain: $PrimaryDomain"
Write-Info "Output Directory: $OutputDir"

if (-not (Test-Path $SharedDriveIdsCsv)) { throw "Shared Drive IDs CSV not found: $SharedDriveIdsCsv" }
if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }

[Environment]::SetEnvironmentVariable('GAM_NUM_THREADS', $GamThreads, 'Process')
$cmdArgs = @($SharedDriveIdsCsv)
if ($DefaultScanningUserEmail) { $cmdArgs += $DefaultScanningUserEmail }

# ============================================================================
# STEP 1: Site inventory + permissions (01d)
# ============================================================================
if (-not $SkipExport) {
    Write-Step "STEP 1: Shared Drive GSites export (01d_run_gam_exports_shareddrive.cmd)"
    $script01d = Join-Path $ScriptDir '01d_run_gam_exports_shareddrive.cmd'
    if (-not (Test-Path $script01d)) { throw "Script not found: $script01d" }

    $result = Invoke-LoggedCmd -CmdScript $script01d -ScriptArgs $cmdArgs -LogPrefix '01d_shareddrive_export'
    if ($result.ExitCode -ne 0) {
        Write-Error-Custom "01d export failed with exit code $($result.ExitCode)"
        Write-LogTail -Path $result.StdErrLog -Label 'stderr tail' -Color Red
        Write-LogTail -Path $result.StdOutLog -Label 'stdout tail' -Color Gray
        throw '01d_run_gam_exports_shareddrive.cmd failed'
    }
    Write-Success 'Shared Drive site inventory + permissions export completed'
}
else {
    Write-Step "STEP 1: Shared Drive GSites export (SKIPPED)"
}

# ============================================================================
# STEP 2: Drive-level settings + organizers (01e)
# ============================================================================
if (-not $SkipMetadata) {
    Write-Step "STEP 2: Shared Drive metadata (01e_list_shareddrive_metadata.cmd)"
    $script01e = Join-Path $ScriptDir '01e_list_shareddrive_metadata.cmd'
    if (-not (Test-Path $script01e)) { throw "Script not found: $script01e" }

    $result = Invoke-LoggedCmd -CmdScript $script01e -ScriptArgs $cmdArgs -LogPrefix '01e_shareddrive_metadata'
    if ($result.ExitCode -ne 0) {
        Write-Error-Custom "01e metadata export failed with exit code $($result.ExitCode)"
        Write-LogTail -Path $result.StdErrLog -Label 'stderr tail' -Color Red
        Write-LogTail -Path $result.StdOutLog -Label 'stdout tail' -Color Gray
        Write-Info "Continuing - Settings/Organizers are not required for scoring."
    }
    else {
        Write-Success 'Shared Drive settings + organizers export completed'
    }
}
else {
    Write-Step "STEP 2: Shared Drive metadata (SKIPPED)"
}

# ============================================================================
# STEP 3: Score (05b)
# ============================================================================
Write-Step "STEP 3: Scoring (05b_score_shareddrive_sites.ps1)"
$scoreScript = Join-Path $ScriptDir '05b_score_shareddrive_sites.ps1'
if (-not (Test-Path $scoreScript)) { throw "Script not found: $scoreScript" }

& $scoreScript -OutputDir $OutputDir -PrimaryDomain $PrimaryDomain
Write-Success 'Shared Drive scoring completed'

# ============================================================================
# SUMMARY
# ============================================================================
Write-Host "`n============================================================" -ForegroundColor Green
Write-Host "  SHARED DRIVE ASSESSMENT COMPLETE" -ForegroundColor Green
Write-Host "============================================================`n" -ForegroundColor Green
Write-Host "OUTPUT FILES:" -ForegroundColor Cyan
Write-Host "  Location: $OutputDir" -ForegroundColor Gray
Write-Host ""

foreach ($file in @('GSites_SharedDrive_Inventory.csv', 'GSites_SharedDrive_Permissions.csv',
                     'GSites_SharedDrive_Settings.csv', 'GSites_SharedDrive_Organizers.csv',
                     'GSites_SharedDrive_Complexity_Report.csv')) {
    $path = Join-Path $OutputDir $file
    if (Test-Path $path) {
        $lineCount = 0
        switch -File $path { default { $lineCount++ } }
        $rowCount = [Math]::Max(0, $lineCount - 1)
        Write-Host "  [OK] $file " -NoNewline -ForegroundColor Green
        Write-Host "($rowCount rows)" -ForegroundColor Gray
    }
    else {
        Write-Host "  [MISSING] $file" -ForegroundColor DarkYellow
    }
}
Write-Host ""
