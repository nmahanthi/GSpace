<#
.SYNOPSIS
    Orchestrates the full GSites -> SPO Permissions Migration workflow.
.DESCRIPTION
    Runs the four stages in sequence:
      1. Export-GSitePermissions.ps1  -> GSite_Permissions_DATE.csv
      2. Export-SPOPermissions.ps1    -> SPO_Permissions_DATE.csv
      3. Compare-Permissions.ps1      -> Permission_Differences_DATE.csv
      4. Fix-SPOPermissions.ps1       -> Fix_Log_DATE.csv

    Each stage can be skipped with the -Skip* switches, e.g. if you have
    already exported GSite permissions and just want to re-compare.

    IMPORTANT:
      - Set Fix.WhatIf = $true in Config.psd1 on the first run to preview changes.
      - Review the Differences CSV before setting WhatIf = $false.
.PARAMETER ConfigPath
    Path to Config.psd1.  Defaults to .\Config.psd1
.PARAMETER SkipGSiteExport
    Skip Step 1 (useful when re-running after a failed SPO export).
.PARAMETER SkipSPOExport
    Skip Step 2.
.PARAMETER SkipCompare
    Skip Step 3.
.PARAMETER SkipFix
    Skip Step 4 (run export + compare only - safe for first-run auditing).
.PARAMETER InputCsvFile
    Path to InputSites.csv listing the specific Google Sites to scan.
    Columns: SiteId, SiteUrl, SiteName, SPOSiteUrl, Notes
    If omitted, Config.psd1 InputSitesFile value is used.
    If both are absent, ALL sites in the tenant are scanned (not recommended for large tenants).
.PARAMETER GSitePermissionsFile
    Supply an existing GSite CSV to override step 1 output.
.PARAMETER SPOPermissionsFile
    Supply an existing SPO CSV to override step 2 output.
.PARAMETER DifferencesFile
    Supply an existing Differences CSV to override step 3 output.
.EXAMPLE
    # Full run - scans only the sites listed in InputSites.csv (WhatIf mode by default)
    .\Start-Migration.ps1 -InputCsvFile ".\InputSites.csv"

    # Use a custom input file with a different name
    .\Start-Migration.ps1 -InputCsvFile ".\MySites_Batch1.csv"

    # Re-run fix stage only after reviewing the differences CSV
    .\Start-Migration.ps1 -SkipGSiteExport -SkipSPOExport -SkipCompare
#>
[CmdletBinding()]
param(
    [string]$ConfigPath            = ".\Config.psd1",
    [string]$InputCsvFile          = "",
    [switch]$SkipGSiteExport,
    [switch]$SkipSPOExport,
    [switch]$SkipCompare,
    [switch]$SkipFix,
    [string]$GSitePermissionsFile  = "",
    [string]$SPOPermissionsFile    = "",
    [string]$DifferencesFile       = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$startTime = Get-Date
$banner = @"

  +---------------------------------------------------------+
  |   GSites -> SPO Permission Migration Toolkit            |
  |   Started : $($startTime.ToString('yyyy-MM-dd HH:mm:ss'))                    |
  +---------------------------------------------------------+
"@
Write-Host $banner -ForegroundColor Cyan

# Validate config exists
if (-not (Test-Path $ConfigPath)) {
    throw "Config file not found: $ConfigPath. Please create Config.psd1 from the template."
}
$cfg = Import-PowerShellDataFile -Path $ConfigPath
Write-Host "Config loaded from: $ConfigPath" -ForegroundColor DarkGray

# Resolve InputCsvFile: parameter > config > warn
if (-not $InputCsvFile -and $cfg.ContainsKey('InputSitesFile')) {
    $InputCsvFile = $cfg.InputSitesFile
}
if ($InputCsvFile -and (Test-Path $InputCsvFile)) {
    $rowCount = (Import-Csv $InputCsvFile).Count
    Write-Host "Input sites file  : $InputCsvFile  ($rowCount sites to scan)" -ForegroundColor DarkGray
} else {
    Write-Warning "InputSites.csv not found at '$InputCsvFile'. ALL tenant sites will be scanned - this may be very slow."
}

# ---------------------------------------------------------------------------
# Helper: run a stage script and capture its returned output file path
# ---------------------------------------------------------------------------
function Invoke-Stage {
    param([string]$Label, [string]$ScriptPath, [hashtable]$Params)
    Write-Host "`n>>> STAGE: $Label" -ForegroundColor Magenta
    if (-not (Test-Path $ScriptPath)) { throw "Script not found: $ScriptPath" }
    $result = & $ScriptPath @Params
    Write-Host "<<< STAGE COMPLETE: $Label" -ForegroundColor Magenta
    return $result
}

$stage = 0

# ---------------------------------------------------------------------------
# Step 1: Export Google Sites permissions
# ---------------------------------------------------------------------------
$stage++
if (-not $SkipGSiteExport) {
    $params = @{ ConfigPath = $ConfigPath }
    if ($InputCsvFile) { $params.InputCsvFile = $InputCsvFile }
    $GSitePermissionsFile = Invoke-Stage "[$stage/4] Export GSite Permissions" ".\Export-GSitePermissions.ps1" $params
} else {
    Write-Host "`n>>> STAGE: [$stage/4] Export GSite Permissions [SKIPPED]" -ForegroundColor DarkGray
}

# ---------------------------------------------------------------------------
# Step 2: Export SharePoint Online permissions
# ---------------------------------------------------------------------------
$stage++
if (-not $SkipSPOExport) {
    $params = @{ ConfigPath = $ConfigPath }
    $SPOPermissionsFile = Invoke-Stage "[$stage/4] Export SPO Permissions" ".\Export-SPOPermissions.ps1" $params
} else {
    Write-Host "`n>>> STAGE: [$stage/4] Export SPO Permissions [SKIPPED]" -ForegroundColor DarkGray
}

# ---------------------------------------------------------------------------
# Step 3: Compare permissions
# ---------------------------------------------------------------------------
$stage++
if (-not $SkipCompare) {
    $params = @{ ConfigPath = $ConfigPath }
    if ($GSitePermissionsFile) { $params.GSitePermissionsFile = $GSitePermissionsFile }
    if ($SPOPermissionsFile)   { $params.SPOPermissionsFile   = $SPOPermissionsFile }
    $DifferencesFile = Invoke-Stage "[$stage/4] Compare Permissions" ".\Compare-Permissions.ps1" $params
} else {
    Write-Host "`n>>> STAGE: [$stage/4] Compare Permissions [SKIPPED]" -ForegroundColor DarkGray
}

# ---------------------------------------------------------------------------
# Step 4: Fix SPO permissions
# ---------------------------------------------------------------------------
$stage++
if (-not $SkipFix) {
    $params = @{ ConfigPath = $ConfigPath }
    if ($DifferencesFile) { $params.DifferencesFile = $DifferencesFile }
    $fixLog = Invoke-Stage "[$stage/4] Fix SPO Permissions" ".\Fix-SPOPermissions.ps1" $params
} else {
    Write-Host "`n>>> STAGE: [$stage/4] Fix SPO Permissions [SKIPPED]" -ForegroundColor DarkGray
    $fixLog = "(skipped)"
}

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
$elapsed = (Get-Date) - $startTime
Write-Host "`n=========================================" -ForegroundColor Cyan
Write-Host "  Migration Run Complete"                   -ForegroundColor Cyan
Write-Host "  Elapsed : $($elapsed.ToString('hh\:mm\:ss'))" -ForegroundColor Cyan
Write-Host "-----------------------------------------" -ForegroundColor DarkGray
Write-Host "  GSite CSV   : $GSitePermissionsFile"
Write-Host "  SPO CSV     : $SPOPermissionsFile"
Write-Host "  Diffs CSV   : $DifferencesFile"
Write-Host "  Fix Log CSV : $fixLog"
Write-Host "=========================================" -ForegroundColor Cyan

if ([bool]$cfg.Fix.WhatIf) {
    Write-Host "`n[!] WhatIf = true in Config.psd1 - no changes were applied to SPO." -ForegroundColor Yellow
    Write-Host "    Review the Differences CSV, then set Fix.WhatIf = `$false and re-run.`n" -ForegroundColor Yellow
}
