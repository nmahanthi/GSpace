# =============================================================================
# Invoke-BulkNavFix.ps1
# Bulk-fixes SharePoint Online navigation for 100 migrated Google Sites.
# Reads SiteMapping.csv → runs Invoke-SPONavFix in parallel batches.
#
# REQUIREMENTS:
#   PowerShell 7+  (for ForEach-Object -Parallel)
#   PnP.PowerShell  →  Install-Module PnP.PowerShell -Scope CurrentUser
#
# USAGE:
#   # Dry-run (no changes written):
#   .\Invoke-BulkNavFix.ps1 -CsvPath .\SiteMapping.csv -WhatIf
#
#   # Fix broken URLs across all 100 sites, 5 at a time:
#   .\Invoke-BulkNavFix.ps1 -CsvPath .\SiteMapping.csv -ThrottleLimit 5
#
#   # Full rebuild across all sites, 3 at a time:
#   .\Invoke-BulkNavFix.ps1 -CsvPath .\SiteMapping.csv -RebuildNavigation -ThrottleLimit 3
# =============================================================================

param(
    [Parameter(Mandatory)][string]$CsvPath,          # Path to SiteMapping.csv
    [int]   $ThrottleLimit = 5,                  # Sites processed in parallel
    [string]$ClientId = "3834b2e7-ab80-45fc-b4c8-ed5c960076b7",  # Azure AD App Client ID
    [string]$TenantId = "",                 # e.g. contoso.onmicrosoft.com
    [switch]$RebuildNavigation,                      # Wipe & rebuild nav on every site
    [switch]$WhatIf,                                 # Simulate — no changes written
    [string]$LogFolder = ".\NavFixLogs"      # Output folder for per-site logs
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

# ── Resolve paths ─────────────────────────────────────────────────────────────
$scriptRoot = $PSScriptRoot
$coreScript = Join-Path $scriptRoot "Fix-SPONavigation.ps1"
$csvFullPath = Resolve-Path $CsvPath -ErrorAction Stop

if (-not (Test-Path $coreScript)) {
    Write-Error "Cannot find Fix-SPONavigation.ps1 alongside this script at: $coreScript"
    exit 1
}

# ── Ensure log folder exists ──────────────────────────────────────────────────
$logDir = Join-Path $scriptRoot $LogFolder
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }

$runStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$summaryLog = Join-Path $logDir "BulkNavFix_Summary_$runStamp.csv"

# ── Load site list ────────────────────────────────────────────────────────────
$sites = Import-Csv -Path $csvFullPath
Write-Host "[INFO] Loaded $($sites.Count) site(s) from $csvFullPath" -ForegroundColor Cyan

if ($sites.Count -eq 0) { Write-Warning "CSV is empty. Nothing to do."; exit 0 }

# ── Parallel execution ────────────────────────────────────────────────────────
Write-Host "[INFO] Starting parallel run with ThrottleLimit=$ThrottleLimit ..." -ForegroundColor Cyan

$results = $sites | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
    $site = $_
    $coreScriptPath = $using:coreScript
    $logDirectory = $using:logDir
    $runStampLocal = $using:runStamp
    $isRebuild = $using:RebuildNavigation
    $isDryRun = $using:WhatIf
    $clientIdLocal = $using:ClientId
    $tenantIdLocal = $using:TenantId

    # Each parallel thread gets its own log file
    $siteSlug = $site.SPOSiteUrl -replace "^https?://[^/]+/sites/", "" -replace "/.*", ""
    $siteLog = Join-Path $logDirectory "${siteSlug}_${runStampLocal}.log"

    try {
        # Dot-source the core script to load Invoke-SPONavFix & Write-Log
        . $coreScriptPath

        $result = Invoke-SPONavFix `
            -SPOSiteUrl    $site.SPOSiteUrl `
            -GSitesBaseUrl $site.GoogleSitesBaseUrl `
            -ClientId      $clientIdLocal `
            -TenantId      $tenantIdLocal `
            -Rebuild       ([bool]$isRebuild) `
            -DryRun        ([bool]$isDryRun) `
            -LogFile       $siteLog

        $result
    }
    catch {
        [PSCustomObject]@{
            Site   = $siteSlug
            Status = "ERROR"
            Error  = $_.ToString()
        }
    }
}

# ── Summary report ────────────────────────────────────────────────────────────
Write-Host "`n========================================" -ForegroundColor White
Write-Host "  BULK NAV FIX — SUMMARY" -ForegroundColor White
Write-Host "========================================" -ForegroundColor White

$results | ForEach-Object {
    $color = switch -Wildcard ($_.Status) {
        "FIXED*" { "Green" }
        "REBUILT" { "Cyan" }
        "FAILED" { "Red" }
        "ERROR" { "Red" }
        default { "Yellow" }
    }
    Write-Host ("  {0,-40} [{1}]" -f $_.Site, $_.Status) -ForegroundColor $color
    if ($_.Error) { Write-Host "    ⚠  $($_.Error)" -ForegroundColor Red }
}

# Write machine-readable summary CSV
$results | Export-Csv -Path $summaryLog -NoTypeInformation
Write-Host "`n[INFO] Summary saved to: $summaryLog" -ForegroundColor Cyan

$failed = @($results | Where-Object { $_.Status -in "FAILED", "ERROR" })
if ($failed.Count -gt 0) {
    Write-Host "[WARN] $($failed.Count) site(s) failed. Check per-site logs in: $logDir" -ForegroundColor Yellow
}
else {
    Write-Host "[SUCCESS] All $($results.Count) site(s) processed successfully." -ForegroundColor Green
}
