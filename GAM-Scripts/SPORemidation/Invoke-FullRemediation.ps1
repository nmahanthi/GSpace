# =============================================================================
# Invoke-FullRemediation.ps1
# Master orchestrator — runs ALL three remediation fixes across 100 sites:
#   Step 1 : Fix broken navigation URLs  (Fix-SPONavigation.ps1)
#   Step 2 : Switch nav to cascading/vertical (Set-SPONavigationCascading.ps1)
#   Step 3 : Fix page text formatting on ALL pages (Fix-SPOPageFormatting.ps1)
#
# AUTHENTICATION — Single sign-in for all sites:
#   A browser window opens ONCE before processing starts. After you sign in,
#   MSAL caches the token. Every subsequent Connect-PnPOnline (per site) reuses
#   that cache silently — no more browser prompts for the remaining 99 sites.
#
# REQUIREMENTS:
#   PowerShell 7+  (pwsh)
#   PnP.PowerShell  ->  Install-Module PnP.PowerShell -Scope CurrentUser
#   All four scripts must be in the same folder.
#
# DRY RUN (simulate, no changes written):
#   pwsh -ExecutionPolicy Bypass -File .\Invoke-FullRemediation.ps1 `
#        -CsvPath .\SiteMapping.csv `
#        -ClientId "3834b2e7-ab80-45fc-b4c8-ed5c960076b7" -DryRun
#
# FULL RUN (applies all fixes):
#   pwsh -ExecutionPolicy Bypass -File .\Invoke-FullRemediation.ps1 `
#        -CsvPath .\SiteMapping.csv `
#        -ClientId "3834b2e7-ab80-45fc-b4c8-ed5c960076b7"
#
# NOTE: Use -DryRun (not -WhatIf) to simulate. PowerShell intercepts -WhatIf
#       as a built-in common parameter which prevents it from reaching the script.
# =============================================================================

param(
    [Parameter(Mandatory)][string]$CsvPath,
    [string]$ClientId = "3834b2e7-ab80-45fc-b4c8-ed5c960076b7",
    [string]$TenantId = "",
    [switch]$DryRun       # Use this instead of -WhatIf for safe simulation
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

# ── Resolve sibling script paths ──────────────────────────────────────────────
$scriptRoot = $PSScriptRoot
$navScript = Join-Path $scriptRoot "Fix-SPONavigation.ps1"
$pageScript = Join-Path $scriptRoot "Fix-SPOPageFormatting.ps1"

foreach ($s in @($navScript, $pageScript)) {
    if (-not (Test-Path $s)) { Write-Error "Missing: $s"; exit 1 }
}

# ── Log folder ────────────────────────────────────────────────────────────────
$logDir = Join-Path $scriptRoot "RemediationLogs"
$runStamp = Get-Date -Format "yyyyMMdd_HHmmss"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }

$csvSummary = Join-Path $logDir "FullRemediation_$runStamp.csv"
$htmlReport = Join-Path $logDir "FullRemediation_$runStamp.html"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        "WARN" { "Yellow" }; "ERROR" { "Red" }
        "SUCCESS" { "Green" }; default { "Cyan" }
    }
    Write-Host "[$ts][$Level] $Message" -ForegroundColor $color
}

# ── Load CSV ──────────────────────────────────────────────────────────────────
$sites = Import-Csv -Path (Resolve-Path $CsvPath -ErrorAction Stop)
Write-Log "Loaded $($sites.Count) site(s). DryRun=$($DryRun.IsPresent)"

# ── Load helper functions ──────────────────────────────────────────────────────
. $navScript
. $pageScript

# ── Single sign-in — browser opens ONCE here, token cached for all sites ──────
$firstSite = $sites[0]
$firstHost = ([uri]$firstSite.SPOSiteUrl).Host
$firstTenant = if ($TenantId) { $TenantId } else { ($firstHost -replace "\.sharepoint\.com$", "") + ".onmicrosoft.com" }

Write-Log "Authenticating — a browser window will open ONCE. Sign in and it will be reused for all $($sites.Count) site(s)." "INFO"
try {
    Connect-PnPOnline -Url $firstSite.SPOSiteUrl -ClientId $ClientId -Tenant $firstTenant -Interactive -ErrorAction Stop
    Write-Log "Authentication successful. Token cached — no further prompts needed." "SUCCESS"
    Disconnect-PnPOnline
}
catch {
    Write-Log "Pre-authentication failed: $_" "ERROR"
    exit 1
}

# ── Per-site remediation block (sequential) ───────────────────────────────────
$results = $sites | ForEach-Object {

    $site = $_
    $isDryRun = $DryRun.IsPresent
    $spoHost = ([uri]$site.SPOSiteUrl).Host
    $tenantName = if ($TenantId) { $TenantId } else { ($spoHost -replace "\.sharepoint\.com$", "") + ".onmicrosoft.com" }
    $siteSlug = $site.SPOSiteUrl -replace "^https?://[^/]+/sites/", "" -replace "/.*", ""
    $siteLog = Join-Path $logDir "${siteSlug}_${runStamp}.log"

    $result = [PSCustomObject]@{
        SiteName     = $site.SiteName
        SiteSlug     = $siteSlug
        NavFix       = "PENDING"
        CascadingFix = "PENDING"
        PagesFixed   = 0
        PagesScanned = 0
        PagesFailed  = 0
        Error        = ""
        LogFile      = $siteLog
    }

    # ── Single Connect per site (MSAL reuses cached token — no browser prompt) ─
    try {
        Write-Log "[$siteSlug] Connecting to $($site.SPOSiteUrl) ..."
        Add-Content $siteLog "[$((Get-Date -f 'yyyy-MM-dd HH:mm:ss'))][INFO] Connecting to $($site.SPOSiteUrl) ..."
        Connect-PnPOnline -Url $site.SPOSiteUrl -ClientId $ClientId -Tenant $tenantName -Interactive -ErrorAction Stop
        Write-Log "[$siteSlug] Connected." "SUCCESS"
        Add-Content $siteLog "[$((Get-Date -f 'yyyy-MM-dd HH:mm:ss'))][SUCCESS] Connected."
    }
    catch {
        $result.NavFix = $result.CascadingFix = "CONNECT_FAILED"
        $result.Error = $_.ToString()
        Write-Log "[$siteSlug] Connection failed: $_" "ERROR"
        Add-Content $siteLog "[$((Get-Date -f 'yyyy-MM-dd HH:mm:ss'))][ERROR] Connection failed: $_"
        return $result
    }

    # ── STEP 1: Fix broken navigation URLs (reuses current connection) ────────
    try {
        Add-Content $siteLog "[$((Get-Date -f 'yyyy-MM-dd HH:mm:ss'))][INFO] STEP 1: Nav URL fix ..."
        $navOut = Invoke-SPONavFix `
            -SPOSiteUrl          $site.SPOSiteUrl `
            -GSitesBaseUrl       $site.GoogleSitesBaseUrl `
            -ClientId            $ClientId `
            -TenantId            $tenantName `
            -Rebuild             ($site.RebuildNavigation -eq "TRUE") `
            -DryRun              $isDryRun `
            -LogFile             $siteLog `
            -UseCurrentConnection
        $result.NavFix = $navOut.Status
        Add-Content $siteLog "[$((Get-Date -f 'yyyy-MM-dd HH:mm:ss'))][INFO] Nav fix result: $($navOut.Status)"
    }
    catch {
        $result.NavFix = "ERROR"
        Add-Content $siteLog "[$((Get-Date -f 'yyyy-MM-dd HH:mm:ss'))][ERROR] Nav fix failed: $_"
    }

    # ── STEP 2: Switch to cascading navigation ────────────────────────────────
    try {
        Add-Content $siteLog "[$((Get-Date -f 'yyyy-MM-dd HH:mm:ss'))][INFO] STEP 2: Cascading nav fix ..."
        $webProps = Invoke-PnPSPRestMethod -Method Get `
            -Url "/_api/web?`$select=MegaMenuEnabled"
        if ($webProps.MegaMenuEnabled) {
            if (-not $isDryRun) {
                Invoke-PnPSPRestMethod -Method Patch -Url "/_api/web" `
                    -Content @{ "MegaMenuEnabled" = $false }
                $result.CascadingFix = "FIXED"
            }
            else {
                $result.CascadingFix = "DRYRUN"
            }
        }
        else {
            $result.CascadingFix = "ALREADY_CASCADING"
        }
        Add-Content $siteLog "[$((Get-Date -f 'yyyy-MM-dd HH:mm:ss'))][INFO] Cascading result: $($result.CascadingFix)"
    }
    catch {
        $result.CascadingFix = "ERROR"
        Add-Content $siteLog "[$((Get-Date -f 'yyyy-MM-dd HH:mm:ss'))][ERROR] Cascading fix failed: $_"
    }

    # ── STEP 3: Fix page text formatting on ALL pages ─────────────────────────
    try {
        Add-Content $siteLog "[$((Get-Date -f 'yyyy-MM-dd HH:mm:ss'))][INFO] STEP 3: Page formatting fix ..."
        $pages = Get-PnPListItem -List "SitePages" -Fields "FileLeafRef" |
        Where-Object { $_["FileLeafRef"] -like "*.aspx" }
        $result.PagesScanned = $pages.Count
        Add-Content $siteLog "[$((Get-Date -f 'yyyy-MM-dd HH:mm:ss'))][INFO] Found $($pages.Count) pages."

        foreach ($p in $pages) {
            $pageName = $p["FileLeafRef"]
            try {
                $pageOut = Repair-SPOPage -SPOSiteUrl $site.SPOSiteUrl `
                    -SPOPageName $pageName -DryRun $isDryRun
                if ($pageOut.Status -like "FIXED*" -or $pageOut.Status -like "DRYRUN*") {
                    $result.PagesFixed++
                }
                Add-Content $siteLog "[$((Get-Date -f 'yyyy-MM-dd HH:mm:ss'))][INFO] Page $pageName : $($pageOut.Status)"
            }
            catch {
                $result.PagesFailed++
                Add-Content $siteLog "[$((Get-Date -f 'yyyy-MM-dd HH:mm:ss'))][ERROR] Page $pageName failed: $_"
            }
        }
    }
    catch {
        Add-Content $siteLog "[$((Get-Date -f 'yyyy-MM-dd HH:mm:ss'))][ERROR] Page list failed: $_"
    }

    # ── Disconnect ────────────────────────────────────────────────────────────
    Disconnect-PnPOnline
    Add-Content $siteLog "[$((Get-Date -f 'yyyy-MM-dd HH:mm:ss'))][INFO] Disconnected. Done."

    return $result
}

# =============================================================================
# Save CSV summary
# =============================================================================
$results | Export-Csv -Path $csvSummary -NoTypeInformation
Write-Log "CSV summary saved: $csvSummary" "SUCCESS"

# =============================================================================
# HTML Report Generator
# =============================================================================
function New-HtmlReport {
    param([object[]]$Data, [string]$OutputPath, [string]$RunStamp, [bool]$IsDryRun)

    $modeLabel = if ($IsDryRun) { " — DRY RUN (no changes written)" } else { "" }
    $totalSites = $Data.Count
    $totalFixed = @($Data | Where-Object { $_.NavFix -like "FIXED*" -or $_.CascadingFix -eq "FIXED" -or $_.PagesFixed -gt 0 }).Count
    $totalFailed = @($Data | Where-Object { $_.NavFix -like "*FAILED*" -or $_.NavFix -eq "ERROR" -or $_.CascadingFix -eq "ERROR" -or $_.PagesFailed -gt 0 }).Count

    $rows = foreach ($r in $Data) {
        $navColor = switch -Wildcard ($r.NavFix) { "FIXED*" { "#d4edda" } "ERROR" { "#f8d7da" } "CONNECT*" { "#f8d7da" } default { "#fff3cd" } }
        $cascColor = switch ($r.CascadingFix) { "FIXED" { "#d4edda" } "ALREADY_CASCADING" { "#d1ecf1" } "ERROR" { "#f8d7da" } default { "#fff3cd" } }
        $pageColor = if ($r.PagesFailed -gt 0) { "#f8d7da" } elseif ($r.PagesFixed -gt 0) { "#d4edda" } else { "#d1ecf1" }
        $logLink = if (Test-Path $r.LogFile) { "<a href='$($r.LogFile)'>View Log</a>" } else { "—" }

        "<tr>
          <td>$($r.SiteName)</td>
          <td>$($r.SiteSlug)</td>
          <td style='background:$navColor'>$($r.NavFix)</td>
          <td style='background:$cascColor'>$($r.CascadingFix)</td>
          <td style='background:$pageColor'>$($r.PagesFixed) / $($r.PagesScanned) fixed ($($r.PagesFailed) failed)</td>
          <td>$($r.Error)</td>
          <td>$logLink</td>
        </tr>"
    }

    $html = @"
<!DOCTYPE html>
<html><head><meta charset='utf-8'>
<title>SPO Full Remediation Report — $RunStamp</title>
<style>
  body { font-family: Segoe UI, Arial, sans-serif; margin: 30px; color: #212529; }
  h1   { color: #0078d4; }
  .badge { display:inline-block; padding:4px 10px; border-radius:12px; font-size:.85em; margin:2px; }
  .ok    { background:#d4edda; color:#155724; }
  .warn  { background:#fff3cd; color:#856404; }
  .err   { background:#f8d7da; color:#721c24; }
  table  { border-collapse:collapse; width:100%; margin-top:20px; font-size:.9em; }
  th     { background:#0078d4; color:#fff; padding:8px 12px; text-align:left; }
  td     { padding:7px 12px; border-bottom:1px solid #dee2e6; }
  tr:hover td { background:#f1f3f5; }
  a      { color:#0078d4; }
</style></head><body>
<h1>SPO Full Remediation Report$modeLabel</h1>
<p>Run: <strong>$RunStamp</strong> &nbsp;|&nbsp;
   Sites: <strong>$totalSites</strong> &nbsp;|&nbsp;
   <span class='badge ok'>Fixed/Updated: $totalFixed</span>
   <span class='badge err'>Issues: $totalFailed</span></p>
<table>
  <tr><th>Site Name</th><th>Slug</th><th>Nav URL Fix</th><th>Cascading Nav</th><th>Pages (Formatting)</th><th>Error</th><th>Log</th></tr>
  $($rows -join "`n")
</table>
<p style='margin-top:20px;font-size:.8em;color:#6c757d'>Generated by Invoke-FullRemediation.ps1 on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
</body></html>
"@
    $html | Out-File -FilePath $OutputPath -Encoding utf8
}

# =============================================================================
# Generate HTML report + Console summary
# =============================================================================
New-HtmlReport -Data $results -OutputPath $htmlReport -RunStamp $runStamp -IsDryRun $DryRun.IsPresent
Write-Log "HTML report saved: $htmlReport" "SUCCESS"

Write-Host ""
Write-Host "================================================================" -ForegroundColor White
Write-Host "  FULL REMEDIATION COMPLETE$(if ($DryRun) { ' [DRY RUN]' })" -ForegroundColor White
Write-Host "================================================================" -ForegroundColor White

$results | ForEach-Object {
    $overall = if ($_.NavFix -like "*ERROR*" -or $_.CascadingFix -eq "ERROR" -or $_.PagesFailed -gt 0) { "Red" }
    elseif ($_.NavFix -like "FIXED*" -or $_.CascadingFix -eq "FIXED" -or $_.PagesFixed -gt 0) { "Green" }
    else { "Cyan" }
    Write-Host ("  {0,-35}  Nav:{1,-12}  Casc:{2,-20}  Pages:{3}/{4} fixed" -f `
            $_.SiteSlug, $_.NavFix, $_.CascadingFix, $_.PagesFixed, $_.PagesScanned) -ForegroundColor $overall
}

Write-Host ""
Write-Host "  Reports:" -ForegroundColor White
Write-Host "    CSV  -> $csvSummary" -ForegroundColor Gray
Write-Host "    HTML -> $htmlReport" -ForegroundColor Gray
Write-Host "================================================================" -ForegroundColor White
