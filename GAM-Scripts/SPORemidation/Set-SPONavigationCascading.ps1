# =============================================================================
# Set-SPONavigationCascading.ps1
# Switches SPO top-navigation from Mega Menu (horizontal/side-by-side)
# to Cascading (vertical dropdown) — matching Google Sites behaviour.
#
# REQUIREMENTS:
#   PowerShell 7+   (pwsh)
#   PnP.PowerShell  ->  Install-Module PnP.PowerShell -Scope CurrentUser
#
# SINGLE SITE:
#   pwsh -ExecutionPolicy Bypass -File .\Set-SPONavigationCascading.ps1 `
#        -SiteUrl "https://mngenvmcap908272.sharepoint.com/sites/ftcmigrationtestsite" `
#        -ClientId "3834b2e7-ab80-45fc-b4c8-ed5c960076b7"
#
# BULK (100 sites from CSV):
#   pwsh -ExecutionPolicy Bypass -File .\Set-SPONavigationCascading.ps1 `
#        -CsvPath .\SiteMapping.csv `
#        -ClientId "3834b2e7-ab80-45fc-b4c8-ed5c960076b7" `
#        -ThrottleLimit 5
#
# DRY RUN (no changes written):
#   Add -WhatIf to any of the above
# =============================================================================

param(
    [string]$SiteUrl = "",   # Single site mode
    [string]$CsvPath = "",   # Bulk mode — path to SiteMapping.csv
    [string]$ClientId = "3834b2e7-ab80-45fc-b4c8-ed5c960076b7",
    [string]$TenantId = "",   # Optional — auto-derived from SiteUrl if blank
    [int]   $ThrottleLimit = 5,   # Parallel sites (bulk mode only)
    [switch]$WhatIf               # Simulate — no changes written
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

# ── Helper ────────────────────────────────────────────────────────────────────
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        "WARN" { "Yellow" }; "ERROR" { "Red" }
        "SUCCESS" { "Green" }; default { "Cyan" }
    }
    Write-Host "[$ts][$Level] $Message" -ForegroundColor $color
}

# ── Core function — disable mega menu on ONE site ────────────────────────────
function Set-CascadingNav {
    param(
        [string]$SPOSiteUrl,
        [string]$AppClientId,
        [string]$AppTenantId,
        [bool]  $DryRun = $false
    )

    # Auto-derive tenant if not provided
    $spoHost = ([uri]$SPOSiteUrl).Host
    $tenantName = if ($AppTenantId) { $AppTenantId }
    else { ($spoHost -replace "\.sharepoint\.com$", "") + ".onmicrosoft.com" }
    $siteSlug = $SPOSiteUrl -replace "^https?://[^/]+/sites/", "" -replace "/.*", ""

    Write-Log "[$siteSlug] Connecting (Tenant: $tenantName) ..."
    try {
        Connect-PnPOnline -Url $SPOSiteUrl -ClientId $AppClientId `
            -Tenant $tenantName -Interactive -ErrorAction Stop
        Write-Log "[$siteSlug] Connected." "SUCCESS"
    }
    catch {
        Write-Log "[$siteSlug] Connection failed: $_" "ERROR"
        return [PSCustomObject]@{ Site = $siteSlug; Status = "FAILED"; Error = $_.ToString() }
    }

    try {
        # Read current MegaMenuEnabled value
        $webProps = Invoke-PnPSPRestMethod -Method Get `
            -Url "/_api/web?`$select=MegaMenuEnabled,Title"
        $current = $webProps.MegaMenuEnabled

        Write-Log "[$siteSlug] Current MegaMenuEnabled = $current"

        if (-not $current) {
            Write-Log "[$siteSlug] Already cascading (MegaMenuEnabled=false). No change needed." "SUCCESS"
            Disconnect-PnPOnline
            return [PSCustomObject]@{ Site = $siteSlug; Status = "ALREADY_CASCADING"; Error = "" }
        }

        if ($DryRun) {
            Write-Log "[$siteSlug] [DryRun] Would set MegaMenuEnabled = false (cascading)." "WARN"
            Disconnect-PnPOnline
            return [PSCustomObject]@{ Site = $siteSlug; Status = "DRYRUN"; Error = "" }
        }

        # Switch to cascading navigation (vertical dropdowns)
        Invoke-PnPSPRestMethod -Method Patch -Url "/_api/web" `
            -Content @{ "MegaMenuEnabled" = $false }

        # Verify the change was applied
        $verify = Invoke-PnPSPRestMethod -Method Get `
            -Url "/_api/web?`$select=MegaMenuEnabled"

        if (-not $verify.MegaMenuEnabled) {
            Write-Log "[$siteSlug] Navigation switched to CASCADING (vertical dropdowns)." "SUCCESS"
            Disconnect-PnPOnline
            return [PSCustomObject]@{ Site = $siteSlug; Status = "FIXED"; Error = "" }
        }
        else {
            Write-Log "[$siteSlug] Patch sent but MegaMenuEnabled is still true. Check permissions." "WARN"
            Disconnect-PnPOnline
            return [PSCustomObject]@{ Site = $siteSlug; Status = "WARN_NOT_APPLIED"; Error = "" }
        }
    }
    catch {
        Write-Log "[$siteSlug] Error: $_" "ERROR"
        Disconnect-PnPOnline
        return [PSCustomObject]@{ Site = $siteSlug; Status = "ERROR"; Error = $_.ToString() }
    }
}

# =============================================================================
# SINGLE SITE MODE
# =============================================================================
if ($SiteUrl) {
    $result = Set-CascadingNav -SPOSiteUrl $SiteUrl -AppClientId $ClientId `
        -AppTenantId $TenantId -DryRun $WhatIf.IsPresent
    Write-Log "Result: $($result.Status)" $(if ($result.Status -eq "FIXED") { "SUCCESS" } else { "INFO" })
    if ($result.Error) { Write-Log "Error detail: $($result.Error)" "ERROR" }
    exit
}

# =============================================================================
# BULK MODE — reads SiteMapping.csv, runs in parallel
# =============================================================================
if (-not $CsvPath) {
    Write-Log "Provide either -SiteUrl (single) or -CsvPath (bulk)." "ERROR"
    exit 1
}

$csvFull = Resolve-Path $CsvPath -ErrorAction Stop
$sites = Import-Csv -Path $csvFull
Write-Log "Loaded $($sites.Count) site(s) from $csvFull"

$logDir = Join-Path $PSScriptRoot "NavFixLogs"
$runStamp = Get-Date -Format "yyyyMMdd_HHmmss"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
$summaryPath = Join-Path $logDir "CascadingFix_Summary_$runStamp.csv"

$results = $sites | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
    $site = $_
    $scriptPath = $using:PSCommandPath
    $clientIdLocal = $using:ClientId
    $tenantIdLocal = $using:TenantId
    $isDryRun = $using:WhatIf

    # Dot-source to load Set-CascadingNav and Write-Log
    . $scriptPath

    Set-CascadingNav `
        -SPOSiteUrl  $site.SPOSiteUrl `
        -AppClientId $clientIdLocal `
        -AppTenantId $tenantIdLocal `
        -DryRun      ([bool]$isDryRun)
}

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host "`n========================================" -ForegroundColor White
Write-Host "  CASCADING NAV FIX — SUMMARY"           -ForegroundColor White
Write-Host "========================================" -ForegroundColor White

$results | ForEach-Object {
    $color = switch ($_.Status) {
        "FIXED" { "Green" }
        "ALREADY_CASCADING" { "Cyan" }
        "DRYRUN" { "Yellow" }
        default { "Red" }
    }
    Write-Host ("  {0,-40} [{1}]" -f $_.Site, $_.Status) -ForegroundColor $color
    if ($_.Error) { Write-Host "    ERROR: $($_.Error)" -ForegroundColor Red }
}

$results | Export-Csv -Path $summaryPath -NoTypeInformation
Write-Log "Summary saved to $summaryPath"

$fixed = @($results | Where-Object { $_.Status -eq "FIXED" }).Count
$skipped = @($results | Where-Object { $_.Status -eq "ALREADY_CASCADING" }).Count
$failed = @($results | Where-Object { $_.Status -in "FAILED", "ERROR", "WARN_NOT_APPLIED" }).Count

Write-Log "Fixed: $fixed  |  Already cascading: $skipped  |  Failed: $failed" `
$(if ($failed -gt 0) { "WARN" } else { "SUCCESS" })
