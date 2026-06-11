# =============================================================================
# Fix-SPOPageFormatting.ps1
# Fixes text-formatting issues in SPO SitePages after Google Sites migration:
#   1. Nested ordered list sub-items rendered as letters (a,b) -> numbers (1,2)
#   2. Bullet list markers rendered as squares -> round circles
#
# REQUIREMENTS:
#   PowerShell 7+  (pwsh)
#   PnP.PowerShell -> Install-Module PnP.PowerShell -Scope CurrentUser
#
# SINGLE PAGE:
#   pwsh -ExecutionPolicy Bypass -File .\Fix-SPOPageFormatting.ps1 `
#        -SiteUrl "https://mngenvmcap908272.sharepoint.com/sites/ftcmigrationtestsite" `
#        -PageName "Text-Formatting.aspx" `
#        -ClientId "3834b2e7-ab80-45fc-b4c8-ed5c960076b7"
#
# ALL PAGES IN A SITE:
#   pwsh -ExecutionPolicy Bypass -File .\Fix-SPOPageFormatting.ps1 `
#        -SiteUrl "https://mngenvmcap908272.sharepoint.com/sites/ftcmigrationtestsite" `
#        -ClientId "3834b2e7-ab80-45fc-b4c8-ed5c960076b7" `
#        -AllPages
#
# BULK (all sites from CSV):
#   pwsh -ExecutionPolicy Bypass -File .\Fix-SPOPageFormatting.ps1 `
#        -CsvPath .\SiteMapping.csv `
#        -ClientId "3834b2e7-ab80-45fc-b4c8-ed5c960076b7" `
#        -AllPages
#
# Add -WhatIf to any command to simulate without writing changes.
# =============================================================================

param(
    [string]$SiteUrl = "",
    [string]$PageName = "",        # e.g. "Text-Formatting.aspx" — single page mode
    [string]$CsvPath = "",        # Bulk mode: path to SiteMapping.csv
    [string]$ClientId = "3834b2e7-ab80-45fc-b4c8-ed5c960076b7",
    [string]$TenantId = "",
    [int]   $ThrottleLimit = 5,
    [switch]$AllPages,              # Process every page in SitePages library
    [switch]$WhatIf
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

# ── HTML fix rules ────────────────────────────────────────────────────────────
# Returns corrected HTML and a count of changes made
function Repair-PageHtml {
    param([string]$Html)

    $original = $Html
    $changes = 0

    # Fix 1: Nested ordered list — lower-alpha (a,b,c) -> decimal (1,2,3)
    # Matches: <ol ...style="...list-style-type:lower-alpha...">
    $fixed = $Html -replace '(?i)(list-style-type\s*:\s*)lower-alpha', '${1}decimal'
    if ($fixed -ne $Html) { $changes++; $Html = $fixed }

    # Fix 2: Nested ordered list — type="a" attribute -> type="1"
    $fixed = $Html -replace '(?i)<ol([^>]*)\btype="a"', '<ol${1}type="1"'
    if ($fixed -ne $Html) { $changes++; $Html = $fixed }

    # Fix 3: Bullet list — square markers -> disc (round circles)
    $fixed = $Html -replace '(?i)(list-style-type\s*:\s*)square', '${1}disc'
    if ($fixed -ne $Html) { $changes++; $Html = $fixed }

    return [PSCustomObject]@{ Html = $Html; Changes = $changes; Modified = ($Html -ne $original) }
}

# ── Fix one page ──────────────────────────────────────────────────────────────
function Repair-SPOPage {
    param([string]$SPOSiteUrl, [string]$SPOPageName, [bool]$DryRun)

    $siteSlug = $SPOSiteUrl -replace "^https?://[^/]+/sites/", "" -replace "/.*", ""
    Write-Log "[$siteSlug] Processing page: $SPOPageName"

    try {
        $page = Get-PnPPage -Identity $SPOPageName -ErrorAction Stop
    }
    catch {
        Write-Log "[$siteSlug] Could not load page '$SPOPageName': $_" "ERROR"
        return [PSCustomObject]@{ Page = $SPOPageName; Status = "PAGE_NOT_FOUND"; Changes = 0 }
    }

    # Get all text web parts on the page
    $textParts = $page.Controls | Where-Object {
        $_.GetType().Name -like "*PageText*" -or $_.PSObject.Properties['Text']
    }
    if (-not $textParts) {
        Write-Log "[$siteSlug] No text web parts found on '$SPOPageName'." "WARN"
        return [PSCustomObject]@{ Page = $SPOPageName; Status = "NO_TEXT_PARTS"; Changes = 0 }
    }

    $totalChanges = 0
    $partIndex = 0

    foreach ($part in $textParts) {
        $partIndex++
        $partText = if ($part.PSObject.Properties['Text']) { $part.Text } else { $null }
        if ([string]::IsNullOrWhiteSpace($partText)) { continue }

        $result = Repair-PageHtml -Html $partText

        if ($result.Modified) {
            Write-Log "[$siteSlug]   WebPart #$partIndex — $($result.Changes) fix(es) applied." "WARN"
            if (-not $DryRun) {
                $part.Text = $result.Html
            }
            else {
                Write-Log "[$siteSlug]   [DryRun] Would update WebPart #$partIndex." "WARN"
            }
            $totalChanges += $result.Changes
        }
        else {
            Write-Log "[$siteSlug]   WebPart #$partIndex — OK (no formatting issues)."
        }
    }

    if ($totalChanges -gt 0 -and -not $DryRun) {
        $page.Save()    | Out-Null
        $page.Publish("Auto-fixed text formatting after Google Sites migration") | Out-Null
        Write-Log "[$siteSlug] Page '$SPOPageName' saved and published." "SUCCESS"
        return [PSCustomObject]@{ Page = $SPOPageName; Status = "FIXED($totalChanges)"; Changes = $totalChanges }
    }
    elseif ($totalChanges -gt 0 -and $DryRun) {
        return [PSCustomObject]@{ Page = $SPOPageName; Status = "DRYRUN($totalChanges)"; Changes = $totalChanges }
    }
    else {
        Write-Log "[$siteSlug] Page '$SPOPageName' — no formatting issues found." "SUCCESS"
        return [PSCustomObject]@{ Page = $SPOPageName; Status = "OK"; Changes = 0 }
    }
}

# ── Fix all pages in one site ─────────────────────────────────────────────────
function Invoke-SiteFormatFix {
    param([string]$SPOSiteUrl, [string]$AppClientId, [string]$AppTenantId,
        [string]$SinglePage, [bool]$ProcessAll, [bool]$DryRun)

    $siteSlug = $SPOSiteUrl -replace "^https?://[^/]+/sites/", "" -replace "/.*", ""
    $spoHost = ([uri]$SPOSiteUrl).Host
    $tenantName = if ($AppTenantId) { $AppTenantId }
    else { ($spoHost -replace "\.sharepoint\.com$", "") + ".onmicrosoft.com" }

    Write-Log "[$siteSlug] Connecting (Tenant: $tenantName) ..."
    try {
        Connect-PnPOnline -Url $SPOSiteUrl -ClientId $AppClientId `
            -Tenant $tenantName -Interactive -ErrorAction Stop
        Write-Log "[$siteSlug] Connected." "SUCCESS"
    }
    catch {
        Write-Log "[$siteSlug] Connection failed: $_" "ERROR"
        return @([PSCustomObject]@{ Page = "N/A"; Status = "CONNECT_FAILED"; Changes = 0 })
    }

    $pageResults = @()

    if ($SinglePage) {
        $pageResults += Repair-SPOPage -SPOSiteUrl $SPOSiteUrl -SPOPageName $SinglePage -DryRun $DryRun
    }
    elseif ($ProcessAll) {
        Write-Log "[$siteSlug] Fetching all pages from SitePages ..."
        $pages = Get-PnPListItem -List "SitePages" -Fields "FileLeafRef" |
        Where-Object { $_["FileLeafRef"] -like "*.aspx" }
        Write-Log "[$siteSlug] Found $($pages.Count) page(s)."
        foreach ($p in $pages) {
            $pageResults += Repair-SPOPage -SPOSiteUrl $SPOSiteUrl `
                -SPOPageName $p["FileLeafRef"] -DryRun $DryRun
        }
    }

    Disconnect-PnPOnline
    return $pageResults
}

# =============================================================================
# ENTRY POINT — only executes when this script is run directly (not dot-sourced)
# =============================================================================
$isDotSourced = $MyInvocation.InvocationName -eq '.'
if ($isDotSourced) { return }   # loaded as a library — expose functions only

if ($SiteUrl) {
    if (-not $PageName -and -not $AllPages) {
        Write-Log "Provide -PageName <page.aspx> or -AllPages to process all pages." "ERROR"
        exit 1
    }

    $results = Invoke-SiteFormatFix `
        -SPOSiteUrl   $SiteUrl `
        -AppClientId  $ClientId `
        -AppTenantId  $TenantId `
        -SinglePage   $PageName `
        -ProcessAll   $AllPages.IsPresent `
        -DryRun       $WhatIf.IsPresent

    Write-Host "`n========================================" -ForegroundColor White
    Write-Host "  PAGE FORMATTING FIX — RESULTS"         -ForegroundColor White
    Write-Host "========================================" -ForegroundColor White
    $results | ForEach-Object {
        $color = switch -Wildcard ($_.Status) {
            "FIXED*" { "Green" }
            "DRYRUN*" { "Yellow" }
            "OK" { "Cyan" }
            default { "Red" }
        }
        Write-Host ("  {0,-45} [{1}]" -f $_.Page, $_.Status) -ForegroundColor $color
    }
    exit
}

# =============================================================================
# BULK MODE — reads SiteMapping.csv, processes in parallel
# =============================================================================
if (-not $CsvPath) {
    Write-Log "Provide -SiteUrl (single site) or -CsvPath (bulk)." "ERROR"
    exit 1
}

$csvFull = Resolve-Path $CsvPath -ErrorAction Stop
$sites = Import-Csv -Path $csvFull
$logDir = Join-Path $PSScriptRoot "NavFixLogs"
$runStamp = Get-Date -Format "yyyyMMdd_HHmmss"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
$summaryPath = Join-Path $logDir "PageFormattingFix_Summary_$runStamp.csv"

Write-Log "Loaded $($sites.Count) site(s) — processing with ThrottleLimit=$ThrottleLimit ..."

$allResults = $sites | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
    $site = $_
    $scriptPath = $using:PSCommandPath
    $clientIdLocal = $using:ClientId
    $tenantIdLocal = $using:TenantId
    $pageLocal = $using:PageName
    $allPagesLocal = $using:AllPages
    $isDryRun = $using:WhatIf

    . $scriptPath

    Invoke-SiteFormatFix `
        -SPOSiteUrl  $site.SPOSiteUrl `
        -AppClientId $clientIdLocal `
        -AppTenantId $tenantIdLocal `
        -SinglePage  $pageLocal `
        -ProcessAll  ([bool]$allPagesLocal) `
        -DryRun      ([bool]$isDryRun)
}

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host "`n========================================" -ForegroundColor White
Write-Host "  PAGE FORMATTING FIX — SUMMARY"         -ForegroundColor White
Write-Host "========================================" -ForegroundColor White

$allResults | ForEach-Object {
    $color = switch -Wildcard ($_.Status) {
        "FIXED*" { "Green" }
        "DRYRUN*" { "Yellow" }
        "OK" { "Cyan" }
        default { "Red" }
    }
    Write-Host ("  {0,-45} [{1}]" -f $_.Page, $_.Status) -ForegroundColor $color
}

$allResults | Export-Csv -Path $summaryPath -NoTypeInformation
$fixed = @($allResults | Where-Object { $_.Status -like "FIXED*" }).Count
$skipped = @($allResults | Where-Object { $_.Status -eq "OK" }).Count
$failed = @($allResults | Where-Object { $_.Status -notlike "FIXED*" -and $_.Status -ne "OK" -and $_.Status -notlike "DRYRUN*" }).Count
Write-Log "Fixed: $fixed  |  Already OK: $skipped  |  Issues: $failed" `
$(if ($failed -gt 0) { "WARN" } else { "SUCCESS" })
Write-Log "Summary saved to: $summaryPath"
