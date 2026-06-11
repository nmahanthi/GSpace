<#
.SYNOPSIS
    End-to-end orchestrator: Google Sites auth/crawl -> extract embeds -> build mapping -> add to SPO.

.PARAMETER SPOUrl
    Target SharePoint Online site URL (e.g. https://tenant.sharepoint.com/sites/mysite).

.PARAMETER Gam7Path
    Path to your local gam7 folder containing 02_save_playwright_auth.js and 03_crawl_sites.js.

.PARAMETER ClientId
    Azure AD App Registration ClientId for PnP PowerShell auth.

.PARAMETER TenantId
    Optional tenant name. If omitted, derived from SPO URL.

.PARAMETER SkipGoogleAuth
    Skip re-running the Google auth step if state.json is already fresh.

.PARAMETER SkipCrawl
    Skip the crawl; use existing gam7/output CSVs.

.PARAMETER DryRun
    Preview only; do not write changes to SPO.

.EXAMPLE
    .\Invoke-FullEmbedRemediation.ps1 -SPOUrl "https://mngenvmcap908272.sharepoint.com/sites/ftcmigrationtestsite"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$SPOUrl,
    [string]$Gam7Path = "C:\Users\v-nmahanthi\OneDrive - Microsoft\Documents\gam7",
    [string]$ClientId = "3834b2e7-ab80-45fc-b4c8-ed5c960076b7",
    [string]$TenantId = "",
    [switch]$SkipGoogleAuth,
    [switch]$SkipCrawl,
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"
$Gam7Path = Resolve-Path $Gam7Path
$OutputDir = Join-Path $Gam7Path "output"
$MappingCsv = Join-Path $PSScriptRoot "EmbedMapping.csv"

function Test-Prereq {
    param([string]$Cmd, [string]$Name)
    if (-not (Get-Command $Cmd -ErrorAction SilentlyContinue)) { throw "$Name is required but not found in PATH." }
}

Write-Host "`n=== PREREQUISITE CHECKS ===" -ForegroundColor Cyan
Test-Prereq "node" "Node.js"
Test-Prereq "npm" "npm"
if (-not (Get-Module -ListAvailable -Name "PnP.PowerShell")) {
    throw "PnP.PowerShell is required. Install: Install-Module PnP.PowerShell -Force"
}
Import-Module PnP.PowerShell
$spoHost = ([uri]$SPOUrl).Host
$tenantName = if ($TenantId) { $TenantId } else { ($spoHost -replace "\.sharepoint\.com$","") + ".onmicrosoft.com" }
Write-Host "SPO: $SPOUrl | Tenant: $tenantName | Gam7: $Gam7Path" -ForegroundColor Green

# ── STEP 1: Google Auth ─────────────────────────────────────────────────────
if (-not $SkipGoogleAuth) {
    Write-Host "`n=== STEP 1: Google Sites Browser Auth ===" -ForegroundColor Cyan
    Write-Host "A Chromium browser window will open. Sign in to Google with the account that can access the target Sites." -ForegroundColor Yellow
    Write-Host "After sign-in completes and a site loads, return here and press Enter to save the auth state." -ForegroundColor Yellow
    Push-Location $Gam7Path
    & node "02_save_playwright_auth.js"
    if ($LASTEXITCODE -ne 0) { throw "Auth script failed." }
    Pop-Location
    Write-Host "Auth state saved." -ForegroundColor Green
}

# ── STEP 2: Crawl Google Sites ────────────────────────────────────────────
if (-not $SkipCrawl) {
    Write-Host "`n=== STEP 2: Crawling Google Sites ===" -ForegroundColor Cyan
    Push-Location $Gam7Path
    & node "03_crawl_sites.js"
    if ($LASTEXITCODE -ne 0) { throw "Crawl script failed." }
    Pop-Location
    Write-Host "Crawl complete. Outputs in $OutputDir" -ForegroundColor Green
}

# ── STEP 3: Extract Embeds from HTML snapshots ────────────────────────────
Write-Host "`n=== STEP 3: Extracting embeds from crawl snapshots ===" -ForegroundColor Cyan
$HtmlDir = Join-Path $OutputDir "html"
$EmbedsCsv = Join-Path $OutputDir "08_Embeds.csv"
$PagesCsv = Join-Path $OutputDir "07_Pages.csv"

if (-not (Test-Path $EmbedsCsv)) { throw "Embeds CSV not found: $EmbedsCsv. Re-run crawl." }

# Load crawl outputs
$embedRows = Import-Csv $EmbedsCsv
$pagesRows = if (Test-Path $PagesCsv) { Import-Csv $PagesCsv } else { @() }

# Filter real content embeds (exclude auth frames, links, images)
$realEmbeds = $embedRows | Where-Object {
    ($_.ItemKind -in @('iframe','embed')) -and
    ($_.ArtifactUrl -notmatch 'accounts\.google\.com|bscframe|recaptcha|google\.com/signin')
}

Write-Host "Found $($realEmbeds.Count) content embed(s) across $($pagesRows.Count) page(s)." -ForegroundColor Green

if ($realEmbeds.Count -eq 0) {
    Write-Warning "No content embeds discovered. If the crawl only shows auth pages, the Google session may still be expired or the Sites may be restricted."
    return
}

# ── STEP 4: Get SPO pages ──────────────────────────────────────────────────
Write-Host "`n=== STEP 4: Reading SPO Site Pages ===" -ForegroundColor Cyan
Write-Host "A browser window may open for Microsoft sign-in. Please authenticate if prompted." -ForegroundColor Yellow
Connect-PnPOnline -Url $SPOUrl -ClientId $ClientId -Tenant $tenantName -Interactive -ErrorAction Stop
$spoPages = Get-PnPListItem -List "SitePages" -PageSize 500 | ForEach-Object {
    [pscustomobject]@{
        FileName = $_.FieldValues["FileLeafRef"]
        Title    = $_.FieldValues["Title"]
    }
}
Write-Host "Retrieved $($spoPages.Count) SPO page(s)." -ForegroundColor Green

# ── STEP 5: Build Mapping CSV ────────────────────────────────────────────────
Write-Host "`n=== STEP 5: Building Embed Mapping CSV ===" -ForegroundColor Cyan
$mapping = [System.Collections.Generic.List[pscustomobject]]::new()
$orderTracker = @{}

foreach ($row in $realEmbeds) {
    $gsTitle = ($row.PageTitle -replace '\s+',' ').Trim()
    if ($gsTitle -match '^Error\s+\d') { $gsTitle = ($pagesRows | Where-Object { $_.PageUrl -eq $row.PageUrl } | Select-Object -First 1).PageTitle }

    # Match GS page title to SPO page (by Title or by FileName stem)
    $match = $spoPages | Where-Object {
        ($_.Title -and $_.Title.Trim() -eq $gsTitle) -or
        ($_.FileName -replace '\.aspx$','' -replace '[-_]',' ').Trim() -eq ($gsTitle -replace '[-_]',' ').Trim()
    } | Select-Object -First 1

    if (-not $match) {
        # Try case-insensitive contains match as fallback
        $match = $spoPages | Where-Object {
            $_.Title -and ($gsTitle -like "*$($_.Title.Trim())*" -or $_.Title.Trim() -like "*$gsTitle*")
        } | Select-Object -First 1
    }

    if (-not $match) {
        Write-Warning "No SPO page match for Google Sites page '$gsTitle'. Skipping embed: $($row.ArtifactUrl)"
        continue
    }

    $key = $match.FileName
    if (-not $orderTracker.ContainsKey($key)) { $orderTracker[$key] = 0 } else { $orderTracker[$key] += 1 }

    $mapping.Add([pscustomobject]@{
        PageName      = $match.FileName
        EmbedUrl      = $row.ArtifactUrl
        SectionIndex  = 1
        ColumnIndex   = 1
        Order         = $orderTracker[$key]
        GSitePageTitle= $gsTitle
        ArtifactType  = $row.ArtifactType
    })
}

if ($mapping.Count -eq 0) { throw "No embeds could be mapped to SPO pages. Check page titles match between Google Sites and SPO." }

$mapping | Export-Csv -Path $MappingCsv -NoTypeInformation
Write-Host "Mapping written to: $MappingCsv ($($mapping.Count) rows)" -ForegroundColor Green
$mapping | Format-Table -AutoSize

# ── STEP 6: Apply to SPO ───────────────────────────────────────────────────
Write-Host "`n=== STEP 6: Adding Embed Web Parts to SPO ===" -ForegroundColor Cyan
$addScript = Join-Path $PSScriptRoot "Add-SPOYouTubeWebParts.ps1"
if (-not (Test-Path $addScript)) { throw "Add-SPOYouTubeWebParts.ps1 not found in $PSScriptRoot" }

$invokeParams = @{
    SiteUrl     = $SPOUrl
    MappingCsv  = $MappingCsv
    ClientId    = $ClientId
    TenantId    = $tenantName
}
if ($DryRun) { $invokeParams['DryRun'] = $true }

& $addScript @invokeParams

Disconnect-PnPOnline
Write-Host "`nDone. Review $MappingCsv before re-running without -DryRun if needed." -ForegroundColor Cyan
