<#
.SYNOPSIS
    Runs the enhanced Google Sites embed crawl using Playwright.

.DESCRIPTION
    Wrapper for 03_crawl_sites_enhanced.js. Ensures auth is fresh, then crawls
    the selected Google Sites for embedded content (iframes, YouTube, Maps, Drive, etc.)
    with shadow-DOM traversal, lazy-load scrolling, and subpage BFS navigation.

    Prerequisites (one-time setup in this folder):
      npm install
      npx playwright install chromium
      node Save-GoogleAuth.js   # sign in to Google; re-run if session expires

    Three ways to specify what to crawl (pick one):
      1. -SitesCsv  — Recommended. CSV with SiteUrl (required) and SiteName (optional).
                       Use SelectedSites.csv in this folder as a template.
      2. -SiteUrl   — Crawl a single site by URL.
      3. Neither    — Falls back to output\02_GSites_Inventory_Detailed.csv.

.PARAMETER SitesCsv
    Path to a CSV listing sites to crawl.
    Required columns : SiteUrl  (Google Sites URL, editor or published)
    Optional columns : SiteName (friendly label; defaults to the URL if omitted)

.PARAMETER SiteUrl
    Crawl a single Google Sites URL instead of a CSV.

.PARAMETER MaxPages
    Maximum subpages to crawl per site (default: 200).
    Lower this (e.g. 20) for a quick test run.

.PARAMETER DryRun
    Validates prerequisites and CSV without running the crawl.

.EXAMPLE
    # Crawl a hand-picked selection of sites
    .\Run-EnhancedCrawl.ps1 -SitesCsv ".\SelectedSites.csv"

.EXAMPLE
    # Quick test: crawl one site, limit to 10 pages
    .\Run-EnhancedCrawl.ps1 -SiteUrl "https://sites.google.com/d/<id>/p/<page>" -MaxPages 10

.EXAMPLE
    # Validate the CSV and prerequisites without actually crawling
    .\Run-EnhancedCrawl.ps1 -SitesCsv ".\SelectedSites.csv" -DryRun
#>
[CmdletBinding()]
param(
    [string]$SiteUrl  = "",
    [string]$SitesCsv = "",
    [int]   $MaxPages = 200,
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"
$EnhancedScript = Join-Path $PSScriptRoot "03_crawl_sites_enhanced.js"

if (-not (Test-Path $EnhancedScript)) {
    throw "Crawl script not found: $EnhancedScript"
}
if (-not (Get-Command "node" -ErrorAction SilentlyContinue)) {
    throw "Node.js is required. Install from https://nodejs.org/"
}
# Verify node_modules are installed locally
if (-not (Test-Path (Join-Path $PSScriptRoot "node_modules"))) {
    throw "node_modules not found. Run the following in this folder first:`n  npm install`n  npx playwright install chromium"
}

# Determine what to crawl
$crawlTarget = ""
if ($SiteUrl) {
    $crawlTarget = $SiteUrl
    Write-Host "Target: Single site URL -> $SiteUrl" -ForegroundColor Cyan
}
elseif ($SitesCsv) {
    # Resolve and validate the CSV before handing it to Node
    if (-not (Test-Path $SitesCsv)) {
        throw "SitesCsv not found: $SitesCsv"
    }
    $csvRows = Import-Csv -Path $SitesCsv
    if ($csvRows.Count -eq 0) {
        throw "SitesCsv is empty: $SitesCsv"
    }
    if (-not ($csvRows[0].PSObject.Properties.Name -contains 'SiteUrl')) {
        throw "SitesCsv must have a 'SiteUrl' column. Found: $($csvRows[0].PSObject.Properties.Name -join ', ')"
    }
    $validRows = $csvRows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.SiteUrl) }
    Write-Host "CSV validated: $($validRows.Count) site(s) found in $SitesCsv" -ForegroundColor Green
    if ($DryRun) {
        Write-Host "`nSites that would be crawled:" -ForegroundColor Cyan
        $validRows | Format-Table SiteUrl, SiteName -AutoSize
    }
    $crawlTarget = (Resolve-Path $SitesCsv).Path
    Write-Host "Target: Selected sites CSV -> $crawlTarget" -ForegroundColor Cyan
}
else {
    Write-Host "Target: Default inventory (gam7/output/02_GSites_Inventory_Detailed.csv)" -ForegroundColor Cyan
}

$authFile = Join-Path $PSScriptRoot ".auth\state.json"
if (-not (Test-Path $authFile)) {
    Write-Host "Auth state missing. Run this first:" -ForegroundColor Yellow
    Write-Host "  cd `"$PSScriptRoot`"" -ForegroundColor Cyan
    Write-Host "  node Save-GoogleAuth.js" -ForegroundColor Cyan
    return
}

$authAge = (Get-Date) - (Get-Item $authFile).LastWriteTime
Write-Host "Auth state age: $($authAge.TotalHours.ToString('F1')) hours" -ForegroundColor Cyan
if ($authAge.TotalHours -gt 24) {
    Write-Warning "Auth state is older than 24 hours. Google session may have expired."
    Write-Host "Re-run: node Save-GoogleAuth.js (in $PSScriptRoot)" -ForegroundColor Yellow
}

if ($DryRun) {
    Write-Host "DRY RUN - prerequisites OK. Ready to crawl." -ForegroundColor Green
    return
}

Write-Host "`n=== Running Enhanced Google Sites Crawl ===" -ForegroundColor Cyan
Push-Location $PSScriptRoot
& node "03_crawl_sites_enhanced.js" $crawlTarget $MaxPages
$exitCode = $LASTEXITCODE
Pop-Location

if ($exitCode -ne 0) {
    Write-Warning "Crawl exited with code $exitCode — check errors above. Partial results may still have been saved."
}

Write-Host "`n=== Results ===" -ForegroundColor Green
$outputDir  = Join-Path $PSScriptRoot "output"
$embedsFile = Join-Path $outputDir "08_Embeds_Enhanced.csv"
$pagesFile  = Join-Path $outputDir "07_Pages_Enhanced.csv"
Write-Host "Output folder : $outputDir" -ForegroundColor Cyan
Write-Host "Embeds report : $embedsFile" -ForegroundColor Cyan
Write-Host "Pages report  : $pagesFile"  -ForegroundColor Cyan

if (Test-Path $embedsFile) {
    $rows = Import-Csv $embedsFile
    Write-Host "Embeds found  : $($rows.Count)" -ForegroundColor $(if ($rows.Count -gt 0) { 'Green' } else { 'Yellow' })
    if ($rows.Count -gt 0) {
        $rows | Format-Table PageTitle, ItemKind, ArtifactType, ArtifactUrl -AutoSize
    }
    else {
        Write-Warning "No embeds detected. Possible causes:`n  - Google auth expired (re-run: node Save-GoogleAuth.js)`n  - Sites are private / restricted`n  - Embeds use unsupported formats (check html snapshots in $outputDir\html)"
    }
}
else {
    Write-Warning "Embeds CSV not found at: $embedsFile`nThe crawl may have failed before writing any output. Check the errors above."
}
