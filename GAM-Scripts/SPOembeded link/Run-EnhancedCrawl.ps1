<#
.SYNOPSIS
    Runs the enhanced Google Sites embed crawl using Playwright.

.DESCRIPTION
    Wrapper for 03_crawl_sites_enhanced.js. Ensures auth is fresh, then runs
    the crawl with shadow-DOM traversal, lazy-load scrolling, and deeper
    embed detection (data-src, srcdoc, YouTube patterns).

.PARAMETER Gam7Path
    Path to your gam7 folder containing .auth/state.json.

.PARAMETER DryRun
    If specified, only validates prerequisites without running.

.EXAMPLE
    .\Run-EnhancedCrawl.ps1
#>
[CmdletBinding()]
param(
    [string]$SiteUrl = "",
    [string]$SitesCsv = "",
    [string]$Gam7Path = "C:\Users\v-nmahanthi\OneDrive - Microsoft\Documents\gam7",
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"
$Gam7Path = Resolve-Path $Gam7Path
$EnhancedScript = Join-Path $PSScriptRoot "03_crawl_sites_enhanced.js"

if (-not (Test-Path $EnhancedScript)) {
    throw "Enhanced crawl script not found: $EnhancedScript"
}
if (-not (Get-Command "node" -ErrorAction SilentlyContinue)) {
    throw "Node.js is required. Install from https://nodejs.org/"
}

# Determine what to crawl
$crawlTarget = ""
if ($SiteUrl) {
    $crawlTarget = $SiteUrl
    Write-Host "Target: Direct site URL -> $SiteUrl" -ForegroundColor Cyan
}
elseif ($SitesCsv) {
    $crawlTarget = $SitesCsv
    Write-Host "Target: Custom inventory CSV -> $SitesCsv" -ForegroundColor Cyan
}
else {
    Write-Host "Target: Default inventory (gam7/output/02_GSites_Inventory_Detailed.csv)" -ForegroundColor Cyan
}

$authFile = Join-Path $Gam7Path ".auth\state.json"
if (-not (Test-Path $authFile)) {
    Write-Host "Auth state missing. Run this first:" -ForegroundColor Yellow
    Write-Host "  cd `"$Gam7Path`"" -ForegroundColor Cyan
    Write-Host "  node 02_save_playwright_auth.js" -ForegroundColor Cyan
    Write-Host "Then sign in to Google in the browser window and press Enter." -ForegroundColor Yellow
    return
}

$authAge = (Get-Date) - (Get-Item $authFile).LastWriteTime
Write-Host "Auth state age: $($authAge.TotalHours.ToString('F1')) hours" -ForegroundColor Cyan
if ($authAge.TotalHours -gt 24) {
    Write-Warning "Auth state is older than 24 hours. Google session may have expired."
    Write-Host "Re-run: node 02_save_playwright_auth.js (in $Gam7Path)" -ForegroundColor Yellow
}

if ($DryRun) {
    Write-Host "DRY RUN - prerequisites OK. Ready to crawl." -ForegroundColor Green
    return
}

Write-Host "`n=== Running Enhanced Google Sites Crawl ===" -ForegroundColor Cyan
$destScript = Join-Path $Gam7Path "03_crawl_sites_enhanced.js"
Copy-Item -Path $EnhancedScript -Destination $destScript -Force
Push-Location $Gam7Path
& node "03_crawl_sites_enhanced.js" $crawlTarget
$exitCode = $LASTEXITCODE
Pop-Location

if ($exitCode -ne 0) {
    throw "Enhanced crawl failed with exit code $exitCode"
}

Write-Host "`n=== Results ===" -ForegroundColor Green
$outputDir = Join-Path $Gam7Path "output"
$embedsFile = Join-Path $outputDir "08_Embeds_Enhanced.csv"
if (Test-Path $embedsFile) {
    $rows = Import-Csv $embedsFile
    Write-Host "Embeds found: $($rows.Count)" -ForegroundColor Cyan
    if ($rows.Count -gt 0) {
        $rows | Format-Table PageTitle, ItemKind, ArtifactType, ArtifactUrl -AutoSize
    }
    else {
        Write-Warning "No embeds detected. Possible causes:`n  - Google auth expired`n  - Sites are private / restricted`n  - Embeds use unsupported formats (see html snapshots in $outputDir\html)"
    }
}
else {
    Write-Warning "Embeds CSV not generated. Check console output above."
}
