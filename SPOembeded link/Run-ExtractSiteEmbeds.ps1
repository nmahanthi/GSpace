<#
.SYNOPSIS
    Extracts embedded content (YouTube, Maps, etc.) from a single Google Sites page.

.DESCRIPTION
    Uses Playwright to load the page with the saved Google session, scrolls to
    trigger lazy-loaded embeds, and saves all discovered embed URLs to a CSV.

    Prerequisites (one-time setup in this folder):
      npm install
      npx playwright install chromium
      node Save-GoogleAuth.js   # sign in to Google; re-run if session expires

.PARAMETER SiteUrl
    The Google Sites URL to inspect (editor or published view).

.PARAMETER OutputCsv
    Output CSV path. Defaults to ExtractedEmbeds.csv in the current directory.

.EXAMPLE
    .\Run-ExtractSiteEmbeds.ps1 -SiteUrl "https://sites.google.com/d/19JMNQ5.../p/1TQ4yy.../edit"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$SiteUrl,
    [string]$OutputCsv = "ExtractedEmbeds.csv"
)

$ErrorActionPreference = "Stop"
$ScriptPath = Join-Path $PSScriptRoot "Extract-SiteEmbeds-Playwright.js"

if (-not (Test-Path $ScriptPath)) {
    throw "Extractor script not found: $ScriptPath"
}
if (-not (Get-Command "node" -ErrorAction SilentlyContinue)) {
    throw "Node.js is required. Install from https://nodejs.org/"
}
if (-not (Test-Path (Join-Path $PSScriptRoot "node_modules"))) {
    throw "node_modules not found. Run the following in this folder first:`n  npm install`n  npx playwright install chromium"
}

Push-Location $PSScriptRoot
Write-Host "Extracting embeds from: $SiteUrl" -ForegroundColor Cyan
& node "Extract-SiteEmbeds-Playwright.js" "$SiteUrl" "$OutputCsv"
$exit = $LASTEXITCODE
Pop-Location

if ($exit -ne 0) { throw "Extraction failed with exit code $exit" }

Write-Host "`nExtraction complete. Results:" -ForegroundColor Green
if (Test-Path $OutputCsv) {
    Import-Csv $OutputCsv | Format-Table -AutoSize
} else {
    Write-Warning "Output CSV not found."
}
