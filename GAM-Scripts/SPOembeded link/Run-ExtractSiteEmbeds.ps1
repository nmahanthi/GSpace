<#
.SYNOPSIS
    Extracts embedded content (YouTube, Maps, etc.) from a single Google Sites page.

.DESCRIPTION
    Copies the Playwright extractor into the gam7 folder (where node_modules live),
    launches a visible browser for sign-in if needed, scrolls the page to trigger
    lazy-loaded embeds, and saves results to CSV.

.PARAMETER SiteUrl
    The Google Sites URL to inspect (can be /edit or public view).

.PARAMETER Gam7Path
    Path to your gam7 folder containing Playwright and node_modules.

.PARAMETER OutputCsv
    Output CSV path. Defaults to ExtractedEmbeds.csv in the current directory.

.EXAMPLE
    .\Run-ExtractSiteEmbeds.ps1 -SiteUrl "https://sites.google.com/d/19JMNQ5kCG9VAAaWjrW65uH_G5--5r3un/p/1TQ4yy3DBedUJ43SP-ZYz1DubSv50Iu7e/edit"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$SiteUrl,
    [string]$Gam7Path = "C:\Users\v-nmahanthi\OneDrive - Microsoft\Documents\gam7",
    [string]$OutputCsv = "ExtractedEmbeds.csv"
)

$ErrorActionPreference = "Stop"
$Gam7Path = Resolve-Path $Gam7Path
$ScriptPath = Join-Path $PSScriptRoot "Extract-SiteEmbeds-Playwright.js"

if (-not (Test-Path $ScriptPath)) {
    throw "Extractor script not found: $ScriptPath"
}
if (-not (Get-Command "node" -ErrorAction SilentlyContinue)) {
    throw "Node.js is required."
}

$dest = Join-Path $Gam7Path "Extract-SiteEmbeds-Playwright.js"
Copy-Item -Path $ScriptPath -Destination $dest -Force

Push-Location $Gam7Path
Write-Host "Extracting embeds from: $SiteUrl" -ForegroundColor Cyan
Write-Host "A Chromium browser will open. Sign in to Google if prompted, then wait for the script to finish." -ForegroundColor Yellow
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
