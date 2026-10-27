<#
.SYNOPSIS
    Builds the CUSTOMER edition of the licence review app: a staged folder and a ZIP.
.DESCRIPTION
    Copies only what a customer needs. It deliberately never copies partner-terms.csv (your fees and margins),
    and fails if that file ends up in the package. The installer (installer.iss) is built from the staged folder.
.EXAMPLE
    ./Build-CustomerPackage.ps1 -Version 1.0.0
#>
[CmdletBinding()]
param(
    [string]$Version = '1.0.0',
    [switch]$BundleGraphModule,   # copy Microsoft.Graph.Authentication into the package so install and first run need no internet (needed for Microsoft Store)
    [string]$OutDir = (Join-Path $PSScriptRoot 'dist')
)
$ErrorActionPreference = 'Stop'
$src = Split-Path $PSScriptRoot -Parent
$stage = Join-Path $OutDir 'M365LicenseReview'
if (Test-Path $OutDir) { Remove-Item $OutDir -Recurse -Force }
New-Item -ItemType Directory -Path $stage -Force | Out-Null

$include = 'Get-M365LicenseSavings.ps1', 'Start-M365LicenseGUI.ps1', 'Start-M365LicenseGUI.cmd', 'prices.csv', 'customer-prices.csv'
foreach ($f in $include) {
    $p = Join-Path $src $f
    if (-not (Test-Path $p)) { throw "Missing $p" }
    Copy-Item $p $stage
}
if ($BundleGraphModule) {
    Write-Output 'Bundling Microsoft.Graph.Authentication ...'
    New-Item -ItemType Directory -Path (Join-Path $stage 'Modules') -Force | Out-Null
    try { Save-Module -Name Microsoft.Graph.Authentication -Path (Join-Path $stage 'Modules') -Repository PSGallery -ErrorAction Stop }
    catch {
        # Gallery unreachable: copy the newest installed copy instead
        $m = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $m) { throw "Could not download or find Microsoft.Graph.Authentication: $($_.Exception.Message)" }
        Write-Warning "Gallery download failed; bundling installed version $($m.Version) from $($m.ModuleBase)"
        $dest = Join-Path $stage "Modules/Microsoft.Graph.Authentication/$($m.Version)"
        New-Item -ItemType Directory -Path $dest -Force | Out-Null
        Copy-Item (Join-Path $m.ModuleBase '*') $dest -Recurse -Force
    }
    if (-not (Get-ChildItem (Join-Path $stage 'Modules') -Recurse -Filter 'Microsoft.Graph.Authentication.psd1')) { throw 'Graph module was not bundled.' }
}
Set-Content -Path (Join-Path $stage 'version.txt') -Value $Version -Encoding ASCII

# Safety checks: nothing internal may ship
$banned = Get-ChildItem $stage -Recurse -File | Where-Object { $_.FullName -notmatch '[\\/]Modules[\\/]' } | Where-Object { $_.Name -match 'partner|\.json$|output' }
if ($banned) { throw "Internal file(s) in package: $($banned.Name -join ', ')" }

$zip = Join-Path $OutDir "M365LicenseReview-$Version.zip"
Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $zip
Write-Output "Staged: $stage"
Write-Output "ZIP:    $zip"
Get-ChildItem $stage | Select-Object Name, Length | Format-Table -AutoSize | Out-String | Write-Output
