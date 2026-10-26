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
    [string]$OutDir = (Join-Path $PSScriptRoot 'dist')
)
$ErrorActionPreference = 'Stop'
$src = Split-Path $PSScriptRoot -Parent
$stage = Join-Path $OutDir 'M365LicenseReview'
if (Test-Path $OutDir) { Remove-Item $OutDir -Recurse -Force }
New-Item -ItemType Directory -Path $stage -Force | Out-Null

$include = 'Get-M365LicenseSavings.ps1', 'Start-M365LicenseGUI.ps1', 'Start-M365LicenseGUI.cmd', 'prices.csv'
foreach ($f in $include) {
    $p = Join-Path $src $f
    if (-not (Test-Path $p)) { throw "Missing $p" }
    Copy-Item $p $stage
}
Set-Content -Path (Join-Path $stage 'version.txt') -Value $Version -Encoding ASCII

# Safety checks: nothing internal may ship
$banned = Get-ChildItem $stage -Recurse -File | Where-Object { $_.Name -match 'partner|\.json$|output' }
if ($banned) { throw "Internal file(s) in package: $($banned.Name -join ', ')" }

$zip = Join-Path $OutDir "M365LicenseReview-$Version.zip"
Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $zip
Write-Output "Staged: $stage"
Write-Output "ZIP:    $zip"
Get-ChildItem $stage | Select-Object Name, Length | Format-Table -AutoSize | Out-String | Write-Output
