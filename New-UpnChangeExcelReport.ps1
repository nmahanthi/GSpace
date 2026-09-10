<#
.SYNOPSIS
    Turns a UpnChangeReport CSV (from Update-UserUpn.ps1) into a formatted
    Excel workbook, enriched with each user's OneDrive data size.

.DESCRIPTION
    Reads the UPN-change report CSV, splits the Timestamp column into
    separate Date/Time columns, and (optionally) looks up each user's
    OneDrive storage usage via PnP.PowerShell so the report shows how much
    data is associated with each migrated account.

    Produces a two-sheet .xlsx:
      - "UPN Change Report": one row per user, colour-coded by Status,
        with Date, Time, and DataSizeGB columns, as an Excel table with
        autofilter and frozen header row.
      - "Summary": counts by status, total users processed, and total
        OneDrive data size across all rows, plus the report generation
        timestamp.

    Requires the ImportExcel module (already used elsewhere in this repo,
    e.g. M365/SPO-Migration-Assessment.ps1):
        Install-Module ImportExcel -Scope CurrentUser

    OneDrive size lookup requires PnP.PowerShell and a SharePoint admin
    connection. If -AdminUrl/-ClientId are not supplied, the report is
    still generated, just without the DataSizeGB column populated.

.PARAMETER UpnChangeReportCsv
    Path to the UpnChangeReport_*.csv produced by Update-UserUpn.ps1.

.PARAMETER OutputXlsx
    Path for the output workbook. Defaults to
    .\UpnChangeReport_<timestamp>.xlsx alongside the input CSV.

.PARAMETER AdminUrl
    SharePoint admin center URL, e.g. https://kaaratec-admin.sharepoint.com.
    Required to populate the DataSizeGB column.

.PARAMETER ClientId
    Entra ID app registration (client) ID for PnP -Interactive login
    (required since PnP dropped its shared multi-tenant app in Sept 2024).
    Register one with:
        Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP.PowerShell" -Tenant yourtenant.onmicrosoft.com

.EXAMPLE
    .\New-UpnChangeExcelReport.ps1 -UpnChangeReportCsv .\UpnChangeReport_20260910_101500.csv `
        -AdminUrl https://kaaratec-admin.sharepoint.com -ClientId <appId>

.EXAMPLE
    # Without OneDrive size lookup
    .\New-UpnChangeExcelReport.ps1 -UpnChangeReportCsv .\UpnChangeReport_20260910_101500.csv
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$UpnChangeReportCsv,

    [Parameter(Mandatory = $false)]
    [string]$OutputXlsx,

    [Parameter(Mandatory = $false)]
    [string]$AdminUrl,

    [Parameter(Mandatory = $false)]
    [string]$ClientId
)

$ErrorActionPreference = 'Stop'

function Write-Ok    { param([string]$Text) Write-Host "   OK   $Text" -ForegroundColor Green }
function Write-Warn2 { param([string]$Text) Write-Host "   WARN $Text" -ForegroundColor Yellow }
function Write-Step  { param([string]$Text) Write-Host "-> $Text" -ForegroundColor Yellow }

if (-not (Test-Path $UpnChangeReportCsv)) { throw "UpnChangeReportCsv not found: $UpnChangeReportCsv" }
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    throw "Required module 'ImportExcel' is not installed. Run: Install-Module ImportExcel -Scope CurrentUser"
}
Import-Module ImportExcel -ErrorAction Stop

if (-not $OutputXlsx) {
    $dir = Split-Path -Parent (Resolve-Path $UpnChangeReportCsv)
    $OutputXlsx = Join-Path $dir "UpnChangeReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
}

$rows = Import-Csv -Path $UpnChangeReportCsv
if ($rows.Count -eq 0) { throw "No rows found in $UpnChangeReportCsv" }

function Get-OneDrivePath {
    param([string]$Upn)
    if (-not $Upn) { return $null }
    return ($Upn.Replace('@', '_').Replace('.', '_'))
}

# --- Optional: OneDrive size lookup via PnP ---
$sizeByPath = @{}
$lookupEnabled = $false
if ($AdminUrl -and $ClientId) {
    Write-Step 'Connecting to SharePoint Online (PnP) for OneDrive size lookup'
    try {
        Connect-PnPOnline -Url $AdminUrl -Interactive -ClientId $ClientId
        Write-Ok 'Connected via PnP'
        Write-Step 'Retrieving OneDrive site sizes (this may take a while)'
        $odSites = Get-PnPTenantSite -IncludeOneDriveSites -Filter "Url -like '-my.sharepoint.com/personal/'" -Detailed
        foreach ($s in $odSites) {
            $path = ($s.Url -split '/personal/')[-1].TrimEnd('/')
            if ($path) { $sizeByPath[$path] = $s.StorageUsageCurrent }  # MB
        }
        Write-Ok "Retrieved sizes for $($sizeByPath.Count) OneDrive site(s)"
        $lookupEnabled = $true
    } catch {
        Write-Warn2 "OneDrive size lookup skipped - PnP connection/query failed: $($_.Exception.Message)"
    }
} else {
    Write-Warn2 'AdminUrl/ClientId not supplied - report will be generated without DataSizeGB.'
}

# --- Build enriched rows ---
Write-Step 'Building enriched report rows'
$enriched = foreach ($row in $rows) {
    $dt = $null
    if ($row.Timestamp) { [datetime]::TryParse($row.Timestamp, [ref]$dt) | Out-Null }

    $dataSizeGB = $null
    if ($lookupEnabled) {
        $upnForPath = if ($row.NewUserPrincipalName) { $row.NewUserPrincipalName } else { $row.OldUserPrincipalName }
        $path = Get-OneDrivePath -Upn $upnForPath
        if ($path -and $sizeByPath.ContainsKey($path)) {
            $dataSizeGB = [math]::Round($sizeByPath[$path] / 1024, 2)  # MB -> GB
        }
    }

    [PSCustomObject]@{
        Date                 = if ($dt) { $dt.ToString('yyyy-MM-dd') } else { '' }
        Time                 = if ($dt) { $dt.ToString('HH:mm:ss') } else { '' }
        DisplayName          = $row.DisplayName
        OldUserPrincipalName = $row.OldUserPrincipalName
        NewUserPrincipalName = $row.NewUserPrincipalName
        Status               = $row.Status
        DataSizeGB           = $dataSizeGB
        Detail               = $row.Detail
        Error                = $row.Error
        NotificationStatus   = $row.NotificationStatus
        NotificationDetail   = $row.NotificationDetail
    }
}

# --- Summary sheet data ---
$statusCounts = $enriched | Group-Object Status | Sort-Object Count -Descending
$totalSizeGB  = [math]::Round((($enriched | Where-Object { $_.DataSizeGB } | Measure-Object DataSizeGB -Sum).Sum), 2)

$summary = [ordered]@{
    'Report Generated'          = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    'Source Report'             = (Split-Path -Leaf $UpnChangeReportCsv)
    'Total Users in Report'     = $enriched.Count
}
foreach ($sc in $statusCounts) { $summary["Status: $($sc.Name)"] = $sc.Count }
$summary['Total OneDrive Data Size (GB)'] = if ($lookupEnabled) { $totalSizeGB } else { 'N/A (size lookup not run)' }

$summaryData = $summary.Keys | ForEach-Object { [PSCustomObject]@{ Metric = $_; Value = $summary[$_] } }

# --- Write workbook ---
Write-Step "Writing workbook: $OutputXlsx"
$xl = $summaryData | Export-Excel -Path $OutputXlsx -WorksheetName 'Summary' `
    -TableName 'Summary' -TableStyle Medium2 -AutoSize -FreezeTopRow -PassThru

$enriched | Export-Excel -ExcelPackage $xl -WorksheetName 'UPN Change Report' `
    -TableName 'UpnChanges' -TableStyle Medium6 -AutoSize -FreezeTopRow -AutoFilter

# Conditional formatting on the Status column
$ws       = $xl.Workbook.Worksheets['UPN Change Report']
$colNames = ($enriched | Select-Object -First 1).PSObject.Properties.Name
$statusIdx = [array]::IndexOf($colNames, 'Status') + 1
if ($statusIdx -gt 0 -and $enriched.Count -gt 0) {
    $endRow = $enriched.Count + 1
    $colLetter = [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($statusIdx)
    $range = "${colLetter}2:${colLetter}${endRow}"
    Add-ConditionalFormatting -WorkSheet $ws -Address $range -RuleType Equal -ConditionValue '"Success"' `
        -BackgroundColor ([System.Drawing.Color]::FromArgb(16,124,16)) -ForegroundColor ([System.Drawing.Color]::White)
    Add-ConditionalFormatting -WorkSheet $ws -Address $range -RuleType Equal -ConditionValue '"Failed"' `
        -BackgroundColor ([System.Drawing.Color]::FromArgb(209,52,56)) -ForegroundColor ([System.Drawing.Color]::White)
    Add-ConditionalFormatting -WorkSheet $ws -Address $range -RuleType Equal -ConditionValue '"Skipped"' `
        -BackgroundColor ([System.Drawing.Color]::FromArgb(255,140,0)) -ForegroundColor ([System.Drawing.Color]::Black)
    Add-ConditionalFormatting -WorkSheet $ws -Address $range -RuleType Equal -ConditionValue '"WhatIf"' `
        -BackgroundColor ([System.Drawing.Color]::FromArgb(128,128,128)) -ForegroundColor ([System.Drawing.Color]::White)
}

Close-ExcelPackage $xl
Write-Ok "Excel report saved: $OutputXlsx"

if ($AdminUrl -and $ClientId) {
    try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch { }
}

Write-Host ''
Write-Host '========================================' -ForegroundColor Cyan
Write-Host 'SUMMARY' -ForegroundColor Cyan
Write-Host '========================================' -ForegroundColor Cyan
$statusCounts | ForEach-Object { Write-Host ("  {0,-10} {1}" -f $_.Name, $_.Count) }
if ($lookupEnabled) { Write-Host "  Total OneDrive data: $totalSizeGB GB" }
Write-Host "Workbook: $OutputXlsx" -ForegroundColor Green
