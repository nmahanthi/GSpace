<#
.SYNOPSIS
    Lists all web parts on all modern pages in an SPO site to diagnose embed/content issues.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$SiteUrl,
    [string]$ClientId = "3834b2e7-ab80-45fc-b4c8-ed5c960076b7",
    [string]$TenantId = "",
    [string]$OutputCsv = "SPOWebParts.csv"
)

Import-Module PnP.PowerShell
$spoHost = ([uri]$SiteUrl).Host
$tenantName = if ($TenantId) { $TenantId } else { ($spoHost -replace "\.sharepoint\.com$", "") + ".onmicrosoft.com" }

Write-Host "Connecting to $SiteUrl..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $tenantName -Interactive -ErrorAction Stop

$Pages = Get-PnPListItem -List "SitePages" -PageSize 500
$AllParts = [System.Collections.Generic.List[object]]::new()

foreach ($pageItem in $Pages) {
    $fileName = $pageItem.FieldValues["FileLeafRef"]
    Write-Host "Inspecting: $fileName" -ForegroundColor Green
    try {
        $page = Get-PnPPage -Identity $fileName -ErrorAction Stop
        foreach ($wp in $page.Controls) {
            $props = $wp.PropertiesJson | ConvertFrom-Json -ErrorAction SilentlyContinue
            $AllParts.Add([pscustomobject]@{
                PageName     = $fileName
                WebPartType  = $wp.Type.Name
                Title        = $wp.Title
                Section      = $wp.Section.Order
                Column       = $wp.Column.Order
                Order        = $wp.Order
                Properties   = if ($wp.PropertiesJson.Length -gt 500) { $wp.PropertiesJson.Substring(0,500) + "..." } else { $wp.PropertiesJson }
            })
        }
    } catch {
        Write-Warning "Could not load page '$fileName': $_"
    }
}

Write-Host "Found $($AllParts.Count) web part(s) across $($Pages.Count) page(s)." -ForegroundColor Cyan
$AllParts | Export-Csv -Path $OutputCsv -NoTypeInformation
Write-Host "Exported to $OutputCsv" -ForegroundColor Cyan

Disconnect-PnPOnline
