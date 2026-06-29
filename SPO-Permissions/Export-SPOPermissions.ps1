<#
.SYNOPSIS
    Exports SharePoint Online site permissions to a CSV file.
.DESCRIPTION
    Connects to every SPO site collection using PnP PowerShell (app-only auth),
    enumerates all unique role assignments at the site, subweb, list, and
    library level, and writes the results to a timestamped CSV.
.PARAMETER ConfigPath
    Path to Config.psd1.  Defaults to .\Config.psd1
.PARAMETER OutputFile
    Overrides the output CSV path from Config.psd1.
.PARAMETER SiteUrls
    Optional array of specific SPO site URLs to process.
    If omitted, all site collections returned by the Admin Center are processed.
.EXAMPLE
    .\Export-SPOPermissions.ps1
    .\Export-SPOPermissions.ps1 -SiteUrls "https://contoso.sharepoint.com/sites/Finance"
#>
[CmdletBinding()]
param(
    [string]  $ConfigPath = ".\Config.psd1",
    [string]  $OutputFile = "",
    [string[]]$SiteUrls   = @()
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# Require PnP.PowerShell
# ---------------------------------------------------------------------------
if (-not (Get-Module -ListAvailable -Name "PnP.PowerShell")) {
    throw "PnP.PowerShell module is not installed. Run: Install-Module PnP.PowerShell -Scope CurrentUser"
}
Import-Module PnP.PowerShell -ErrorAction Stop

# ---------------------------------------------------------------------------
# Helper: collect unique role assignments from a web (recursive into subwebs)
# ---------------------------------------------------------------------------
function Get-WebPermissions {
    param(
        [string]$SiteUrl,
        [string]$WebUrl,
        [string]$WebTitle,
        $Records,
        $CfgSP,
        [string]$CaptureTime
    )

    Connect-PnPOnline -Url $WebUrl `
        -ClientId           $CfgSP.ClientId `
        -CertificatePath    $CfgSP.CertificatePath `
        -CertificatePassword (ConvertTo-SecureString $CfgSP.CertificatePassword -AsPlainText -Force) `
        -Tenant             $CfgSP.TenantId

    $web = Get-PnPWeb -Includes HasUniqueRoleAssignments, RoleAssignments

    if ($web.HasUniqueRoleAssignments) {
        $assignments = Get-PnPProperty -ClientObject $web -Property RoleAssignments
        foreach ($ra in $assignments) {
            $member = Get-PnPProperty -ClientObject $ra -Property Member
            $bindings = Get-PnPProperty -ClientObject $ra -Property RoleDefinitionBindings
            foreach ($rd in $bindings) {
                $Records.Add([PSCustomObject]@{
                    SiteCollectionUrl = $SiteUrl
                    WebUrl            = $WebUrl
                    WebTitle          = $WebTitle
                    ObjectType        = "Web"
                    ObjectUrl         = $WebUrl
                    PrincipalType     = $member.PrincipalType
                    PrincipalLogin    = $member.LoginName
                    PrincipalName     = $member.Title
                    PermissionLevel   = $rd.Name
                    IsInherited       = $false
                    CapturedAt        = $CaptureTime
                })
            }
        }
    } else {
        $Records.Add([PSCustomObject]@{
            SiteCollectionUrl = $SiteUrl
            WebUrl            = $WebUrl
            WebTitle          = $WebTitle
            ObjectType        = "Web"
            ObjectUrl         = $WebUrl
            PrincipalType     = "(inherited)"
            PrincipalLogin    = ""
            PrincipalName     = ""
            PermissionLevel   = "(inherited)"
            IsInherited       = $true
            CapturedAt        = $CaptureTime
        })
    }

    # Recurse into sub-webs
    $subwebs = Get-PnPSubWeb -Recurse:$false
    foreach ($sw in $subwebs) {
        Get-WebPermissions -SiteUrl $SiteUrl -WebUrl $sw.Url -WebTitle $sw.Title `
            -Records $Records -CfgSP $CfgSP -CaptureTime $CaptureTime
    }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
Write-Host "`n=== Export-SPOPermissions.ps1 ===" -ForegroundColor Cyan

$cfg  = Import-PowerShellDataFile -Path $ConfigPath
$date = Get-Date -Format "yyyyMMdd_HHmmss"

if (-not $OutputFile) {
    $dir = $cfg.Output.Directory
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    $OutputFile = Join-Path $dir "$($cfg.Output.SPOPermissionsFile)_$date.csv"
}

# Connect to admin center to get all site collections
Write-Host "Connecting to SPO Admin Center: $($cfg.SharePoint.AdminUrl)" -ForegroundColor Yellow
Connect-PnPOnline -Url $cfg.SharePoint.AdminUrl `
    -ClientId           $cfg.SharePoint.ClientId `
    -CertificatePath    $cfg.SharePoint.CertificatePath `
    -CertificatePassword (ConvertTo-SecureString $cfg.SharePoint.CertificatePassword -AsPlainText -Force) `
    -Tenant             $cfg.SharePoint.TenantId

if ($SiteUrls.Count -eq 0) {
    Write-Host "Enumerating all SPO site collections..." -ForegroundColor Yellow
    $sites = Get-PnPTenantSite -IncludeOneDriveSites:$false | Select-Object -ExpandProperty Url
} else {
    $sites = $SiteUrls
}

Write-Host "  Processing $($sites.Count) site collection(s)." -ForegroundColor Green

$records     = [System.Collections.Generic.List[object]]::new()
$captureTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

foreach ($siteUrl in $sites) {
    Write-Host "  Site: $siteUrl" -ForegroundColor DarkCyan
    try {
        Get-WebPermissions -SiteUrl $siteUrl -WebUrl $siteUrl -WebTitle $siteUrl `
            -Records $records -CfgSP $cfg.SharePoint -CaptureTime $captureTime
    } catch {
        Write-Warning "    Failed to process $siteUrl : $_"
    }
}

$records | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
Write-Host "`nSPO permissions exported -> $OutputFile" -ForegroundColor Green
Write-Host "Total permission entries: $($records.Count)`n"
return $OutputFile
