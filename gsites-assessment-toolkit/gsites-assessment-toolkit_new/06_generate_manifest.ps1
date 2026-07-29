<#
.SYNOPSIS
    Build a single consolidated per-site manifest from the assessment reports.

.DESCRIPTION
    Joins GSites_Inventory_Detailed.csv, GSites_Permissions.csv, Pages.csv,
    Embeds.csv, ExternalDomains.csv and Complexity_Report.csv (all keyed by
    the site's Drive file id) into one row per site, so a reviewer has a
    single file to scan instead of cross-referencing six reports.

    Works for both pipelines - just point -OutputDir / the *File params at
    the relevant folder and file names:
      - Main My-Drive pipeline (defaults below)
      - Standalone Shared Drive pipeline (05b), e.g.:
          .\06_generate_manifest.ps1 -PrimaryDomain "rocheua.com" `
              -InventoryFile 'GSites_SharedDrive_Inventory.csv' `
              -PermissionsFile 'GSites_SharedDrive_Permissions.csv' `
              -PagesFile 'GSites_SharedDrive_Pages.csv' `
              -EmbedsFile 'GSites_SharedDrive_Embeds.csv' `
              -ExternalDomainsFile 'GSites_SharedDrive_ExternalDomains.csv' `
              -ComplexityReportFile 'GSites_SharedDrive_Complexity_Report.csv' `
              -ManifestFile 'GSites_SharedDrive_Manifest.csv'

.PARAMETER OutputDir
    Folder containing the report CSVs. Defaults to .\output.

.PARAMETER PrimaryDomain
    Your primary domain. Grantees on this domain are labeled "internal";
    everything else (other domains, "anyone") is labeled "external".

.EXAMPLE
    .\06_generate_manifest.ps1 -PrimaryDomain "rocheua.com"
#>

param(
    [string]$OutputDir = "$PSScriptRoot\output",
    [string]$PrimaryDomain = '',

    [string]$InventoryFile = 'GSites_Inventory_Detailed.csv',
    [string]$PermissionsFile = 'GSites_Permissions.csv',
    [string]$PagesFile = 'Pages.csv',
    [string]$EmbedsFile = 'Embeds.csv',
    [string]$ExternalDomainsFile = 'ExternalDomains.csv',
    [string]$ComplexityReportFile = 'Complexity_Report.csv',
    [string]$ManifestFile = 'GSites_Manifest.csv'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-SafeProperty {
    param(
        [Parameter(Mandatory = $true)][psobject]$InputObject,
        [Parameter(Mandatory = $true)][string[]]$PropertyNames
    )
    foreach ($prop in $PropertyNames) {
        $p = $InputObject.psobject.Properties[$prop]
        if ($null -ne $p -and -not [string]::IsNullOrWhiteSpace($p.Value)) {
            return [string]$p.Value
        }
    }
    return $null
}

function Import-IfExists {
    param([string]$Path)
    if (Test-Path $Path) { return @(Import-Csv $Path) }
    return @()
}

$inventoryPath = Join-Path $OutputDir $InventoryFile
if (-not (Test-Path $inventoryPath)) {
    throw "Inventory file not found: $inventoryPath"
}

$sites = @(Import-Csv $inventoryPath)
$permissions = Import-IfExists (Join-Path $OutputDir $PermissionsFile)
$pages = Import-IfExists (Join-Path $OutputDir $PagesFile)
$embeds = Import-IfExists (Join-Path $OutputDir $EmbedsFile)
$externalDomains = Import-IfExists (Join-Path $OutputDir $ExternalDomainsFile)
$complexity = Import-IfExists (Join-Path $OutputDir $ComplexityReportFile)

$manifest = New-Object System.Collections.Generic.List[object]

foreach ($site in $sites) {
    $siteId = Get-SafeProperty -InputObject $site -PropertyNames @('id')
    $siteName = Get-SafeProperty -InputObject $site -PropertyNames @('name')
    $siteUrl = Get-SafeProperty -InputObject $site -PropertyNames @('webViewLink', 'webviewlink')

    $sitePerms = @($permissions | Where-Object { (Get-SafeProperty -InputObject $_ -PropertyNames @('id')) -eq $siteId })
    $sitePages = @($pages | Where-Object { (Get-SafeProperty -InputObject $_ -PropertyNames @('SiteId')) -eq $siteId })
    $siteEmbeds = @($embeds | Where-Object { (Get-SafeProperty -InputObject $_ -PropertyNames @('SiteId')) -eq $siteId })
    $siteDomains = @($externalDomains | Where-Object { (Get-SafeProperty -InputObject $_ -PropertyNames @('SiteId')) -eq $siteId })
    $siteScore = $complexity | Where-Object { (Get-SafeProperty -InputObject $_ -PropertyNames @('SiteId')) -eq $siteId } | Select-Object -First 1

    # Grantees: "type:identity:role", internal vs. external relative to PrimaryDomain
    $grantees = foreach ($perm in $sitePerms) {
        $type = Get-SafeProperty -InputObject $perm -PropertyNames @('permission.type')
        $role = Get-SafeProperty -InputObject $perm -PropertyNames @('permission.role')
        $identity = Get-SafeProperty -InputObject $perm -PropertyNames @('permission.emailAddress', 'permission.domain', 'permission.displayName')
        if (-not $identity) { $identity = $type }
        $scope = 'internal'
        if ($type -eq 'anyone') { $scope = 'external' }
        elseif ($identity -and $identity -match '@' -and $PrimaryDomain -and $identity -notmatch [regex]::Escape("@$PrimaryDomain")) { $scope = 'external' }
        elseif ($type -eq 'domain' -and $PrimaryDomain -and $identity -notmatch [regex]::Escape($PrimaryDomain)) { $scope = 'external' }
        "$type`:$identity`:$role`:$scope"
    }
    $grantees = @($grantees | Sort-Object -Unique)
    $externalGrantees = @($grantees | Where-Object { $_ -match ':external$' })

    $embedTypes = @($siteEmbeds | ForEach-Object { Get-SafeProperty -InputObject $_ -PropertyNames @('ArtifactType') } | Where-Object { $_ } | Sort-Object -Unique)
    $domainList = @($siteDomains | ForEach-Object { Get-SafeProperty -InputObject $_ -PropertyNames @('ExternalDomain') } | Where-Object { $_ } | Sort-Object -Unique)
    $crawlErrors = @($sitePages | Where-Object { (Get-SafeProperty -InputObject $_ -PropertyNames @('CrawlStatus')) -notlike 'Success*' }).Count

    $manifest.Add([pscustomobject]@{
            SiteId              = $siteId
            SiteName            = $siteName
            SiteUrl             = $siteUrl
            Owner               = Get-SafeProperty -InputObject $site -PropertyNames @('owners.emailAddress', 'Owner')
            OwnerDisplayName    = Get-SafeProperty -InputObject $site -PropertyNames @('owners.displayName')
            CreatedTime         = Get-SafeProperty -InputObject $site -PropertyNames @('createdTime')
            ModifiedTime        = Get-SafeProperty -InputObject $site -PropertyNames @('modifiedTime')
            SizeBytes           = Get-SafeProperty -InputObject $site -PropertyNames @('size')
            Shared              = Get-SafeProperty -InputObject $site -PropertyNames @('shared')
            PageCount           = $sitePages.Count
            CrawlErrorPages     = $crawlErrors
            EmbedCount          = $siteEmbeds.Count
            EmbedTypes          = ($embedTypes -join '; ')
            ExternalDomainCount = $domainList.Count
            ExternalDomains     = ($domainList -join '; ')
            PermissionRows      = $sitePerms.Count
            ExternalGranteeCount = $externalGrantees.Count
            Grantees            = ($grantees -join '; ')
            TotalScore          = if ($siteScore) { $siteScore.TotalScore } else { '' }
            Rating              = if ($siteScore) { $siteScore.Rating } else { '' }
            Recommendation      = if ($siteScore) { $siteScore.Recommendation } else { '' }
        })
}

$manifest | Export-Csv -NoTypeInformation -Path (Join-Path $OutputDir $ManifestFile)
Write-Host "[OK] Manifest generated: $ManifestFile ($($manifest.Count) sites)" -ForegroundColor Green
