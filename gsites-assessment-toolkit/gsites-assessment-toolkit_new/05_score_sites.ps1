param(
    [string]$OutputDir = "$PSScriptRoot\output",
    [string]$PrimaryDomain = ''   # Set your domain here, e.g. 'yourcompany.com', or pass -PrimaryDomain at runtime
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$sites = Import-Csv (Join-Path $OutputDir 'GSites_Inventory_Detailed.csv')
$permissions = @()
$pages = @()
$embeds = @()
$externalDomains = @()

if (Test-Path (Join-Path $OutputDir 'GSites_Permissions.csv')) { $permissions = Import-Csv (Join-Path $OutputDir 'GSites_Permissions.csv') }
if (Test-Path (Join-Path $OutputDir 'Pages.csv')) { $pages = Import-Csv (Join-Path $OutputDir 'Pages.csv') }
if (Test-Path (Join-Path $OutputDir 'Embeds.csv')) { $embeds = Import-Csv (Join-Path $OutputDir 'Embeds.csv') }
if (Test-Path (Join-Path $OutputDir 'ExternalDomains.csv')) { $externalDomains = Import-Csv (Join-Path $OutputDir 'ExternalDomains.csv') }

$report = New-Object System.Collections.Generic.List[object]

foreach ($site in $sites) {
    $siteId = $site.id
    $siteName = $site.name
    $siteUrl = $site.webviewlink

    $sitePages = @($pages | Where-Object { $_.SiteId -eq $siteId })
    $siteEmbeds = @($embeds | Where-Object { $_.SiteId -eq $siteId })
    $siteDomains = @($externalDomains | Where-Object { $_.SiteId -eq $siteId })
    $sitePerms = @($permissions | Where-Object { $_.id -eq $siteId })

    $pageCount = $sitePages.Count
    $embedCount = $siteEmbeds.Count
    $pagesWithEmbeds = @($siteEmbeds | Group-Object PageUrl).Count
    $maxDepth = 0
    if ($sitePages.Count -gt 0) {
        $maxDepth = (($sitePages | Measure-Object Depth -Maximum).Maximum)
    }
    $errorPages = @($sitePages | Where-Object { $_.CrawlStatus -notlike 'Success*' }).Count
    $externalDomainCount = @($siteDomains | Group-Object ExternalDomain).Count

    $permRowsText = $sitePerms | ForEach-Object { (($_.PSObject.Properties.Value | Where-Object { $_ }) -join ' | ').ToLowerInvariant() }
    $publicCount = @($permRowsText | Where-Object { $_ -match '\banyone\b' }).Count
    $domainCount = @($permRowsText | Where-Object { $_ -match '\bdomain\b' }).Count
    $groupCount = @($permRowsText | Where-Object { $_ -match '\bgroup\b' }).Count
    $userCount = @($permRowsText | Where-Object { $_ -match '\buser\b' }).Count
    $externalCount = @($permRowsText | Where-Object { $_ -match '@' -and $_ -notmatch [regex]::Escape("@$PrimaryDomain") }).Count

    $structurePoints = [Math]::Min(20, $pageCount) + [Math]::Min(10, $maxDepth * 2) + [Math]::Min(10, $errorPages * 3)
    $embedPoints = [Math]::Min(20, $embedCount) + [Math]::Min(10, $pagesWithEmbeds * 2) + [Math]::Min(10, $externalDomainCount * 2)
    $securityPoints = [Math]::Min(10, $publicCount * 5) + [Math]::Min(10, $externalCount * 2 + $domainCount)

    $totalScore = $structurePoints + $embedPoints + $securityPoints
    $rating = if ($totalScore -le 25) { 'Low' } elseif ($totalScore -le 50) { 'Medium' } elseif ($totalScore -le 75) { 'High' } else { 'Very High' }
    $recommendation = if ($totalScore -le 25) { 'Direct migration candidate' } elseif ($totalScore -le 50) { 'Partial rebuild likely' } elseif ($totalScore -le 75) { 'Manual redesign needed' } else { 'High-risk; assess separately' }

    $report.Add([pscustomobject]@{
            SiteId                = $siteId
            SiteName              = $siteName
            SiteUrl               = $siteUrl
            PageCount             = $pageCount
            MaxDepth              = $maxDepth
            ErrorPages            = $errorPages
            EmbedCount            = $embedCount
            PagesWithEmbeds       = $pagesWithEmbeds
            ExternalDomainCount   = $externalDomainCount
            PermissionRows        = $sitePerms.Count
            PublicPermissionRows  = $publicCount
            DomainPermissionRows  = $domainCount
            GroupPermissionRows   = $groupCount
            UserPermissionRows    = $userCount
            ExternalPrincipalRows = $externalCount
            StructurePoints       = $structurePoints
            EmbedPoints           = $embedPoints
            SecurityPoints        = $securityPoints
            TotalScore            = $totalScore
            Rating                = $rating
            Recommendation        = $recommendation
        })
}

$report | Export-Csv -NoTypeInformation -Path (Join-Path $OutputDir 'Complexity_Report.csv')
Write-Host 'Scoring completed. Output: Complexity_Report.csv'
