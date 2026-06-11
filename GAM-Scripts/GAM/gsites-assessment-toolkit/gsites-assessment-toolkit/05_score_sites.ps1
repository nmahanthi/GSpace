param(
    [string]$OutputDir = "$PSScriptRoot\output",
    [string]$PrimaryDomain = 'roche.com'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$sites = Import-Csv (Join-Path $OutputDir '02_GSites_Inventory_Detailed.csv')
$permissions = @()
$pages = @()
$embeds = @()
$externalDomains = @()
$sheetEnrichment = @()
$formEnrichment = @()
$scriptEnrichment = @()

if (Test-Path (Join-Path $OutputDir '03_GSites_Permissions.csv')) { $permissions = Import-Csv (Join-Path $OutputDir '03_GSites_Permissions.csv') }
if (Test-Path (Join-Path $OutputDir '07_Pages.csv')) { $pages = Import-Csv (Join-Path $OutputDir '07_Pages.csv') }
if (Test-Path (Join-Path $OutputDir '08_Embeds.csv')) { $embeds = Import-Csv (Join-Path $OutputDir '08_Embeds.csv') }
if (Test-Path (Join-Path $OutputDir '09_ExternalDomains.csv')) { $externalDomains = Import-Csv (Join-Path $OutputDir '09_ExternalDomains.csv') }
if (Test-Path (Join-Path $OutputDir '11_Sheets_Enrichment.csv')) { $sheetEnrichment = Import-Csv (Join-Path $OutputDir '11_Sheets_Enrichment.csv') }
if (Test-Path (Join-Path $OutputDir '12_Forms_Enrichment.csv')) { $formEnrichment = Import-Csv (Join-Path $OutputDir '12_Forms_Enrichment.csv') }
if (Test-Path (Join-Path $OutputDir '13_Scripts_Enrichment.csv')) { $scriptEnrichment = Import-Csv (Join-Path $OutputDir '13_Scripts_Enrichment.csv') }

$report = New-Object System.Collections.Generic.List[object]

foreach ($site in $sites) {
    $siteId = $site.id
    $siteName = $site.name
    $siteUrl = $site.webviewlink

    $sitePages = @($pages | Where-Object { $_.SiteId -eq $siteId })
    $siteEmbeds = @($embeds | Where-Object { $_.SiteId -eq $siteId })
    $siteDomains = @($externalDomains | Where-Object { $_.SiteId -eq $siteId })
    $sitePerms = @($permissions | Where-Object { $_.id -eq $siteId })

    $sheetMeasure = $sheetEnrichment | Where-Object { $_.SiteId -eq $siteId } | Measure-Object ComplexityPoints -Sum
    $siteSheetPoints = if ($sheetMeasure -and $sheetMeasure.Sum) { $sheetMeasure.Sum } else { 0 }

    $formMeasure = $formEnrichment | Where-Object { $_.SiteId -eq $siteId } | Measure-Object ComplexityPoints -Sum
    $siteFormPoints = if ($formMeasure -and $formMeasure.Sum) { $formMeasure.Sum } else { 0 }

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
    $artifactPoints = [Math]::Min(15, [int]$siteSheetPoints) + [Math]::Min(15, [int]$siteFormPoints)

    $totalScore = $structurePoints + $embedPoints + $securityPoints + $artifactPoints
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
            ArtifactPoints        = $artifactPoints
            TotalScore            = $totalScore
            Rating                = $rating
            Recommendation        = $recommendation
        })
}

$report | Export-Csv -NoTypeInformation -Path (Join-Path $OutputDir '14_Complexity_Report.csv')
Write-Host 'Scoring completed. Output: 14_Complexity_Report.csv'
