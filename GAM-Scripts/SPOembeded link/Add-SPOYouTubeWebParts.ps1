<#
.SYNOPSIS
    Bulk-adds Embed/YouTube web parts to SharePoint Online modern pages from a CSV mapping file.

.DESCRIPTION
    Reads a CSV containing PageName, EmbedUrl, and optional positioning info, then uses
    PnP PowerShell to add the correct web part type to each page. Supports -WhatIf / -DryRun
    to preview changes without applying them.

.PARAMETER SiteUrl
    The SharePoint Online site URL.

.PARAMETER MappingCsv
    Path to CSV with columns: PageName, EmbedUrl, SectionIndex (opt), ColumnIndex (opt), Order (opt).

.PARAMETER Publish
    If specified, publishes the page after adding the web part. Default leaves page checked out.

.PARAMETER DryRun
    If specified, no changes are made; the script only logs what it would do.

.EXAMPLE
    .\Add-SPOYouTubeWebParts.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/migrated" -MappingCsv "embeds.csv"

.EXAMPLE
    .\Add-SPOYouTubeWebParts.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/migrated" -MappingCsv "embeds.csv" -DryRun
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $true)]
    [string]$MappingCsv,

    [Parameter(Mandatory = $false)]
    [string]$ClientId = "3834b2e7-ab80-45fc-b4c8-ed5c960076b7",

    [Parameter(Mandatory = $false)]
    [string]$TenantId = "",

    [switch]$Publish,

    [switch]$DryRun
)

if (-not (Get-Module -ListAvailable -Name "PnP.PowerShell")) {
    throw "PnP.PowerShell module is required. Install it with: Install-Module PnP.PowerShell -Force"
}

Import-Module PnP.PowerShell

if (-not (Test-Path -Path $MappingCsv)) {
    throw "Mapping CSV not found: $MappingCsv"
}

$Mappings = Import-Csv -Path $MappingCsv
Write-Host "Loaded $($Mappings.Count) row(s) from $MappingCsv" -ForegroundColor Cyan

if ($DryRun) {
    Write-Host "DRY RUN MODE - no changes will be applied." -ForegroundColor Magenta
}

$spoHost = ([uri]$SiteUrl).Host
$tenantName = if ($TenantId) { $TenantId } else { ($spoHost -replace "\.sharepoint\.com$", "") + ".onmicrosoft.com" }

Write-Host "Connecting to: $SiteUrl (Tenant: $tenantName)" -ForegroundColor Cyan
Write-Host "A browser window may open for sign-in. Please authenticate if prompted." -ForegroundColor Yellow
try {
    Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $tenantName -Interactive -ErrorAction Stop
}
catch {
    throw "Failed to connect to SPO: $_"
}

foreach ($row in $Mappings) {
    $pageName = $row.PageName
    $embedUrl = $row.EmbedUrl
    $sectionIndex = if ([string]::IsNullOrWhiteSpace($row.SectionIndex)) { 1 } else { [int]$row.SectionIndex }
    $columnIndex = if ([string]::IsNullOrWhiteSpace($row.ColumnIndex)) { 1 } else { [int]$row.ColumnIndex }
    $order = if ([string]::IsNullOrWhiteSpace($row.Order)) { 0 } else { [int]$row.Order }

    if ([string]::IsNullOrWhiteSpace($pageName) -or [string]::IsNullOrWhiteSpace($embedUrl)) {
        Write-Warning "Skipping row with missing PageName or EmbedUrl"
        continue
    }

    # Use ContentEmbed for all embeds with iframe HTML strings.
    # The ContentEmbed web part expects embedCode to contain a full <iframe> HTML snippet.
    # Ampersands in HTML attribute values must be escaped as &amp; to avoid embed web part parser errors.
    $webPartType = "ContentEmbed"
    $escapedUrl = $embedUrl -replace '&', '&amp;'

    if ($embedUrl -match "youtube\.com|youtu\.be") {
        $videoId = $null
        if ($embedUrl -match 'v=([a-zA-Z0-9_-]+)') { $videoId = $Matches[1] }
        elseif ($embedUrl -match 'youtu\.be/([a-zA-Z0-9_-]+)') { $videoId = $Matches[1] }
        elseif ($embedUrl -match 'embed/([a-zA-Z0-9_-]+)') { $videoId = $Matches[1] }

        if ($videoId) {
            $embedCode = '<iframe width="560" height="315" src="https://www.youtube.com/embed/' + $videoId + '" title="YouTube video player" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe>'
        }
        else {
            Write-Warning "Could not extract YouTube video ID from $embedUrl; using raw URL in iframe."
            $embedCode = '<iframe src="' + $escapedUrl + '" width="600" height="450" style="border:0;" allowfullscreen="" loading="lazy"></iframe>'
        }
    }
    else {
        $embedCode = '<iframe src="' + $escapedUrl + '" width="600" height="450" style="border:0;" allowfullscreen="" loading="lazy"></iframe>'
    }

    $props = @{ embedCode = $embedCode }

    $action = "Add $webPartType to '$pageName' (Section=$sectionIndex, Column=$columnIndex)"

    if ($DryRun) {
        Write-Host "[DRY RUN] Would $action with URL: $embedUrl" -ForegroundColor Magenta
        continue
    }

    if ($PSCmdlet.ShouldProcess($pageName, "Add $webPartType web part")) {
        try {
            $page = Get-PnPPage -Identity $pageName -ErrorAction Stop

            # Determine effective section to use
            $effectiveSection = $sectionIndex
            $maxSection = ($page.Sections | Measure-Object).Count

            # If target section is beyond current count, append new sections
            if ($effectiveSection -gt $maxSection) {
                for ($s = $maxSection + 1; $s -le $effectiveSection; $s++) {
                    Add-PnPPageSection -Page $page -SectionTemplate OneColumn -ErrorAction SilentlyContinue
                }
            }

            # If target section exists and is full-width, append a standard OneColumn section instead
            if ($effectiveSection -le $maxSection) {
                $targetSec = $page.Sections[$effectiveSection - 1]
                if ($targetSec.Type -eq "OneColumnFullWidth") {
                    Write-Host "Target section $effectiveSection is full-width; appending a new OneColumn section for the embed." -ForegroundColor Yellow
                    Add-PnPPageSection -Page $page -SectionTemplate OneColumn -ErrorAction SilentlyContinue
                    $effectiveSection = ($page.Sections | Measure-Object).Count
                }
            }

            $jsonProps = $props | ConvertTo-Json -Compress

            # Add the web part (retry once if full-width conflict occurs)
            try {
                Add-PnPPageWebPart -Page $page `
                    -DefaultWebPartType $webPartType `
                    -WebPartProperties $jsonProps `
                    -Section $effectiveSection `
                    -Column $columnIndex `
                    -Order $order
            }
            catch {
                if ($_ -match "one column full width section" -or $_ -match "text controls inside a one column full width") {
                    Write-Host "Full-width section conflict; appending a new OneColumn section and retrying." -ForegroundColor Yellow
                    Add-PnPPageSection -Page $page -SectionTemplate OneColumn -ErrorAction SilentlyContinue
                    $effectiveSection = ($page.Sections | Measure-Object).Count
                    Add-PnPPageWebPart -Page $page `
                        -DefaultWebPartType $webPartType `
                        -WebPartProperties $jsonProps `
                        -Section $effectiveSection `
                        -Column $columnIndex `
                        -Order $order
                }
                else {
                    throw $_
                }
            }

            if ($Publish) {
                $page.Publish("Published via Add-SPOYouTubeWebParts.ps1")
                Write-Host "Added $webPartType and published '$pageName'." -ForegroundColor Green
            }
            else {
                Write-Host "Added $webPartType to '$pageName' (not published)." -ForegroundColor Green
            }
        }
        catch {
            Write-Warning "Failed to update '$pageName': $_"
        }
    }
}

Disconnect-PnPOnline
Write-Host "Done." -ForegroundColor Cyan
