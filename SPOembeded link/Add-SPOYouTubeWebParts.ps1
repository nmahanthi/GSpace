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
    $pageName        = $row.PageName
    $embedUrl        = $row.EmbedUrl
    $sectionIndex    = if ([string]::IsNullOrWhiteSpace($row.SectionIndex))    { 1 }        else { [int]$row.SectionIndex }
    $columnIndex     = if ([string]::IsNullOrWhiteSpace($row.ColumnIndex))     { 1 }        else { [int]$row.ColumnIndex }
    $order           = if ([string]::IsNullOrWhiteSpace($row.Order))           { 0 }        else { [int]$row.Order }
    $embedWidth      = if ([string]::IsNullOrWhiteSpace($row.EmbedWidth)  -or [int]$row.EmbedWidth  -le 0) { 600 } else { [int]$row.EmbedWidth }
    $embedHeight     = if ([string]::IsNullOrWhiteSpace($row.EmbedHeight) -or [int]$row.EmbedHeight -le 0) { 450 } else { [int]$row.EmbedHeight }
    $hAlign          = if ([string]::IsNullOrWhiteSpace($row.HorizontalAlign)) { 'center' } else { $row.HorizontalAlign.ToLower() }
    $sectionTemplate = if ([string]::IsNullOrWhiteSpace($row.SectionTemplate)) { 'OneColumn'} else { $row.SectionTemplate }

    # Validate required fields
    if ([string]::IsNullOrWhiteSpace($pageName) -or [string]::IsNullOrWhiteSpace($embedUrl)) {
        Write-Warning "Skipping row with missing PageName or EmbedUrl"
        continue
    }

    # Alignment CSS: centre uses auto margins; left/right uses no margin on the leading side
    $marginStyle = switch ($hAlign) {
        'left'  { 'display:block;margin:0;' }
        'right' { 'display:block;margin:0 0 0 auto;' }
        default { 'display:block;margin:0 auto;' }   # center
    }

    # Build the iframe embed code using actual captured dimensions and alignment
    $webPartType = "ContentEmbed"
    $escapedUrl  = $embedUrl -replace '&', '&amp;'

    if ($embedUrl -match "youtube\.com|youtu\.be") {
        $videoId = $null
        if      ($embedUrl -match 'v=([a-zA-Z0-9_-]+)')          { $videoId = $Matches[1] }
        elseif  ($embedUrl -match 'youtu\.be/([a-zA-Z0-9_-]+)')  { $videoId = $Matches[1] }
        elseif  ($embedUrl -match 'embed/([a-zA-Z0-9_-]+)')      { $videoId = $Matches[1] }

        if ($videoId) {
            $embedCode = '<iframe width="' + $embedWidth + '" height="' + $embedHeight + '" style="' + $marginStyle + '" src="https://www.youtube.com/embed/' + $videoId + '" title="YouTube video player" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe>'
        } else {
            Write-Warning "Could not extract YouTube video ID from $embedUrl; using raw URL in iframe."
            $embedCode = '<iframe src="' + $escapedUrl + '" width="' + $embedWidth + '" height="' + $embedHeight + '" style="' + $marginStyle + 'border:0;" allowfullscreen="" loading="lazy"></iframe>'
        }
    } else {
        $embedCode = '<iframe src="' + $escapedUrl + '" width="' + $embedWidth + '" height="' + $embedHeight + '" style="' + $marginStyle + 'border:0;" allowfullscreen="" loading="lazy"></iframe>'
    }

    $props  = @{ embedCode = $embedCode }
    $action = "Add $webPartType to '$pageName' (Section=$sectionIndex Col=$columnIndex Align=$hAlign ${embedWidth}x${embedHeight} Template=$sectionTemplate)"

    if ($DryRun) {
        Write-Host "[DRY RUN] $action" -ForegroundColor Magenta
        Write-Host "          URL: $embedUrl" -ForegroundColor DarkGray
        continue
    }

    if ($PSCmdlet.ShouldProcess($pageName, "Add $webPartType web part")) {
        try {
            $page       = Get-PnPPage -Identity $pageName -ErrorAction Stop
            $maxSection = ($page.Sections | Measure-Object).Count

            # Always append a new section so each embed gets its own row,
            # using the template that matches the original Google Sites layout.
            Add-PnPPageSection -Page $page -SectionTemplate $sectionTemplate -ErrorAction SilentlyContinue
            $effectiveSection = ($page.Sections | Measure-Object).Count

            # Safety: if the new section turned out full-width, replace with OneColumn
            $newSec = $page.Sections[$effectiveSection - 1]
            if ($newSec.Type -eq "OneColumnFullWidth") {
                Write-Host "  Section is full-width; switching to OneColumn for embed." -ForegroundColor Yellow
                Add-PnPPageSection -Page $page -SectionTemplate OneColumn -ErrorAction SilentlyContinue
                $effectiveSection = ($page.Sections | Measure-Object).Count
            }

            $jsonProps = $props | ConvertTo-Json -Compress

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
                    Write-Host "  Full-width conflict on retry; falling back to OneColumn." -ForegroundColor Yellow
                    Add-PnPPageSection -Page $page -SectionTemplate OneColumn -ErrorAction SilentlyContinue
                    $effectiveSection = ($page.Sections | Measure-Object).Count
                    Add-PnPPageWebPart -Page $page `
                        -DefaultWebPartType $webPartType `
                        -WebPartProperties $jsonProps `
                        -Section $effectiveSection `
                        -Column 1 `
                        -Order $order
                } else { throw $_ }
            }

            if ($Publish) {
                $page.Publish("Published via Add-SPOYouTubeWebParts.ps1")
                Write-Host "Added + published $webPartType on '$pageName' ($hAlign, ${embedWidth}x${embedHeight})." -ForegroundColor Green
            } else {
                Write-Host "Added $webPartType to '$pageName' ($hAlign, ${embedWidth}x${embedHeight}, $sectionTemplate col $columnIndex)." -ForegroundColor Green
            }
        }
        catch {
            Write-Warning "Failed to update '$pageName': $_"
        }
    }
}

Disconnect-PnPOnline
Write-Host "Done." -ForegroundColor Cyan
