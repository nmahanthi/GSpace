<#
.SYNOPSIS
    Extracts all embedded content URLs (iframes, embeds, etc.) from a Google Sites HTML export.

.DESCRIPTION
    Scans HTML files exported from Google Sites and extracts iframe src, embed src,
    and Google Sites-specific embed URLs. Outputs a CSV inventory that can be used
    to rebuild embeds in SharePoint Online.

.PARAMETER ExportPath
    Path to the Google Sites HTML export folder.

.PARAMETER OutputCsv
    Path to the output CSV file. Defaults to GsitesEmbeds.csv in the current directory.

.EXAMPLE
    .\Scan-GSitesEmbeds.ps1 -ExportPath "C:\GsitesExport" -OutputCsv "C:\Embeds.csv"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ExportPath,

    [Parameter(Mandatory = $false)]
    [string]$OutputCsv = "GsitesEmbeds.csv"
)

# Validate export path
if (-not (Test-Path -Path $ExportPath)) {
    throw "Export path not found: $ExportPath"
}

Write-Host "Scanning Google Sites export at: $ExportPath" -ForegroundColor Cyan

$HtmlFiles = Get-ChildItem -Path $ExportPath -Recurse -Filter "*.html"
Write-Host "Found $($HtmlFiles.Count) HTML file(s)." -ForegroundColor Green

$Embeds = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($file in $HtmlFiles) {
    try {
        $content = Get-Content -Path $file.FullName -Raw -ErrorAction Stop
    } catch {
        Write-Warning "Failed to read $($file.FullName): $_"
        continue
    }

    # Pattern 1: Standard iframe src
    $iframeMatches = [regex]::Matches($content, '<iframe[^>]*?\s+src=["\'']([^"\''>]+)["\''][^>]*?>')
    foreach ($m in $iframeMatches) {
        $url = $m.Groups[1].Value.Trim()
        $embedType = if ($url -match "youtube|youtu\.be") { "YouTube" }
                     elseif ($url -match "google\.com/maps|maps\.google") { "GoogleMaps" }
                     elseif ($url -match "drive\.google") { "GoogleDrive" }
                     elseif ($url -match "docs\.google") { "GoogleDocs" }
                     else { "GenericIframe" }

        $Embeds.Add([PSCustomObject]@{
            SourceFile  = $file.FullName.Substring($ExportPath.Length).TrimStart('\','/')
            FileName    = $file.Name
            PatternType = "iframe"
            EmbedUrl    = $url
            EmbedType   = $embedType
            ContextHtml = ($m.Value -replace '\s+', ' ').Substring(0, [Math]::Min(200, $m.Value.Length))
        })
    }

    # Pattern 2: embed tag src (e.g. Flash, video)
    $embedMatches = [regex]::Matches($content, '<embed[^>]*?\s+src=["\'']([^"\''>]+)["\''][^>]*?>')
    foreach ($m in $embedMatches) {
        $Embeds.Add([PSCustomObject]@{
            SourceFile  = $file.FullName.Substring($ExportPath.Length).TrimStart('\','/')
            FileName    = $file.Name
            PatternType = "embed"
            EmbedUrl    = $m.Groups[1].Value.Trim()
            EmbedType   = "EmbedTag"
            ContextHtml = ($m.Value -replace '\s+', ' ').Substring(0, [Math]::Min(200, $m.Value.Length))
        })
    }

    # Pattern 3: Google Sites embed URLs (data-url or ng-non-bindable blocks)
    $gsitesMatches = [regex]::Matches($content, 'data-url=["\'']([^"\''>]+)["\'']')
    foreach ($m in $gsitesMatches) {
        $url = $m.Groups[1].Value.Trim()
        if ($url -match '^https?://') {
            $Embeds.Add([PSCustomObject]@{
                SourceFile  = $file.FullName.Substring($ExportPath.Length).TrimStart('\','/')
                FileName    = $file.Name
                PatternType = "gsites-data-url"
                EmbedUrl    = $url
                EmbedType   = "GoogleSitesEmbed"
                ContextHtml = ($m.Value -replace '\s+', ' ').Substring(0, [Math]::Min(200, $m.Value.Length))
            })
        }
    }
}

# Remove duplicates based on SourceFile + EmbedUrl
$UniqueEmbeds = $Embeds | Sort-Object SourceFile, EmbedUrl -Unique

Write-Host "Found $($UniqueEmbeds.Count) unique embed(s)." -ForegroundColor Green

$UniqueEmbeds | Export-Csv -Path $OutputCsv -NoTypeInformation
Write-Host "Results exported to: $OutputCsv" -ForegroundColor Cyan
