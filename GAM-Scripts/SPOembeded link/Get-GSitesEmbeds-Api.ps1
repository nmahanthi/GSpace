<#
.SYNOPSIS
    Extracts embeds from a New Google Sites page using the Google Sites API (v1).

.DESCRIPTION
    Calls sites.googleapis.com/v1 to read the page structure, extracts
    embeddedContent nodes (YouTube, Maps, etc.), and outputs a CSV.

.PARAMETER SiteId
    Google Sites site ID (from URL: /d/<SiteId>/p/<PageId>).

.PARAMETER PageId
    Google Sites page ID (from URL: /d/<SiteId>/p/<PageId>).

.PARAMETER AccessToken
    A valid Google OAuth access token with the scope:
    https://www.googleapis.com/auth/sites.readonly

    You can get one from Google Cloud OAuth Playground:
    https://developers.google.com/oauthplayground

.PARAMETER OutputCsv
    Output CSV path.

.EXAMPLE
    .\Get-GSitesEmbeds-Api.ps1 `
        -SiteId "19JMNQ5kCG9VAAaWjrW65uH_G5--5r3un" `
        -PageId "1TQ4yy3DBedUJ43SP-ZYz1DubSv50Iu7e" `
        -AccessToken "ya29.a0AfH6SMB..."
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$SiteId,
    [Parameter(Mandatory)][string]$PageId,
    [Parameter(Mandatory)][string]$AccessToken,
    [string]$OutputCsv = "GSitesApiEmbeds.csv"
)

$headers = @{ Authorization = "Bearer $AccessToken"; Accept = "application/json" }

function Invoke-GSitesApi {
    param([string]$Uri)
    try {
        return Invoke-RestMethod -Uri $Uri -Headers $headers -Method GET -ContentType "application/json"
    } catch {
        Write-Warning "API call failed: $($_.Exception.Message)"
        return $null
    }
}

Write-Host "Fetching page from Google Sites API..." -ForegroundColor Cyan
$page = Invoke-GSitesApi -Uri "https://sites.googleapis.com/v1/sites/$SiteId/pages/$PageId"
if (-not $page) { throw "Failed to fetch page. Check SiteId, PageId, and AccessToken." }

$embeds = [System.Collections.Generic.List[pscustomobject]]::new()
function Extract-Content {
    param($node, $path = "")
    if ($null -eq $node) { return }

    # Check for embeddedContent
    if ($node.embeddedContent) {
        $ec = $node.embeddedContent
        $embeds.Add([pscustomobject]@{
            Path       = $path
            Type       = $ec.type
            Url        = $ec.url
            EmbedUrl   = $ec.embedUrl
            Title      = $ec.title
            Height     = $ec.height
            Width      = $ec.width
        })
    }

    # Recurse into child content
    if ($node.content) {
        $i = 0
        foreach ($child in $node.content) {
            Extract-Content -node $child -path "$path/content[$i]"
            $i++
        }
    }
    if ($node.children) {
        $i = 0
        foreach ($child in $node.children) {
            Extract-Content -node $child -path "$path/children[$i]"
            $i++
        }
    }
    if ($node.section) {
        $i = 0
        foreach ($child in $node.section) {
            Extract-Content -node $child -path "$path/section[$i]"
            $i++
        }
    }
}

if ($page.content) { Extract-Content -node $page.content -path "content" }
if ($page.children) { Extract-Content -node $page.children -path "children" }

Write-Host "Found $($embeds.Count) embed(s) via API." -ForegroundColor Green
if ($embeds.Count -gt 0) {
    $embeds | Format-Table -AutoSize
    $embeds | Export-Csv -Path $OutputCsv -NoTypeInformation
    Write-Host "Saved to: $OutputCsv" -ForegroundColor Cyan
} else {
    Write-Warning "No embeddedContent nodes found in API response. The page may have no embeds, or the API structure may differ."
    Write-Host "Raw response keys:" -ForegroundColor Yellow
    if ($page.PSObject.Properties) { $page.PSObject.Properties.Name }
}
