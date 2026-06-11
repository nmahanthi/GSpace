<#
.SYNOPSIS
    Scans SharePoint Online modern site pages for empty or placeholder embed web parts.

.DESCRIPTION
    Connects to a SharePoint Online site and enumerates modern pages in the Site Pages library.
    Identifies Embed web parts with no URL, or pages that contain placeholder text commonly
    left by migration tools (e.g., "[Embed]", "iframe", "not supported").

.PARAMETER SiteUrl
    The SharePoint Online site URL to scan (e.g., https://contoso.sharepoint.com/sites/migrated).

.PARAMETER OutputCsv
    Path to the output CSV file. Defaults to SPOEmptyEmbeds.csv.

.PARAMETER PlaceholderPatterns
    Array of regex patterns that indicate a broken/placeholder embed. Defaults cover common
    Cloudiway / migration tool placeholders.

.EXAMPLE
    .\Find-SPOEmptyEmbeds.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/migrated"

.EXAMPLE
    .\Find-SPOEmptyEmbeds.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/migrated" -OutputCsv "C:\Audit.csv"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false)]
    [string]$OutputCsv = "SPOEmptyEmbeds.csv",

    [Parameter(Mandatory = $false)]
    [string]$ClientId = "3834b2e7-ab80-45fc-b4c8-ed5c960076b7",

    [Parameter(Mandatory = $false)]
    [string]$TenantId = "",

    [Parameter(Mandatory = $false)]
    [string[]]$PlaceholderPatterns = @(
        '\[embed\]',
        '\[iframe\]',
        'not supported',
        'cannot display',
        'placeholder',
        'google\.com\/embed',
        'youtube\.com\/embed',
        '<iframe'
    )
)

# Ensure PnP.PowerShell is available
if (-not (Get-Module -ListAvailable -Name "PnP.PowerShell")) {
    throw "PnP.PowerShell module is required. Install it with: Install-Module PnP.PowerShell -Force"
}

Import-Module PnP.PowerShell

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

$Pages = Get-PnPListItem -List "SitePages" -PageSize 500
Write-Host "Retrieved $($Pages.Count) page(s) from Site Pages." -ForegroundColor Green

$Issues = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($pageItem in $Pages) {
    $fileName = $pageItem.FieldValues["FileLeafRef"]
    Write-Verbose "Scanning page: $fileName"

    try {
        $page = Get-PnPPage -Identity $fileName -ErrorAction Stop
    }
    catch {
        Write-Warning "Could not load page '$fileName': $_"
        continue
    }

    # Check web parts on the page
    $webParts = $page.Controls
    foreach ($wp in $webParts) {
        $json = $wp.PropertiesJson | ConvertFrom-Json -ErrorAction SilentlyContinue

        # Detect empty Embed web parts
        if ($wp.Type.Name -eq "Embed") {
            $embedCode = $json.embedCode
            $isEmpty = [string]::IsNullOrWhiteSpace($embedCode)

            # Also check for placeholders inside embedCode
            $isPlaceholder = $false
            if (-not $isEmpty) {
                foreach ($pattern in $PlaceholderPatterns) {
                    if ($embedCode -match $pattern) {
                        $isPlaceholder = $true
                        break
                    }
                }
            }

            if ($isEmpty -or $isPlaceholder) {
                $Issues.Add([PSCustomObject]@{
                        PageName     = $fileName
                        PageUrl      = "$SiteUrl/SitePages/$fileName"
                        WebPartType  = $wp.Type.Name
                        WebPartTitle = $wp.Title
                        SectionIndex = $wp.Section.Order
                        ColumnIndex  = $wp.Column.Order
                        Issue        = if ($isEmpty) { "EmptyEmbedUrl" } else { "PlaceholderEmbed" }
                        CurrentValue = if ($embedCode.Length -gt 200) { $embedCode.Substring(0, 200) + "..." } else { $embedCode }
                    })
            }
        }

        # Detect placeholder text in Text web parts (migration tools sometimes dump iframe HTML as text)
        if ($wp.Type.Name -eq "Text") {
            $text = $json.text
            if (-not [string]::IsNullOrWhiteSpace($text)) {
                foreach ($pattern in $PlaceholderPatterns) {
                    if ($text -match $pattern) {
                        $Issues.Add([PSCustomObject]@{
                                PageName     = $fileName
                                PageUrl      = "$SiteUrl/SitePages/$fileName"
                                WebPartType  = $wp.Type.Name
                                WebPartTitle = $wp.Title
                                SectionIndex = $wp.Section.Order
                                ColumnIndex  = $wp.Column.Order
                                Issue        = "PlaceholderText"
                                CurrentValue = if ($text.Length -gt 200) { $text.Substring(0, 200) + "..." } else { $text }
                            })
                        break
                    }
                }
            }
        }
    }
}

Write-Host "Found $($Issues.Count) issue(s)." -ForegroundColor $(if ($Issues.Count -gt 0) { "Yellow" } else { "Green" })

$Issues | Export-Csv -Path $OutputCsv -NoTypeInformation
Write-Host "Results exported to: $OutputCsv" -ForegroundColor Cyan

Disconnect-PnPOnline
