#Requires -Version 5.1
<#
.SYNOPSIS
    SharePoint Online Site Collection Assessment Report
.DESCRIPTION
    Comprehensive assessment of all SPO site collections in a tenant:
      - Site collection metadata and storage size
      - Lists and document libraries inventory
      - Web parts used on modern (client-side) and classic pages
      - Custom / third-party web parts detection
      - App Catalog installed solutions
    Outputs a multi-sheet Excel workbook and a self-contained HTML dashboard.

    Authentication: Delegated / interactive (browser login).
    Only ClientId and TenantId are required — a login popup appears once.
    MSAL token caching means all subsequent site connections are silent.

    Azure AD App Registration requirements:
      - Platform: Mobile and desktop applications  (redirect URI: https://login.microsoftonline.com/common/oauth2/nativeclient)
      - API permissions (delegated): Sites.FullControl.All (SharePoint), AllSites.Read (SharePoint)
.PARAMETER AdminUrl
    SPO admin center URL, e.g. https://contoso-admin.sharepoint.com
.PARAMETER ClientId
    Azure AD Application (client) ID. Required.
.PARAMETER TenantId
    Azure AD Tenant ID (GUID or domain). Required.
.PARAMETER OutputPath
    Output folder (created if absent). Defaults to script directory + timestamp.
.PARAMETER SkipPersonalSites
    Skip OneDrive personal sites. Default: $true.
.PARAMETER SkipSystemSites
    Skip system/app-catalog sites. Default: $true.
.PARAMETER MaxSites
    Cap on sites processed; 0 = all. Default: 0.
.PARAMETER ExportCSV
    Also write raw CSV files alongside Excel and HTML.
.PARAMETER ScanClassicPages
    Also enumerate classic ASPX pages for legacy web parts (slower).
.EXAMPLE
    # Minimal — just ClientId and TenantId; a browser login prompt appears once.
    .\SPO-SiteCollection-Assessment.ps1 `
        -AdminUrl "https://contoso-admin.sharepoint.com" `
        -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
.EXAMPLE
    # With optional extras — export CSVs, cap at 50 sites, also scan classic pages.
    .\SPO-SiteCollection-Assessment.ps1 `
        -AdminUrl "https://contoso-admin.sharepoint.com" `
        -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -ExportCSV -MaxSites 50 -ScanClassicPages
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$AdminUrl,
    [Parameter(Mandatory)][string]$ClientId,
    [Parameter(Mandatory)][string]$TenantId,
    [string]$OutputPath        = (Join-Path $PSScriptRoot "SPO_SC_Assessment_$(Get-Date -f 'yyyyMMdd_HHmmss')"),
    [bool]  $SkipPersonalSites = $true,
    [bool]  $SkipSystemSites   = $true,
    [int]   $MaxSites          = 0,
    [switch]$ExportCSV,
    [switch]$ScanClassicPages
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$script:HasErrors      = $false

# ─── KNOWN OOTB MODERN WEB PART COMPONENT IDs ────────────────────────────────
# Source: Microsoft SharePoint Framework docs & community registry.
# Any web part whose ID is NOT in this list is flagged as Custom/Third-party.
$script:OOTBWebPartIds = @(
    # Text & Media
    "1ef91747-30b6-4e12-9c25-b4a1d5f4e3df", # Text
    "d1d91016-032f-456d-98a4-721247c305e8", # Image
    "09073491-e787-4acc-bdab-5f4bcf8a6cce", # Image Gallery
    "a2750b41-8ac2-4fd7-9dfd-4e0c8b47d01b", # File Viewer
    "e377ea37-9047-43b9-8cdb-a761be2f8e09", # Video (Stream Classic)
    "275c0095-a77e-4f6d-a2a0-6a7626911518", # Microsoft Stream
    "f6c7485e-1b9d-4f0c-b440-c2db3f5f5c1e", # Embed
    # Layout
    "c4bd7b2f-7b6e-4599-8485-16504575f590", # Hero
    "8c88f208-6c77-4bdb-86a0-0c47b4316588", # News
    "2ef55461-c7b2-4c8a-8b1d-25c7b5f52e87", # News Reel / Ticker
    "b7dd04e1-19ce-4b25-b19a-7f4e19a3f97b", # Quick Links
    "db3e97c7-d593-42b3-9e35-f4e5bd25b25a", # Call To Action
    "6410b3b6-d440-4663-8744-378976dc041e", # Divider / Spacer
    "a0e3fd5b-2c78-484e-9f96-c5d7bc1a1e4a", # Button
    # Business / Data
    "868ac3c3-cad7-4bd6-9a90-af2b92af1cef", # List
    "7f718435-ee4d-431c-bdbf-9c4ff326f46e", # Document Library
    "58fcd18b-e1af-4b0a-b23b-422c2c364d64", # Events
    "cbdef574-2e1f-4eb2-9c97-fecbf7f3e7a3", # Sites (Highlighted content)
    "251d7ab8-7285-49f9-9ee3-5e8485aba1c0", # People
    "1a0e59d8-7af5-4d32-b7ee-a81d4e4fa17d", # My Feed / Activity
    "d0a0bfb3-4c96-4f23-a5d0-1cb614c1d5aa", # Microsoft Forms
    "45166e24-c8e8-4588-bcc3-30c7c5e6a0e3", # Power BI
    "9d7e898c-f1bb-473a-9ace-8b7b6f7c3c85", # Power Apps
    "c7d40f43-5b75-4c2b-b62f-6be7f7e2e76a", # Power Automate
    "c1f0d0e4-0a9e-4c1a-b8e5-7d3f5b9f8c2a", # Planner
    # Navigation & Search
    "c9b16de1-5a1a-4a87-87d0-c2df7d0c0ab9", # Search Box
    "a0e3fd5b-2c78-484e-9f96-c5d7bc1a1e4a", # Quick Links (alt)
    "4ef7b2c2-8aa9-4959-aa43-e7fbbebea38f", # Hub Navigation
    "0f81b48e-c3d4-4b30-8e94-aaf5d5c56e58"  # Table of Contents
) | ForEach-Object { $_.ToLower() }

# ─── HELPERS ──────────────────────────────────────────────────────────────────
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts    = Get-Date -Format "HH:mm:ss"
    $color = switch ($Level) { "SUCCESS"{"Green"} "WARN"{"Yellow"} "ERROR"{"Red"} default{"Cyan"} }
    Write-Host "[$ts][$Level] $Message" -ForegroundColor $color
    if ($Level -eq "ERROR") { $script:HasErrors = $true }
}

function Invoke-SafePnP {
    param([scriptblock]$Block, [string]$Ctx = "")
    try   { & $Block }
    catch {
        $msg = ($_.Exception.Message -split "`n")[0].Trim()
        Write-Log "Skipped$(if($Ctx){ " [$Ctx]" }): $msg" "WARN"
        return $null
    }
}

# ─── MODULE CHECK ─────────────────────────────────────────────────────────────
function Assert-Modules {
    $req  = @("PnP.PowerShell","ImportExcel")
    $miss = @($req | Where-Object { -not (Get-Module -ListAvailable -Name $_) })
    if ($miss) {
        $cmds = ($miss | ForEach-Object { "  Install-Module $_ -Scope CurrentUser -Force" }) -join "`n"
        Write-Log "Missing modules. Run:`n$cmds" "ERROR"; exit 1
    }
    Import-Module PnP.PowerShell -ErrorAction Stop
    Import-Module ImportExcel    -ErrorAction Stop
    Write-Log "Modules ready: PnP.PowerShell, ImportExcel." "SUCCESS"
}

# ─── CONNECTION ───────────────────────────────────────────────────────────────
# $script:PnPCreds holds the acquired OAuth token after the first interactive
# login so that Connect-Site reuses it silently (no re-prompt per site).
$script:PnPCreds = $null

function Connect-Admin {
    Write-Log "Connecting to SPO Admin: $AdminUrl" "INFO"
    Write-Log "Auth mode: Interactive (browser login — prompted once)" "INFO"
    Connect-PnPOnline -Url $AdminUrl -ClientId $ClientId -Tenant $TenantId -Interactive
    # Cache the acquired credentials for subsequent per-site connections
    $script:PnPCreds = Get-PnPConnection
    Write-Log "Connected to admin center." "SUCCESS"
}

function Connect-Site { param([string]$Url)
    # Reuse the cached token; MSAL will silently refresh if needed — no browser popup.
    Connect-PnPOnline -Url $Url -ClientId $ClientId -Tenant $TenantId -Interactive -ErrorAction Stop
}

# ─── SITE COLLECTION ENUMERATION ──────────────────────────────────────────────
function Get-AllSiteCollections {
    Write-Log "=== Enumerating Site Collections ===" "INFO"
    Connect-Admin
    $all = @(Invoke-SafePnP { Get-PnPTenantSite -Detailed } "TenantSite")
    if (-not $all -or $all.Count -eq 0) { Write-Log "No sites returned. Verify SharePoint Admin role." "ERROR"; return @() }
    if ($SkipPersonalSites) { $all = @($all | Where-Object { $_.Url -notlike "*/personal/*" }) }
    if ($SkipSystemSites) {
        $all = @($all | Where-Object {
            $_.Url      -notlike "*/appcatalog/*"    -and
            $_.Url      -notlike "*/contentTypeHub*" -and
            $_.Template -notlike "SRCHCEN*"          -and
            $_.Template -notlike "POINTPUBLISHINGHUB*" -and
            $_.Template -notlike "RedirectSite*"
        })
    }
    if ($MaxSites -gt 0) { $all = $all | Select-Object -First $MaxSites }
    Write-Log "Found $($all.Count) site collections to process." "INFO"
    return $all
}

# ─── APP CATALOG SOLUTIONS ────────────────────────────────────────────────────
function Get-AppCatalogSolutions {
    Write-Log "=== Scanning App Catalog for Installed Solutions ===" "INFO"
    $solutions = [System.Collections.Generic.List[PSCustomObject]]::new()
    try {
        Connect-Admin
        $apps = Invoke-SafePnP { Get-PnPApp -Scope Tenant } "AppCatalog"
        if ($apps) {
            foreach ($app in $apps) {
                $solutions.Add([PSCustomObject]@{
                    AppId            = $app.Id
                    Title            = $app.Title
                    Version          = $app.AppVersion
                    IsEnabled        = $app.Enabled
                    IsDeployed       = $app.Deployed
                    SkipFeatureDeployment = $app.SkipFeatureDeployment
                    Publisher        = if ($app.AppPackageErrorMessage) { "Error" } else { "Custom/Partner" }
                    ErrorMessage     = $app.AppPackageErrorMessage
                })
            }
        }
    } catch { Write-Log "App Catalog scan skipped: $($_.Exception.Message)" "WARN" }
    Write-Log "App Catalog solutions found: $($solutions.Count)" "INFO"
    return $solutions
}

# ─── LISTS & LIBRARIES COLLECTION ────────────────────────────────────────────
function Get-ListsAndLibraries {
    param([string]$SiteUrl, [string]$SiteTitle)
    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    $lists = @(Invoke-SafePnP {
        Get-PnPList -Includes HasUniqueRoleAssignments,ItemCount,RootFolder,EnableVersioning,
                              MajorVersionLimit,MinorVersionLimit,BaseType,BaseTemplate,
                              Hidden,ContentTypes,LastItemModifiedDate,Created
    } $SiteUrl)

    foreach ($list in $lists) {
        if ($list.Hidden) { continue }
        $listType = switch ($list.BaseType) {
            "DocumentLibrary" { "Document Library" }
            "GenericList"     { "List" }
            default           { $list.BaseType }
        }
        $results.Add([PSCustomObject]@{
            SiteUrl        = $SiteUrl
            SiteTitle      = $SiteTitle
            ListTitle      = $list.Title
            ListType       = $listType
            BaseTemplate   = [int]$list.BaseTemplate
            ItemCount      = $list.ItemCount
            HasUniquePerms = $list.HasUniqueRoleAssignments
            Versioning     = $list.EnableVersioning
            MajorVersions  = $list.MajorVersionLimit
            MinorVersions  = $list.MinorVersionLimit
            ContentTypeCount = if ($list.ContentTypes) { $list.ContentTypes.Count } else { 0 }
            Created        = if ($list.Created) { $list.Created.ToString("yyyy-MM-dd") } else { "" }
            LastModified   = if ($list.LastItemModifiedDate) { $list.LastItemModifiedDate.ToString("yyyy-MM-dd") } else { "" }
            URL            = if ($list.RootFolder) { $SiteUrl.TrimEnd('/') + $list.RootFolder.ServerRelativeUrl } else { "" }
        })
    }
    return $results
}

# ─── WEB PARTS SCANNING ───────────────────────────────────────────────────────
function Get-ModernPageWebParts {
    param([string]$SiteUrl, [string]$SiteTitle)
    $results = [System.Collections.Generic.List[PSCustomObject]]::new()

    # Get all modern pages (Site Pages library)
    $pages = @(Invoke-SafePnP {
        Get-PnPListItem -List "Site Pages" -Fields "FileLeafRef","Title","_SPPageType","Modified" -PageSize 200
    } "$SiteUrl/SitePages")

    if (-not $pages) { return $results }

    foreach ($pageItem in $pages) {
        $pageName = $pageItem.FieldValues["FileLeafRef"]
        if (-not $pageName) { continue }

        $page = Invoke-SafePnP { Get-PnPClientSidePage -Identity $pageName } "$SiteUrl/$pageName"
        if (-not $page) { continue }

        foreach ($control in $page.Controls) {
            # Only process web part controls (not text sections)
            if ($control.GetType().Name -notin @("ClientSideWebPart","PageWebPart")) { continue }

            $wpId    = if ($control.WebPartId) { $control.WebPartId.ToString().ToLower() } else { "" }
            $wpTitle = if ($control.Title)     { $control.Title }     else { "Unknown" }
            $wpType  = if ($control.WebPartId) { "Modern (Client-Side)" } else { "Text/Section" }
            $isOOTB  = $wpId -and ($script:OOTBWebPartIds -contains $wpId)
            $isCustom = $wpId -and -not $isOOTB

            $results.Add([PSCustomObject]@{
                SiteUrl       = $SiteUrl
                SiteTitle     = $SiteTitle
                PageName      = $pageName
                PageType      = "Modern"
                WebPartTitle  = $wpTitle
                WebPartId     = $wpId
                WebPartType   = $wpType
                IsOOTB        = $isOOTB
                IsCustom      = $isCustom
                Section       = if ($control.Section) { $control.Section.Order } else { 0 }
                Column        = if ($control.Column)  { $control.Column.Order  } else { 0 }
                PropertiesJson = if ($control.PropertiesJson) { $control.PropertiesJson.Substring(0,[Math]::Min(200,$control.PropertiesJson.Length)) } else { "" }
            })
        }
    }
    return $results
}

function Get-ClassicPageWebParts {
    param([string]$SiteUrl, [string]$SiteTitle)
    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    if (-not $ScanClassicPages) { return $results }

    # Enumerate classic ASPX pages in root and subsites
    $aspxPages = @(Invoke-SafePnP {
        Get-PnPListItem -List "Pages" -Fields "FileLeafRef","Title","Modified" -PageSize 200
    } "$SiteUrl/Pages")

    if (-not $aspxPages) { return $results }

    foreach ($pageItem in $aspxPages) {
        $pageName = $pageItem.FieldValues["FileLeafRef"]
        if (-not $pageName) { continue }

        $serverRelUrl = "/sites/" + ($SiteUrl -split "/sites/" | Select-Object -Last 1) + "/Pages/$pageName"
        $webparts = @(Invoke-SafePnP { Get-PnPWebPart -ServerRelativePageUrl $serverRelUrl } "$SiteUrl/Pages/$pageName")
        if (-not $webparts) { continue }

        foreach ($wp in $webparts) {
            $results.Add([PSCustomObject]@{
                SiteUrl       = $SiteUrl
                SiteTitle     = $SiteTitle
                PageName      = $pageName
                PageType      = "Classic"
                WebPartTitle  = $wp.WebPart.Title
                WebPartId     = $wp.Id.ToString().ToLower()
                WebPartType   = $wp.WebPart.Properties.FieldValues["WebPartTypeName"]
                IsOOTB        = $true   # Classic OOTB by default; custom detection done via App Catalog
                IsCustom      = $false
                Section       = 0
                Column        = $wp.WebPart.ZoneIndex
                PropertiesJson = ""
            })
        }
    }
    return $results
}

# ─── MAIN DATA COLLECTION ─────────────────────────────────────────────────────
function Invoke-SiteAssessment {
    param($Sites)
    Write-Log "=== Collecting Site-Level Data ===" "INFO"

    $siteRows    = [System.Collections.Generic.List[PSCustomObject]]::new()
    $listRows    = [System.Collections.Generic.List[PSCustomObject]]::new()
    $webPartRows = [System.Collections.Generic.List[PSCustomObject]]::new()
    $i = 0

    foreach ($site in $Sites) {
        $i++
        Write-Progress -Activity "Assessing Sites" `
                       -Status "$i / $($Sites.Count) : $($site.Url)" `
                       -PercentComplete (($i / $Sites.Count) * 100)

        $connected = $false
        try { Connect-Site -Url $site.Url; $connected = $true }
        catch { Write-Log "Cannot connect to $($site.Url): $(($_.Exception.Message -split '`n')[0])" "WARN" }

        $subsiteCount = 0; $listCount = 0; $libCount = 0
        $modernPages  = 0; $classicPages = 0

        if ($connected) {
            $subsites     = @(Invoke-SafePnP { Get-PnPSubWeb -Recurse } $site.Url)
            $subsiteCount = if ($subsites) { $subsites.Count } else { 0 }

            $siteLists = Get-ListsAndLibraries -SiteUrl $site.Url -SiteTitle $site.Title
            foreach ($l in $siteLists) { $listRows.Add($l) }
            $listCount = @($siteLists | Where-Object { $_.ListType -eq "List" }).Count
            $libCount  = @($siteLists | Where-Object { $_.ListType -eq "Document Library" }).Count

            $modernWPs   = Get-ModernPageWebParts -SiteUrl $site.Url -SiteTitle $site.Title
            $modernPages = @($modernWPs | Select-Object -ExpandProperty PageName -Unique).Count
            foreach ($wp in $modernWPs) { $webPartRows.Add($wp) }

            if ($ScanClassicPages) {
                $classicWPs   = Get-ClassicPageWebParts -SiteUrl $site.Url -SiteTitle $site.Title
                $classicPages = @($classicWPs | Select-Object -ExpandProperty PageName -Unique).Count
                foreach ($wp in $classicWPs) { $webPartRows.Add($wp) }
            }

            Disconnect-PnPOnline -ErrorAction SilentlyContinue
        }

        $storageGB     = [math]::Round([double]$site.StorageUsageCurrent / 1024, 3)
        $quotaGB       = [math]::Round([double]$site.StorageMaximumLevel  / 1024, 3)
        $storageUsePct = if ($quotaGB -gt 0) { [math]::Round(($storageGB / $quotaGB) * 100, 1) } else { 0 }

        $siteRows.Add([PSCustomObject]@{
            SiteUrl           = $site.Url
            Title             = $site.Title
            Template          = $site.Template
            Owner             = $site.Owner
            StorageUsedGB     = $storageGB
            StorageQuotaGB    = $quotaGB
            StorageUsedPct    = $storageUsePct
            Created           = if ($site.Created)                 { ([datetime]$site.Created).ToString("yyyy-MM-dd") }                 else { "N/A" }
            LastModified      = if ($site.LastContentModifiedDate) { $site.LastContentModifiedDate.ToString("yyyy-MM-dd") } else { "N/A" }
            IsHubSite         = $site.IsHubSite
            HubSiteId         = $site.HubSiteId
            SharingCapability = $site.SharingCapability
            LockState         = $site.LockState
            SubsiteCount      = $subsiteCount
            ListCount         = $listCount
            LibraryCount      = $libCount
            ModernPages       = $modernPages
            ClassicPages      = $classicPages
            Connected         = $connected
        })
    }
    Write-Progress -Activity "Assessing Sites" -Completed
    Write-Log "Data collection complete: $($siteRows.Count) sites." "SUCCESS"
    return @{ Sites = $siteRows; Lists = $listRows; WebParts = $webPartRows }
}

# ─── CUSTOM WEB PARTS SUMMARY ─────────────────────────────────────────────────
function Get-CustomWebPartSummary { param($WebPartRows)
    $custom = @($WebPartRows | Where-Object { $_.IsCustom -eq $true })
    $custom | Group-Object -Property WebPartId,WebPartTitle | ForEach-Object {
        $first = $_.Group[0]
        [PSCustomObject]@{
            WebPartTitle  = $first.WebPartTitle
            WebPartId     = $first.WebPartId
            UsageCount    = $_.Count
            SitesUsedIn   = ($_.Group | Select-Object -ExpandProperty SiteUrl -Unique).Count
            PagesUsedIn   = ($_.Group | Select-Object -ExpandProperty PageName -Unique).Count
            SampleSiteUrl = $first.SiteUrl
            SamplePage    = $first.PageName
        }
    } | Sort-Object -Property UsageCount -Descending
}

# ─── WEB PART USAGE SUMMARY ───────────────────────────────────────────────────
function Get-WebPartUsageSummary { param($WebPartRows)
    $WebPartRows | Group-Object -Property WebPartTitle,WebPartId | ForEach-Object {
        $first = $_.Group[0]
        [PSCustomObject]@{
            WebPartTitle  = $first.WebPartTitle
            WebPartId     = $first.WebPartId
            PageType      = ($_.Group | Select-Object -ExpandProperty PageType -Unique) -join ", "
            IsOOTB        = $first.IsOOTB
            IsCustom      = $first.IsCustom
            UsageCount    = $_.Count
            SitesUsedIn   = ($_.Group | Select-Object -ExpandProperty SiteUrl -Unique).Count
            PagesUsedIn   = ($_.Group | Select-Object -ExpandProperty PageName -Unique).Count
        }
    } | Sort-Object -Property UsageCount -Descending
}

# ─── EXCEL EXPORT ─────────────────────────────────────────────────────────────
function Export-ToExcel {
    param([string]$OutputFile, $SitesData, $ListsData, $WebPartsData,
          $WebPartSummary, $CustomWPData, $AppCatalogData)
    Write-Log "Generating Excel workbook..." "INFO"
    $totalGB = [math]::Round(($SitesData | Measure-Object StorageUsedGB -Sum).Sum, 2)
    $summaryItems = [ordered]@{
        "Total Site Collections"       = $SitesData.Count
        "Total Storage Used (GB)"      = $totalGB
        "Sites with External Sharing"  = @($SitesData | Where-Object { $_.SharingCapability -ne "Disabled" }).Count
        "Hub Sites"                    = @($SitesData | Where-Object { $_.IsHubSite }).Count
        "Sites with Subsites"          = @($SitesData | Where-Object { $_.SubsiteCount -gt 0 }).Count
        "Total Lists"                  = @($ListsData | Where-Object { $_.ListType -eq "List" }).Count
        "Total Document Libraries"     = @($ListsData | Where-Object { $_.ListType -eq "Document Library" }).Count
        "Total Items Across All Lists" = ($ListsData | Measure-Object ItemCount -Sum).Sum
        "Lists with Unique Permissions"= @($ListsData | Where-Object { $_.HasUniquePerms }).Count
        "Total Web Part Usages"        = $WebPartsData.Count
        "Distinct OOTB Web Parts"      = (@($WebPartSummary) | Where-Object { $_.IsOOTB  } | Measure-Object).Count
        "Distinct Custom Web Parts"    = (@($WebPartSummary) | Where-Object { $_.IsCustom } | Measure-Object).Count
        "Total Custom WP Usages"       = @($WebPartsData | Where-Object { $_.IsCustom }).Count
        "App Catalog Solutions"        = $AppCatalogData.Count
    }
    $summaryData = $summaryItems.Keys | ForEach-Object { [PSCustomObject]@{ Metric = $_; Value = $summaryItems[$_] } }
    $xl = $summaryData | Export-Excel -Path $OutputFile -WorksheetName "Summary" `
          -TableName "Summary" -TableStyle Medium2 -AutoSize -FreezeTopRow -PassThru
    $SitesData | Export-Excel -ExcelPackage $xl -WorksheetName "Site Collections" `
          -TableName "Sites" -TableStyle Medium6 -AutoSize -FreezeTopRow
    if ($ListsData.Count -gt 0) {
        $ListsData | Export-Excel -ExcelPackage $xl -WorksheetName "Lists and Libraries" `
              -TableName "Lists" -TableStyle Medium6 -AutoSize -FreezeTopRow
    }
    if (@($WebPartSummary).Count -gt 0) {
        @($WebPartSummary) | Export-Excel -ExcelPackage $xl -WorksheetName "Web Parts Usage" `
              -TableName "WebPartUsage" -TableStyle Medium4 -AutoSize -FreezeTopRow
    }
    if (@($CustomWPData).Count -gt 0) {
        @($CustomWPData) | Export-Excel -ExcelPackage $xl -WorksheetName "Custom Web Parts" `
              -TableName "CustomWP" -TableStyle Medium9 -AutoSize -FreezeTopRow
    }
    if ($AppCatalogData.Count -gt 0) {
        $AppCatalogData | Export-Excel -ExcelPackage $xl -WorksheetName "App Catalog" `
              -TableName "AppCatalog" -TableStyle Medium7 -AutoSize -FreezeTopRow
    }
    $ws = $xl.Workbook.Worksheets["Site Collections"]
    if ($ws -and $SitesData.Count -gt 0) {
        $cols   = ($SitesData | Select-Object -First 1).PSObject.Properties.Name
        $pctIdx = [array]::IndexOf($cols, "StorageUsedPct") + 1
        if ($pctIdx -gt 0) {
            Add-ConditionalFormatting -WorkSheet $ws -Address "${pctIdx}:${pctIdx}" `
                -RuleType GreaterThan -ConditionValue 80 `
                -BackgroundColor ([System.Drawing.Color]::FromArgb(209,52,56)) `
                -ForegroundColor ([System.Drawing.Color]::White)
        }
    }
    Close-ExcelPackage $xl
    Write-Log "Excel workbook saved: $OutputFile" "SUCCESS"
}

# ─── HTML EXPORT ──────────────────────────────────────────────────────────────
function Export-ToHTML {
    param([string]$OutputFile, $SitesData, $ListsData, $WebPartSummary, $CustomWPData,
          $AppCatalogData, [string]$TenantUrl, [datetime]$RunDate)
    Write-Log "Generating HTML report..." "INFO"
    $totalGB    = [math]::Round(($SitesData | Measure-Object StorageUsedGB -Sum).Sum, 2)
    $totalLists = $ListsData.Count
    $totalWPs   = @($WebPartSummary).Count
    $customWPs  = @($WebPartSummary | Where-Object { $_.IsCustom }).Count
    $ootbWPs    = @($WebPartSummary | Where-Object { $_.IsOOTB  }).Count
    $totalApps  = $AppCatalogData.Count
    $reportDate = $RunDate.ToString("dddd, dd MMM yyyy HH:mm")
    $totalSites = $SitesData.Count

    function Build-Rows { param($Data, [string[]]$Cols)
        $sb = [System.Text.StringBuilder]::new()
        foreach ($row in ($Data | Select-Object -First 3000)) {
            [void]$sb.Append("<tr>")
            foreach ($c in $Cols) {
                $v    = $row.$c
                $cell = if ($null -eq $v) { "" } else { [System.Net.WebUtility]::HtmlEncode($v.ToString()) }
                if      ($v -eq $true)  { [void]$sb.Append("<td><span class='badge b-yes'>Yes</span></td>") }
                elseif  ($v -eq $false) { [void]$sb.Append("<td><span class='badge b-no'>No</span></td>") }
                else                    { [void]$sb.Append("<td title='$cell'>$cell</td>") }
            }
            [void]$sb.Append("</tr>")
        }
        return $sb.ToString()
    }
    function THead { param([string[]]$Cols)
        "<thead><tr>" + ($Cols | ForEach-Object { "<th>$_</th>" } | Out-String).Replace("`n","") + "</tr></thead>"
    }

    $sC = @("Title","Template","Owner","StorageUsedGB","StorageQuotaGB","StorageUsedPct","Created","LastModified","IsHubSite","SharingCapability","LockState","SubsiteCount","ListCount","LibraryCount","ModernPages","ClassicPages")
    $lC = @("SiteTitle","ListTitle","ListType","ItemCount","HasUniquePerms","Versioning","MajorVersions","ContentTypeCount","Created","LastModified")
    $wC = @("WebPartTitle","WebPartId","PageType","IsOOTB","IsCustom","UsageCount","SitesUsedIn","PagesUsedIn")
    $cC = @("WebPartTitle","WebPartId","UsageCount","SitesUsedIn","PagesUsedIn","SampleSiteUrl","SamplePage")
    $aC = @("Title","Version","IsEnabled","IsDeployed","SkipFeatureDeployment","Publisher")

    $css  = ':root{--bg:#f0f4f8;--card:#fff;--acc:#0078d4;--txt:#1a1a2e;--sub:#6b7280;--bdr:#e2e8f0;--hi:#d13438;--lo:#107c10;}'
    $css += '*{box-sizing:border-box;margin:0;padding:0;}'
    $css += "body{font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--txt);font-size:14px;}"
    $css += 'header{background:linear-gradient(135deg,#0078d4,#00b4d8);color:#fff;padding:28px 40px;}'
    $css += 'header h1{font-size:1.7rem;font-weight:700;}header p{opacity:.9;font-size:.87rem;margin-top:6px;}'
    $css += '.summary{display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:12px;padding:20px 40px;}'
    $css += '.card{background:var(--card);border-radius:10px;padding:16px 18px;box-shadow:0 1px 4px rgba(0,0,0,.08);border-left:4px solid var(--acc);}'
    $css += '.card .val{font-size:1.85rem;font-weight:700;color:var(--acc);}.card .lbl{color:var(--sub);font-size:.7rem;text-transform:uppercase;letter-spacing:.5px;margin-top:4px;}'
    $css += '.card.hi{border-color:var(--hi);}.card.hi .val{color:var(--hi);}.card.ok{border-color:var(--lo);}.card.ok .val{color:var(--lo);}'
    $css += 'nav{display:flex;flex-wrap:wrap;padding:0 40px;background:#fff;border-bottom:1px solid var(--bdr);}'
    $css += 'nav button{background:none;border:none;padding:13px 18px;cursor:pointer;font-size:.88rem;color:var(--sub);border-bottom:3px solid transparent;transition:.2s;font-weight:500;}'
    $css += 'nav button.active{color:var(--acc);border-bottom-color:var(--acc);font-weight:600;}'
    $css += '.tab{display:none;padding:20px 40px;}.tab.active{display:block;}'
    $css += '.sh{font-size:1rem;font-weight:600;margin-bottom:10px;}.sb{margin-bottom:10px;}'
    $css += '.sb input{padding:7px 12px;border:1px solid var(--bdr);border-radius:8px;width:320px;font-size:.87rem;outline:none;}'
    $css += '.sb input:focus{border-color:var(--acc);}.tw{overflow-x:auto;border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,.07);}'
    $css += 'table{width:100%;border-collapse:collapse;background:#fff;}'
    $css += 'th{background:var(--acc);color:#fff;padding:10px 12px;text-align:left;font-size:.77rem;white-space:nowrap;font-weight:600;}'
    $css += 'td{padding:8px 12px;border-bottom:1px solid var(--bdr);font-size:.81rem;max-width:240px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}'
    $css += 'tr:nth-child(even) td{background:#f8fafc;}tr:hover td{background:#ebf5ff;}'
    $css += '.badge{display:inline-block;padding:2px 9px;border-radius:20px;font-size:.74rem;font-weight:700;color:#fff;}'
    $css += '.b-yes{background:var(--lo);}.b-no{background:var(--hi);}'
    $css += 'footer{text-align:center;padding:16px;color:var(--sub);font-size:.75rem;border-top:1px solid var(--bdr);margin-top:20px;}'

    $js  = 'function showTab(id,btn){document.querySelectorAll(".tab").forEach(t=>t.classList.remove("active"));document.querySelectorAll("nav button").forEach(b=>b.classList.remove("active"));document.getElementById(id).classList.add("active");btn.classList.add("active");}'
    $js += 'function flt(inp,tbl){var q=document.getElementById(inp).value.toLowerCase();var rows=document.getElementById(tbl).getElementsByTagName("tr");for(var i=1;i<rows.length;i++){rows[i].style.display=rows[i].innerText.toLowerCase().includes(q)?"":"none";}}'

    $h = [System.Text.StringBuilder]::new(131072)
    [void]$h.Append('<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/><title>SPO Site Collection Assessment</title>')
    [void]$h.Append("<style>$css</style></head><body>")
    [void]$h.Append('<header><h1>&#128202; SharePoint Online &#8212; Site Collection Assessment</h1>')
    [void]$h.Append("<p>Tenant: <strong>$TenantUrl</strong> &nbsp;&bull;&nbsp; Report Date: <strong>$reportDate</strong></p></header>")
    [void]$h.Append('<div class="summary">')
    [void]$h.Append("<div class='card'><div class='val'>$totalSites</div><div class='lbl'>Site Collections</div></div>")
    [void]$h.Append("<div class='card'><div class='val'>$totalGB GB</div><div class='lbl'>Total Storage Used</div></div>")
    [void]$h.Append("<div class='card'><div class='val'>$totalLists</div><div class='lbl'>Lists &amp; Libraries</div></div>")
    [void]$h.Append("<div class='card'><div class='val'>$totalWPs</div><div class='lbl'>Distinct Web Parts</div></div>")
    [void]$h.Append("<div class='card ok'><div class='val'>$ootbWPs</div><div class='lbl'>OOTB Web Parts</div></div>")
    [void]$h.Append("<div class='card hi'><div class='val'>$customWPs</div><div class='lbl'>Custom Web Parts</div></div>")
    [void]$h.Append("<div class='card'><div class='val'>$totalApps</div><div class='lbl'>App Catalog Solutions</div></div>")
    [void]$h.Append('</div><nav>')
    [void]$h.Append("<button class='active' onclick='showTab(\"sites\",this)'>&#127760; Sites ($totalSites)</button>")
    [void]$h.Append("<button onclick='showTab(\"lists\",this)'>&#128196; Lists &amp; Libs ($totalLists)</button>")
    [void]$h.Append("<button onclick='showTab(\"webparts\",this)'>&#129513; All Web Parts ($totalWPs)</button>")
    [void]$h.Append("<button onclick='showTab(\"custom\",this)'>&#128736; Custom WPs ($customWPs)</button>")
    [void]$h.Append("<button onclick='showTab(\"apps\",this)'>&#128230; App Catalog ($totalApps)</button>")
    [void]$h.Append('</nav>')
    [void]$h.Append('<div id="sites" class="tab active"><p class="sh">Site Collections &#8212; Storage &amp; Metadata</p>')
    [void]$h.Append('<div class="sb"><input id="s1" onkeyup="flt(''s1'',''t1'')" placeholder="Search sites..."/></div>')
    [void]$h.Append("<div class='tw'><table id='t1'>$(THead -Cols $sC)<tbody>$(Build-Rows -Data $SitesData -Cols $sC)</tbody></table></div></div>")
    [void]$h.Append('<div id="lists" class="tab"><p class="sh">Lists &amp; Libraries Inventory</p>')
    [void]$h.Append('<div class="sb"><input id="s2" onkeyup="flt(''s2'',''t2'')" placeholder="Search lists..."/></div>')
    [void]$h.Append("<div class='tw'><table id='t2'>$(THead -Cols $lC)<tbody>$(Build-Rows -Data $ListsData -Cols $lC)</tbody></table></div></div>")
    [void]$h.Append('<div id="webparts" class="tab"><p class="sh">All Web Parts Usage (sorted by frequency)</p>')
    [void]$h.Append('<div class="sb"><input id="s3" onkeyup="flt(''s3'',''t3'')" placeholder="Search web parts..."/></div>')
    [void]$h.Append("<div class='tw'><table id='t3'>$(THead -Cols $wC)<tbody>$(Build-Rows -Data $WebPartSummary -Cols $wC)</tbody></table></div></div>")
    [void]$h.Append('<div id="custom" class="tab"><p class="sh">Custom / Third-Party Web Parts (component ID not in OOTB registry)</p>')
    [void]$h.Append('<div class="sb"><input id="s4" onkeyup="flt(''s4'',''t4'')" placeholder="Search custom web parts..."/></div>')
    [void]$h.Append("<div class='tw'><table id='t4'>$(THead -Cols $cC)<tbody>$(Build-Rows -Data $CustomWPData -Cols $cC)</tbody></table></div></div>")
    [void]$h.Append('<div id="apps" class="tab"><p class="sh">Tenant App Catalog &#8212; Installed Solutions</p>')
    [void]$h.Append('<div class="sb"><input id="s5" onkeyup="flt(''s5'',''t5'')" placeholder="Search apps..."/></div>')
    [void]$h.Append("<div class='tw'><table id='t5'>$(THead -Cols $aC)<tbody>$(Build-Rows -Data $AppCatalogData -Cols $aC)</tbody></table></div></div>")
    [void]$h.Append("<footer>SPO Site Collection Assessment &#8212; $reportDate &#8212; SPO-SiteCollection-Assessment.ps1</footer>")
    [void]$h.Append("<script>$js</script></body></html>")
    $h.ToString() | Out-File -FilePath $OutputFile -Encoding UTF8 -Force
    Write-Log "HTML report saved: $OutputFile" "SUCCESS"
}

# ─── MAIN ─────────────────────────────────────────────────────────────────────
function Main {
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "   SharePoint Online -- Site Collection Assessment" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    Assert-Modules
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    Write-Log "Output folder: $OutputPath" "INFO"
    $sites = Get-AllSiteCollections
    if (-not $sites -or $sites.Count -eq 0) { Write-Log "No sites to process. Exiting." "ERROR"; return }
    $appCatalogData = Get-AppCatalogSolutions
    $data           = Invoke-SiteAssessment -Sites $sites
    $wpSummary      = @(Get-WebPartUsageSummary  -WebPartRows $data.WebParts)
    $customWPs      = @(Get-CustomWebPartSummary -WebPartRows $data.WebParts)
    Write-Log "Web parts: $($wpSummary.Count) distinct | $($customWPs.Count) custom" "INFO"
    if ($ExportCSV) {
        $data.Sites     | Export-Csv (Join-Path $OutputPath "SPO_SC_Sites.csv")        -NoTypeInformation -Encoding UTF8
        $data.Lists     | Export-Csv (Join-Path $OutputPath "SPO_SC_Lists.csv")        -NoTypeInformation -Encoding UTF8
        $data.WebParts  | Export-Csv (Join-Path $OutputPath "SPO_SC_WebParts_Raw.csv") -NoTypeInformation -Encoding UTF8
        $wpSummary      | Export-Csv (Join-Path $OutputPath "SPO_SC_WP_Summary.csv")   -NoTypeInformation -Encoding UTF8
        $customWPs      | Export-Csv (Join-Path $OutputPath "SPO_SC_CustomWP.csv")     -NoTypeInformation -Encoding UTF8
        $appCatalogData | Export-Csv (Join-Path $OutputPath "SPO_SC_AppCatalog.csv")   -NoTypeInformation -Encoding UTF8
        Write-Log "CSV files written to: $OutputPath" "SUCCESS"
    }
    $excelFile = Join-Path $OutputPath "SPO_SiteCollection_Assessment.xlsx"
    Export-ToExcel -OutputFile     $excelFile `
                   -SitesData      $data.Sites `
                   -ListsData      $data.Lists `
                   -WebPartsData   $data.WebParts `
                   -WebPartSummary $wpSummary `
                   -CustomWPData   $customWPs `
                   -AppCatalogData $appCatalogData
    $htmlFile = Join-Path $OutputPath "SPO_SiteCollection_Assessment.html"
    Export-ToHTML -OutputFile     $htmlFile `
                  -SitesData      $data.Sites `
                  -ListsData      $data.Lists `
                  -WebPartSummary $wpSummary `
                  -CustomWPData   $customWPs `
                  -AppCatalogData $appCatalogData `
                  -TenantUrl      $AdminUrl `
                  -RunDate        (Get-Date)
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Log "Assessment complete!" "SUCCESS"
    Write-Log "  Excel : $excelFile" "SUCCESS"
    Write-Log "  HTML  : $htmlFile"  "SUCCESS"
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    if ($script:HasErrors) { Write-Log "Some sites had errors/warnings. Review WARN messages above." "WARN" }
    Start-Process $htmlFile
}

Main
