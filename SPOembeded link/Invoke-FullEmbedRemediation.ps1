<#
.SYNOPSIS
    End-to-end multi-site orchestrator: crawls all Google Sites, then migrates embeds
    to each corresponding SharePoint Online site.

.DESCRIPTION
    Reads a mapping CSV (default: SelectedSites.csv) where every row defines:
      - SiteUrl      Google Sites source URL (crawled for embedded content)
      - SiteName     Friendly label used in log output and per-site report filenames
      - SPOSiteUrl   Destination SharePoint Online site URL

    The script runs one unified crawl across all Google Sites, then loops through
    each source->destination pair to apply the discovered embeds to the correct SPO site.

    Prerequisites (one-time setup in this folder):
      npm install
      npx playwright install chromium
      node Save-GoogleAuth.js

.PARAMETER MappingCsv
    Path to the source->destination mapping CSV.
    Required columns : SiteUrl, SPOSiteUrl
    Optional columns : SiteName
    Defaults to SelectedSites.csv in this folder.

.PARAMETER ClientId
    Azure AD App Registration ClientId for PnP PowerShell auth.

.PARAMETER TenantId
    Tenant name (e.g. contoso.onmicrosoft.com). Derived from the first SPOSiteUrl if omitted.

.PARAMETER SkipGoogleAuth
    Skip the Google sign-in step if .auth\state.json is already fresh.

.PARAMETER SkipCrawl
    Skip the crawl and use existing output\08_Embeds_Enhanced.csv.

.PARAMETER DryRun
    Preview only -- no changes are written to SharePoint.

.EXAMPLE
    .\Invoke-FullEmbedRemediation.ps1

.EXAMPLE
    .\Invoke-FullEmbedRemediation.ps1 -MappingCsv ".\SelectedSites.csv" -DryRun

.EXAMPLE
    .\Invoke-FullEmbedRemediation.ps1 -SkipCrawl -DryRun
#>
[CmdletBinding()]
param(
    [string]$MappingCsv     = "",
    [string]$ClientId       = "3834b2e7-ab80-45fc-b4c8-ed5c960076b7",
    [string]$TenantId       = "",
    [switch]$SkipGoogleAuth,
    [switch]$SkipCrawl,
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"
$OutputDir = Join-Path $PSScriptRoot "output"
if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir | Out-Null }

# ── PREREQUISITE CHECKS ───────────────────────────────────────────────────────
Write-Host "`n=== PREREQUISITE CHECKS ===" -ForegroundColor Cyan

function Test-Prereq { param([string]$Cmd,[string]$Name)
    if (-not (Get-Command $Cmd -ErrorAction SilentlyContinue)) { throw "$Name is required but not found in PATH." }
}
Test-Prereq "node" "Node.js"
Test-Prereq "npm"  "npm"

if (-not (Test-Path (Join-Path $PSScriptRoot "node_modules"))) {
    throw "node_modules not found. Run in this folder first:`n  npm install`n  npx playwright install chromium"
}
if (-not (Get-Module -ListAvailable -Name "PnP.PowerShell")) {
    throw "PnP.PowerShell is required. Install: Install-Module PnP.PowerShell -Force"
}
Import-Module PnP.PowerShell

# ── LOAD AND VALIDATE MAPPING CSV ─────────────────────────────────────────────
$csvPath = if ($MappingCsv -and (Test-Path $MappingCsv)) {
    (Resolve-Path $MappingCsv).Path
} elseif (Test-Path (Join-Path $PSScriptRoot "SelectedSites.csv")) {
    (Resolve-Path (Join-Path $PSScriptRoot "SelectedSites.csv")).Path
} else {
    throw "No mapping CSV found. Create SelectedSites.csv or pass -MappingCsv. Required columns: SiteUrl, SPOSiteUrl"
}

$siteMap = Import-Csv $csvPath
Write-Host "Mapping CSV : $csvPath" -ForegroundColor Cyan
Write-Host "Site pairs  : $($siteMap.Count)" -ForegroundColor Cyan

$cols = $siteMap[0].PSObject.Properties.Name
if ($cols -notcontains 'SiteUrl')    { throw "CSV is missing required column 'SiteUrl'." }
if ($cols -notcontains 'SPOSiteUrl') { throw "CSV is missing required column 'SPOSiteUrl'. Add the destination SPO URL for each row." }

$validPairs = $siteMap | Where-Object {
    -not [string]::IsNullOrWhiteSpace($_.SiteUrl) -and
    -not [string]::IsNullOrWhiteSpace($_.SPOSiteUrl)
}
if ($validPairs.Count -eq 0) { throw "No valid source->destination pairs found in $csvPath." }

# Derive tenant from first SPO URL (overridden by -TenantId)
$firstSPO    = $validPairs[0].SPOSiteUrl
$spoHost     = ([uri]$firstSPO).Host
$tenantName  = if ($TenantId) { $TenantId } else { ($spoHost -replace "\.sharepoint\.com$","") + ".onmicrosoft.com" }
Write-Host "Tenant      : $tenantName" -ForegroundColor Green

$validPairs | Format-Table @{L='Google Site (Source)';E={if ($_.SiteName) {$_.SiteName} else {$_.SiteUrl}}},
                            @{L='SPO Site (Destination)';E={$_.SPOSiteUrl}} -AutoSize

# ── STEP 1: Google Auth ───────────────────────────────────────────────────────
if (-not $SkipGoogleAuth) {
    Write-Host "`n=== STEP 1: Google Sites Browser Auth ===" -ForegroundColor Cyan
    Write-Host "A Chromium browser window will open. Sign in to Google, then press Enter here." -ForegroundColor Yellow
    Push-Location $PSScriptRoot
    & node "Save-GoogleAuth.js"
    if ($LASTEXITCODE -ne 0) { throw "Auth script failed." }
    Pop-Location
    Write-Host "Auth state saved." -ForegroundColor Green
}

# ── STEP 2: Crawl all Google Sites ───────────────────────────────────────────
if (-not $SkipCrawl) {
    Write-Host "`n=== STEP 2: Crawling Google Sites ===" -ForegroundColor Cyan
    Write-Host "Crawling $($validPairs.Count) site(s) listed in: $csvPath" -ForegroundColor Cyan
    Push-Location $PSScriptRoot
    & node "03_crawl_sites_enhanced.js" $csvPath
    $crawlExit = $LASTEXITCODE
    Pop-Location
    if ($crawlExit -ne 0) {
        Write-Warning "Crawl exited with code $crawlExit - partial results may still be available."
    } else {
        Write-Host "Crawl complete. Outputs in: $OutputDir" -ForegroundColor Green
    }
}

# ── STEP 3: Load crawl output ─────────────────────────────────────────────────
Write-Host "`n=== STEP 3: Loading crawl output ===" -ForegroundColor Cyan
$EmbedsCsv = Join-Path $OutputDir "08_Embeds_Enhanced.csv"
$PagesCsv  = Join-Path $OutputDir "07_Pages_Enhanced.csv"

if (-not (Test-Path $EmbedsCsv)) { throw "Embeds CSV not found: $EmbedsCsv. Re-run without -SkipCrawl." }

$allEmbeds  = Import-Csv $EmbedsCsv
$allPages   = if (Test-Path $PagesCsv) { Import-Csv $PagesCsv } else { @() }

$allRealEmbeds = $allEmbeds | Where-Object {
    ($_.ItemKind -in @('iframe','embed')) -and
    ($_.ArtifactUrl -notmatch 'accounts\.google\.com|bscframe|recaptcha|google\.com/signin')
}
Write-Host "Total content embeds found across all sites: $($allRealEmbeds.Count)" -ForegroundColor Green

# ── STEPS 4-6: Per-site: match embeds -> connect SPO -> apply web parts ───────
$addScript      = Join-Path $PSScriptRoot "Add-SPOYouTubeWebParts.ps1"
if (-not (Test-Path $addScript)) { throw "Add-SPOYouTubeWebParts.ps1 not found in $PSScriptRoot" }

$masterMapping  = [System.Collections.Generic.List[pscustomobject]]::new()
$siteResults    = [System.Collections.Generic.List[pscustomobject]]::new()
$siteIndex      = 0

foreach ($pair in $validPairs) {
    $siteIndex++
    $googleUrl  = $pair.SiteUrl.Trim()
    $spoUrl     = $pair.SPOSiteUrl.Trim()
    $siteName   = if (-not [string]::IsNullOrWhiteSpace($pair.SiteName)) { $pair.SiteName.Trim() } else { $googleUrl }
    $safeLabel  = $siteName -replace '[\\/:*?"<>|]','-'

    Write-Host "`n=== [$siteIndex/$($validPairs.Count)] Processing: $siteName ===" -ForegroundColor Cyan
    Write-Host "  Source : $googleUrl" -ForegroundColor DarkCyan
    Write-Host "  Dest   : $spoUrl"   -ForegroundColor DarkCyan

    # Filter embeds belonging to this Google Site
    $siteEmbeds = $allRealEmbeds | Where-Object { $_.SiteUrl -eq $googleUrl }
    Write-Host "  Embeds for this site: $($siteEmbeds.Count)" -ForegroundColor $(if ($siteEmbeds.Count -gt 0) {'Green'} else {'Yellow'})

    if ($siteEmbeds.Count -eq 0) {
        Write-Warning "  No embeds found for '$siteName'. Skipping SPO update."
        $siteResults.Add([pscustomobject]@{
            SiteName    = $siteName
            GoogleUrl   = $googleUrl
            SPOUrl      = $spoUrl
            EmbedsFound = 0
            EmbedsMapped= 0
            Status      = 'Skipped - no embeds'
        })
        continue
    }

    # ── Connect to this SPO site ──────────────────────────────────────────────
    Write-Host "  Connecting to SPO..." -ForegroundColor Cyan
    try {
        Connect-PnPOnline -Url $spoUrl -ClientId $ClientId -Tenant $tenantName -Interactive -ErrorAction Stop
    } catch {
        Write-Warning "  Failed to connect to $spoUrl : $_"
        $siteResults.Add([pscustomobject]@{
            SiteName    = $siteName; GoogleUrl = $googleUrl; SPOUrl = $spoUrl
            EmbedsFound = $siteEmbeds.Count; EmbedsMapped = 0; Status = "Connection failed: $_"
        })
        continue
    }

    # ── Get SPO pages for this site ───────────────────────────────────────────
    $spoPages = Get-PnPListItem -List "SitePages" -PageSize 500 | ForEach-Object {
        [pscustomobject]@{
            FileName = $_.FieldValues["FileLeafRef"]
            Title    = $_.FieldValues["Title"]
        }
    }
    Write-Host "  SPO pages retrieved: $($spoPages.Count)" -ForegroundColor Cyan

    # ── Build embed mapping for this site ────────────────────────────────────
    $siteMapping   = [System.Collections.Generic.List[pscustomobject]]::new()
    $orderTracker  = @{}

    foreach ($row in $siteEmbeds) {
        $gsTitle = ($row.PageTitle -replace '\s+',' ').Trim()
        if ($gsTitle -match '^Error\s+\d') {
            $gsTitle = ($allPages | Where-Object { $_.PageUrl -eq $row.PageUrl } | Select-Object -First 1).PageTitle
        }

        # Exact match on Title or filename stem
        $match = $spoPages | Where-Object {
            ($_.Title -and $_.Title.Trim() -eq $gsTitle) -or
            ($_.FileName -replace '\.aspx$','' -replace '[-_]',' ').Trim() -eq ($gsTitle -replace '[-_]',' ').Trim()
        } | Select-Object -First 1

        # Fallback: partial contains match
        if (-not $match) {
            $match = $spoPages | Where-Object {
                $_.Title -and ($gsTitle -like "*$($_.Title.Trim())*" -or $_.Title.Trim() -like "*$gsTitle*")
            } | Select-Object -First 1
        }

        if (-not $match) {
            Write-Warning "  No SPO page match for '$gsTitle' in '$siteName'. Skipping: $($row.ArtifactUrl)"
            continue
        }

        $key = $match.FileName
        if (-not $orderTracker.ContainsKey($key)) { $orderTracker[$key] = 0 } else { $orderTracker[$key] += 1 }

        $mapRow = [pscustomobject]@{
            SiteName       = $siteName
            GoogleSiteUrl  = $googleUrl
            SPOSiteUrl     = $spoUrl
            PageName       = $match.FileName
            EmbedUrl       = $row.ArtifactUrl
            SectionIndex   = 1
            ColumnIndex    = 1
            Order          = $orderTracker[$key]
            GSitePageTitle = $gsTitle
            ArtifactType   = $row.ArtifactType
        }
        $siteMapping.Add($mapRow)
        $masterMapping.Add($mapRow)
    }

    Write-Host "  Embeds mapped: $($siteMapping.Count) / $($siteEmbeds.Count)" -ForegroundColor $(if ($siteMapping.Count -gt 0) {'Green'} else {'Yellow'})

    # ── Write per-site mapping CSV ────────────────────────────────────────────
    $perSiteCsv = Join-Path $OutputDir "EmbedMapping_$safeLabel.csv"
    $siteMapping | Export-Csv -Path $perSiteCsv -NoTypeInformation
    Write-Host "  Per-site mapping: $perSiteCsv" -ForegroundColor Cyan

    if ($siteMapping.Count -eq 0) {
        Write-Warning "  No embeds could be matched to SPO pages for '$siteName'."
        Disconnect-PnPOnline
        $siteResults.Add([pscustomobject]@{
            SiteName = $siteName; GoogleUrl = $googleUrl; SPOUrl = $spoUrl
            EmbedsFound = $siteEmbeds.Count; EmbedsMapped = 0; Status = 'No page title matches'
        })
        continue
    }

    $siteMapping | Format-Table PageName, ArtifactType, EmbedUrl -AutoSize

    # ── Apply web parts to this SPO site ─────────────────────────────────────
    if (-not $DryRun) {
        $invokeParams = @{
            SiteUrl    = $spoUrl
            MappingCsv = $perSiteCsv
            ClientId   = $ClientId
            TenantId   = $tenantName
        }
        # Add-SPOYouTubeWebParts connects internally; disconnect first to avoid double-connect
        Disconnect-PnPOnline
        & $addScript @invokeParams
    } else {
        Write-Host "  [DRY RUN] Would apply $($siteMapping.Count) embed(s) to $spoUrl" -ForegroundColor Magenta
        Disconnect-PnPOnline
    }

    $siteResults.Add([pscustomobject]@{
        SiteName    = $siteName
        GoogleUrl   = $googleUrl
        SPOUrl      = $spoUrl
        EmbedsFound = $siteEmbeds.Count
        EmbedsMapped= $siteMapping.Count
        Status      = if ($DryRun) { 'DryRun - not applied' } else { 'Applied' }
    })
}

# ── Write master mapping CSV ───────────────────────────────────────────────────
$masterCsv = Join-Path $PSScriptRoot "EmbedMapping.csv"
$masterMapping | Export-Csv -Path $masterCsv -NoTypeInformation
Write-Host "`nMaster mapping written to: $masterCsv ($($masterMapping.Count) total rows)" -ForegroundColor Cyan

# ── Final summary ─────────────────────────────────────────────────────────────
Write-Host "`n=== REMEDIATION SUMMARY ===" -ForegroundColor Cyan
$siteResults | Format-Table SiteName, EmbedsFound, EmbedsMapped, Status -AutoSize

$totalApplied = ($siteResults | Measure-Object -Property EmbedsMapped -Sum).Sum
Write-Host "Sites processed : $($siteResults.Count)" -ForegroundColor Green
Write-Host "Total embeds    : $($masterMapping.Count) mapped / $($allRealEmbeds.Count) found" -ForegroundColor Green
if ($DryRun) {
    Write-Host "DRY RUN complete. Re-run without -DryRun to apply changes." -ForegroundColor Magenta
} else {
    Write-Host "Done. Per-site reports are in: $OutputDir" -ForegroundColor Green
}
