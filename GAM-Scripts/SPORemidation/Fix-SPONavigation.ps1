# =============================================================================
# Fix-SPONavigation.ps1
# Fixes broken navigation links after Google Sites → SharePoint Online migration
# Can be run standalone (single site) OR dot-sourced by Invoke-BulkNavFix.ps1
# Uses PnP PowerShell: Install-Module PnP.PowerShell -Scope CurrentUser
# =============================================================================

param(
    [Parameter(Mandatory = $false)]
    [string]$SiteUrl,                  # e.g. https://mngenvmcap908272.sharepoint.com/sites/ftcmigrationtestsite

    [Parameter(Mandatory = $false)]
    [string]$GoogleSitesBaseUrl = "",  # e.g. https://sites.google.com/censftmigsme.microsoft-int.com/ftcmigrationtestsite

    [Parameter(Mandatory = $false)]
    [string]$ClientId = "3834b2e7-ab80-45fc-b4c8-ed5c960076b7",  # Azure AD App Registration Client ID

    [Parameter(Mandatory = $false)]
    [string]$TenantId = "",            

    [switch]$RebuildNavigation,        
    [switch]$WhatIf                   
)

# ── Helper: timestamped console log ──────────────────────────────────────────
function Write-Log {
    param([string]$Message, [string]$Level = "INFO", [string]$LogFile = "")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts][$Level] $Message"
    $color = switch ($Level) {
        "WARN" { "Yellow" }
        "ERROR" { "Red" }
        "SUCCESS" { "Green" }
        default { "Cyan" }
    }
    Write-Host $line -ForegroundColor $color
    if ($LogFile) { Add-Content -Path $LogFile -Value $line }
}

# =============================================================================
# Invoke-SPONavFix — core function (used standalone AND by bulk runner)
# =============================================================================
function Invoke-SPONavFix {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SPOSiteUrl,
        [string]$GSitesBaseUrl = "",
        [string]$ClientId = "3834b2e7-ab80-45fc-b4c8-ed5c960076b7",
        [string]$TenantId = "",
        [bool]  $Rebuild = $false,
        [bool]  $DryRun = $false,
        [string]$LogFile = "",
        [switch]$UseCurrentConnection   # When set, skips Connect/Disconnect (orchestrator owns the session)
    )

    # Derive the site-relative prefix dynamically, e.g. "/sites/hrportal"
    $siteSlug = $SPOSiteUrl -replace "^https?://[^/]+/sites/", "" -replace "/.*", ""
    $siteBase = "/sites/$siteSlug"
    $sitePagesBase = "$siteBase/SitePages"

    # ── Default nav structure — mirrors Google Sites top-nav ─────────────────
    # Edit child pages per site, or pass a custom $NavStructure from the caller.
    $NavStructure = @(
        @{ Title = "Home"; Url = "$siteBase/"; Children = @() },
        @{ Title = "Text Formatting"; Url = "$sitePagesBase/Text-Formatting.aspx"; Children = @() },
        @{ Title = "Resources"; Url = "$sitePagesBase/Resources.aspx";
            Children = @(
                @{ Title = "File Cabinet"; Url = "$sitePagesBase/File-Cabinet.aspx"; Children = @() },
                @{ Title = "Announcements"; Url = "$sitePagesBase/Announcements.aspx"; Children = @() }
            )
        },
        @{ Title = "HR Policies"; Url = "$sitePagesBase/HR-Policies.aspx";
            Children = @(
                @{ Title = "Onboarding"; Url = "$sitePagesBase/Onboarding.aspx"; Children = @() },
                @{ Title = "Benefits"; Url = "$sitePagesBase/Benefits.aspx"; Children = @() }
            )
        },
        @{ Title = "IT Help"; Url = "$sitePagesBase/IT-Help.aspx";
            Children = @(
                @{ Title = "Support"; Url = "$sitePagesBase/Support.aspx"; Children = @() }
            )
        },
        @{ Title = "Projects"; Url = "$sitePagesBase/Projects.aspx"; Children = @() },
        @{ Title = "Team Wiki"; Url = "$sitePagesBase/Team-Wiki.aspx";
            Children = @(
                @{ Title = "Guidelines"; Url = "$sitePagesBase/Guidelines.aspx"; Children = @() }
            )
        }
    )

    # ── Connect ───────────────────────────────────────────────────────────────
    # Auto-derive tenant from SPO URL if not explicitly supplied
    # e.g. https://mngenvmcap908272.sharepoint.com/... -> mngenvmcap908272.onmicrosoft.com
    $tenantName = if ($TenantId) {
        $TenantId
    }
    else {
        $spoHost = ([uri]$SPOSiteUrl).Host                     # mngenvmcap908272.sharepoint.com
        $prefix = $spoHost -replace "\.sharepoint\.com$", ""  # mngenvmcap908272
        "$prefix.onmicrosoft.com"
    }

    if (-not $UseCurrentConnection) {
        Write-Log "[$siteSlug] Connecting to $SPOSiteUrl (Tenant: $tenantName | ClientId: $ClientId) ..." -LogFile $LogFile
        try {
            Connect-PnPOnline -Url $SPOSiteUrl -ClientId $ClientId -Tenant $tenantName -Interactive -ErrorAction Stop
            Write-Log "[$siteSlug] Connected." "SUCCESS" $LogFile
        }
        catch {
            Write-Log "[$siteSlug] Connection failed: $_" "ERROR" $LogFile
            return [PSCustomObject]@{ Site = $siteSlug; Status = "FAILED"; Error = $_.ToString() }
        }
    }

    # ── MODE A: Full Rebuild ──────────────────────────────────────────────────
    if ($Rebuild) {
        Write-Log "[$siteSlug] Rebuild mode — removing existing top-nav nodes ..." -LogFile $LogFile
        if (-not $DryRun) {
            Get-PnPNavigationNode -Location TopNavigationBar |
            ForEach-Object { Remove-PnPNavigationNode -Identity $_.Id -Force }
        }
        else {
            Write-Log "[$siteSlug] [DryRun] Would remove all top-nav nodes." "WARN" $LogFile
        }
        foreach ($item in $NavStructure) {
            Write-Log "[$siteSlug]  + $($item.Title) -> $($item.Url)" -LogFile $LogFile
            if (-not $DryRun) {
                $parentNode = Add-PnPNavigationNode -Location TopNavigationBar `
                    -Title $item.Title -Url $item.Url
                foreach ($child in $item.Children) {
                    Write-Log "[$siteSlug]      -> $($child.Title) -> $($child.Url)" -LogFile $LogFile
                    Add-PnPNavigationNode -Location TopNavigationBar `
                        -Title $child.Title -Url $child.Url -Parent $parentNode.Id | Out-Null
                }
            }
            else {
                Write-Log "[$siteSlug] [DryRun] Would add $($item.Title) + $($item.Children.Count) child(ren)." "WARN" $LogFile
            }
        }
        Write-Log "[$siteSlug] Rebuild complete." "SUCCESS" $LogFile
        if (-not $UseCurrentConnection) { Disconnect-PnPOnline }
        return [PSCustomObject]@{ Site = $siteSlug; Status = "REBUILT"; Error = "" }
    }

    # ── MODE B: Auto-fix broken URLs ─────────────────────────────────────────
    Write-Log "[$siteSlug] Scanning navigation for broken/Google-Sites URLs ..." -LogFile $LogFile
    $script:fixCount = 0

    function Repair-NavNode {
        param($Node, [string]$ParentTitle = "")
        # Safely read properties — strict mode throws if the property is missing on certain node types
        $url = if ($Node.PSObject.Properties['Url']) { $Node.Url } else { "" }
        $title = if ($Node.PSObject.Properties['Title']) { $Node.Title } else { "" }
        $label = if ($ParentTitle) { "$ParentTitle > $title" } else { $title }
        $newUrl = $url; $isBroken = $false

        if ($url -match "sites\.google\.com") {
            # Rule 1: Google URL
            $path = ($url -replace "^https?://[^/]+/[^/]+/[^/]+", "").TrimStart("/")
            $segments = $path.Split("/") | ForEach-Object {
                (($_ -replace "-", " ") | ForEach-Object { (Get-Culture).TextInfo.ToTitleCase($_) }) -replace " ", "-"
            }
            $newUrl = "$sitePagesBase/" + ($segments -join "/") + ".aspx"
            $isBroken = $true
            Write-Log "[$siteSlug] BROKEN(GSite): [$label] -> $newUrl" "WARN" $LogFile
        }
        elseif ([string]::IsNullOrWhiteSpace($url) -or $url -eq "#") {
            # Rule 2: Empty/placeholder
            $newUrl = "$sitePagesBase/$($title -replace '\s+','-').aspx"
            $isBroken = $true
            Write-Log "[$siteSlug] BROKEN(Empty): [$label] -> $newUrl" "WARN" $LogFile
        }
        elseif ($GSitesBaseUrl -and $url -like "*$GSitesBaseUrl*") {
            # Rule 3: Old base URL
            $newUrl = $url -replace [regex]::Escape($GSitesBaseUrl), ""
            $newUrl = "$sitePagesBase/" + $newUrl.TrimStart("/") + ".aspx"
            $isBroken = $true
            Write-Log "[$siteSlug] BROKEN(OldBase): [$label] -> $newUrl" "WARN" $LogFile
        }
        else {
            Write-Log "[$siteSlug] OK: [$label] $url" -LogFile $LogFile
        }

        if ($isBroken) {
            if (-not $DryRun) {
                if ([string]::IsNullOrWhiteSpace($title)) {
                    # Ghost node — no title, no Id. Find it via REST and delete by its REST Id.
                    $restNavNodes = Invoke-PnPSPRestMethod -Method Get `
                        -Url "/_api/web/navigation/topnavigationbar?`$select=Id,Title,Url"
                    $orphans = $restNavNodes.value | Where-Object {
                        [string]::IsNullOrWhiteSpace($_.Title) -and [string]::IsNullOrWhiteSpace($_.Url)
                    }
                    if ($orphans) {
                        foreach ($orphan in $orphans) {
                            Invoke-PnPSPRestMethod -Method Delete `
                                -Url "/_api/web/navigation/topnavigationbar/getById($($orphan.Id))"
                            Write-Log "[$siteSlug] [REMOVED] Ghost node id=$($orphan.Id)" "SUCCESS" $LogFile
                        }
                    }
                    else {
                        Write-Log "[$siteSlug] [SKIP] Ghost node not found via REST (already removed?)" "WARN" $LogFile
                    }
                }
                else {
                    # Set-PnPNavigationNode does not exist in PnP.PowerShell 1.x — use REST PATCH
                    Invoke-PnPSPRestMethod -Method Patch `
                        -Url "/_api/web/navigation/topnavigationbar/getById($($Node.Id))" `
                        -Content @{
                        "__metadata" = @{ "type" = "SP.NavigationNode" }
                        "Url"        = $newUrl
                        "Title"      = $title
                    }
                    Write-Log "[$siteSlug] [FIXED] [$label] -> $newUrl" "SUCCESS" $LogFile
                }
            }
            else {
                if ([string]::IsNullOrWhiteSpace($title)) {
                    Write-Log "[$siteSlug] [DryRun] Would REMOVE ghost node (no title/url/id)" "WARN" $LogFile
                }
                else {
                    Write-Log "[$siteSlug] [DryRun] Would fix [$label] -> $newUrl" "WARN" $LogFile
                }
            }
            $script:fixCount++
        }
        if ($Node.PSObject.Properties['Children'] -and $Node.Children) {
            foreach ($child in $Node.Children) { Repair-NavNode $child $title }
        }
    }

    Get-PnPNavigationNode -Location TopNavigationBar -Tree |
    ForEach-Object { Repair-NavNode $_ }

    Write-Log "[$siteSlug] Done. $($script:fixCount) node(s) fixed." "SUCCESS" $LogFile
    if (-not $UseCurrentConnection) { Disconnect-PnPOnline }
    return [PSCustomObject]@{ Site = $siteSlug; Status = "FIXED($($script:fixCount))"; Error = "" }
}

# =============================================================================
# Standalone entry-point (ignored when dot-sourced by bulk runner)
# =============================================================================
if ($SiteUrl) {
    Invoke-SPONavFix `
        -SPOSiteUrl    $SiteUrl `
        -GSitesBaseUrl $GoogleSitesBaseUrl `
        -ClientId      $ClientId `
        -TenantId      $TenantId `
        -Rebuild       $RebuildNavigation.IsPresent `
        -DryRun        $WhatIf.IsPresent
}

