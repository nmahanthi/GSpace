<#
.SYNOPSIS
    Complete Google Sites Non-Prod Assessment Orchestrator

.DESCRIPTION
    This script orchestrates all steps of the Google Sites assessment:
    1. GAM exports (inventory, permissions)
    2. Node.js dependency installation
    3. Browser authentication (manual step)
    4. Site crawling with Playwright
    5. Complexity scoring

.PARAMETER PrimaryDomain
    Your primary domain (e.g., "rocheua.com") for external permission detection

.PARAMETER MaxPagesPerSite
    Maximum pages to crawl per site (default: 200)

.PARAMETER MaxSites
    Maximum number of sites to crawl in this run (default: 0 = all sites).
    Use with -SiteOffset to process sites in batches, e.g. -MaxSites 10 -SiteOffset 0
    for the first 10, then -MaxSites 10 -SiteOffset 10 for the next 10.

.PARAMETER SiteOffset
    Number of sites to skip from the start of the inventory before crawling (default: 0).
    Use with -MaxSites to implement batched runs.

.PARAMETER SkipDependencyCheck
    Skip Node.js dependency installation check

.PARAMETER SkipGAMExport
    Skip GAM export step (use existing output files)

.PARAMETER SkipBrowserAuth
    Skip browser authentication step (use existing .auth/state.json)

.PARAMETER SkipCrawl
    Skip site crawling step (use existing crawl output)

.PARAMETER UseApiExtract
    Use the Sites API v1 (03b_api_extract_embeds.js) instead of the Playwright
    browser crawler (03_crawl_sites.js) to identify embedded content.
    Much faster - minutes vs hours. Requires a valid -AccessToken or GCP_ACCESS_TOKEN.
    Does not need a browser, Playwright, or a saved auth session.

.PARAMETER AccessToken
    OAuth 2.0 access token for Google APIs (if not provided, will try to get from gcloud)

.EXAMPLE
    .\Run-FullAssessment.ps1 -PrimaryDomain "rocheua.com"

.EXAMPLE
    .\Run-FullAssessment.ps1 -PrimaryDomain "rocheua.com" -SkipGAMExport -SkipBrowserAuth

.EXAMPLE
    .\Run-FullAssessment.ps1 -PrimaryDomain "rocheua.com" -AccessToken "ya29.a0AfB_..." -MaxPagesPerSite 100

.EXAMPLE
    # First batch: crawl sites 1-10
    .\Run-FullAssessment.ps1 -PrimaryDomain "rocheua.com" -SkipGAMExport -SkipBrowserAuth -MaxSites 10 -SiteOffset 0

.EXAMPLE
    # Second batch: crawl sites 11-20
    .\Run-FullAssessment.ps1 -PrimaryDomain "rocheua.com" -SkipGAMExport -SkipBrowserAuth -MaxSites 10 -SiteOffset 10

.EXAMPLE
    # Use the fast Sites API extractor instead of the Playwright crawler
    .\Run-FullAssessment.ps1 -PrimaryDomain "rocheua.com" -SkipGAMExport -SkipBrowserAuth -UseApiExtract -AccessToken "ya29...."

.PARAMETER SelectedSitesCsv
    Path to a CSV file containing a list of specific site names or URLs to process.
    Only these sites will be crawled and scored. Useful for large
    tenants where processing all sites is not feasible. The CSV must contain
    a column named SiteName, name, Name, SiteUrl, url, URL, or SiteURL.
    If you provide Google Sites URLs, the site name will be extracted from the URL path.

    Example CSV (by name):
        SiteName
        My First Site
        Another Site

    Example CSV (by URL):
        SiteUrl
        https://sites.google.com/yourdomain.com/my-first-site
        https://sites.google.com/yourdomain.com/another-site

.PARAMETER InventoryCsv
    Path to an existing GSites_Inventory_Detailed.csv file. Use this when you
    already have the inventory from a previous GAM export and want to skip
    re-running GAM. The file will be copied into the output/ folder so all
    downstream scripts can find it.

.PARAMETER SitesAdminEmail
    Email address of the account that will call the Sites API in Step 4A
    (i.e. the account used for "gcloud auth login" / -AccessToken). Before
    Step 4A runs, the toolkit uses GAM's elevated access to grant this
    account Reader access to every site in the inventory. This is required
    because the Sites API v1 rejects requests from accounts - even domain
    admins - that lack explicit Drive-level access to the file. If omitted,
    this pre-grant step is skipped and Step 4A may return 403 errors.

.PARAMETER SkipGrantAccess
    Skip the bulk Drive-access grant step that normally runs before Step 4A.

.PARAMETER ServiceAccountKeyPath
    Path to a service account key JSON file (e.g. GAM's oauth2service.json)
    authorized for domain-wide delegation. When provided (or auto-detected
    next to gam.exe / GAMCFGDIR / the toolkit folder), the toolkit mints its
    own OAuth token with the sites.readonly + drive.readonly scopes by
    impersonating -ImpersonateEmail, instead of using `gcloud auth
    print-access-token`. This is required because gcloud's own OAuth client
    can never be granted Sites API scopes. The service account's Client ID
    must first be authorized for these scopes in Admin Console > Security >
    API controls > Domain-wide delegation.

.PARAMETER ImpersonateEmail
    Workspace user/admin email to impersonate when minting a token via
    -ServiceAccountKeyPath. Defaults to -SitesAdminEmail if not provided.

.PARAMETER SkipPublishedUrls
    Skip Step 4A (Sites API published-URL lookup) entirely, e.g. when it
    keeps failing with 403/scope/API-disabled errors that can't be resolved
    quickly. Step 4B still runs normally - the Playwright crawler and the
    Sites API v1 extractor both already fall back to each site's edit URL
    (webViewLink) when no published-URL data is available, so results are
    still produced, just using edit URLs instead of published ones.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$PrimaryDomain,

    [int]$MaxPagesPerSite = 200,

    # Batching: limit how many sites are crawled per run
    [int]$MaxSites = 0,       # 0 = no limit (all sites)
    [int]$SiteOffset = 0,     # skip this many sites from the start

    [switch]$SkipDependencyCheck,
    [switch]$SkipGAMExport,
    [switch]$SkipBrowserAuth,
    [switch]$SkipCrawl,
    [switch]$SkipGrantAccess,
    [switch]$SkipPublishedUrls,

    # Use Sites API v1 extractor instead of Playwright crawler for Step 4B
    [switch]$UseApiExtract,

    [string]$AccessToken,

    # Filter to a specific list of sites (for large tenants)
    [string]$SelectedSitesCsv,

    # Provide an existing inventory CSV if skipping GAM export
    [string]$InventoryCsv,

    # Filter to a specific list of target users
    [string]$TargetUsersCsv,

    # Account that will call the Sites API - gets bulk-granted Reader access
    [string]$SitesAdminEmail,

    # Service-account-based token minting (Option A fix for scope-limited gcloud tokens)
    [string]$ServiceAccountKeyPath,
    [string]$ImpersonateEmail
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ScriptDir = $PSScriptRoot
$OutputDir = Join-Path $ScriptDir 'output'
$LogsDir = Join-Path $ScriptDir 'logs'
$AuthDir = Join-Path $ScriptDir '.auth'
$AuthFile = Join-Path $AuthDir 'state.json'

# Color output functions
function Write-Step {
    param([string]$Message)
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "  $Message" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
}

function Write-Success {
    param([string]$Message)
    Write-Host "[OK] $Message" -ForegroundColor Green
}

function Write-Info {
    param([string]$Message)
    Write-Host "[INFO] $Message" -ForegroundColor Yellow
}

function Write-Error-Custom {
    param([string]$Message)
    Write-Host "[ERROR] $Message" -ForegroundColor Red
}

function Get-CsvRowCount {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path $Path)) {
        return 0
    }

    $rows = @(Import-Csv $Path)
    return $rows.Count
}

function Get-SafeProperty {
    param(
        [Parameter(Mandatory = $true)][psobject]$InputObject,
        [Parameter(Mandatory = $true)][string[]]$PropertyNames
    )
    foreach ($prop in $PropertyNames) {
        $p = $InputObject.psobject.Properties[$prop]
        if ($null -ne $p -and -not [string]::IsNullOrWhiteSpace($p.Value)) {
            return [string]$p.Value
        }
    }
    return $null
}

function Normalize-CsvHeaders {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path $Path)) { return }
    $lines = @(Get-Content $Path)
    if ($lines.Count -eq 0) { return }
    # Remove numeric array indices like .0. .1. from the header row
    $lines[0] = $lines[0] -replace '\.[0-9]+\.', '.'
    $lines | Set-Content $Path
}

function Filter-InventoryBySelectedSites {
    param(
        [Parameter(Mandatory = $true)][string]$InventoryPath,
        [Parameter(Mandatory = $true)][string]$SelectedSitesCsvPath
    )

    if (-not (Test-Path $SelectedSitesCsvPath)) {
        throw "Selected sites CSV not found: $SelectedSitesCsvPath"
    }

    $selected = @(Import-Csv $SelectedSitesCsvPath)
    $selectedNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $selected) {
        $name = Get-SafeProperty -InputObject $row -PropertyNames @('SiteName', 'name', 'Name', 'SITENAME', 'SiteUrl', 'url', 'URL', 'SiteURL')
        $name = Extract-SiteNameFromValue -Value $name
        if (-not [string]::IsNullOrWhiteSpace($name)) { $selectedNames.Add([string]$name.Trim()) | Out-Null }
    }

    if ($selectedNames.Count -eq 0) {
        throw "No site names found in $SelectedSitesCsvPath. Expected column: SiteName, name, Name, SITENAME, SiteUrl, url, URL, or SiteURL."
    }

    Write-Info "Filtering inventory to $($selectedNames.Count) selected site name(s)..."

    $inventory = @(Import-Csv $InventoryPath)
    $filtered = @($inventory | Where-Object {
        $name = Get-SafeProperty -InputObject $_ -PropertyNames @('name', 'SiteName', 'Name', 'SITENAME')

        $urlName = $null
        $webviewlink = Get-SafeProperty -InputObject $_ -PropertyNames @('webviewlink')
        if (-not [string]::IsNullOrWhiteSpace($webviewlink)) {
            $urlName = Extract-SiteNameFromValue -Value $webviewlink
        }

        $selectedNames.Contains([string]$name) -or $selectedNames.Contains([string]$urlName)
    })

    if ($filtered.Count -eq 0) {
        throw "None of the selected site names were found in the inventory. Check the names in $SelectedSitesCsvPath. If you provided URLs, make sure the site name in the URL matches the Drive file name."
    }

    # Backup original inventory if not already backed up
    $backupPath = "$InventoryPath.full"
    if (-not (Test-Path $backupPath)) {
        Copy-Item $InventoryPath $backupPath
        Write-Success "Backed up full inventory to $(Split-Path $backupPath -Leaf)"
    }

    $filtered | Export-Csv -NoTypeInformation -Path $InventoryPath
    Write-Success "Filtered inventory written to $(Split-Path $InventoryPath -Leaf) ($($filtered.Count) sites)"

    # Also filter published URLs if they exist to avoid unnecessary API calls
    $publishedUrlsPath = Join-Path (Split-Path $InventoryPath) 'Sites_Published_URLs.csv'
    if (Test-Path $publishedUrlsPath) {
        $published = @(Import-Csv $publishedUrlsPath)
        $filteredPublished = @($published | Where-Object {
            $name = Get-SafeProperty -InputObject $_ -PropertyNames @('SiteName', 'name')
            $selectedNames.Contains([string]$name)
        })
        $filteredPublished | Export-Csv -NoTypeInformation -Path $publishedUrlsPath
        Write-Success "Filtered published URLs to $($filteredPublished.Count) site(s)"
    }

    # Also filter permissions if they exist to keep scoring consistent
    $permissionsPath = Join-Path (Split-Path $InventoryPath) 'GSites_Permissions.csv'
    if (Test-Path $permissionsPath) {
        $perms = @(Import-Csv $permissionsPath)
        $filteredPerms = @($perms | Where-Object {
            $name = Get-SafeProperty -InputObject $_ -PropertyNames @('name', 'SiteName')
            $selectedNames.Contains([string]$name)
        })
        $filteredPerms | Export-Csv -NoTypeInformation -Path $permissionsPath
        Write-Success "Filtered permissions to $($filteredPerms.Count) row(s)"
    }
}

function Extract-SiteNameFromValue {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }

    $Value = $Value.Trim()

    # If it looks like a Google Sites URL, extract the site name (path segment) from it
    # Format: https://sites.google.com/<domain>/<site-name>[/<page-path>]
    # Also supports old Apps format: https://sites.google.com/a/<domain>/<site-name>
    if ($Value -match '^https?://sites\.google\.com/(?:a/)?[^/]+/([^/]+)') {
        return $Matches[1]
    }

    return $Value
}

function Build-GamNameFilter {
    param([Parameter(Mandatory = $true)][string]$SelectedSitesCsvPath)

    if (-not (Test-Path $SelectedSitesCsvPath)) {
        return $null
    }

    $selected = @(Import-Csv $SelectedSitesCsvPath)
    $names = [System.Collections.Generic.List[string]]::new()
    foreach ($row in $selected) {
        $name = Get-SafeProperty -InputObject $row -PropertyNames @('SiteName', 'name', 'Name', 'SITENAME', 'SiteUrl', 'url', 'URL', 'SiteURL')
        $name = Extract-SiteNameFromValue -Value $name
        if (-not [string]::IsNullOrWhiteSpace($name)) { $names.Add([string]$name.Trim()) | Out-Null }
    }

    if ($names.Count -eq 0) {
        Write-Info "No site names found in $SelectedSitesCsvPath. Check the column header (expected SiteName, name, Name, SITENAME, SiteUrl, url, URL, or SiteURL)."
        return $null
    }

    Write-Info "Found $($names.Count) site name(s) in selected sites CSV."

    # Escape single quotes for Google Drive API queries (use \')
    $filterParts = foreach ($name in $names) {
        $escaped = $name -replace "'", "\'"
        "name='$escaped'"
    }

    $filter = $filterParts -join ' or '

    # Google Drive API query length limit is about 1000 chars.
    # Base query for sites is ~60 chars, so keep filter under ~800.
    if ($filter.Length -gt 800) {
        Write-Info "Too many selected sites for a GAM name filter. Running full tenant scan."
        return $null
    }

    Write-Info "GAM Sites filter: $filter"
    return $filter
}

function Build-GamTargetUsersFile {
    param(
        [Parameter(Mandatory = $true)][string]$CsvPath,
        [Parameter(Mandatory = $true)][string]$OutputDir
    )

    if (-not (Test-Path $CsvPath)) {
        return $null
    }

    $selected = @(Import-Csv $CsvPath)
    $emails = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $selected) {
        $email = Get-SafeProperty -InputObject $row -PropertyNames @('Owner', 'OwnerEmail', 'User', 'UserEmail', 'Email', 'primaryEmail')

        if (-not [string]::IsNullOrWhiteSpace($email) -and $email -match '@') {
            $emails.Add([string]$email.Trim()) | Out-Null
        }
    }

    if ($emails.Count -eq 0) {
        return $null
    }

    Write-Info "Found $($emails.Count) specific user email(s) in selected sites CSV. Restricting GAM scan to these users."

    $targetFile = Join-Path $OutputDir 'gam_target_users.csv'
    $csvData = $emails | ForEach-Object { [PSCustomObject]@{ Email = $_ } }
    $csvData | Export-Csv -Path $targetFile -NoTypeInformation
    return $targetFile
}

function Write-LogTail {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Label,
        [ConsoleColor]$Color = [ConsoleColor]::Gray,
        [int]$Tail = 20
    )

    if (-not (Test-Path $Path)) {
        return
    }

    Write-Info "$Label"
    $lines = @(Get-Content $Path -Tail $Tail)
    foreach ($line in $lines) {
        Write-Host "    $line" -ForegroundColor $Color
    }
}

function Invoke-LoggedProcess {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [string[]]$ArgumentList = @(),
        [string]$WorkingDirectory = $ScriptDir,
        [Parameter(Mandatory = $true)][string]$LogPrefix
    )

    if (-not (Test-Path $LogsDir)) {
        New-Item -ItemType Directory -Path $LogsDir -Force | Out-Null
    }

    $safeLogPrefix = $LogPrefix -replace '[^a-zA-Z0-9._-]', '_'
    $stdoutLog = Join-Path $LogsDir "${safeLogPrefix}_stdout.log"
    $stderrLog = Join-Path $LogsDir "${safeLogPrefix}_stderr.log"

    if (Test-Path $stdoutLog) { Remove-Item $stdoutLog -Force }
    if (Test-Path $stderrLog) { Remove-Item $stderrLog -Force }

    $proc = Start-Process `
        -FilePath $FilePath `
        -ArgumentList $ArgumentList `
        -WorkingDirectory $WorkingDirectory `
        -NoNewWindow `
        -Wait `
        -PassThru `
        -RedirectStandardOutput $stdoutLog `
        -RedirectStandardError $stderrLog

    return [pscustomobject]@{
        ExitCode  = $proc.ExitCode
        StdOutLog = $stdoutLog
        StdErrLog = $stderrLog
    }
}

# Start assessment
Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "  Google Sites Non-Prod Assessment - Full Orchestrator" -ForegroundColor Cyan
Write-Host "============================================================`n" -ForegroundColor Cyan

Write-Info "Primary Domain: $PrimaryDomain"
Write-Info "Max Pages Per Site: $MaxPagesPerSite"
Write-Info "Output Directory: $OutputDir"
Write-Host ""

# Create output and logs directories
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    Write-Success "Created output directory"
}

if (-not (Test-Path $LogsDir)) {
    New-Item -ItemType Directory -Path $LogsDir -Force | Out-Null
    Write-Success "Created logs directory"
}

# ============================================================================
# STEP 1: GAM EXPORTS
# ============================================================================
if (-not $SkipGAMExport) {
    Write-Step "STEP 1: GAM Exports"

    $gamExportScript = Join-Path $ScriptDir '01_run_gam_exports.cmd'
    if (-not (Test-Path $gamExportScript)) {
        throw "GAM export script not found: $gamExportScript"
    }

    Write-Info "Running GAM exports..."
    Write-Info "Progress output is being written to log files. Please wait..."

    # If selected sites are specified, build a GAM name filter for the Sites export
    if ($SelectedSitesCsv) {
        $gamFilter = Build-GamNameFilter -SelectedSitesCsvPath $SelectedSitesCsv
        if ($gamFilter) {
            [Environment]::SetEnvironmentVariable('GAM_SITES_FILTER', $gamFilter, 'Process')
            Write-Info "Applying GAM Sites name filter to reduce tenant scan scope."
        }

        # Check if users/owners were provided to avoid scanning all users
        $gamTargetFile = Build-GamTargetUsersFile -CsvPath $SelectedSitesCsv -OutputDir $OutputDir
        if ($gamTargetFile) {
            [Environment]::SetEnvironmentVariable('GAM_TARGET_FILE', $gamTargetFile, 'Process')
        }
    } elseif ($TargetUsersCsv) {
        $gamTargetFile = Build-GamTargetUsersFile -CsvPath $TargetUsersCsv -OutputDir $OutputDir
        if ($gamTargetFile) {
            [Environment]::SetEnvironmentVariable('GAM_TARGET_FILE', $gamTargetFile, 'Process')
        } else {
            throw "No user emails found in $TargetUsersCsv. Expected column: Email, User, Owner, or OwnerEmail."
        }
    }

    $gamResult = Invoke-LoggedProcess -FilePath 'cmd.exe' -ArgumentList @('/c', "`"$gamExportScript`"") -WorkingDirectory $ScriptDir -LogPrefix '01_gam_exports'

    if ($gamResult.ExitCode -ne 0) {
        Write-Error-Custom "GAM export failed with exit code $($gamResult.ExitCode)"
        Write-LogTail -Path $gamResult.StdErrLog -Label 'GAM stderr tail (last 20 lines)' -Color Red
        Write-LogTail -Path $gamResult.StdOutLog -Label 'GAM stdout tail (last 20 lines)' -Color Gray
        if ($SelectedSitesCsv -and $gamFilter) {
            throw @"
GAM export failed. This may be because the selected site names do not match the Drive file names.
The filter used was: $gamFilter

If your CSV contains Google Sites URLs, the script extracts the site name from the URL path.
For example: https://sites.google.com/domain.com/site-name -> site-name

If the extracted names do not match the actual Drive file names, provide the exact Drive file names
(the site titles) in your CSV instead of URLs.
"@
        }
        throw 'GAM export failed'
    }

    Write-Success 'GAM exports completed'
    Write-Info "  Logs saved to: $LogsDir"

    # Verify output files
    $requiredFiles = @(
        'GSites_Inventory_Min.csv',
        'GSites_Inventory_Detailed.csv',
        'GSites_Permissions.csv'
    )

    foreach ($file in $requiredFiles) {
        $filePath = Join-Path $OutputDir $file
        if (Test-Path $filePath) {
            $rowCount = Get-CsvRowCount -Path $filePath
            Write-Info "  [OK] $file ($rowCount rows)"
            Normalize-CsvHeaders -Path $filePath
        }
        else {
            Write-Error-Custom "  [MISSING] $file"
        }
    }
}
else {
    Write-Step "STEP 1: GAM Exports (SKIPPED)"
}

# Filter inventory to selected sites if a CSV is provided
if ($SelectedSitesCsv) {
    $inventoryFile = Join-Path $OutputDir 'GSites_Inventory_Detailed.csv'

    # If an external inventory CSV was provided, copy it to the output folder
    if ($InventoryCsv -and (Test-Path $InventoryCsv)) {
        if (-not (Test-Path $OutputDir)) {
            New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
        }
        Copy-Item $InventoryCsv $inventoryFile -Force
        Write-Success "Copied inventory from $InventoryCsv to output folder"
    }

    if (-not (Test-Path $inventoryFile)) {
        if ($SkipGAMExport) {
            throw @"
Inventory file not found: $inventoryFile

You used -SkipGAMExport but the inventory file does not exist yet.
On a fresh run you must let the script run GAM export to build the inventory.

Options:
  1. Remove -SkipGAMExport from your command (recommended for fresh runs)
     Example: .\Run-FullAssessment.ps1 -PrimaryDomain "$PrimaryDomain" -SelectedSitesCsv "$SelectedSitesCsv"

  2. If you already have an inventory file somewhere else, use -InventoryCsv:
     Example: .\Run-FullAssessment.ps1 -PrimaryDomain "$PrimaryDomain" -SelectedSitesCsv "$SelectedSitesCsv" -InventoryCsv "C:\path\to\GSites_Inventory_Detailed.csv"
"@
        }
        throw "Inventory file not found: $inventoryFile. GAM export may have failed or produced no sites."
    }
    Filter-InventoryBySelectedSites -InventoryPath $inventoryFile -SelectedSitesCsvPath $SelectedSitesCsv
}

# ============================================================================
# STEP 2: NODE.JS DEPENDENCY CHECK
# ============================================================================
if (-not $SkipDependencyCheck) {
    Write-Step "STEP 2: Node.js Dependency Check"

    # Check if Node.js is installed
    try {
        $nodeVersion = & node --version 2>&1
        Write-Success "Node.js installed: $nodeVersion"
    }
    catch {
        Write-Error-Custom "Node.js is not installed!"
        Write-Info "Please install Node.js from https://nodejs.org/"
        throw "Node.js is required for this assessment"
    }

    # Check if npm is installed
    try {
        $npmVersion = & npm --version 2>&1
        Write-Success "npm installed: $npmVersion"
    }
    catch {
        Write-Error-Custom "npm is not installed!"
        throw "npm is required for this assessment"
    }

    # Check if package.json exists
    $packageJson = Join-Path $ScriptDir 'package.json'
    if (-not (Test-Path $packageJson)) {
        Write-Info "Initializing npm project..."
        Push-Location $ScriptDir
        & npm init -y | Out-Null
        Pop-Location
        Write-Success "npm project initialized"
    }

    # Install dependencies
    Write-Info "Installing Node.js dependencies (playwright, csv-parse, csv-stringify)..."
    Push-Location $ScriptDir
    & npm install playwright csv-parse csv-stringify 2>&1 | Out-Null
    Pop-Location
    Write-Success "Node.js dependencies installed"

    # Install Playwright browsers
    Write-Info "Installing Playwright Chromium browser..."
    Push-Location $ScriptDir
    & npx playwright install chromium 2>&1 | Out-Null
    Pop-Location
    Write-Success "Playwright Chromium installed"

}
else {
    Write-Step "STEP 2: Node.js Dependency Check (SKIPPED)"
}

# ============================================================================
# STEP 3: BROWSER AUTHENTICATION
# ============================================================================
if (-not $SkipBrowserAuth) {
    Write-Step "STEP 3: Browser Authentication"

    if (Test-Path $AuthFile) {
        Write-Info "Existing authentication found at: $AuthFile"
        $response = Read-Host "Do you want to re-authenticate? (y/N)"
        if ($response -ne 'y' -and $response -ne 'Y') {
            Write-Success "Using existing authentication"
        }
        else {
            Remove-Item $AuthFile -Force
            Write-Info "Deleted existing authentication"
        }
    }

    if (-not (Test-Path $AuthFile)) {
        Write-Info "Launching browser for authentication..."
        Write-Host ""
        Write-Host "INSTRUCTIONS:" -ForegroundColor Yellow
        Write-Host "  1. A browser window will open" -ForegroundColor Gray
        Write-Host "  2. Sign in with your Google account" -ForegroundColor Gray
        Write-Host "  3. Navigate to a Google Site to verify access" -ForegroundColor Gray
        Write-Host "  4. Return to this window and press Enter" -ForegroundColor Gray
        Write-Host ""

        $authScript = Join-Path $ScriptDir '02_save_playwright_auth.js'
        Push-Location $ScriptDir
        & node $authScript "https://sites.google.com/"
        Pop-Location

        if (Test-Path $AuthFile) {
            Write-Success "Browser authentication saved"
        }
        else {
            throw "Browser authentication failed - auth file not created"
        }
    }
}
else {
    Write-Step "STEP 3: Browser Authentication (SKIPPED)"
    if (-not (Test-Path $AuthFile)) {
        Write-Error-Custom "Authentication file not found: $AuthFile"
        throw "Cannot skip browser authentication - no existing auth file found"
    }
}

# ============================================================================
# STEP 3.5: GRANT SITES API ACCESS (Option B fix for Step 4A 403 errors)
# ============================================================================
if (-not $SkipCrawl -and -not $SkipGrantAccess) {
    if ([string]::IsNullOrWhiteSpace($SitesAdminEmail)) {
        Write-Step "STEP 3.5: Grant Sites API Access (SKIPPED - no -SitesAdminEmail provided)"
        Write-Info "Step 4A may fail with 403 Forbidden unless the calling account already"
        Write-Info "has Drive-level access to every site. Pass -SitesAdminEmail to fix this."
    }
    else {
        Write-Step "STEP 3.5: Grant Sites API Access"

        $grantScript = Join-Path $ScriptDir '01b_grant_site_access.cmd'
        if (-not (Test-Path $grantScript)) {
            throw "Grant access script not found: $grantScript"
        }

        $inventoryFile = Join-Path $OutputDir 'GSites_Inventory_Detailed.csv'
        if (-not (Test-Path $inventoryFile)) {
            throw "Inventory file not found: $inventoryFile. Run GAM export first."
        }

        Write-Info "Granting $SitesAdminEmail Reader access to all sites via GAM (Option B fix for 403 errors)..."
        [Environment]::SetEnvironmentVariable('SITES_ADMIN_EMAIL', $SitesAdminEmail, 'Process')

        $grantResult = Invoke-LoggedProcess -FilePath 'cmd.exe' -ArgumentList @('/c', "`"$grantScript`"") -WorkingDirectory $ScriptDir -LogPrefix '01b_grant_access'

        if ($grantResult.ExitCode -ne 0) {
            Write-Error-Custom "Grant access step failed with exit code $($grantResult.ExitCode)"
            Write-LogTail -Path $grantResult.StdErrLog -Label 'Grant access stderr tail (last 20 lines)' -Color Red
            Write-LogTail -Path $grantResult.StdOutLog -Label 'Grant access stdout tail (last 20 lines)' -Color Gray
            Write-Info "Continuing anyway - Step 4A will report any remaining 403 errors per-site."
        }
        else {
            Write-Success "Access grant step completed"
            Write-Info "  Logs saved to: $LogsDir"
        }
    }
}

# ============================================================================
# STEP 4: GET PUBLISHED URLs (NEW)
# ============================================================================
if ($SkipPublishedUrls -and -not $SkipCrawl) {
    Write-Step "STEP 4A: Get Published URLs from Sites API (SKIPPED)"
    Write-Info "Step 4B will use each site's edit URL (webViewLink) instead."
}
elseif (-not $SkipCrawl) {
    Write-Step "STEP 4A: Get Published URLs from Sites API"

    $inventoryFile = Join-Path $OutputDir 'GSites_Inventory_Detailed.csv'
    if (-not (Test-Path $inventoryFile)) {
        throw "Inventory file not found: $inventoryFile. Run GAM export first."
    }

    $siteCount = @(Import-Csv $inventoryFile).Count
    Write-Info "Found $siteCount sites"

    if ($siteCount -eq 0) {
        Write-Info "No sites found - skipping published URL retrieval"
    }
    else {
        # Try to get OAuth token for Sites API
        $tokenAvailable = $false
        $sitesApiToken = $null

        if (-not [string]::IsNullOrWhiteSpace($AccessToken)) {
            $sitesApiToken = $AccessToken
            $tokenAvailable = $true
            Write-Success "Using provided access token"
        }
        elseif (-not [string]::IsNullOrWhiteSpace($ServiceAccountKeyPath)) {
            # Option A fix: mint a token via domain-wide delegation using a
            # service account key (e.g. GAM's oauth2service.json), which can
            # be authorized for the sites.readonly scope. gcloud's own OAuth
            # client can never be granted that scope, no matter how it is
            # re-authenticated - see GAM_PATH_FIX.md / CUSTOMER_SETUP.md.
            $impersonate = if (-not [string]::IsNullOrWhiteSpace($ImpersonateEmail)) { $ImpersonateEmail } else { $SitesAdminEmail }

            if ([string]::IsNullOrWhiteSpace($impersonate)) {
                Write-Info "ServiceAccountKeyPath was provided but no -ImpersonateEmail or -SitesAdminEmail was given."
                Write-Info "Skipping token minting - cannot impersonate without a target email."
            }
            elseif (-not (Test-Path $ServiceAccountKeyPath)) {
                Write-Info "Service account key file not found: $ServiceAccountKeyPath"
            }
            else {
                Write-Info "Minting Sites API token via service account (impersonating $impersonate)..."
                $tokenScript = Join-Path $ScriptDir 'get_service_account_token.js'
                Push-Location $ScriptDir
                $saTokenOutput = & node $tokenScript $ServiceAccountKeyPath $impersonate 2>&1
                $saTokenExit = $LASTEXITCODE
                Pop-Location

                if ($saTokenExit -eq 0 -and $saTokenOutput -match '^ya29\.') {
                    $sitesApiToken = ([string]$saTokenOutput).Trim()
                    $tokenAvailable = $true
                    Write-Success "Access token obtained via service account (domain-wide delegation)"
                }
                else {
                    Write-Info "Service account token minting failed:"
                    ($saTokenOutput | Out-String) -split "`n" | ForEach-Object {
                        if ($_.Trim()) { Write-Info "  $_" }
                    }
                }
            }
        }
        else {
            Write-Info "No access token provided, checking for gcloud CLI..."
            try {
                $gcloudCheck = Get-Command gcloud -ErrorAction SilentlyContinue
                if ($null -eq $gcloudCheck) {
                    Write-Info "gcloud CLI not found in PATH"
                }
                else {
                    # Call gcloud directly via .NET Process instead of Start-Job.
                    # Start-Job spins up a brand-new PowerShell host just to run gcloud,
                    # which alone can take longer than the old 5s timeout on a loaded
                    # machine / VPN, causing this step to be silently skipped and
                    # Sites_Published_URLs.csv to never be written. Process.WaitForExit
                    # gives an accurate timeout on the gcloud call itself with no extra
                    # host-spin-up overhead.
                    $gcloudTimeoutSeconds = 25
                    Write-Info "Attempting to get token from gcloud (timeout: $gcloudTimeoutSeconds seconds)..."

                    # gcloud is typically a .cmd shim on Windows (not a .exe), and
                    # ProcessStartInfo/CreateProcess (UseShellExecute=false) cannot
                    # launch .cmd/.bat files directly - only cmd.exe or PowerShell's
                    # own command resolution can. Route through cmd.exe /c so this
                    # works regardless of whether gcloud resolves to a .cmd, .bat,
                    # or .exe.
                    $psi = New-Object System.Diagnostics.ProcessStartInfo
                    $psi.FileName = 'cmd.exe'
                    $psi.Arguments = '/c gcloud auth print-access-token'
                    $psi.RedirectStandardOutput = $true
                    $psi.RedirectStandardError = $true
                    $psi.UseShellExecute = $false
                    $psi.CreateNoWindow = $true

                    $proc = New-Object System.Diagnostics.Process
                    $proc.StartInfo = $psi
                    $proc.Start() | Out-Null

                    $stdoutTask = $proc.StandardOutput.ReadToEndAsync()
                    $stderrTask = $proc.StandardError.ReadToEndAsync()

                    if ($proc.WaitForExit($gcloudTimeoutSeconds * 1000)) {
                        $stdout = $stdoutTask.GetAwaiter().GetResult()
                        $stderr = $stderrTask.GetAwaiter().GetResult()
                        # Trim removes all \r\n that gcloud appends to its output.
                        # Without this the Bearer header becomes "Bearer ya29.xxx\r\n" -> HTTP 401.
                        $sitesApiToken = ($stdout + $stderr).Trim()

                        if ($sitesApiToken -match '^ya29\.' -and $sitesApiToken.Length -gt 20) {
                            Write-Success "Access token obtained from gcloud"
                            $tokenAvailable = $true
                        }
                        elseif ($sitesApiToken -match 'ERROR') {
                            Write-Info "gcloud auth not configured or expired - run: gcloud auth login"
                        }
                        else {
                            Write-Info "Unexpected gcloud output - could not extract token"
                        }
                    }
                    else {
                        Write-Info "gcloud command timed out (>$gcloudTimeoutSeconds s)"
                        try { $proc.Kill() } catch {}
                    }
                }
            }
            catch {
                Write-Info "Could not get access token from gcloud: $_"
            }
        }

        if ($tokenAvailable) {
            Write-Info "Fetching published URLs from Sites API..."
            $publishedUrlScript = Join-Path $ScriptDir '03a_get_published_urls.js'
            Push-Location $ScriptDir
            $env:GCP_ACCESS_TOKEN = $sitesApiToken
            & node $publishedUrlScript
            Pop-Location

            $publishedUrlsFile = Join-Path $OutputDir 'Sites_Published_URLs.csv'
            if (Test-Path $publishedUrlsFile) {
                $publishedData = @(Import-Csv $publishedUrlsFile | Where-Object { $_.PublishedUrl -and $_.PublishedUrl -ne '' })
                $publishedCount = $publishedData.Count
                Write-Success "Published URLs retrieved: $publishedCount of $siteCount sites"
                Write-Info "  Output: Sites_Published_URLs.csv"

                if ($publishedCount -eq 0) {
                    Write-Host ""
                    Write-Host "========================================" -ForegroundColor Yellow
                    Write-Host "  OAuth Scope Issue Detected" -ForegroundColor Yellow
                    Write-Host "========================================" -ForegroundColor Yellow
                    Write-Host ""
                    Write-Host "The access token doesn't have Sites API scope." -ForegroundColor Gray
                    Write-Host "Re-authenticate with the correct scope:" -ForegroundColor Yellow
                    Write-Host ""
                    Write-Host "  gcloud auth login" -ForegroundColor White
                    Write-Host ""
                    Write-Host "Then re-run the assessment." -ForegroundColor Gray
                    Write-Host ""
                }
            }
            else {
                Write-Info "Published URLs file not created - Sites API may have failed"
            }
        }
        else {
            Write-Host ""
            Write-Host "========================================" -ForegroundColor Yellow
            Write-Host "  Published URLs Not Retrieved" -ForegroundColor Yellow
            Write-Host "========================================" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "No OAuth token available for Sites API." -ForegroundColor Gray
            Write-Host "The crawler will attempt to use edit URLs, which may result in 403 errors." -ForegroundColor Gray
            Write-Host ""
            Write-Host "To get published URLs, provide an OAuth token:" -ForegroundColor Yellow
            Write-Host "  gcloud auth login" -ForegroundColor White
            Write-Host "  .\\Run-FullAssessment.ps1 -PrimaryDomain 'rocheua.com'" -ForegroundColor White
            Write-Host ""
            Write-Info "Continuing with crawl using edit URLs..."
            Write-Host ""
        }
    }
}

# ============================================================================
# STEP 4B: SITE CRAWLING
# ============================================================================
if (-not $SkipCrawl) {
    Write-Step "STEP 4B: Site Crawling with Playwright"

    $inventoryFile = Join-Path $OutputDir 'GSites_Inventory_Detailed.csv'
    $siteCount = @(Import-Csv $inventoryFile).Count
    Write-Info "Total sites in inventory : $siteCount"
    Write-Info "Max pages per site       : $MaxPagesPerSite"
    if ($SiteOffset -gt 0) { Write-Info "Site offset (skip first) : $SiteOffset" }
    if ($MaxSites -gt 0) { Write-Info "Max sites this run       : $MaxSites" }

    if ($siteCount -eq 0) {
        Write-Info "No sites found - skipping crawl"
    }
    else {
        # Calculate which slice will actually be processed and report it clearly
        $effectiveOffset = [Math]::Min($SiteOffset, $siteCount)
        $remaining = $siteCount - $effectiveOffset
        $effectiveCount = if ($MaxSites -gt 0) { [Math]::Min($MaxSites, $remaining) } else { $remaining }
        Write-Info "Sites that will be processed this run: $effectiveCount (sites $($effectiveOffset + 1) - $($effectiveOffset + $effectiveCount) of $siteCount)"

        if ($UseApiExtract) {
            # -- Fast path: Sites API v1 - no browser required ----------------
            Write-Info "Mode: Sites API v1 extractor (03b_api_extract_embeds.js)"

            if ([string]::IsNullOrWhiteSpace($AccessToken)) {
                Write-Error-Custom "UseApiExtract requires an OAuth access token."
                Write-Info "  Provide -AccessToken or set the GCP_ACCESS_TOKEN env var."
                throw "No access token for API extract mode."
            }

            $apiExtractScript = Join-Path $ScriptDir '03b_api_extract_embeds.js'
            if (-not (Test-Path $apiExtractScript)) {
                throw "API extract script not found: $apiExtractScript"
            }

            Push-Location $ScriptDir
            $env:GCP_ACCESS_TOKEN = $AccessToken
            $env:MAX_SITES = $MaxSites
            $env:SITE_OFFSET = $SiteOffset
            & node $apiExtractScript
            Pop-Location

            Write-Success "API embed extraction completed"
        }
        else {
            # -- Standard path: Playwright browser crawler ---------------------
            Write-Info "Mode: Playwright browser crawler (03_crawl_sites.js)"
            Write-Info "Starting crawl (this may take a while)..."

            $crawlScript = Join-Path $ScriptDir '03_crawl_sites.js'
            Push-Location $ScriptDir
            $env:MAX_PAGES_PER_SITE = $MaxPagesPerSite
            $env:MAX_SITES = $MaxSites
            $env:SITE_OFFSET = $SiteOffset
            & node $crawlScript
            Pop-Location

            Write-Success "Site crawling completed"
        }

        # Verify crawl output
        $crawlOutputFiles = @(
            'Pages.csv',
            'Embeds.csv',
            'ExternalDomains.csv'
        )

        foreach ($file in $crawlOutputFiles) {
            $filePath = Join-Path $OutputDir $file
            if (Test-Path $filePath) {
                $rows = @(Import-Csv $filePath)
                $rowCount = $rows.Count
                Write-Info "  [OK] $file ($rowCount rows)"
            }
            else {
                Write-Error-Custom "  [MISSING] $file"
            }
        }
    }
}
else {
    Write-Step "STEP 4: Get Published URLs & Site Crawling (SKIPPED)"
}

# ============================================================================
# STEP 5: COMPLEXITY SCORING
# ============================================================================
Write-Step "STEP 5: Complexity Scoring"

$scoreScript = Join-Path $ScriptDir '05_score_sites.ps1'
Write-Info "Generating complexity report..."

& pwsh -ExecutionPolicy Bypass -File $scoreScript -OutputDir $OutputDir -PrimaryDomain $PrimaryDomain

Write-Success "Complexity scoring completed"

$reportFile = Join-Path $OutputDir 'Complexity_Report.csv'
if (Test-Path $reportFile) {
    $report = @(Import-Csv $reportFile)
    Write-Info "   Complexity_Report.csv ($($report.Count) sites)"

    # Summary statistics
    Write-Host ""
    Write-Host "COMPLEXITY SUMMARY:" -ForegroundColor Cyan
    $lowCount = @($report | Where-Object { $_.Rating -eq 'Low' }).Count
    $mediumCount = @($report | Where-Object { $_.Rating -eq 'Medium' }).Count
    $highCount = @($report | Where-Object { $_.Rating -eq 'High' }).Count
    $veryHighCount = @($report | Where-Object { $_.Rating -eq 'Very High' }).Count

    Write-Host "  Low:       $lowCount sites" -ForegroundColor Green
    Write-Host "  Medium:    $mediumCount sites" -ForegroundColor Yellow
    Write-Host "  High:      $highCount sites" -ForegroundColor DarkYellow
    Write-Host "  Very High: $veryHighCount sites" -ForegroundColor Red
}

# ============================================================================
# FINAL SUMMARY
# ============================================================================
Write-Host "`n============================================================" -ForegroundColor Green
Write-Host "  ASSESSMENT COMPLETE!" -ForegroundColor Green
Write-Host "============================================================`n" -ForegroundColor Green

Write-Host "OUTPUT FILES:" -ForegroundColor Cyan
Write-Host "  Location: $OutputDir" -ForegroundColor Gray
Write-Host ""

$allOutputFiles = Get-ChildItem $OutputDir -Filter "*.csv" | Sort-Object Name
foreach ($file in $allOutputFiles) {
    # Stream-count lines - fast even for million-row files; Import-Csv was loading all data into memory
    $lineCount = 0
    switch -File $file.FullName { default { $lineCount++ } }
    $rowCount = [Math]::Max(0, $lineCount - 1)  # subtract CSV header row
    Write-Host "  [OK] $($file.Name) " -NoNewline -ForegroundColor Green
    Write-Host "($rowCount rows)" -ForegroundColor Gray
}

Write-Host ""
Write-Host "NEXT STEPS:" -ForegroundColor Cyan
Write-Host "  1. Review the complexity report: Complexity_Report.csv" -ForegroundColor Gray
Write-Host "  2. Analyze high-complexity sites for migration planning" -ForegroundColor Gray
Write-Host ""

Write-Success "Assessment completed successfully!"
Write-Host ""

