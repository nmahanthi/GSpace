<#
.SYNOPSIS
    Complete Google Sites Non-Prod Assessment Orchestrator

.DESCRIPTION
    This script orchestrates all 5 steps of the Google Sites assessment:
    1. GAM exports (inventory, permissions, artifacts)
    2. Node.js dependency installation
    3. Browser authentication (manual step)
    4. Site crawling with Playwright
    5. Artifact enrichment with Google APIs
    6. Complexity scoring

.PARAMETER PrimaryDomain
    Your primary domain (e.g., "rocheua.com") for external permission detection

.PARAMETER MaxPagesPerSite
    Maximum pages to crawl per site (default: 200)

.PARAMETER SkipDependencyCheck
    Skip Node.js dependency installation check

.PARAMETER SkipGAMExport
    Skip GAM export step (use existing output files)

.PARAMETER SkipBrowserAuth
    Skip browser authentication step (use existing .auth/state.json)

.PARAMETER SkipCrawl
    Skip site crawling step (use existing crawl output)

.PARAMETER SkipEnrichment
    Skip artifact enrichment step

.PARAMETER AccessToken
    OAuth 2.0 access token for Google APIs (if not provided, will try to get from gcloud)

.EXAMPLE
    .\Run-FullAssessment.ps1 -PrimaryDomain "rocheua.com"

.EXAMPLE
    .\Run-FullAssessment.ps1 -PrimaryDomain "rocheua.com" -SkipGAMExport -SkipBrowserAuth

.EXAMPLE
    .\Run-FullAssessment.ps1 -PrimaryDomain "rocheua.com" -AccessToken "ya29.a0AfB_..." -MaxPagesPerSite 100
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$PrimaryDomain,
    
    [int]$MaxPagesPerSite = 200,
    
    [switch]$SkipDependencyCheck,
    [switch]$SkipGAMExport,
    [switch]$SkipBrowserAuth,
    [switch]$SkipCrawl,
    [switch]$SkipEnrichment,
    
    [string]$AccessToken
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

    $gamResult = Invoke-LoggedProcess -FilePath 'cmd.exe' -ArgumentList @('/c', "`"$gamExportScript`"") -WorkingDirectory $ScriptDir -LogPrefix '01_gam_exports'

    if ($gamResult.ExitCode -ne 0) {
        Write-Error-Custom "GAM export failed with exit code $($gamResult.ExitCode)"
        Write-LogTail -Path $gamResult.StdErrLog -Label 'GAM stderr tail (last 20 lines)' -Color Red
        Write-LogTail -Path $gamResult.StdOutLog -Label 'GAM stdout tail (last 20 lines)' -Color Gray
        throw 'GAM export failed'
    }

    Write-Success 'GAM exports completed'
    Write-Info "  Logs saved to: $LogsDir"

    # Verify output files
    $requiredFiles = @(
        '01_GSites_Inventory_Min.csv',
        '02_GSites_Inventory_Detailed.csv',
        '03_GSites_Permissions.csv',
        '04_Candidate_Sheets.csv',
        '05_Candidate_Forms.csv',
        '06_Candidate_Scripts.csv'
    )

    foreach ($file in $requiredFiles) {
        $filePath = Join-Path $OutputDir $file
        if (Test-Path $filePath) {
            $rowCount = Get-CsvRowCount -Path $filePath
            Write-Info "  [OK] $file ($rowCount rows)"
        }
        else {
            Write-Error-Custom "  [MISSING] $file"
        }
    }
}
else {
    Write-Step "STEP 1: GAM Exports (SKIPPED)"
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
# STEP 4: GET PUBLISHED URLs (NEW)
# ============================================================================
if (-not $SkipCrawl) {
    Write-Step "STEP 4A: Get Published URLs from Sites API"

    $inventoryFile = Join-Path $OutputDir '02_GSites_Inventory_Detailed.csv'
    if (-not (Test-Path $inventoryFile)) {
        throw "Inventory file not found: $inventoryFile. Run GAM export first."
    }

    $siteCount = (Import-Csv $inventoryFile).Count
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
        else {
            Write-Info "No access token provided, checking for gcloud CLI..."
            try {
                $gcloudCheck = Get-Command gcloud -ErrorAction SilentlyContinue
                if ($null -eq $gcloudCheck) {
                    Write-Info "gcloud CLI not found in PATH"
                }
                else {
                    Write-Info "Attempting to get token from gcloud (timeout: 5 seconds)..."
                    $job = Start-Job -ScriptBlock { gcloud auth print-access-token 2>&1 }
                    $completed = Wait-Job -Job $job -Timeout 5

                    if ($null -ne $completed) {
                        $sitesApiToken = Receive-Job -Job $job
                        Remove-Job -Job $job -Force

                        if (-not [string]::IsNullOrWhiteSpace($sitesApiToken) -and $sitesApiToken -notmatch 'ERROR') {
                            Write-Success "Access token obtained from gcloud"
                            $tokenAvailable = $true
                        }
                        else {
                            Write-Info "gcloud auth not configured or expired"
                        }
                    }
                    else {
                        Write-Info "gcloud command timed out"
                        Remove-Job -Job $job -Force
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

            $publishedUrlsFile = Join-Path $OutputDir '02a_Sites_Published_URLs.csv'
            if (Test-Path $publishedUrlsFile) {
                $publishedData = @(Import-Csv $publishedUrlsFile | Where-Object { $_.PublishedUrl -and $_.PublishedUrl -ne '' })
                $publishedCount = $publishedData.Count
                Write-Success "Published URLs retrieved: $publishedCount of $siteCount sites"
                Write-Info "  Output: 02a_Sites_Published_URLs.csv"

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

    $inventoryFile = Join-Path $OutputDir '02_GSites_Inventory_Detailed.csv'
    $siteCount = (Import-Csv $inventoryFile).Count
    Write-Info "Max pages per site: $MaxPagesPerSite"

    if ($siteCount -eq 0) {
        Write-Info "No sites found - skipping crawl"
    }
    else {
        Write-Info "Starting crawl (this may take a while)..."

        $crawlScript = Join-Path $ScriptDir '03_crawl_sites.js'
        Push-Location $ScriptDir
        $env:MAX_PAGES_PER_SITE = $MaxPagesPerSite
        & node $crawlScript
        Pop-Location

        Write-Success "Site crawling completed"

        # Verify crawl output
        $crawlOutputFiles = @(
            '07_Pages.csv',
            '08_Embeds.csv',
            '09_ExternalDomains.csv',
            '10_NetworkRequests.csv'
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
# STEP 5: ARTIFACT ENRICHMENT
# ============================================================================
if (-not $SkipEnrichment) {
    Write-Step "STEP 5: Artifact Enrichment"

    # Get OAuth token
    $tokenAvailable = $false
    if ([string]::IsNullOrWhiteSpace($AccessToken)) {
        Write-Info "No access token provided, checking for gcloud CLI..."
        try {
            $gcloudCheck = Get-Command gcloud -ErrorAction SilentlyContinue
            if ($null -eq $gcloudCheck) {
                Write-Info "gcloud CLI not found in PATH - skipping token retrieval"
            }
            else {
                Write-Info "Attempting to get token from gcloud (timeout: 5 seconds)..."
                $job = Start-Job -ScriptBlock { gcloud auth print-access-token 2>&1 }
                $completed = Wait-Job -Job $job -Timeout 5

                if ($null -ne $completed) {
                    $AccessToken = Receive-Job -Job $job
                    Remove-Job -Job $job -Force

                    if (-not [string]::IsNullOrWhiteSpace($AccessToken) -and $AccessToken -notmatch 'ERROR') {
                        Write-Success "Access token obtained from gcloud"
                        $tokenAvailable = $true
                    }
                    else {
                        Write-Info "gcloud auth not configured or expired"
                    }
                }
                else {
                    Write-Info "gcloud command timed out - skipping token retrieval"
                    Remove-Job -Job $job -Force
                }
            }
        }
        catch {
            Write-Info "Could not get access token from gcloud: $_"
        }
    }
    else {
        $tokenAvailable = $true
        Write-Success "Using provided access token"
    }

    if (-not $tokenAvailable) {
        Write-Host ""
        Write-Host "========================================" -ForegroundColor Yellow
        Write-Host "  Artifact Enrichment Skipped" -ForegroundColor Yellow
        Write-Host "========================================" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "No OAuth 2.0 access token available." -ForegroundColor Gray
        Write-Host "Artifact enrichment provides detailed metadata for Sheets, Forms, and Scripts." -ForegroundColor Gray
        Write-Host ""
        Write-Host "To enable enrichment, use one of these methods:" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  Method 1: Use gcloud CLI" -ForegroundColor Cyan
        Write-Host "    gcloud auth login" -ForegroundColor White
        Write-Host "    gcloud config set project YOUR_PROJECT_ID" -ForegroundColor White
        Write-Host "    .\Run-FullAssessment.ps1 -PrimaryDomain 'rocheua.com'" -ForegroundColor White
        Write-Host ""
        Write-Host "  Method 2: Provide token directly" -ForegroundColor Cyan
        Write-Host "    `$token = (gcloud auth print-access-token)" -ForegroundColor White
        Write-Host "    .\Run-FullAssessment.ps1 -PrimaryDomain 'rocheua.com' -AccessToken `$token" -ForegroundColor White
        Write-Host ""
        Write-Host "  Method 3: Skip enrichment (current behavior)" -ForegroundColor Cyan
        Write-Host "    .\Run-FullAssessment.ps1 -PrimaryDomain 'rocheua.com' -SkipEnrichment" -ForegroundColor White
        Write-Host ""
        Write-Info "Continuing without artifact enrichment..."
        Write-Host ""
    }
    else {
        $enrichScript = Join-Path $ScriptDir '04_enrich_artifacts.ps1'
        Write-Info "Running artifact enrichment..."

        & pwsh -ExecutionPolicy Bypass -File $enrichScript -AccessToken $AccessToken -OutputDir $OutputDir

        Write-Success "Artifact enrichment completed"
    }

    # Verify enrichment output (only if token was available)
    if ($tokenAvailable) {
        $enrichOutputFiles = @(
            '11_Sheets_Enrichment.csv',
            '12_Forms_Enrichment.csv',
            '13_Scripts_Enrichment.csv'
        )

        foreach ($file in $enrichOutputFiles) {
            $filePath = Join-Path $OutputDir $file
            if (Test-Path $filePath) {
                $rowCount = Get-CsvRowCount -Path $filePath
                Write-Info "  [OK] $file ($rowCount rows)"
            }
            else {
                Write-Info "  [SKIP] $file (not created - no artifacts found)"
            }
        }
    }
}
else {
    Write-Step "STEP 5: Artifact Enrichment (SKIPPED)"
}

# ============================================================================
# STEP 6: COMPLEXITY SCORING
# ============================================================================
Write-Step "STEP 6: Complexity Scoring"

$scoreScript = Join-Path $ScriptDir '05_score_sites.ps1'
Write-Info "Generating complexity report..."

& pwsh -ExecutionPolicy Bypass -File $scoreScript -OutputDir $OutputDir -PrimaryDomain $PrimaryDomain

Write-Success "Complexity scoring completed"

$reportFile = Join-Path $OutputDir '14_Complexity_Report.csv'
if (Test-Path $reportFile) {
    $report = Import-Csv $reportFile
    Write-Info "   14_Complexity_Report.csv ($($report.Count) sites)"

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
    $rows = @(Import-Csv $file.FullName)
    $rowCount = $rows.Count
    Write-Host "  [OK] $($file.Name) " -NoNewline -ForegroundColor Green
    Write-Host "($rowCount rows)" -ForegroundColor Gray
}

Write-Host ""
Write-Host "NEXT STEPS:" -ForegroundColor Cyan
Write-Host "  1. Review the complexity report: 14_Complexity_Report.csv" -ForegroundColor Gray
Write-Host "  2. Analyze high-complexity sites for migration planning" -ForegroundColor Gray
Write-Host "  3. Review HTML snapshots in output/html/ folder" -ForegroundColor Gray
Write-Host ""

Write-Success "Assessment completed successfully!"
Write-Host ""

