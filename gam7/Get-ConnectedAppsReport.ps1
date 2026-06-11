    #Requires -Version 5.1
<#
.SYNOPSIS
    Google Workspace - Connected Apps & Integration Usage Report (GAM / GAMADV-XTD3 / GAM7)

.DESCRIPTION
    Generates a comprehensive HTML + CSV report covering:
      [1] All domain users (status, last login)
      [2] OAuth tokens  - third-party apps connected per user, scopes granted
      [3] App summary   - distinct apps aggregated across all users
      [4] Suspended users still holding OAuth tokens (security flag)
      [5] Marketplace apps - from Cloud Identity Policies (gam print policies)
      [6] Service accounts with Domain-Wide Delegation (gam print svcaccts)
      [7] Admin roles & delegated admins (gam print admins)
      [8] Last activity  (optional -IncludeLastActivity)
      [9] Chat Spaces    (optional -IncludeChatSpaces, needs GAMADV-XTD3/GAM7)

.PARAMETER GamPath
    Full path to gam.exe / gam binary. Auto-detected if omitted.

.PARAMETER InputCsv
    Path to a CSV file listing the users to scan (one user per row).
    The script auto-detects the email column from these names (first match wins):
        primaryEmail, PrimaryEmail, email, Email, userEmail, UserEmail, user, User
    If omitted, the script scans every user in the domain.

.PARAMETER OutputDir
    Where to save all output files. Defaults to .\ConnectedAppsReport_<timestamp>

.PARAMETER IncludeLastActivity
    Collect per-user Drive last-active and last-login timestamps via Reports API.
    Adds noticeable runtime on large domains.

.PARAMETER IncludeChatSpaces
    Collect Google Chat Spaces + members via admin access.
    Requires GAMADV-XTD3 / GAM7 and the Chat admin API scope.

.PARAMETER MaxUsers
    Cap users processed (0 = all). Useful for testing.

.PARAMETER DwdTimeoutSeconds
    Per-strategy timeout (seconds) for the four Domain-Wide Delegation GAM commands.
    If a command does not finish within this limit it is skipped and the next strategy is tried.
    Default: 60. Increase on slow networks; decrease to fail-fast.

.EXAMPLE
    # Basic run
    .\Get-ConnectedAppsReport.ps1

    # Full run with all optional sections
    .\Get-ConnectedAppsReport.ps1 -IncludeLastActivity -IncludeChatSpaces

    # Specify a custom GAM path and output folder
    .\Get-ConnectedAppsReport.ps1 -GamPath "C:\GAM7\gam.exe" -OutputDir "C:\Reports"

    # Test mode - first 30 users only
    .\Get-ConnectedAppsReport.ps1 -MaxUsers 30

    # Scan only the users listed in a CSV file (column: primaryEmail / email / userEmail / ...)
    .\Get-ConnectedAppsReport.ps1 -InputCsv ".\users_to_scan.csv"
#>

[CmdletBinding()]
param(
    [string]$GamPath = "",
    [string]$OutputDir = "",
    [string]$InputCsv = "",
    [switch]$IncludeLastActivity,
    [switch]$IncludeChatSpaces,
    [int]   $MaxUsers = 0,
    [int]   $DwdTimeoutSeconds = 60   # per-strategy timeout for Domain-Wide Delegation GAM calls
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"   # don't abort on non-fatal GAM stderr

# ======================================================
#  HELPERS
# ======================================================
function Write-Banner { param([string]$t); $l = "=" * 64; Write-Host "`n$l`n  $t`n$l" -ForegroundColor Cyan }
function Write-Step { param([string]$t); Write-Host "  [$(Get-Date -f 'HH:mm:ss')] $t" -ForegroundColor Yellow }
function Write-OK { param([string]$t); Write-Host "  [OK]   $t" -ForegroundColor Green }
function Write-Warn { param([string]$t); Write-Host "  [WARN] $t" -ForegroundColor DarkYellow }
function Write-Fail { param([string]$t); Write-Host "  [FAIL] $t" -ForegroundColor Red }

function SafeCsv {
    param([string]$Path)
    if (Test-Path $Path) {
        $raw = Get-Content $Path -Raw -ErrorAction SilentlyContinue
        if ($raw -and $raw.Trim().Length -gt 5) {
            try { return Import-Csv $Path } catch {}
        }
    }
    return @()
}

# Run a GAM command in a background job and return its output.
# Returns $null if the command does not finish within $TimeoutSec seconds.
function Invoke-GamTimeout {
    param([string[]]$GamArgs, [int]$TimeoutSec = 60)
    $gamBin = $script:GAM
    $job = Start-Job -ScriptBlock {
        param($bin, $a)
        & $bin @a 2>&1
    } -ArgumentList $gamBin, (, $GamArgs)

    $done = Wait-Job $job -Timeout $TimeoutSec
    if ($done) {
        $out = Receive-Job $job
        Remove-Job $job -Force
        return $out
    }
    # Timed out – kill the job cleanly
    Stop-Job  $job
    Remove-Job $job -Force
    return $null   # caller checks for $null to detect timeout
}

# ======================================================
#  FIND GAM
# ======================================================
Clear-Host
Write-Host @"

  +--------------------------------------------------------------+
  |  Google Workspace - Connected Apps & Integration Report      |
  |  Powered by GAM / GAMADV-XTD3 / GAM7                        |
  +--------------------------------------------------------------+

"@ -ForegroundColor Cyan

Write-Banner "0. Locating GAM"

# $PSScriptRoot = folder the .ps1 lives in (same as where GAM is when run from the GAM folder)
# $PWD          = current working directory at time of execution
$candidates = @(
    $GamPath,
    ".\gam.exe",                                              # current directory  <-- most common case
    ".\gam",                                                  # Linux/macOS binary in current dir
    "$PSScriptRoot\gam.exe",                                  # same folder as the script
    "$PSScriptRoot\gam",
    "$PWD\gam.exe",                                           # explicit PWD
    "$PWD\gam",
    "gam",                                                    # on PATH
    "$env:USERPROFILE\AppData\Local\GAM7\gam.exe",
    "$env:USERPROFILE\AppData\Local\GAMADV-XTD3\gam.exe",
    "C:\GAM7\gam.exe",
    "C:\GAMADV-XTD3\gam.exe",
    "C:\GAM6\gam.exe",
    "C:\GAM\gam.exe"
)

$script:GAM = $null
foreach ($c in $candidates) {
    if (-not $c) { continue }
    $found = Get-Command $c -ErrorAction SilentlyContinue
    if ($found) { $script:GAM = $found.Source; break }
    if (Test-Path $c) { $script:GAM = (Resolve-Path $c).Path; break }
}

if (-not $script:GAM) {
    Write-Fail "GAM not found. Tried:"
    $candidates | Where-Object { $_ } | ForEach-Object { Write-Host "    $_" -ForegroundColor Gray }
    Write-Host ""
    Write-Host "  Options:" -ForegroundColor Yellow
    Write-Host "    1. Run this script from inside your GAM folder" -ForegroundColor White
    Write-Host "    2. Pass the full path:  .\Get-ConnectedAppsReport.ps1 -GamPath 'C:\GAM7\gam.exe'" -ForegroundColor White
    Write-Host "    3. Install GAM7: https://github.com/taers232c/GAMADV-XTD3" -ForegroundColor White
    exit 1
}

$verLines = & $script:GAM version 2>&1 | Select-Object -First 5
$isGAM7 = @($verLines | Where-Object { $_ -match "GAM7|GAMADV" }).Count -gt 0
$verString = ($verLines | Where-Object { $_ -match "\d+\.\d+" } | Select-Object -First 1) -replace "^\s+", ""
Write-OK "GAM binary : $script:GAM"
Write-OK "Version    : $verString  $(if($isGAM7){'[GAM7/GAMADV-XTD3]'}else{'[Standard GAM]'})"

if ($IncludeChatSpaces -and -not $isGAM7) {
    Write-Warn "Chat Spaces need GAM7/GAMADV-XTD3 - disabling -IncludeChatSpaces."
    $IncludeChatSpaces = [switch]$false
}

if (-not $OutputDir) {
    $OutputDir = Join-Path $PWD "ConnectedAppsReport_$(Get-Date -f 'yyyyMMdd_HHmmss')"
}
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
Write-OK "Output dir : $OutputDir"
$T0 = Get-Date

# ======================================================
#  SECTION 1 - USERS
# ======================================================
Write-Banner "1. Domain Users"

$usersFile = Join-Path $OutputDir "users.csv"

if ($InputCsv) {
    if (-not (Test-Path $InputCsv)) {
        Write-Fail "InputCsv not found: $InputCsv"
        exit 1
    }
    Write-Step "Reading user list from: $InputCsv"
    $inputRows = @(Import-Csv $InputCsv)
    if ($inputRows.Count -eq 0) {
        Write-Fail "InputCsv is empty: $InputCsv"
        exit 1
    }

    # Detect the email column - try common header names first, then any column whose value looks like an email
    $emailCol = $null
    foreach ($cn in @('primaryEmail', 'PrimaryEmail', 'email', 'Email', 'userEmail', 'UserEmail', 'user', 'User')) {
        if ($inputRows[0].PSObject.Properties[$cn]) { $emailCol = $cn; break }
    }
    if (-not $emailCol) {
        foreach ($p in $inputRows[0].PSObject.Properties.Name) {
            if ("$($inputRows[0].$p)" -match "@") { $emailCol = $p; break }
        }
    }
    if (-not $emailCol) {
        Write-Fail "No email column detected in $InputCsv. Add a 'primaryEmail' column."
        exit 1
    }

    $emails = @($inputRows | ForEach-Object { "$($_.$emailCol)".Trim() } |
        Where-Object { $_ -match "@" } | Select-Object -Unique)
    if ($emails.Count -eq 0) {
        Write-Fail "No valid email addresses found in column '$emailCol' of $InputCsv."
        exit 1
    }
    Write-OK "Loaded $($emails.Count) user(s) from $InputCsv (column: $emailCol)"

    # Write a select-file (one email per line) and ask GAM to print only those users
    $selectFile = Join-Path $OutputDir "input_users.txt"
    $emails | Out-File $selectFile -Encoding ASCII

    Write-Step "gam print users select file <input> fields primaryEmail,name,lastLoginTime,suspended,orgUnitPath"
    & $script:GAM print users select file $selectFile fields "primaryEmail,name,lastLoginTime,suspended,orgUnitPath" 2>$null |
    Out-File $usersFile -Encoding UTF8

    $allUsers = SafeCsv $usersFile
    if ($allUsers.Count -eq 0) {
        Write-Warn "'gam print users select file' returned nothing - falling back to a minimal users.csv built from the input emails."
        $fallback = $emails | ForEach-Object {
            [PSCustomObject]@{ primaryEmail = $_; name = ""; lastLoginTime = ""; suspended = "False"; orgUnitPath = "" }
        }
        $fallback | Export-Csv $usersFile -NoTypeInformation -Force
        $allUsers = SafeCsv $usersFile
    }
}
else {
    Write-Step "gam print users fields primaryEmail,name,lastLoginTime,suspended,orgUnitPath"
    & $script:GAM print users fields "primaryEmail,name,lastLoginTime,suspended,orgUnitPath" 2>$null |
    Out-File $usersFile -Encoding UTF8
    $allUsers = SafeCsv $usersFile
}

if ($allUsers.Count -eq 0) {
    Write-Fail "Could not retrieve users. Verify GAM authorisation: gam oauth info"
    exit 1
}
if ($MaxUsers -gt 0 -and $allUsers.Count -gt $MaxUsers) {
    Write-Warn "MaxUsers=$MaxUsers - trimming to first $MaxUsers users (test mode)."
    $allUsers = $allUsers | Select-Object -First $MaxUsers
    $allUsers | Export-Csv $usersFile -NoTypeInformation -Force
}
Write-OK "$($allUsers.Count) users retrieved."

# ======================================================
#  SECTION 2 - OAUTH TOKENS
# ======================================================
Write-Banner "2. OAuth Tokens (Third-Party Connected Apps)"

$tokensFile = Join-Path $OutputDir "tokens_all_users.csv"
Write-Step "Fetching OAuth tokens for all users (may take several minutes on large domains)..."

& $script:GAM redirect csv $tokensFile multiprocess csv $usersFile `
    gam user "~primaryEmail" print tokens 2>$null | Out-Null

$tokenData = SafeCsv $tokensFile
Write-OK "$($tokenData.Count) token records retrieved."

# Token audit log - gives us the first-authorization date per app per user.
# gam report token returns the Admin SDK token activity log.
$tokenAuditData = @()
$tokenAuditFile = Join-Path $OutputDir "token_audit.csv"
Write-Step "Fetching token authorization audit log (last 365 days)..."
$auditRaw = Invoke-GamTimeout @("report", "token", "daysago", "365") 90
if ($null -eq $auditRaw) {
    Write-Warn "Token audit log timed out - CreatedDate will show N/A."
}
else {
    $auditErr = @($auditRaw | Where-Object { $_ -match "ERROR|403|401|unknown command|Invalid" })
    if ($auditErr.Count -gt 0) {
        Write-Warn "Token audit log not accessible - CreatedDate will show N/A."
    }
    else {
        $auditRaw | Out-File $tokenAuditFile -Encoding UTF8
        $tokenAuditData = SafeCsv $tokenAuditFile
        Write-OK "$($tokenAuditData.Count) token audit events retrieved."
    }
}

# GAM7 uses column "user"; older GAM uses "userEmail". Auto-detect.
# GAM7 uses "displayText"; some versions use "appName". Auto-detect.
$colUser = "user"
$colApp = "displayText"
if ($tokenData.Count -gt 0) {
    $sampleProps = $tokenData[0].PSObject.Properties.Name
    if ($sampleProps -contains "userEmail") { $colUser = "userEmail" }
    elseif ($sampleProps -contains "user") { $colUser = "user" }
    if ($sampleProps -contains "displayText") { $colApp = "displayText" }
    elseif ($sampleProps -contains "appName") { $colApp = "appName" }
    Write-OK "Token columns detected: user='$colUser'  app='$colApp'"
}

# Aggregate by app
$appAggFile = Join-Path $OutputDir "apps_aggregated.csv"
Write-Step "Aggregating by app name..."

& $script:GAM redirect csv $appAggFile all users print tokens `
    aggregateusersby displaytext 2>$null | Out-Null

$appAggData = SafeCsv $appAggFile

if ($appAggData.Count -eq 0 -and $tokenData.Count -gt 0) {
    Write-Warn "aggregateusersby not supported on this GAM version - building aggregation manually."
    $appAggData = @($tokenData |
        Where-Object { $_.$colApp -and $_.$colApp.Trim() -ne "" } |
        Group-Object -Property { $_.$colApp } |
        ForEach-Object {
            $uniqueUsers = @($_.Group | ForEach-Object { $_.$colUser } | Select-Object -Unique)
            [PSCustomObject]@{
                AppName   = $_.Name
                ClientId  = ($_.Group | Select-Object -First 1).clientId
                UserCount = $uniqueUsers.Count
                Users     = $uniqueUsers -join "; "
                Scopes    = ($_.Group | Select-Object -First 1).scopes
            }
        } | Sort-Object UserCount -Descending)
    $appAggData | Export-Csv $appAggFile -NoTypeInformation
}
Write-OK "$($appAggData.Count) distinct apps found."

# ======================================================
#  SECTION 3 - MARKETPLACE APPS
#
#  Strategy (tried in order):
#  [A] gam print appdetails  - Chrome Mgmt API (GAM7/GAMADV, shows installed Marketplace apps)
#  [B] gam print policies type app_access_settings - Cloud Identity Policies
#  [C] gam print policies filter "workspace_marketplace" - older filter syntax
#  [D] Build from OAuth token data as best-effort fallback
# ======================================================
Write-Banner "3. Marketplace Apps"

$marketplaceFile = Join-Path $OutputDir "marketplace_apps.csv"
$marketplaceData = @()
$mktMethod = "none"

# --- Strategy A: gam print appdetails (most reliable on GAM7/GAMADV) ---
Write-Step "[1/4] Trying: gam print appdetails type CHROME_EXTENSION (Chrome Mgmt / App Details API)..."
$mktRawA = & $script:GAM print appdetails type CHROME_EXTENSION 2>&1
if (-not ($mktRawA | Where-Object { $_ -match "ERROR|403|401|unknown|invalid" })) {
    $mktRawA | Out-File $marketplaceFile -Encoding UTF8
    $tryA = SafeCsv $marketplaceFile
    if ($tryA.Count -gt 0) { $marketplaceData = $tryA; $mktMethod = "appdetails-chrome" }
}

# Also try ANDROID and WEB app types and merge (only when CHROME_EXTENSION query succeeded)
if ($marketplaceData.Count -gt 0) {
    foreach ($appType in @("ANDROID", "WEB", "GOOGLE_WORKSPACE_APP")) {
        $rawT = & $script:GAM print appdetails type $appType 2>&1
        if (-not ($rawT | Where-Object { $_ -match "ERROR|403|401|unknown|invalid" })) {
            $tmpF = Join-Path $OutputDir "mkt_tmp_$appType.csv"
            $rawT | Out-File $tmpF -Encoding UTF8
            $tmpD = SafeCsv $tmpF
            if ($tmpD.Count -gt 0) {
                $marketplaceData += $tmpD
                $mktMethod = "appdetails-multi"
            }
            Remove-Item $tmpF -Force -ErrorAction SilentlyContinue
        }
    }
}

# --- Strategy B: gam print policies type app_access_settings ---
if ($marketplaceData.Count -eq 0) {
    Write-Step "[2/4] Trying: gam print policies type app_access_settings..."
    $mktRawB = & $script:GAM print policies type app_access_settings 2>&1
    if (-not ($mktRawB | Where-Object { $_ -match "ERROR|403|401" })) {
        $mktRawB | Out-File $marketplaceFile -Encoding UTF8
        $tryB = SafeCsv $marketplaceFile
        if ($tryB.Count -gt 0) { $marketplaceData = $tryB; $mktMethod = "policies-app_access" }
    }
}

# --- Strategy C: gam print policies (all) and filter for marketplace rows ---
if ($marketplaceData.Count -eq 0) {
    Write-Step "[3/4] Trying: gam print policies (all, then filter for marketplace/app rows)..."
    $mktRawC = & $script:GAM print policies 2>&1
    if (-not ($mktRawC | Where-Object { $_ -match "ERROR|403|401" })) {
        $tmpAllF = Join-Path $OutputDir "policies_all.csv"
        $mktRawC | Out-File $tmpAllF -Encoding UTF8
        $allPolicies = SafeCsv $tmpAllF
        if ($allPolicies.Count -gt 0) {
            # Save full policies file
            $allPolicies | Export-Csv (Join-Path $OutputDir "policies_all.csv") -NoTypeInformation -Force
            # Filter for marketplace/app-related rows (check property values directly - no CSV round-trip)
            $mktFiltered = $allPolicies | Where-Object {
                ($_.PSObject.Properties.Value -join "|") -match "marketplace|appId|app_access|chrome_app|workspace_app"
            }
            if ($mktFiltered.Count -gt 0) {
                $marketplaceData = $mktFiltered
                $mktMethod = "policies-filtered"
            }
            else {
                # No marketplace-specific rows — save all policies as the output
                $marketplaceData = $allPolicies
                $mktMethod = "policies-all"
                Write-Warn "No marketplace-specific policies found. Saving all $($allPolicies.Count) policy rows."
            }
        }
    }
}

# --- Strategy D: Build from OAuth token data (best-effort fallback) ---
if ($marketplaceData.Count -eq 0 -and $tokenData.Count -gt 0) {
    Write-Step "[4/4] Falling back: building app list from OAuth token data..."
    $marketplaceData = @($tokenData |
        Where-Object { $_.$colApp -and $_.$colApp.Trim() -ne "" } |
        Group-Object -Property { $_.$colApp } |
        ForEach-Object {
            $uniqueUsers = @($_.Group | ForEach-Object { $_.$colUser } | Select-Object -Unique)
            [PSCustomObject]@{
                AppName         = $_.Name
                ClientId        = ($_.Group | Select-Object -First 1).clientId
                AuthorizedUsers = $uniqueUsers.Count
                ScopesGranted   = ($_.Group | Select-Object -First 1).scopes
                DataSource      = "OAuth tokens (no Marketplace API access)"
            }
        } | Sort-Object AuthorizedUsers -Descending)
    $mktMethod = "oauth-fallback"
    Write-Warn "Using OAuth token data as Marketplace fallback. To get real Marketplace data:"
    Write-Warn "  Run: gam oauth update  ->  enable 'Chrome Management API - AppDetails read only'"
    Write-Warn "  OR ensure 'Cloud Identity - Policies' scope is authorised."
}

if ($marketplaceData.Count -gt 0) {
    $marketplaceData | Export-Csv $marketplaceFile -NoTypeInformation -Force
    Write-OK "$($marketplaceData.Count) Marketplace/app row(s) retrieved (method: $mktMethod)."
}
else {
    $marketplaceData = @()   # ensure it is always a proper array
    Write-Warn "No Marketplace data retrieved via any method."
    Write-Warn "To enable: run 'gam oauth update' and authorise:"
    Write-Warn "  - 'Chrome Management API - AppDetails read only'"
    Write-Warn "  - 'Cloud Identity - Policies'"
}

# ======================================================
#  SECTION 4 - SERVICE ACCOUNTS (Domain-Wide Delegation)
#
#  Strategy (tried in order):
#  [A] gam print domainwidedelegation  - most direct command (GAM7/GAMADV)
#  [B] gam print svcaccts              - older GAM command
#  [C] gam show domainwidedelegation   - show variant if print fails
#  [D] gam info domain (extract DwD info if available)
# ======================================================
Write-Banner "4. Service Accounts with Domain-Wide Delegation"

$svcAcctFile = Join-Path $OutputDir "service_accounts_dwd.csv"
$svcAcctData = @()
$svcMethod = "none"
$svcErrors = $false

# --- Strategy A: gam print domainwidedelegation (GAM7 / GAMADV-XTD3 preferred) ---
Write-Step "[1/4] Trying: gam print domainwidedelegation... (timeout: ${DwdTimeoutSeconds}s)"
$dwdRawA = Invoke-GamTimeout @("print", "domainwidedelegation") $DwdTimeoutSeconds
if ($null -eq $dwdRawA) {
    Write-Warn "[1/4] Timed out after ${DwdTimeoutSeconds}s - skipping to next strategy."
}
else {
    $dwdErrA = $dwdRawA | Where-Object { $_ -match "ERROR|403|401|unknown command|Invalid" }
    if (-not $dwdErrA) {
        $dwdRawA | Out-File $svcAcctFile -Encoding UTF8
        $tryA = SafeCsv $svcAcctFile
        if ($tryA.Count -gt 0) { $svcAcctData = $tryA; $svcMethod = "domainwidedelegation" }
        elseif (@($dwdRawA | Where-Object { $_ -and $_.Trim() -ne "" }).Count -le 2) {
            $svcAcctData = @()
            $svcMethod = "domainwidedelegation-empty"
        }
    }
}

# --- Strategy B: gam print svcaccts ---
if ($svcAcctData.Count -eq 0 -and $svcMethod -ne "domainwidedelegation-empty") {
    Write-Step "[2/4] Trying: gam print svcaccts... (timeout: ${DwdTimeoutSeconds}s)"
    $dwdRawB = Invoke-GamTimeout @("print", "svcaccts") $DwdTimeoutSeconds
    if ($null -eq $dwdRawB) {
        Write-Warn "[2/4] Timed out after ${DwdTimeoutSeconds}s - skipping to next strategy."
    }
    else {
        $dwdErrB = $dwdRawB | Where-Object { $_ -match "ERROR|403|401|unknown command|Invalid" }
        if (-not $dwdErrB) {
            $dwdRawB | Out-File $svcAcctFile -Encoding UTF8
            $tryB = SafeCsv $svcAcctFile
            if ($tryB.Count -gt 0) { $svcAcctData = $tryB; $svcMethod = "svcaccts" }
        }
    }
}

# --- Strategy C: gam show domainwidedelegation -> convert text to CSV ---
if ($svcAcctData.Count -eq 0 -and $svcMethod -ne "domainwidedelegation-empty") {
    Write-Step "[3/4] Trying: gam show domainwidedelegation (text output)... (timeout: ${DwdTimeoutSeconds}s)"
    $dwdRawC = Invoke-GamTimeout @("show", "domainwidedelegation") $DwdTimeoutSeconds
    if ($null -eq $dwdRawC) {
        Write-Warn "[3/4] Timed out after ${DwdTimeoutSeconds}s - skipping to next strategy."
    }
    else {
        $dwdErrC = $dwdRawC | Where-Object { $_ -match "ERROR|403|401|unknown command|Invalid" }
        if (-not $dwdErrC) {
            # Parse text output into structured objects
            $parsed = @()
            $current = $null
            foreach ($line in $dwdRawC) {
                if ($line -match "^Client ID:\s*(.+)$") {
                    if ($current) { $parsed += $current }
                    $current = [PSCustomObject]@{ ClientID = ""; DisplayName = ""; Scopes = ""; IssuedTo = "" }
                    $current.ClientID = $Matches[1].Trim()
                }
                elseif ($line -match "^\s*Display Name:\s*(.+)$" -and $current) {
                    $current.DisplayName = $Matches[1].Trim()
                }
                elseif ($line -match "^\s*Issued To:\s*(.+)$" -and $current) {
                    $current.IssuedTo = $Matches[1].Trim()
                }
                elseif ($line -match "^\s*Scope[s]?:\s*(.+)$" -and $current) {
                    $current.Scopes = $Matches[1].Trim()
                }
                elseif ($line -match "^\s+https://" -and $current) {
                    $current.Scopes += "; " + $line.Trim()
                }
            }
            if ($current) { $parsed += $current }

            if ($parsed.Count -gt 0) {
                $svcAcctData = $parsed
                $svcAcctData | Export-Csv $svcAcctFile -NoTypeInformation -Force
                $svcMethod = "show-domainwidedelegation-parsed"
            }
            else {
                # Output came back but was unparseable - save raw text
                $dwdRawC | Out-File $svcAcctFile -Encoding UTF8
                $svcMethod = "show-text-only"
            }
        }
    }
}

# --- Strategy D: gam info domain (check if DwD info embedded) ---
if ($svcAcctData.Count -eq 0 -and $svcMethod -ne "domainwidedelegation-empty") {
    Write-Step "[4/4] Trying: gam info domain (extracting DwD mentions)... (timeout: ${DwdTimeoutSeconds}s)"
    $dwdRawD = Invoke-GamTimeout @("info", "domain") $DwdTimeoutSeconds
    if ($null -eq $dwdRawD) {
        Write-Warn "[4/4] Timed out after ${DwdTimeoutSeconds}s."
        $svcErrors = $true
    }
    else {
        $dwdLines = $dwdRawD | Where-Object { $_ -match "delegation|svcacct|serviceaccount|clientid|DwD" }
        if (@($dwdLines).Count -gt 0) {
            $dwdLines | Out-File $svcAcctFile -Encoding UTF8
            $svcMethod = "info-domain-partial"
            Write-Warn "Only partial DwD info available via 'gam info domain'. See service_accounts_dwd.csv."
        }
        else {
            $svcErrors = $true
        }
    }
}

# Report outcome
if ($svcAcctData.Count -gt 0) {
    Write-OK "$($svcAcctData.Count) DwD service account(s) found (method: $svcMethod)."
    Write-Warn "SECURITY: Each DwD service account can impersonate ANY user in your domain. Review carefully."
}
elseif ($svcMethod -eq "domainwidedelegation-empty") {
    Write-OK "No DwD service accounts configured (domain-wide delegation list is empty)."
}
else {
    Write-Warn "Could not retrieve DwD service accounts via any method."
    Write-Warn "To enable, run:  gam oauth update"
    Write-Warn "Then authorise:  'Admin SDK - Other' and 'Cloud Identity - Policies' scopes."
    Write-Warn "Also verify at:  Admin Console > Security > API Controls > Domain-wide Delegation"
}

# ======================================================
#  SECTION 5 - ADMIN ROLES
# ======================================================
Write-Banner "5. Admin Roles & Delegated Admins"

$adminsFile = Join-Path $OutputDir "admins.csv"
$adminsData = @()

Write-Step "gam print admins"
$admRaw = & $script:GAM print admins 2>&1
$admErrors = $admRaw | Where-Object { $_ -match "ERROR|403|401" }
if ($admErrors) {
    Write-Warn "Admin query error: $($admErrors -join '; ')"
}
else {
    $admRaw | Out-File $adminsFile -Encoding UTF8
    $adminsData = SafeCsv $adminsFile
    Write-OK "$($adminsData.Count) admin role assignment(s) retrieved."
}

# ======================================================
#  SECTION 6 - LAST ACTIVITY (optional)
# ======================================================
$activityData = @()
Write-Banner "6. Last Activity$(if(-not $IncludeLastActivity){' [SKIPPED]'})"

if ($IncludeLastActivity) {
    $activityFile = Join-Path $OutputDir "last_activity.csv"
    Write-Step "Fetching Drive + login activity (Reports API - may be slow on large domains)..."
    & $script:GAM redirect csv $activityFile multiprocess csv $usersFile `
        gam report user user "~primaryEmail" `
        parameters "drive:timestamp_last_active_usage,accounts:last_login_time" 2>$null | Out-Null
    $activityData = SafeCsv $activityFile
    Write-OK "Activity data for $($activityData.Count) users."
}
else {
    Write-Warn "Skipped. Use -IncludeLastActivity to collect per-user Drive/login timestamps."
}

# ======================================================
#  SECTION 7 - CHAT SPACES (optional)
# ======================================================
$chatSpaceData = @()
$chatMemberData = @()
Write-Banner "7. Google Chat Spaces$(if(-not $IncludeChatSpaces){' [SKIPPED]'})"

if ($IncludeChatSpaces) {
    $chatSpacesFile = Join-Path $OutputDir "chat_spaces.csv"
    $chatMembersFile = Join-Path $OutputDir "chat_members.csv"

    Write-Step "gam print chatspaces asadmin"
    & $script:GAM redirect csv $chatSpacesFile print chatspaces asadmin 2>$null | Out-Null
    $chatSpaceData = SafeCsv $chatSpacesFile
    Write-OK "$($chatSpaceData.Count) Chat Spaces found."

    if ($chatSpaceData.Count -gt 0) {
        Write-Step "Fetching members for all spaces..."
        & $script:GAM redirect csv $chatMembersFile multiprocess csv $chatSpacesFile `
            gam print chatmembers "~name" asadmin 2>$null | Out-Null
        $chatMemberData = SafeCsv $chatMembersFile
        Write-OK "$($chatMemberData.Count) membership records."
    }
}
else {
    Write-Warn "Skipped. Use -IncludeChatSpaces (requires GAM7/GAMADV-XTD3) to collect Chat data."
}

# ======================================================
#  BUILD HTML REPORT
# ======================================================
Write-Banner "8. Building HTML Report"

Add-Type -AssemblyName System.Web

function To-HtmlTable {
    param([object[]]$Data, [int]$Limit = 500, [string]$Empty = "No data available.")
    if (-not $Data -or $Data.Count -eq 0) { return "<p class='empty'>$Empty</p>" }
    $rows = if ($Data.Count -gt $Limit) { $Data | Select-Object -First $Limit } else { $Data }
    $props = $rows[0].PSObject.Properties.Name
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append("<table><thead><tr>")
    foreach ($p in $props) { [void]$sb.Append("<th>$([System.Web.HttpUtility]::HtmlEncode($p))</th>") }
    [void]$sb.Append("</tr></thead><tbody>")
    foreach ($row in $rows) {
        [void]$sb.Append("<tr>")
        foreach ($p in $props) {
            $v = "$($row.$p)"
            if ($v.Length -gt 350) { $v = $v.Substring(0, 347) + "..." }
            [void]$sb.Append("<td>$([System.Web.HttpUtility]::HtmlEncode($v))</td>")
        }
        [void]$sb.Append("</tr>")
    }
    [void]$sb.Append("</tbody></table>")
    if ($Data.Count -gt $Limit) {
        [void]$sb.Append("<p class='note'>Showing first $Limit of $($Data.Count) rows. See CSV for full data.</p>")
    }
    return $sb.ToString()
}

# Stats
$totalUsers = $allUsers.Count
$activeUsers = ($allUsers | Where-Object { $_.suspended -ne "True" }).Count
$suspUsers = $totalUsers - $activeUsers
$totalApps = $appAggData.Count
$totalTokens = $tokenData.Count
$totalMktApps = $marketplaceData.Count
$totalSvcAccts = $svcAcctData.Count
$totalSpaces = $chatSpaceData.Count
$reportDate = Get-Date -Format "dddd dd MMM yyyy, HH:mm:ss"
$elapsed = [math]::Round(((Get-Date) - $T0).TotalSeconds, 1)

$suspEmails = @($allUsers | Where-Object { $_.suspended -eq "True" } | Select-Object -ExpandProperty primaryEmail)
# @() wrapper prevents if-else from returning $null when the pipeline emits an empty array
$suspWithTokens = @(if ($tokenData.Count -gt 0 -and $suspEmails.Count -gt 0) {
        $tokenData | Where-Object { $suspEmails -contains $_.$colUser } | Select-Object -First 100
    })

$perUserSummary = @(if ($tokenData.Count -gt 0) {
        $tokenData | Group-Object -Property { $_.$colUser } | ForEach-Object {
            [PSCustomObject]@{
                User     = $_.Name
                AppCount = $_.Count
                Apps     = (@($_.Group | ForEach-Object { $_.$colApp } | Select-Object -Unique) -join " | ")
            }
        } | Sort-Object AppCount -Descending | Select-Object -First 100
    })

$topApps = @(if ($appAggData.Count -gt 0) {
        $appAggData | Select-Object -First 25 | ForEach-Object {
            # Use PSObject.Properties lookup so strict mode never throws on missing columns
            $tName = if ($_.PSObject.Properties['AppName'] -and $_.AppName) { $_.AppName }
            elseif ($_.PSObject.Properties['displayText'] -and $_.displayText) { $_.displayText }
            else { "N/A" }
            $tCid = if ($_.PSObject.Properties['ClientId'] -and $_.ClientId) { $_.ClientId }
            elseif ($_.PSObject.Properties['clientId'] -and $_.clientId) { $_.clientId }
            else { "N/A" }
            $tCount = if ($_.PSObject.Properties['UserCount'] -and $_.UserCount) { $_.UserCount }
            else { "N/A" }
            [PSCustomObject]@{ AppName = $tName; ClientId = $tCid; UserCount = $tCount }
        }
    })

# ── Per-app Detail Summary: AppName, Department, CreatedDate, LastLogin, LastActivity, Members
# Build fast email-keyed lookup tables so we don't loop through arrays repeatedly
$userLookup = @{}
foreach ($u in $allUsers) { $userLookup[$u.primaryEmail] = $u }

$activityLookup = @{}
if ($activityData.Count -gt 0) {
    foreach ($a in $activityData) {
        $em = ""
        foreach ($cn in @('email', 'user_email', 'userEmail', 'User Email')) {
            if ($a.PSObject.Properties[$cn] -and $a.$cn) { $em = $a.$cn; break }
        }
        if ($em) { $activityLookup[$em] = $a }
    }
}

# Build a lookup from ClientId -> earliest authorize event time (from audit log)
# gam report token returns rows with varying column names across versions; try all known patterns
$auditByClientId = @{}
if ($tokenAuditData.Count -gt 0) {
    # Discover which column holds the time and which holds the client ID
    $sampleAudit = $tokenAuditData[0]
    $auditTimeCol = $sampleAudit.PSObject.Properties.Name |
    Where-Object { $_ -match "^time$|^id\.time$|^date$|^timestamp$" } | Select-Object -First 1
    $auditCidCols = $sampleAudit.PSObject.Properties.Name |
    Where-Object { $_ -match "client_id|clientId|client\.id" }

    foreach ($ev in $tokenAuditData) {
        # Only count "authorize" events (not revoke)
        $evVals = $ev.PSObject.Properties.Value -join "|"
        if ($evVals -match "revoke") { continue }

        $cKey = ""
        foreach ($cc in $auditCidCols) {
            if ($ev.PSObject.Properties[$cc] -and $ev.$cc) { $cKey = $ev.$cc.Trim(); break }
        }
        # Fallback: scan all values for a known googleapis client-id pattern
        if (-not $cKey) {
            $cKey = ($ev.PSObject.Properties.Value | Where-Object { $_ -match "\.apps\.googleusercontent\.com$" } | Select-Object -First 1)
        }
        if (-not $cKey) { continue }

        if ($auditTimeCol -and $ev.PSObject.Properties[$auditTimeCol] -and $ev.$auditTimeCol) {
            $ts = $ev.$auditTimeCol
            if (-not $auditByClientId.ContainsKey($cKey)) {
                $auditByClientId[$cKey] = $ts
            }
            else {
                try {
                    if ([DateTime]::Parse($ts) -lt [DateTime]::Parse($auditByClientId[$cKey])) {
                        $auditByClientId[$cKey] = $ts
                    }
                }
                catch {}
            }
        }
    }
}

$appDetailSummary = @(if ($tokenData.Count -gt 0) {
        $tokenData |
        Where-Object { $_.$colApp -and $_.$colApp.Trim() -ne "" } |
        Group-Object -Property { $_.$colApp } |
        ForEach-Object {
            $grp = $_.Group
            $first = $grp | Select-Object -First 1
            $cid = if ($first.PSObject.Properties['clientId'] -and $first.clientId) { $first.clientId } else { "N/A" }
            $scps = if ($first.PSObject.Properties['scopes'] -and $first.scopes) { $first.scopes }   else { "N/A" }

            $emails = @($grp | ForEach-Object { $_.$colUser } | Select-Object -Unique | Sort-Object)

            # Department/Org unit paths of every user who granted this app (de-duplicated)
            $department = (@($emails | ForEach-Object {
                        if ($userLookup.ContainsKey($_) -and $userLookup[$_].PSObject.Properties['orgUnitPath']) {
                            $userLookup[$_].orgUnitPath
                        }
                    } | Select-Object -Unique | Where-Object { $_ } | Sort-Object) -join "; ")
            if (-not $department) { $department = "N/A" }

            # First-authorization date from token audit log
            $createdDate = if ($auditByClientId.ContainsKey($cid)) { $auditByClientId[$cid] } else { "N/A" }

            # Most-recent lastLoginTime among all users who have this app
            $lastLogin = "N/A"
            $latestLogin = [DateTime]::MinValue
            foreach ($em in $emails) {
                if ($userLookup.ContainsKey($em) -and
                    $userLookup[$em].PSObject.Properties['lastLoginTime'] -and
                    $userLookup[$em].lastLoginTime) {
                    try {
                        $d = [DateTime]::Parse($userLookup[$em].lastLoginTime)
                        if ($d -gt $latestLogin) { $latestLogin = $d; $lastLogin = $userLookup[$em].lastLoginTime }
                    }
                    catch {}
                }
            }

            # Most-recent Drive activity (only available when -IncludeLastActivity is used)
            $lastActivity = if ($activityLookup.Count -eq 0) { "Add -IncludeLastActivity flag" } else { "N/A" }
            if ($activityLookup.Count -gt 0) {
                $latestAct = [DateTime]::MinValue
                foreach ($em in $emails) {
                    if ($activityLookup.ContainsKey($em)) {
                        $a = $activityLookup[$em]
                        foreach ($actCol in @('drive:timestamp_last_active_usage', 'drive_last_active', 'lastActivity')) {
                            if ($a.PSObject.Properties[$actCol] -and $a.$actCol) {
                                try {
                                    $d = [DateTime]::Parse($a.$actCol)
                                    if ($d -gt $latestAct) { $latestAct = $d; $lastActivity = $a.$actCol }
                                }
                                catch {}
                                break
                            }
                        }
                    }
                }
            }

            [PSCustomObject]@{
                AppName      = $_.Name
                ClientId     = $cid
                Members      = $emails.Count
                Department   = $department
                CreatedDate  = $createdDate
                LastLogin    = $lastLogin
                LastActivity = $lastActivity
                Scopes       = $scps
                MemberList   = $emails -join "; "
            }
        } | Sort-Object Members -Descending
    })

$appDetailFile = Join-Path $OutputDir "app_detail_summary.csv"
if ($appDetailSummary.Count -gt 0) {
    $appDetailSummary | Export-Csv $appDetailFile -NoTypeInformation -Force
    Write-OK "App detail summary : $($appDetailSummary.Count) apps -> app_detail_summary.csv"
}

$reportFile = Join-Path $OutputDir "ConnectedApps_Report.html"

# ── Pre-compute every HTML fragment that contains logic, multi-line strings,
#    switch statements, em-dashes, or & characters.
#    Inside the @"..."@ heredoc we then just drop in $htmlXxx variables.

# Suspended tokens block
if ($suspWithTokens.Count -gt 0) {
    $htmlSuspAlert = "<div class='warn'>These suspended accounts still hold active OAuth tokens. " +
    "Revoke with: <code>gam user &lt;email&gt; delete tokens</code></div>"
}
else {
    $htmlSuspAlert = "<div class='ok'>No suspended accounts with active OAuth tokens found.</div>"
}
$htmlSuspTable = To-HtmlTable $suspWithTokens -Limit 100 -Empty "No suspended users with active tokens - good!"
$htmlPerUserTbl = To-HtmlTable $perUserSummary -Limit 100
$htmlRawTokTbl = To-HtmlTable $tokenData      -Limit 300
$htmlTopAppsTbl = To-HtmlTable $topApps         -Limit 25  -Empty "No OAuth token data retrieved."
$htmlAppDetailTbl = To-HtmlTable $appDetailSummary -Limit 500 -Empty "No OAuth token data retrieved."

# Activity note shown inside the App Detail section
if ($IncludeLastActivity) {
    $htmlActNote = "<div class='ok'>Drive last-activity timestamps collected via Reports API.</div>"
}
else {
    $htmlActNote = "<div class='warn'>LastActivity shows <em>Re-run with -IncludeLastActivity</em> because " +
    "the Reports API was not called. Re-run with <strong>-IncludeLastActivity</strong> to populate it.</div>"
}

# Marketplace block
if ($marketplaceData.Count -gt 0) {
    $mktNote = switch -Wildcard ($mktMethod) {
        "appdetails*" { "Retrieved via Chrome Management / App Details API (gam print appdetails)." }
        "policies-app_access" { "Retrieved via Cloud Identity Policies API - app access settings." }
        "policies-filtered" { "Retrieved by filtering all Cloud Identity Policies for app-related rows." }
        "policies-all" { "All Cloud Identity Policies returned (no marketplace filter matched). Review manually." }
        "oauth-fallback" {
            "Built from OAuth token data (best-effort). Not the same as the Admin Console Marketplace list. " +
            "To get real data run: gam oauth update and enable Chrome Management API or Cloud Identity Policies scope." 
        }
        default { "Data source: $mktMethod" }
    }
    $htmlMarketplace = "<div class='info'>$mktNote</div>" + (To-HtmlTable $marketplaceData -Limit 200)
}
else {
    $htmlMarketplace = "<div class='warn'>" +
    "No Marketplace or App Details data returned via any method.<br>" +
    "Possible reasons: no Marketplace apps configured, or required API scopes not authorised.<br><br>" +
    "<strong>To fix:</strong> run <code>gam oauth update</code> and authorise:<br>" +
    "Chrome Management API - AppDetails read only, or Cloud Identity - Policies.<br>" +
    "Then re-run this script. Verify apps at: Admin Console &gt; Apps &gt; Google Workspace Marketplace apps." +
    "</div>"
}

# Service accounts / DwD block
if ($svcAcctData.Count -gt 0) {
    $htmlSvcAccts = "<div class='warn'>SECURITY REVIEW REQUIRED: Each DwD service account can impersonate " +
    "any user in the domain. Verify each entry is still needed. Retrieved via: <code>$svcMethod</code></div>" +
    (To-HtmlTable $svcAcctData -Limit 200)
}
elseif ($svcMethod -eq "domainwidedelegation-empty") {
    $htmlSvcAccts = "<div class='ok'>No Domain-Wide Delegation service accounts are configured. This is the safest state.</div>"
}
elseif ($svcMethod -eq "show-text-only") {
    $htmlSvcAccts = "<div class='warn'>DwD data returned as unstructured text. See <strong>service_accounts_dwd.csv</strong>.</div>"
}
else {
    $htmlSvcAccts = "<div class='warn'>" +
    "Could not retrieve DwD service accounts via any of the 4 methods attempted.<br><br>" +
    "<strong>To fix:</strong> run <code>gam oauth update</code> and authorise " +
    "Admin SDK - Other and Cloud Identity - Policies.<br>" +
    "Or verify manually: Admin Console &gt; Security &gt; API Controls &gt; Domain-wide Delegation." +
    "</div>"
}

# Last Activity block
if ($IncludeLastActivity -and $activityData.Count -gt 0) {
    $htmlActivity = To-HtmlTable $activityData -Limit 200
}
elseif ($IncludeLastActivity) {
    $htmlActivity = "<div class='warn'>Activity data was requested but no rows returned. The Reports API may need up to 48h to populate.</div>"
}
else {
    $htmlActivity = "<div class='warn'>Not collected. Re-run with <strong>-IncludeLastActivity</strong> to include Drive last-active and login timestamps.</div>"
}

# Chat Spaces block
if ($IncludeChatSpaces -and $chatSpaceData.Count -gt 0) {
    $htmlChat = To-HtmlTable $chatSpaceData -Limit 200
}
elseif ($IncludeChatSpaces) {
    $htmlChat = "<div class='warn'>Chat Spaces were requested but no data returned. Ensure the Google Chat API - Admin scope is authorised via <code>gam oauth update</code>.</div>"
}
else {
    $htmlChat = "<div class='warn'>Not collected. Re-run with <strong>-IncludeChatSpaces</strong> (requires GAM7/GAMADV-XTD3).</div>"
}

# Optional file rows
$htmlActivityFileRow = ""
$htmlChatFileRows = ""
if ($IncludeLastActivity) {
    $htmlActivityFileRow = "<tr><td>last_activity.csv</td><td>Per-user Drive last-active and login timestamps</td><td>$($activityData.Count)</td></tr>"
}
if ($IncludeChatSpaces) {
    $htmlChatFileRows = "<tr><td>chat_spaces.csv</td><td>Google Chat Spaces</td><td>$($chatSpaceData.Count)</td></tr>" +
    "<tr><td>chat_members.csv</td><td>Chat Space memberships</td><td>$($chatMemberData.Count)</td></tr>"
}

$htmlUsersTbl = To-HtmlTable $allUsers   -Limit 300
$htmlAdminsTbl = To-HtmlTable $adminsData -Limit 200 -Empty "No admin role data retrieved."

# CSS is built as a plain string OUTSIDE the heredoc to prevent PowerShell
# from misinterpreting CSS custom properties (--varname) as the -- operator.
$htmlCss = "<style>`n" +
":root{background-color:#f0f4f8}" + "`n" +
"body{font-family:'Segoe UI',Roboto,Arial,sans-serif;background:#f0f4f8;color:#202124;font-size:14px;margin:0}" + "`n" +
"header{background:linear-gradient(135deg,#1a73e8,#0d47a1);color:#fff;padding:28px 36px}" + "`n" +
"header h1{font-size:1.6rem;font-weight:600}" + "`n" +
"header p{margin-top:5px;opacity:.8;font-size:.88rem}" + "`n" +
".wrap{max-width:1400px;margin:0 auto;padding:24px 20px}" + "`n" +
"*{box-sizing:border-box}" + "`n" +
".cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:14px;margin-bottom:22px}" + "`n" +
".card{background:#fff;border-radius:10px;padding:18px 14px;text-align:center;box-shadow:0 1px 3px rgba(0,0,0,.1);border-top:4px solid #1a73e8}" + "`n" +
".card.g{border-top-color:#34a853}.card.y{border-top-color:#f9ab00}.card.r{border-top-color:#ea4335}.card.p{border-top-color:#9334e6}" + "`n" +
".card-num{font-size:2rem;font-weight:700;color:#1a73e8}" + "`n" +
".card.g .card-num{color:#34a853}.card.y .card-num{color:#e37400}.card.r .card-num{color:#ea4335}.card.p .card-num{color:#9334e6}" + "`n" +
".card-label{font-size:.72rem;color:#5f6368;margin-top:3px;text-transform:uppercase;letter-spacing:.5px}" + "`n" +
".sec{background:#fff;border-radius:10px;padding:22px;margin-bottom:20px;box-shadow:0 1px 3px rgba(0,0,0,.07)}" + "`n" +
".sec h2{font-size:1rem;font-weight:600;color:#1a73e8;border-bottom:2px solid #dadce0;padding-bottom:9px;margin-bottom:16px;display:flex;align-items:center;gap:8px}" + "`n" +
".badge{font-size:.72rem;background:#1a73e8;color:#fff;padding:2px 8px;border-radius:10px}" + "`n" +
".badge.r{background:#ea4335}.badge.g{background:#34a853}.badge.y{background:#f9ab00;color:#333}" + "`n" +
".info{background:#e8f0fe;border-left:4px solid #1a73e8;border-radius:4px;padding:10px 14px;margin-bottom:14px;font-size:.83rem;color:#174ea6}" + "`n" +
".warn{background:#fef7e0;border-left:4px solid #f9ab00;border-radius:4px;padding:10px 14px;margin-bottom:14px;font-size:.83rem;color:#7a4f00}" + "`n" +
".ok{background:#e6f4ea;border-left:4px solid #34a853;border-radius:4px;padding:10px 14px;margin-bottom:14px;font-size:.83rem;color:#1e4620}" + "`n" +
".tabs{display:flex;flex-wrap:wrap;gap:2px;border-bottom:2px solid #dadce0;margin-bottom:0}" + "`n" +
".tab{padding:7px 16px;cursor:pointer;border:none;background:none;font-size:.83rem;color:#5f6368;border-bottom:3px solid transparent;margin-bottom:-2px;border-radius:4px 4px 0 0}" + "`n" +
".tab:hover{background:#f1f3f4;color:#202124}.tab.on{color:#1a73e8;border-bottom-color:#1a73e8;font-weight:600}" + "`n" +
".pane{display:none;padding-top:18px}.pane.on{display:block}" + "`n" +
".tbl{overflow-x:auto}" + "`n" +
"table{width:100%;border-collapse:collapse;font-size:.81rem}" + "`n" +
"thead th{background:#f8f9fa;color:#5f6368;font-weight:600;text-transform:uppercase;font-size:.71rem;padding:9px 11px;text-align:left;border-bottom:2px solid #dadce0;white-space:nowrap}" + "`n" +
"tbody tr:hover{background:#f1f3f4}" + "`n" +
"tbody td{padding:8px 11px;border-bottom:1px solid #dadce0;vertical-align:top;max-width:380px;word-break:break-word}" + "`n" +
"tbody tr:last-child td{border-bottom:none}" + "`n" +
".empty,.note{color:#5f6368;font-style:italic;padding:10px 0;font-size:.83rem}" + "`n" +
"code{background:#f1f3f4;padding:1px 5px;border-radius:3px;font-size:.82rem}" + "`n" +
"footer{text-align:center;padding:20px;color:#5f6368;font-size:.78rem}" + "`n" +
"</style>"

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Connected Apps Report</title>
$htmlCss
</head>
<body>
<header>
  <h1>&#128279; Connected Apps &amp; Integration Usage Report</h1>
  <p>Google Workspace &nbsp;&middot;&nbsp; $reportDate &nbsp;&middot;&nbsp; Runtime: ${elapsed}s</p>
</header>
<div class="wrap">

<div class="cards">
  <div class="card">  <div class="card-num">$totalUsers</div>    <div class="card-label">Total Users</div></div>
  <div class="card g"><div class="card-num">$activeUsers</div>   <div class="card-label">Active Users</div></div>
  <div class="card r"><div class="card-num">$suspUsers</div>     <div class="card-label">Suspended</div></div>
  <div class="card y"><div class="card-num">$totalApps</div>     <div class="card-label">OAuth Apps</div></div>
  <div class="card">  <div class="card-num">$totalTokens</div>   <div class="card-label">Token Records</div></div>
  <div class="card p"><div class="card-num">$totalMktApps</div>  <div class="card-label">Mkt Policies</div></div>
  <div class="card r"><div class="card-num">$totalSvcAccts</div> <div class="card-label">DwD Svc Accounts</div></div>
  <div class="card g"><div class="card-num">$totalSpaces</div>   <div class="card-label">Chat Spaces</div></div>
</div>

<!-- CONNECTED APPS DETAIL -->
<div class="sec">
  <h2>&#128241; Connected Apps Detail <span class="badge">$($appDetailSummary.Count) apps</span></h2>
  <div class="info">
    One row per OAuth-connected app.<br>
    <strong>Members</strong> = number of users who have authorised the app &nbsp;|&nbsp;
    <strong>Department</strong> = Org Unit(s) of those users &nbsp;|&nbsp;
    <strong>CreatedDate</strong> = earliest token-authorization event from the audit log &nbsp;|&nbsp;
    <strong>LastLogin</strong> = most-recent Google login among all app users &nbsp;|&nbsp;
    <strong>LastActivity</strong> = most-recent Drive activity (requires <code>-IncludeLastActivity</code>) &nbsp;|&nbsp;
    <strong>MemberList</strong> = every user who authorised this app.
  </div>
  $htmlActNote
  <div class="tbl">$htmlAppDetailTbl</div>
</div>

<!-- TOKEN DETAIL TABS -->
<div class="sec">
  <h2>&#128269; Detailed Token Data</h2>
  <div class="tabs">
    <button class="tab on" onclick="tab(event,'t-user')">Per User</button>
    <button class="tab"    onclick="tab(event,'t-raw')">All Raw Tokens</button>
    <button class="tab"    onclick="tab(event,'t-susp')">&#9888; Suspended w/ Tokens <span class="badge r">$($suspWithTokens.Count)</span></button>
  </div>
  <div id="t-user" class="pane on">
    <p style="color:#5f6368;font-size:.82rem;margin-bottom:10px">Users ranked by number of OAuth-authorised apps. Top 100 shown.</p>
    <div class="tbl">$htmlPerUserTbl</div>
  </div>
  <div id="t-raw" class="pane">
    <div class="info">Raw OAuth token records - one row per user per app. Up to 300 shown. See tokens_all_users.csv for full data.</div>
    <div class="tbl">$htmlRawTokTbl</div>
  </div>
  <div id="t-susp" class="pane">
    $htmlSuspAlert
    <div class="tbl">$htmlSuspTable</div>
  </div>
</div>

<!-- MARKETPLACE APPS -->
<div class="sec">
  <h2>&#128722; Workspace Marketplace &amp; App Details <span class="badge p">$totalMktApps rows</span></h2>
  $htmlMarketplace
</div>

<!-- SERVICE ACCOUNTS DWD -->
<div class="sec">
  <h2>&#128273; Service Accounts (Domain-Wide Delegation) <span class="badge r">$totalSvcAccts</span></h2>
  $htmlSvcAccts
</div>

<!-- ADMIN ROLES -->
<div class="sec">
  <h2>&#128110; Admin Roles &amp; Delegated Admins <span class="badge">$($adminsData.Count) assignments</span></h2>
  <div class="tbl">$htmlAdminsTbl</div>
</div>

<!-- LAST ACTIVITY -->
<div class="sec">
  <h2>&#9201; Last Activity per User</h2>
  $htmlActivity
</div>

<!-- CHAT SPACES -->
<div class="sec">
  <h2>&#128172; Google Chat Spaces</h2>
  $htmlChat
</div>

<!-- ALL USERS -->
<div class="sec">
  <h2>&#128101; All Domain Users <span class="badge g">$totalUsers</span></h2>
  <div class="tbl">$htmlUsersTbl</div>
</div>

<!-- FILES -->
<div class="sec">
  <h2>&#128193; Output Files</h2>
  <table>
    <thead><tr><th>File</th><th>Description</th><th>Records</th></tr></thead>
    <tbody>
      <tr><td>users.csv</td><td>All domain users with status and login time</td><td>$totalUsers</td></tr>
      <tr><td>tokens_all_users.csv</td><td>Raw OAuth token records (one row per user per app)</td><td>$totalTokens</td></tr>
      <tr><td>apps_aggregated.csv</td><td>Apps aggregated with user lists and scopes</td><td>$totalApps</td></tr>
      <tr><td>app_detail_summary.csv</td><td>Per-app detail: name, org unit, last login, last activity, users, scopes</td><td>$($appDetailSummary.Count)</td></tr>
      <tr><td>marketplace_apps.csv</td><td>Marketplace app policies</td><td>$totalMktApps</td></tr>
      <tr><td>service_accounts_dwd.csv</td><td>GCP service accounts with Domain-Wide Delegation</td><td>$totalSvcAccts</td></tr>
      <tr><td>admins.csv</td><td>Admin role assignments</td><td>$($adminsData.Count)</td></tr>
      $htmlActivityFileRow
      $htmlChatFileRows
      <tr><td>ConnectedApps_Report.html</td><td>This interactive HTML report</td><td>-</td></tr>
    </tbody>
  </table>
</div>

</div>
<footer>Generated by Get-ConnectedAppsReport.ps1 &nbsp;&middot;&nbsp; $reportDate</footer>
<script>
function tab(e,id){
  document.querySelectorAll('.tab').forEach(function(b){b.classList.remove('on');});
  document.querySelectorAll('.pane').forEach(function(p){p.classList.remove('on');});
  e.currentTarget.classList.add('on');
  document.getElementById(id).classList.add('on');
}
</script>
</body>
</html>
"@

$html | Out-File $reportFile -Encoding UTF8
Write-OK "HTML report written: $reportFile"

# ======================================================
#  FINAL SUMMARY
# ======================================================
Write-Banner "Done"
$elapsed2 = [math]::Round(((Get-Date) - $T0).TotalSeconds, 1)

Write-Host ""
Write-Host "  Users                    : $totalUsers  (Active: $activeUsers  |  Suspended: $suspUsers)" -ForegroundColor White
Write-Host "  OAuth Connected Apps     : $totalApps  ($totalTokens token records)" -ForegroundColor White
Write-Host "  Suspended w/ Tokens      : $($suspWithTokens.Count)$(if($suspWithTokens.Count -gt 0){'  <- REVIEW RECOMMENDED'})" -ForegroundColor $(if ($suspWithTokens.Count -gt 0) { 'Yellow' }else { 'White' })
Write-Host "  Marketplace App Policies : $totalMktApps" -ForegroundColor White
Write-Host "  DwD Service Accounts     : $totalSvcAccts$(if($totalSvcAccts -gt 0){'  <- REVIEW EACH ONE'})" -ForegroundColor $(if ($totalSvcAccts -gt 0) { 'Yellow' }else { 'White' })
Write-Host "  Admin Role Assignments   : $($adminsData.Count)" -ForegroundColor White
if ($IncludeChatSpaces) { Write-Host "  Chat Spaces              : $totalSpaces" -ForegroundColor White }
if ($IncludeLastActivity) { Write-Host "  Activity Records         : $($activityData.Count)" -ForegroundColor White }
Write-Host ""

if ($totalMktApps -eq 0) {
    Write-Warn "Marketplace section is empty. If you expected data, run:"
    Write-Host "    gam oauth update" -ForegroundColor Cyan
    Write-Host "  then enable 'Cloud Identity - Policies' API scope and re-run." -ForegroundColor Gray
}
if ($totalSvcAccts -eq 0 -and $svcMethod -eq "none") {
    Write-Warn "Service Accounts: all 4 retrieval methods failed. To fix, run:"
    Write-Host "    gam oauth update" -ForegroundColor Cyan
    Write-Host "  then enable 'Admin SDK - Other' and 'Cloud Identity - Policies' scopes and re-run." -ForegroundColor Gray
}
elseif ($svcMethod -eq "domainwidedelegation-empty") {
    Write-OK "Service Accounts: no DwD entries configured (clean state)."
}

Write-Host ""
Write-Host "  Output folder : $OutputDir" -ForegroundColor Cyan
Write-Host "  HTML report   : $reportFile" -ForegroundColor Cyan
Write-Host ""
Write-OK "Completed in ${elapsed2}s"
Write-Host ""

$open = Read-Host "  Open HTML report in browser? (y/n)"
if ($open -match '^y') { Start-Process $reportFile }
