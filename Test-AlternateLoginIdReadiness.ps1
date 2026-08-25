<#
.SYNOPSIS
    Checks whether migrated users still have their old domain email address
    present as a ProxyAddress, which is required for Entra ID's "Email as an
    alternate login ID" feature to let them sign in with it.

.DESCRIPTION
    Update-UserUpn.ps1 only changes a user's UserPrincipalName - it does not
    touch mailbox ProxyAddresses (aliases). For "Email as an alternate login
    ID" to work, each migrated user needs their pre-migration email (e.g.
    name@kaaratech.com) still listed as a ProxyAddress on their account.

    This script reports Ready / Missing / SkippedSynced for each user, so you
    can confirm readiness before enabling the alternate-login-ID feature (see
    Enable-EmailAlternateLoginId.ps1).

.PARAMETER UserListCsv
    Optional CSV with a UserPrincipalName column (current, i.e. new, UPNs) to
    check. If omitted, checks every non-guest user whose UPN ends with
    -NewDomain.

.PARAMETER NewDomain
    Domain users were migrated to. Default: kaara.ai

.PARAMETER OldDomain
    Domain to look for in ProxyAddresses. Default: kaaratech.com

.PARAMETER OutputCsv
    Path for the report CSV. Defaults to .\AlternateLoginIdReadiness_<timestamp>.csv

.EXAMPLE
    # Check everyone already on kaara.ai
    .\Test-AlternateLoginIdReadiness.ps1

.EXAMPLE
    # Check a specific list
    .\Test-AlternateLoginIdReadiness.ps1 -UserListCsv .\Users.csv
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$UserListCsv,

    [Parameter(Mandatory = $false)]
    [string]$NewDomain = 'kaara.ai',

    [Parameter(Mandatory = $false)]
    [string]$OldDomain = 'kaaratech.com',

    [Parameter(Mandatory = $false)]
    [string]$OutputCsv = ".\AlternateLoginIdReadiness_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

$ErrorActionPreference = 'Stop'

function Write-Ok     { param([string]$Text) Write-Host "   OK   $Text" -ForegroundColor Green }
function Write-Warn   { param([string]$Text) Write-Host "   WARN $Text" -ForegroundColor Yellow }
function Write-Fail   { param([string]$Text) Write-Host "   FAIL $Text" -ForegroundColor Red }
function Write-Detail { param([string]$Text) Write-Host "        $Text" -ForegroundColor Gray }
function Write-Step   { param([string]$Text) Write-Host "-> $Text" -ForegroundColor Yellow }

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
    throw "Required module 'Microsoft.Graph.Users' is not installed. Run: Install-Module Microsoft.Graph.Users -Scope CurrentUser"
}

$connectedHere = $false
try {
    if (-not (Get-MgContext)) {
        Write-Step 'Connecting to Microsoft Graph'
        Connect-MgGraph -Scopes 'User.Read.All' -NoWelcome
        $connectedHere = $true
    } else {
        Write-Detail "Using existing Microsoft Graph session ($((Get-MgContext).Account))"
    }

    $props = 'Id', 'DisplayName', 'UserPrincipalName', 'Mail', 'ProxyAddresses', 'UserType', 'OnPremisesSyncEnabled'

    if ($UserListCsv) {
        if (-not (Test-Path $UserListCsv)) { throw "UserListCsv not found: $UserListCsv" }
        $rows = Import-Csv -Path $UserListCsv
        if (-not ($rows | Get-Member -Name UserPrincipalName)) {
            throw "UserListCsv must contain a UserPrincipalName column."
        }
        Write-Step "Checking $($rows.Count) user(s) from $UserListCsv"
        $users = foreach ($row in $rows) {
            $upn = [string]$row.UserPrincipalName
            if ($upn) { $upn = $upn.Trim() }
            if (-not $upn) { continue }
            try {
                Get-MgUser -UserId $upn -Property $props -ErrorAction Stop
            } catch {
                Write-Warn "Not found in directory, skipped: $upn"
            }
        }
    } else {
        Write-Step "Scanning all users on @$NewDomain"
        $users = Get-MgUser -All -Property $props -ConsistencyLevel eventual |
            Where-Object { $_.UserType -ne 'Guest' -and $_.UserPrincipalName -like "*@$NewDomain" }
    }

    $results = foreach ($u in @($users)) {
        $hasOldAlias = [bool]($u.ProxyAddresses | Where-Object { $_ -match "@$([regex]::Escape($OldDomain))$" })
        $status =
            if ($u.OnPremisesSyncEnabled) { 'SkippedSynced' }
            elseif ($hasOldAlias) { 'Ready' }
            else { 'Missing' }

        switch ($status) {
            'Ready'         { Write-Ok   "$($u.UserPrincipalName): ready ($OldDomain alias present)" }
            'Missing'       { Write-Fail "$($u.UserPrincipalName): no @$OldDomain alias found - alternate login ID will not work for this user" }
            'SkippedSynced' { Write-Detail "$($u.UserPrincipalName): synced from on-premises AD, skipped" }
        }

        [PSCustomObject]@{
            DisplayName         = $u.DisplayName
            UserPrincipalName   = $u.UserPrincipalName
            Mail                = $u.Mail
            OldDomainAliasFound = $hasOldAlias
            ProxyAddresses      = ($u.ProxyAddresses -join '; ')
            Status              = $status
        }
    }

    $results | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "SUMMARY" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    $results | Group-Object Status | ForEach-Object {
        Write-Host ("  {0,-14} {1}" -f $_.Name, $_.Count)
    }
    Write-Host "Report written to: $OutputCsv" -ForegroundColor Green
} finally {
    if ($connectedHere) { try { Disconnect-MgGraph | Out-Null } catch { } }
}
