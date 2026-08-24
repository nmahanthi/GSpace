<#
.SYNOPSIS
    Changes the UserPrincipalName (sign-in UPN) for a list of Entra ID users
    from an old domain to a new domain, and writes a CSV status report.

.DESCRIPTION
    Reads a CSV of users (column: UserPrincipalName, current value e.g.
    naresh.t@kaaratech.com), computes the new UPN by swapping the domain
    suffix for -NewDomain (or uses an explicit NewUserPrincipalName column
    if present), and calls Update-MgUser to change it.

    Defaults to a dry run (WhatIf). Pass -Apply to make real changes; you
    will be prompted to type APPLY to confirm.

    Every row's outcome (Success / Failed / Skipped / WhatIf) is written to
    an ActionLog CSV, along with the reason/error for anything that did not
    succeed.

    NOTE: this script only changes the sign-in UPN. It does not touch
    mailbox primary SMTP address, OneDrive/SharePoint URLs, or shared file
    links - those are separate concerns (see Invoke-M365DomainCutover.ps1
    SetPrimary / AutoReShare phases if you need those too). Changing the
    UPN alone breaks any existing OneDrive/Teams file links tied to the
    old UPN-derived OneDrive URL.

.PARAMETER UserListCsv
    Path to a CSV with a UserPrincipalName column (current UPNs to change).
    An optional NewUserPrincipalName column can be included to override the
    computed target UPN on a per-row basis.

.PARAMETER NewDomain
    Domain to move users to, e.g. kaara.ai. Used to compute NewUPN as
    "<localpart>@<NewDomain>" when NewUserPrincipalName is not supplied.

.PARAMETER OldDomain
    Optional. If supplied, rows whose current UPN is not on this domain are
    flagged as Skipped instead of processed, as a safety check.

.PARAMETER Apply
    Actually perform the UPN changes. Without this switch, the script only
    reports what it would do.

.PARAMETER OutputCsv
    Path for the report CSV. Defaults to .\UpnChangeReport_<timestamp>.csv

.PARAMETER ThrottleMs
    Delay between Graph calls to avoid throttling. Default 250ms.

.EXAMPLE
    # Dry run
    .\Update-UserUpn.ps1 -UserListCsv .\Users.csv -NewDomain kaara.ai

.EXAMPLE
    # Apply for real
    .\Update-UserUpn.ps1 -UserListCsv .\Users.csv -NewDomain kaara.ai -Apply
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$UserListCsv,

    [Parameter(Mandatory = $true)]
    [string]$NewDomain,

    [Parameter(Mandatory = $false)]
    [string]$OldDomain,

    [Parameter(Mandatory = $false)]
    [switch]$Apply,

    [Parameter(Mandatory = $false)]
    [string]$OutputCsv = ".\UpnChangeReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",

    [Parameter(Mandatory = $false)]
    [int]$ThrottleMs = 250
)

$ErrorActionPreference = 'Stop'

function Write-Ok     { param([string]$Text) Write-Host "   OK   $Text" -ForegroundColor Green }
function Write-Warn   { param([string]$Text) Write-Host "   WARN $Text" -ForegroundColor Yellow }
function Write-Fail   { param([string]$Text) Write-Host "   FAIL $Text" -ForegroundColor Red }
function Write-Detail { param([string]$Text) Write-Host "        $Text" -ForegroundColor Gray }
function Write-Step   { param([string]$Text) Write-Host "-> $Text" -ForegroundColor Yellow }

function Add-Result {
    param(
        [string]$DisplayName,
        [string]$OldUpn,
        [string]$NewUpn,
        [string]$Status,
        [string]$Detail = '',
        [string]$ErrorText = ''
    )
    $script:Results.Add([PSCustomObject]@{
        Timestamp             = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        DisplayName           = $DisplayName
        OldUserPrincipalName  = $OldUpn
        NewUserPrincipalName  = $NewUpn
        Status                = $Status
        Detail                = $Detail
        Error                 = $ErrorText
    })
}

if (-not (Test-Path $UserListCsv)) {
    throw "UserListCsv not found: $UserListCsv"
}
$rows = Import-Csv -Path $UserListCsv
if (-not ($rows | Get-Member -Name UserPrincipalName)) {
    throw "UserListCsv must contain a UserPrincipalName column."
}
$hasOverrideColumn = [bool]($rows | Get-Member -Name NewUserPrincipalName)

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
    throw "Required module 'Microsoft.Graph.Users' is not installed. Run: Install-Module Microsoft.Graph.Users -Scope CurrentUser"
}

$script:Results = New-Object System.Collections.Generic.List[object]
$connectedHere = $false

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "UPN DOMAIN CHANGE" -ForegroundColor Cyan
Write-Host "New domain : $NewDomain" -ForegroundColor Cyan
Write-Host "Mode       : $(if ($Apply) { 'APPLY - changes will be made' } else { 'DRY RUN (WhatIf)' })" -ForegroundColor Cyan
Write-Host "Output     : $OutputCsv" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

if ($Apply) {
    $confirm = Read-Host "Type APPLY to confirm changing UPNs for $($rows.Count) user(s)"
    if ($confirm -ne 'APPLY') { Write-Warn 'Aborted by operator.'; return }
}

try {
    if (-not (Get-MgContext)) {
        Write-Step 'Connecting to Microsoft Graph'
        Connect-MgGraph -Scopes 'User.ReadWrite.All' -NoWelcome
        $connectedHere = $true
    } else {
        Write-Detail "Using existing Microsoft Graph session ($((Get-MgContext).Account))"
    }

    $i = 0
    foreach ($row in $rows) {
        $i++
        $oldUpn = [string]$row.UserPrincipalName
        if ($oldUpn) { $oldUpn = $oldUpn.Trim() }
        if (-not $oldUpn) { continue }

        Write-Progress -Activity 'UPN change' -Status $oldUpn -PercentComplete (($i / [Math]::Max($rows.Count, 1)) * 100)

        if ($OldDomain -and ($oldUpn -notlike "*@$OldDomain")) {
            Write-Warn "Not on $OldDomain, skipped: $oldUpn"
            Add-Result '' $oldUpn '' 'Skipped' "UPN is not on $OldDomain"
            continue
        }

        try {
            $user = Get-MgUser -UserId $oldUpn -Property Id, DisplayName, UserPrincipalName, OnPremisesSyncEnabled, AccountEnabled -ErrorAction Stop
        } catch {
            Write-Fail "Not found in directory: $oldUpn"
            Add-Result '' $oldUpn '' 'Failed' 'User not found' $_.Exception.Message
            continue
        }

        if ($user.OnPremisesSyncEnabled) {
            Write-Warn "Synced from on-premises AD, skipped: $oldUpn"
            Add-Result $user.DisplayName $oldUpn '' 'Skipped' 'Synced user - change userPrincipalName in on-premises AD instead'
            continue
        }

        if ($hasOverrideColumn -and $row.NewUserPrincipalName) {
            $newUpn = $row.NewUserPrincipalName.Trim()
        } else {
            $localPart = $oldUpn.Split('@')[0]
            $newUpn = "$localPart@$NewDomain"
        }

        if ($newUpn -ieq $oldUpn) {
            Write-Detail "Already on target domain: $oldUpn"
            Add-Result $user.DisplayName $oldUpn $newUpn 'Skipped' 'UPN already matches target'
            continue
        }

        try {
            $conflict = Get-MgUser -Filter "userPrincipalName eq '$newUpn'" -ErrorAction Stop
        } catch {
            $conflict = $null
        }
        if ($conflict -and $conflict.Id -ne $user.Id) {
            Write-Fail "Target UPN already in use by another account: $newUpn"
            Add-Result $user.DisplayName $oldUpn $newUpn 'Failed' 'Target UPN already assigned to a different user'
            continue
        }

        if (-not $Apply) {
            Write-Detail "WHATIF $oldUpn -> $newUpn"
            Add-Result $user.DisplayName $oldUpn $newUpn 'WhatIf'
            continue
        }

        try {
            Update-MgUser -UserId $user.Id -UserPrincipalName $newUpn -ErrorAction Stop
            Write-Ok "UPN changed: $oldUpn -> $newUpn"
            Add-Result $user.DisplayName $oldUpn $newUpn 'Success'
        } catch {
            Write-Fail "UPN change failed for ${oldUpn}: $($_.Exception.Message)"
            Add-Result $user.DisplayName $oldUpn $newUpn 'Failed' '' $_.Exception.Message
        }
        Start-Sleep -Milliseconds $ThrottleMs
    }
    Write-Progress -Activity 'UPN change' -Completed
} finally {
    if ($script:Results.Count -gt 0) {
        $script:Results | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
    }
    if ($connectedHere) { try { Disconnect-MgGraph | Out-Null } catch { } }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "SUMMARY" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    $script:Results | Group-Object Status | ForEach-Object {
        Write-Host ("  {0,-8} {1}" -f $_.Name, $_.Count)
    }
    Write-Host "Report written to: $OutputCsv" -ForegroundColor Green
}
