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

    Optionally, with -SendNotificationEmail, sends each user an email at
    their existing mailbox address (unaffected by the UPN change) letting
    them know their sign-in ID changed and asking them to sign in with the
    new one going forward. The notification outcome is recorded alongside
    the UPN-change outcome in the same report row.

    Optionally, with -RunAutoReShare, chains into the AutoReShare phase of
    Invoke-M365DomainCutover.ps1 for every user whose UPN change succeeded,
    re-inviting their named-recipient OneDrive/Teams file shares so the
    links work again under the new UPN. Only runs when -Apply was used (a
    dry run has nothing to reshare) and at least one UPN change succeeded.

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

.PARAMETER SendNotificationEmail
    After a successful UPN change, email the user at their existing mailbox
    address to tell them their sign-in ID changed. Requires the connected
    Graph session to include the Mail.Send scope.

.PARAMETER SendFromMailbox
    Mailbox to send notification emails from (uses Graph
    /users/{mailbox}/sendMail). If omitted, sends as the currently
    signed-in account via /me/sendMail.

.PARAMETER SupportContact
    Text shown in the notification email for who to contact if the user has
    trouble signing in, e.g. "the IT helpdesk at helpdesk@kaara.ai".
    Defaults to "your IT administrator".

.PARAMETER RunAutoReShare
    After all UPN changes are applied, automatically run the AutoReShare
    phase of Invoke-M365DomainCutover.ps1 for the users that succeeded.
    Requires -OldDomain, -TenantId and -ClientId (an Entra ID app
    registration with the Files.ReadWrite.All and Sites.ReadWrite.All
    APPLICATION permissions, admin-consented). Only takes effect with -Apply.

.PARAMETER CutoverScriptPath
    Path to Invoke-M365DomainCutover.ps1. Defaults to that script alongside
    this one.

.PARAMETER TenantId
    Entra ID tenant ID, required for -RunAutoReShare (app-only Graph auth).

.PARAMETER ClientId
    Entra ID app registration (client) ID, required for -RunAutoReShare.

.PARAMETER ClientSecret
    Client secret for the app registration, as a SecureString. If omitted,
    Invoke-M365DomainCutover.ps1 will prompt for it interactively.

.EXAMPLE
    # Dry run
    .\Update-UserUpn.ps1 -UserListCsv .\Users.csv -NewDomain kaara.ai

.EXAMPLE
    # Apply for real, and email each user about their new sign-in ID
    .\Update-UserUpn.ps1 -UserListCsv .\Users.csv -NewDomain kaara.ai -Apply `
        -SendNotificationEmail -SendFromMailbox itadmin@kaara.ai `
        -SupportContact "the IT helpdesk at helpdesk@kaara.ai"

.EXAMPLE
    # Apply, notify users, and auto-reshare their files afterwards
    .\Update-UserUpn.ps1 -UserListCsv .\Users.csv -OldDomain kaaratech.com -NewDomain kaara.ai -Apply `
        -SendNotificationEmail -RunAutoReShare -TenantId <tenantId> -ClientId <appId>
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
    [int]$ThrottleMs = 250,

    [Parameter(Mandatory = $false)]
    [switch]$SendNotificationEmail,

    [Parameter(Mandatory = $false)]
    [string]$SendFromMailbox,

    [Parameter(Mandatory = $false)]
    [string]$SupportContact = 'your IT administrator',

    [Parameter(Mandatory = $false)]
    [switch]$RunAutoReShare,

    [Parameter(Mandatory = $false)]
    [string]$CutoverScriptPath = (Join-Path $PSScriptRoot 'Invoke-M365DomainCutover.ps1'),

    [Parameter(Mandatory = $false)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [System.Security.SecureString]$ClientSecret
)

if ($RunAutoReShare) {
    if (-not $OldDomain) { throw '-OldDomain is required when -RunAutoReShare is used.' }
    if (-not $TenantId -or -not $ClientId) { throw '-TenantId and -ClientId are required when -RunAutoReShare is used (app-only Graph auth for AutoReShare).' }
    if (-not (Test-Path $CutoverScriptPath)) { throw "CutoverScriptPath not found: $CutoverScriptPath" }
}

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
        [string]$ErrorText = '',
        [string]$NotificationStatus = '',
        [string]$NotificationDetail = ''
    )
    $script:Results.Add([PSCustomObject]@{
        Timestamp             = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        DisplayName           = $DisplayName
        OldUserPrincipalName  = $OldUpn
        NewUserPrincipalName  = $NewUpn
        Status                = $Status
        Detail                = $Detail
        Error                 = $ErrorText
        NotificationStatus    = $NotificationStatus
        NotificationDetail    = $NotificationDetail
    })
}

function Send-UpnChangeNotification {
    param(
        [string]$DisplayName,
        [string]$OldUpn,
        [string]$NewUpn,
        [string]$ToAddress
    )
    $subject = "Your Microsoft 365 sign-in ID has changed to $NewUpn"
    $body = @"
Dear $DisplayName,

As part of our organization's move to the kaara.ai domain, your Microsoft 365 sign-in ID (User Principal Name) has been updated.

Previous sign-in ID: $OldUpn
New sign-in ID: $NewUpn

Please sign in to Outlook, Teams, OneDrive, and all other Microsoft 365 apps using your new sign-in ID, $NewUpn, from now on. Your password has not changed.

You may be prompted to re-enter your credentials the next time you open these apps. If you are asked to choose or add an account, please select $NewUpn.

If you have any trouble signing in or accessing your files, please contact $SupportContact for assistance.

Thank you for your cooperation during this transition.

Best regards,
IT Team
"@
    $mailRequest = @{
        message         = @{
            subject      = $subject
            body         = @{ contentType = 'Text'; content = $body }
            toRecipients = @(@{ emailAddress = @{ address = $ToAddress } })
        }
        saveToSentItems = $true
    }
    $mailJson = $mailRequest | ConvertTo-Json -Depth 10
    $uri = if ($SendFromMailbox) { "/v1.0/users/$SendFromMailbox/sendMail" } else { '/v1.0/me/sendMail' }
    Invoke-MgGraphRequest -Method POST -Uri $uri -Body $mailJson -ContentType 'application/json' -ErrorAction Stop | Out-Null
}

function Invoke-AutoReShareChain {
    if (-not $Apply) {
        Write-Warn 'Skipping AutoReShare chain: this was a dry run, nothing was changed to reshare.'
        return
    }
    $succeeded = @($script:Results | Where-Object { $_.Status -eq 'Success' -and $_.NewUserPrincipalName })
    if ($succeeded.Count -eq 0) {
        Write-Warn 'Skipping AutoReShare chain: no UPN changes succeeded.'
        return
    }

    Write-Host ''
    Write-Host '========================================' -ForegroundColor Cyan
    Write-Host 'CHAINING: Invoke-M365DomainCutover.ps1 -Phase AutoReShare' -ForegroundColor Cyan
    Write-Host '========================================' -ForegroundColor Cyan

    $outDir = [System.IO.Path]::GetDirectoryName((Resolve-Path $OutputCsv))
    $reshareCsv = Join-Path $outDir "AutoReShareTargets_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $succeeded |
        Select-Object @{ Name = 'UserPrincipalName'; Expression = { $_.NewUserPrincipalName } } |
        Export-Csv -Path $reshareCsv -NoTypeInformation -Encoding UTF8
    Write-Detail "Re-sharing files for $($succeeded.Count) user(s) whose UPN change succeeded: $reshareCsv"

    $cutoverArgs = @{
        Phase       = 'AutoReShare'
        OldDomain   = $OldDomain
        NewDomain   = $NewDomain
        TenantId    = $TenantId
        ClientId    = $ClientId
        UserListCsv = $reshareCsv
        Apply       = $true
    }
    if ($ClientSecret) { $cutoverArgs.ClientSecret = $ClientSecret }

    try {
        & $CutoverScriptPath @cutoverArgs
    } catch {
        Write-Fail "AutoReShare chain failed: $($_.Exception.Message)"
    }
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

$canNotify = $false
try {
    $scopes = @('User.ReadWrite.All')
    if ($SendNotificationEmail) { $scopes += 'Mail.Send' }

    if (-not (Get-MgContext)) {
        Write-Step 'Connecting to Microsoft Graph'
        Connect-MgGraph -Scopes $scopes -NoWelcome
        $connectedHere = $true
    } else {
        Write-Detail "Using existing Microsoft Graph session ($((Get-MgContext).Account))"
    }

    if ($SendNotificationEmail) {
        if (((Get-MgContext).Scopes) -contains 'Mail.Send') {
            $canNotify = $true
        } else {
            Write-Warn 'Current Graph session lacks the Mail.Send scope - notification emails will be skipped. Disconnect-MgGraph and re-run to consent to Mail.Send.'
        }
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
            $user = Get-MgUser -UserId $oldUpn -Property Id, DisplayName, UserPrincipalName, OnPremisesSyncEnabled, AccountEnabled, Mail -ErrorAction Stop
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

        $notifyAddress = if ($user.Mail) { $user.Mail } else { $oldUpn }

        if (-not $Apply) {
            Write-Detail "WHATIF $oldUpn -> $newUpn"
            $notifyStatus = if ($SendNotificationEmail) { 'WhatIf' } else { '' }
            $notifyDetail = if ($SendNotificationEmail) { "Would email $notifyAddress" } else { '' }
            Add-Result $user.DisplayName $oldUpn $newUpn 'WhatIf' '' '' $notifyStatus $notifyDetail
            continue
        }

        try {
            Update-MgUser -UserId $user.Id -UserPrincipalName $newUpn -ErrorAction Stop
            Write-Ok "UPN changed: $oldUpn -> $newUpn"

            $notifyStatus = ''
            $notifyDetail = ''
            if ($SendNotificationEmail) {
                if (-not $canNotify) {
                    $notifyStatus = 'Skipped'
                    $notifyDetail = 'Graph session lacks Mail.Send scope'
                } else {
                    try {
                        Send-UpnChangeNotification -DisplayName $user.DisplayName -OldUpn $oldUpn -NewUpn $newUpn -ToAddress $notifyAddress
                        Write-Ok "Notification emailed to $notifyAddress"
                        $notifyStatus = 'Success'
                        $notifyDetail = "Emailed $notifyAddress"
                    } catch {
                        Write-Fail "Notification email failed for ${notifyAddress}: $($_.Exception.Message)"
                        $notifyStatus = 'Failed'
                        $notifyDetail = $_.Exception.Message
                    }
                }
            }
            Add-Result $user.DisplayName $oldUpn $newUpn 'Success' '' '' $notifyStatus $notifyDetail
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

if ($RunAutoReShare) { Invoke-AutoReShareChain }
