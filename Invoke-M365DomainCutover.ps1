#Requires -Version 5.1

<#
.SYNOPSIS
    Phased cutover of a Microsoft 365 tenant from an old primary domain to a new
    primary domain (e.g. kaaratech.com -> kaara.ai), including an optional
    SharePoint Online tenant rename and a Teams/OneDrive re-share report.
.DESCRIPTION
    The cutover is deliberately split into phases so each one can be run,
    reviewed and rolled forward independently. Nothing is changed unless
    -Apply is supplied; the default is a dry run.

    Phases:
      Precheck   Validates both domains, detects directory-sync (hybrid), counts
                 sites, and confirms the tenant is eligible for a rename.
      Inventory  Baseline CSV of every user, UPN, mailbox address and OneDrive
                 URL before any change is made. Run this first and keep it.
      AddAlias   Adds <user>@<NewDomain> as a secondary SMTP alias only. Safe,
                 reversible, and lets mail arrive on the new domain early.
      SetPrimary Promotes <user>@<NewDomain> to primary SMTP and changes the UPN,
                 retaining the old address as an alias. Cloud-only users only.
      Groups     Repoints M365 groups, distribution groups and shared mailboxes.
      Default    Sets NewDomain as the tenant default domain for new objects.
      Validate   Post-change report: UPN, primary SMTP and OneDrive URL per user.
      ReShare    Reports OneDrive-hosted shared files (including Teams chat
                 files) whose links break after a UPN change and must be
                 re-shared. Report only - there is no supported bulk fix.
      AutoReShare Re-shares OneDrive items that were shared with specific
                 people (named-recipient permissions, including Teams chat
                 file shares) by calling Microsoft Graph's driveItem invite
                 action for the same recipients. This regenerates a working
                 link at the item's current location and emails it to those
                 recipients automatically - no action needed from the file
                 owner. "Anyone in the org" / "Anyone with the link" shares
                 cannot be resolved to recipients by Graph and are reported
                 as NeedsManualReshare instead of being skipped silently.
                 REQUIRES APP-ONLY AUTH (-TenantId, -ClientId, -ClientSecret).
                 Reading another user's drive (/users/{id}/drive) is blocked by
                 Microsoft Graph under delegated auth, even for a Global
                 Administrator signed in interactively - this returns 403
                 Forbidden. It is not a permissions/consent problem; app-only
                 (client credentials) auth is the only supported path.
      Rename     Optional SharePoint tenant rename to <NewSpoName>.sharepoint.com.

    IMPORTANT - what cannot be done:
      SharePoint, OneDrive and Teams file URLs are always <name>.sharepoint.com.
      Vanity domains are not supported, so those URLs can never become kaara.ai.
      The Rename phase can only change the host prefix.

    IMPORTANT - hybrid tenants:
      If users are synced from on-premises AD via Entra Connect, UPNs must be
      changed in on-premises AD (and kaara.ai added as a UPN suffix in AD Domains
      and Trusts). The SetPrimary phase detects synced users and skips them.
.PARAMETER Phase
    One or more phases to run, in the order supplied. Use All for the safe
    read-only set (Precheck, Inventory, Validate, ReShare).
.PARAMETER Apply
    Perform changes. Omit for a dry run that only reports intended actions.
.PARAMETER OldDomain
    Current primary domain, e.g. kaaratech.com.
.PARAMETER NewDomain
    New primary domain, already verified in Entra ID, e.g. kaara.ai.
.PARAMETER AdminUrl
    SharePoint Online admin URL, e.g. https://kaaratech-admin.sharepoint.com.
.PARAMETER ClientId
    Entra ID application (client) ID. Used for PnP interactive sign-in (Since
    Sept 9 2024 PnP PowerShell requires your own app registration. Register once:
      Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP.PowerShell" -Tenant <tenant>.onmicrosoft.com
    ), and, together with -TenantId and -ClientSecret, for app-only
    Graph auth required by the AutoReShare phase. Falls back to
    ENTRAID_APP_ID / ENTRAID_CLIENT_ID / AZURE_CLIENT_ID.
.PARAMETER TenantId
    Entra ID tenant ID or verified domain (e.g. kaaratech.onmicrosoft.com).
    Required for AutoReShare's app-only Graph connection.
.PARAMETER ClientSecret
    Client secret (as a SecureString) of an Entra ID app registration that has
    been granted the Files.ReadWrite.All and Sites.ReadWrite.All Graph
    APPLICATION permissions with admin consent. Required for AutoReShare -
    delegated auth (even as a Global Administrator) cannot read another
    user's OneDrive; Graph returns 403 Forbidden for /users/{id}/drive under
    delegated auth by design. Pass as a SecureString, e.g.:
      -ClientSecret (ConvertTo-SecureString "<secret>" -AsPlainText -Force)
    or omit it and let the script prompt securely.
.PARAMETER UserListCsv
    Optional CSV with a UserPrincipalName column, used to run AddAlias and
    SetPrimary in controlled waves. Omit to target every user on OldDomain.
.PARAMETER NewSpoName
    Host prefix for the Rename phase, e.g. kaara for kaara.sharepoint.com. The
    matching <NewSpoName>.onmicrosoft.com domain must already be verified.
.PARAMETER RenameScheduledUtc
    UTC start time for the rename job. Defaults to 30 minutes from now.
.PARAMETER MaxItemsPerDrive
    Cap on files inspected per OneDrive during the ReShare/AutoReShare phases.
.PARAMETER ReshareMessage
    Message text included in the re-share email sent to recipients during
    AutoReShare. Defaults to a note explaining the domain change.
.PARAMETER ThrottleMs
    Delay between per-user write operations, to stay under service throttling.
.PARAMETER OutputFolder
    Destination for all CSV reports and the action log.
.EXAMPLE
    # 1. Read-only assessment
    .\Invoke-M365DomainCutover.ps1 -Phase Precheck,Inventory -OldDomain kaaratech.com -NewDomain kaara.ai -AdminUrl https://kaaratech-admin.sharepoint.com -ClientId <appId>
.EXAMPLE
    # 2. Dry run the alias phase, then apply it
    .\Invoke-M365DomainCutover.ps1 -Phase AddAlias -OldDomain kaaratech.com -NewDomain kaara.ai
    .\Invoke-M365DomainCutover.ps1 -Phase AddAlias -OldDomain kaaratech.com -NewDomain kaara.ai -Apply
.EXAMPLE
    # 3. Cut over a pilot wave, then validate
    .\Invoke-M365DomainCutover.ps1 -Phase SetPrimary -UserListCsv .\wave1.csv -OldDomain kaaratech.com -NewDomain kaara.ai -Apply
    .\Invoke-M365DomainCutover.ps1 -Phase Validate,ReShare -OldDomain kaaratech.com -NewDomain kaara.ai -AdminUrl https://kaaratech-admin.sharepoint.com -ClientId <appId>
.EXAMPLE
    # 3b. Automatically re-share named-recipient links that broke (app-only auth required), dry run then apply
    .\Invoke-M365DomainCutover.ps1 -Phase AutoReShare -UserListCsv .\wave1.csv -OldDomain kaaratech.com -NewDomain kaara.ai -TenantId <tenantId> -ClientId <appOnlyAppId>
    .\Invoke-M365DomainCutover.ps1 -Phase AutoReShare -UserListCsv .\wave1.csv -OldDomain kaaratech.com -NewDomain kaara.ai -TenantId <tenantId> -ClientId <appOnlyAppId> -Apply
    # (prompts securely for the client secret; or pass -ClientSecret (ConvertTo-SecureString "<secret>" -AsPlainText -Force))
.EXAMPLE
    # 4. Optional SharePoint rename, run on its own change window
    .\Invoke-M365DomainCutover.ps1 -Phase Rename -NewSpoName kaara -OldDomain kaaratech.com -NewDomain kaara.ai -AdminUrl https://kaaratech-admin.sharepoint.com -Apply
.NOTES
    Modules:  Microsoft.Graph, ExchangeOnlineManagement, Microsoft.Online.SharePoint.PowerShell, PnP.PowerShell
              (AutoReShare uses Microsoft Graph REST calls directly via Invoke-MgGraphRequest;
              no extra module beyond Microsoft.Graph is required.)
    Roles:    Global Administrator (or User Admin + Exchange Admin + SharePoint Admin)
              for every phase except AutoReShare, which authenticates as an
              Entra ID app (see below) rather than an admin's own account.

    AutoReShare app-only setup (one-time):
      1. App registrations > New registration (e.g. "M365DomainCutover-AutoReShare").
      2. API permissions > Microsoft Graph > Application permissions:
           User.Read.All, Files.ReadWrite.All, Sites.ReadWrite.All
         Grant admin consent.
      3. Certificates & secrets > Client secrets > New client secret. Copy the
         Value immediately - it is shown only once. Store it in a secret
         manager, not in plain text.
      4. Note the Application (client) ID and the tenant ID.
      5. Pass them as -ClientId, -TenantId to AutoReShare, and either supply
         -ClientSecret as a SecureString or let the script prompt for it.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Precheck', 'Inventory', 'AddAlias', 'SetPrimary', 'Groups', 'Default', 'Validate', 'ReShare', 'AutoReShare', 'Rename', 'All')]
    [string[]]$Phase,

    [Parameter(Mandatory = $false)]
    [switch]$Apply,

    [Parameter(Mandatory = $true)]
    [string]$OldDomain,

    [Parameter(Mandatory = $true)]
    [string]$NewDomain,

    [Parameter(Mandatory = $false)]
    [string]$AdminUrl,

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [System.Security.SecureString]$ClientSecret,

    [Parameter(Mandatory = $false)]
    [string]$UserListCsv,

    [Parameter(Mandatory = $false)]
    [string]$NewSpoName,

    [Parameter(Mandatory = $false)]
    [datetime]$RenameScheduledUtc = (Get-Date).ToUniversalTime().AddMinutes(30),

    [Parameter(Mandatory = $false)]
    [int]$MaxItemsPerDrive = 500,

    [Parameter(Mandatory = $false)]
    [string]$ReshareMessage = "This item is being re-shared with you because our organisation's email domain changed. The previous link no longer works; please use this one going forward.",

    [Parameter(Mandatory = $false)]
    [int]$ThrottleMs = 250,

    [Parameter(Mandatory = $false)]
    [string]$OutputFolder = ".\DomainCutover_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
)

$ErrorActionPreference = 'Stop'
$script:ActionLog = New-Object System.Collections.Generic.List[object]
$script:Connected = @{ Graph = $false; GraphAppOnly = $false; Exchange = $false; SPO = $false; PnP = $false }

# =============================================================================
# Helpers
# =============================================================================

function Write-Banner {
    param([string]$Text)
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host $Text -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
}

function Write-Step   { param([string]$Text) Write-Host "-> $Text" -ForegroundColor Yellow }
function Write-Ok     { param([string]$Text) Write-Host "   OK   $Text" -ForegroundColor Green }
function Write-Warn   { param([string]$Text) Write-Host "   WARN $Text" -ForegroundColor Yellow }
function Write-Fail   { param([string]$Text) Write-Host "   FAIL $Text" -ForegroundColor Red }
function Write-Detail { param([string]$Text) Write-Host "        $Text" -ForegroundColor Gray }

function Add-Action {
    param(
        [string]$PhaseName,
        [string]$Target,
        [string]$Action,
        [string]$Detail,
        [string]$Status,
        [string]$ErrorText = ''
    )
    $script:ActionLog.Add([PSCustomObject]@{
        Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Phase     = $PhaseName
        Target    = $Target
        Action    = $Action
        Detail    = $Detail
        Status    = $Status
        Error     = $ErrorText
    })
}

function Export-Report {
    param([object[]]$Data, [string]$Name)
    if (-not $Data -or $Data.Count -eq 0) {
        Write-Detail "No rows to write for $Name"
        return
    }
    $path = Join-Path $OutputFolder $Name
    $Data | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
    Write-Ok "$Name ($($Data.Count) rows)"
}

function Assert-Module {
    param([string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        throw "Required module '$Name' is not installed. Run: Install-Module $Name -Scope CurrentUser"
    }
}

function Connect-GraphIfNeeded {
    if ($script:Connected.Graph -and -not $script:Connected.GraphAppOnly) { return }
    Assert-Module 'Microsoft.Graph'
    if ($script:Connected.GraphAppOnly) {
        Disconnect-MgGraph | Out-Null
        $script:Connected.GraphAppOnly = $false
    }
    Write-Step 'Connecting to Microsoft Graph'
    $scopes = @('User.ReadWrite.All', 'Domain.ReadWrite.All', 'Group.Read.All', 'Organization.Read.All')
    Connect-MgGraph -Scopes $scopes -NoWelcome
    $script:Connected.Graph = $true
    Write-Ok 'Microsoft Graph (delegated)'
}

function Connect-GraphAppOnlyIfNeeded {
    # AutoReShare must read other users' OneDrive content (/users/{id}/drive).
    # Microsoft Graph rejects this under delegated auth with 403 Forbidden,
    # even for a signed-in Global Administrator - it is only reachable with
    # app-only (client credentials) auth. See:
    # https://learn.microsoft.com/en-us/answers/questions/1103195
    if ($script:Connected.GraphAppOnly) { return }
    Assert-Module 'Microsoft.Graph'
    if (-not $TenantId -or -not $ClientId) {
        throw "AutoReShare requires app-only Graph auth: pass -TenantId, -ClientId and -ClientSecret for an Entra ID app registration granted the Files.ReadWrite.All and Sites.ReadWrite.All APPLICATION permissions with admin consent. Delegated sign-in (even as Global Administrator) cannot read another user's OneDrive and will return 403 Forbidden."
    }
    $secret = $ClientSecret
    if (-not $secret) {
        $secret = Read-Host -Prompt "Client secret for app $ClientId" -AsSecureString
    }
    $credential = New-Object System.Management.Automation.PSCredential ($ClientId, $secret)
    if ($script:Connected.Graph) {
        Disconnect-MgGraph | Out-Null
        $script:Connected.Graph = $false
    }
    Write-Step 'Connecting to Microsoft Graph (app-only)'
    Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $credential -NoWelcome
    $script:Connected.GraphAppOnly = $true
    $script:Connected.Graph = $true
    Write-Ok 'Microsoft Graph (app-only)'
}

function Connect-ExchangeIfNeeded {
    if ($script:Connected.Exchange) { return }
    Assert-Module 'ExchangeOnlineManagement'
    Write-Step 'Connecting to Exchange Online'
    Connect-ExchangeOnline -ShowBanner:$false
    $script:Connected.Exchange = $true
    Write-Ok 'Exchange Online'
}

function Connect-SpoIfNeeded {
    if ($script:Connected.SPO) { return }
    if (-not $AdminUrl) { throw '-AdminUrl is required for this phase.' }
    Assert-Module 'Microsoft.Online.SharePoint.PowerShell'
    Write-Step 'Connecting to SharePoint Online admin'
    Connect-SPOService -Url $AdminUrl
    $script:Connected.SPO = $true
    Write-Ok 'SharePoint Online admin'
}

function Connect-PnPIfNeeded {
    if ($script:Connected.PnP) { return }
    if (-not $AdminUrl) { throw '-AdminUrl is required for this phase.' }
    Assert-Module 'PnP.PowerShell'
    Write-Step 'Connecting to SharePoint Online via PnP'
    $appId = if ($ClientId) { $ClientId }
             elseif ($env:ENTRAID_APP_ID) { $env:ENTRAID_APP_ID }
             elseif ($env:ENTRAID_CLIENT_ID) { $env:ENTRAID_CLIENT_ID }
             elseif ($env:AZURE_CLIENT_ID) { $env:AZURE_CLIENT_ID }
             else { $null }
    if (-not $appId) {
        throw "No ClientId supplied and no ENTRAID_APP_ID/ENTRAID_CLIENT_ID/AZURE_CLIENT_ID environment variable set. Since Sept 9 2024 PnP PowerShell requires your own Entra ID App Registration for -Interactive login. Run Register-PnPEntraIDAppForInteractiveLogin once, then pass -ClientId <appId>."
    }
    Connect-PnPOnline -Url $AdminUrl -Interactive -ClientId $appId
    $script:Connected.PnP = $true
    Write-Ok 'SharePoint Online (PnP)'
}

function Get-TargetUsers {
    # Only needs Get-MgUser, which works under either delegated or app-only
    # auth. Do not force a delegated re-connect if a session (of either kind)
    # is already established - that would evict an app-only session (e.g.
    # from AutoReShare) and pop an unwanted interactive sign-in.
    if (-not $script:Connected.Graph) { Connect-GraphIfNeeded }
    $props = 'Id', 'DisplayName', 'UserPrincipalName', 'Mail', 'ProxyAddresses', 'AccountEnabled', 'OnPremisesSyncEnabled', 'UserType'
    if ($UserListCsv) {
        if (-not (Test-Path $UserListCsv)) { throw "UserListCsv not found: $UserListCsv" }
        $rows = Import-Csv -Path $UserListCsv
        if (-not ($rows | Get-Member -Name UserPrincipalName)) {
            throw "UserListCsv must contain a UserPrincipalName column."
        }
        $users = foreach ($row in $rows) {
            $upn = $row.UserPrincipalName.Trim()
            if (-not $upn) { continue }
            try {
                Get-MgUser -UserId $upn -Property $props
            } catch {
                Write-Warn "Not found in directory, skipped: $upn"
                Add-Action 'Resolve' $upn 'Lookup' 'User not found' 'Skipped' $_.Exception.Message
            }
        }
        return @($users)
    }
    $all = Get-MgUser -All -Property $props -ConsistencyLevel eventual
    return @($all | Where-Object {
        $_.UserType -ne 'Guest' -and $_.UserPrincipalName -like "*@$OldDomain"
    })
}

function Get-NewAddress {
    param([string]$Upn)
    $localPart = $Upn.Split('@')[0]
    return "$localPart@$NewDomain"
}

function Get-OneDrivePathFromUpn {
    param([string]$Upn)
    return $Upn.Replace('@', '_').Replace('.', '_')
}

# =============================================================================
# Phase: Precheck
# =============================================================================

function Invoke-PrecheckPhase {
    Write-Banner 'PHASE: PRECHECK'
    Connect-GraphIfNeeded
    $findings = New-Object System.Collections.Generic.List[object]

    function Add-Finding {
        param([string]$Check, [string]$Result, [string]$Value, [string]$Guidance = '')
        $findings.Add([PSCustomObject]@{ Check = $Check; Result = $Result; Value = $Value; Guidance = $Guidance })
        switch ($Result) {
            'Pass' { Write-Ok "$Check : $Value" }
            'Warn' { Write-Warn "$Check : $Value"; if ($Guidance) { Write-Detail $Guidance } }
            'Fail' { Write-Fail "$Check : $Value"; if ($Guidance) { Write-Detail $Guidance } }
        }
    }

    Write-Step 'Checking verified domains'
    $domains = Get-MgDomain -All

    $old = $domains | Where-Object { $_.Id -eq $OldDomain }
    if (-not $old) {
        Add-Finding 'Old domain present' 'Fail' "$OldDomain not found in tenant" 'Verify the -OldDomain value.'
    } else {
        Add-Finding 'Old domain present' 'Pass' "$OldDomain (default=$($old.IsDefault))"
    }

    $new = $domains | Where-Object { $_.Id -eq $NewDomain }
    if (-not $new) {
        Add-Finding 'New domain present' 'Fail' "$NewDomain not found in tenant" 'Add and verify the domain before running any write phase.'
    } elseif (-not $new.IsVerified) {
        Add-Finding 'New domain verified' 'Fail' "$NewDomain is not verified" 'Complete DNS TXT verification first.'
    } else {
        Add-Finding 'New domain verified' 'Pass' "$NewDomain (default=$($new.IsDefault))"
        $services = @($new.SupportedServices)
        if ($services -notcontains 'Email') {
            Add-Finding 'New domain email-enabled' 'Warn' "SupportedServices: $($services -join ', ')" 'Add MX, SPF, autodiscover and DKIM records for the new domain before promoting primary SMTP.'
        } else {
            Add-Finding 'New domain email-enabled' 'Pass' ($services -join ', ')
        }
    }

    Write-Step 'Checking directory synchronisation'
    $org = Get-MgOrganization
    if ($org.OnPremisesSyncEnabled) {
        Add-Finding 'Directory sync' 'Warn' 'Entra Connect sync is enabled' "UPNs of synced users must be changed in on-premises AD. Add $NewDomain as a UPN suffix in AD Domains and Trusts. The SetPrimary phase will skip synced users."
    } else {
        Add-Finding 'Directory sync' 'Pass' 'Cloud-only tenant'
    }

    $targets = Get-TargetUsers
    $syncedCount = @($targets | Where-Object { $_.OnPremisesSyncEnabled }).Count
    Add-Finding 'Users in scope' 'Pass' "$($targets.Count) total, $syncedCount synced from on-premises"

    if ($AdminUrl) {
        Write-Step 'Checking SharePoint scale and rename eligibility'
        try {
            Connect-SpoIfNeeded
            $siteCount = @(Get-SPOSite -Limit All).Count
            $odbCount  = @(Get-SPOSite -IncludePersonalSite $true -Limit All -Filter "Url -like '-my.sharepoint.com/personal/'").Count
            $total = $siteCount + $odbCount
            if ($total -gt 10000) {
                Add-Finding 'SharePoint scale' 'Warn' "$total sites and OneDrives" 'Advanced Tenant Rename applies above 10,000 sites and requires SharePoint Advanced Management licensing.'
            } else {
                Add-Finding 'SharePoint scale' 'Pass' "$total sites and OneDrives (standard rename path)"
            }
            if ($NewSpoName) {
                $onms = $domains | Where-Object { $_.Id -eq "$NewSpoName.onmicrosoft.com" }
                if (-not $onms) {
                    Add-Finding 'Rename target domain' 'Fail' "$NewSpoName.onmicrosoft.com not found" 'Add and verify this onmicrosoft.com domain before running the Rename phase.'
                } else {
                    Add-Finding 'Rename target domain' 'Pass' "$NewSpoName.onmicrosoft.com verified"
                }
            }
        } catch {
            Add-Finding 'SharePoint scale' 'Warn' 'Could not query SharePoint' $_.Exception.Message
        }
    } else {
        Add-Finding 'SharePoint checks' 'Warn' 'Skipped' 'Supply -AdminUrl to include SharePoint checks.'
    }

    Add-Finding 'File URL domain' 'Warn' 'SharePoint, OneDrive and Teams file URLs stay on *.sharepoint.com' "Vanity domains are unsupported, so file URLs cannot become $NewDomain. The Rename phase can only change the host prefix."

    Export-Report $findings 'Precheck-Findings.csv'
}

# =============================================================================
# Phase: Inventory
# =============================================================================

function Invoke-InventoryPhase {
    Write-Banner 'PHASE: INVENTORY (baseline before any change)'
    Connect-GraphIfNeeded
    Connect-ExchangeIfNeeded

    $targets = Get-TargetUsers
    Write-Step "Building baseline for $($targets.Count) users"

    $rows = New-Object System.Collections.Generic.List[object]
    $i = 0
    foreach ($u in $targets) {
        $i++
        Write-Progress -Activity 'Inventory' -Status $u.UserPrincipalName -PercentComplete (($i / [Math]::Max($targets.Count, 1)) * 100)
        $mbx = $null
        try { $mbx = Get-Mailbox -Identity $u.UserPrincipalName -ErrorAction Stop } catch { }
        $rows.Add([PSCustomObject]@{
            DisplayName        = $u.DisplayName
            UserPrincipalName  = $u.UserPrincipalName
            ProposedUpn        = Get-NewAddress $u.UserPrincipalName
            Mail               = $u.Mail
            PrimarySmtpAddress = $mbx.PrimarySmtpAddress
            AliasAddresses     = if ($mbx) { (($mbx.EmailAddresses | Where-Object { $_ -clike 'smtp:*' }) -join ';') } else { '' }
            RecipientType      = $mbx.RecipientTypeDetails
            AccountEnabled     = $u.AccountEnabled
            OnPremisesSynced   = [bool]$u.OnPremisesSyncEnabled
            CurrentOneDrivePath = "/personal/$(Get-OneDrivePathFromUpn $u.UserPrincipalName)"
            ExpectedOneDrivePath = "/personal/$(Get-OneDrivePathFromUpn (Get-NewAddress $u.UserPrincipalName))"
        })
    }
    Write-Progress -Activity 'Inventory' -Completed
    Export-Report $rows 'Inventory-Users.csv'

    Write-Step 'Inventorying groups and shared mailboxes on the old domain'
    $groupRows = New-Object System.Collections.Generic.List[object]
    foreach ($g in Get-UnifiedGroup -ResultSize Unlimited) {
        if ($g.PrimarySmtpAddress -like "*@$OldDomain") {
            $groupRows.Add([PSCustomObject]@{ Type = 'M365Group'; Name = $g.DisplayName; Identity = $g.Identity; PrimarySmtpAddress = $g.PrimarySmtpAddress; ProposedAddress = Get-NewAddress $g.PrimarySmtpAddress })
        }
    }
    foreach ($d in Get-DistributionGroup -ResultSize Unlimited) {
        if ($d.PrimarySmtpAddress -like "*@$OldDomain") {
            $groupRows.Add([PSCustomObject]@{ Type = 'DistributionGroup'; Name = $d.DisplayName; Identity = $d.Identity; PrimarySmtpAddress = $d.PrimarySmtpAddress; ProposedAddress = Get-NewAddress $d.PrimarySmtpAddress })
        }
    }
    foreach ($s in Get-Mailbox -RecipientTypeDetails SharedMailbox, RoomMailbox, EquipmentMailbox -ResultSize Unlimited) {
        if ($s.PrimarySmtpAddress -like "*@$OldDomain") {
            $groupRows.Add([PSCustomObject]@{ Type = [string]$s.RecipientTypeDetails; Name = $s.DisplayName; Identity = $s.Identity; PrimarySmtpAddress = $s.PrimarySmtpAddress; ProposedAddress = Get-NewAddress $s.PrimarySmtpAddress })
        }
    }
    Export-Report $groupRows 'Inventory-GroupsAndSharedMailboxes.csv'
}

# =============================================================================
# Phase: AddAlias
# =============================================================================

function Invoke-AddAliasPhase {
    Write-Banner 'PHASE: ADDALIAS (add new-domain address as secondary SMTP)'
    Connect-ExchangeIfNeeded
    $targets = Get-TargetUsers
    Write-Step "$($targets.Count) users in scope. Apply=$($Apply.IsPresent)"

    foreach ($u in $targets) {
        $upn = $u.UserPrincipalName
        $newAddr = Get-NewAddress $upn
        try {
            $mbx = Get-Mailbox -Identity $upn -ErrorAction Stop
        } catch {
            Write-Warn "No mailbox, skipped: $upn"
            Add-Action 'AddAlias' $upn 'AddAlias' $newAddr 'Skipped' 'No mailbox'
            continue
        }
        if ($mbx.EmailAddresses -contains "smtp:$newAddr" -or $mbx.EmailAddresses -contains "SMTP:$newAddr") {
            Write-Detail "Already present: $newAddr"
            Add-Action 'AddAlias' $upn 'AddAlias' $newAddr 'AlreadyPresent'
            continue
        }
        if ($mbx.EmailAddressPolicyEnabled) {
            Write-Warn "Email address policy is enabled for $upn; the alias may be overwritten by policy."
        }
        if (-not $Apply) {
            Write-Detail "WHATIF add alias $newAddr to $upn"
            Add-Action 'AddAlias' $upn 'AddAlias' $newAddr 'WhatIf'
            continue
        }
        try {
            Set-Mailbox -Identity $upn -EmailAddresses @{ add = $newAddr } -ErrorAction Stop
            Write-Ok "Alias added: $newAddr"
            Add-Action 'AddAlias' $upn 'AddAlias' $newAddr 'Success'
        } catch {
            Write-Fail "Alias failed for ${upn}: $($_.Exception.Message)"
            Add-Action 'AddAlias' $upn 'AddAlias' $newAddr 'Failed' $_.Exception.Message
        }
        Start-Sleep -Milliseconds $ThrottleMs
    }
}

# =============================================================================
# Phase: SetPrimary
# =============================================================================

function Invoke-SetPrimaryPhase {
    Write-Banner 'PHASE: SETPRIMARY (promote primary SMTP, then change UPN)'
    Connect-GraphIfNeeded
    Connect-ExchangeIfNeeded
    $targets = Get-TargetUsers
    Write-Step "$($targets.Count) users in scope. Apply=$($Apply.IsPresent)"
    Write-Warn 'This changes each OneDrive URL. Previously shared OneDrive and Teams chat file links will 404 and must be re-shared.'

    foreach ($u in $targets) {
        $upn = $u.UserPrincipalName
        $newAddr = Get-NewAddress $upn

        if ($u.OnPremisesSyncEnabled) {
            Write-Warn "Synced from on-premises AD, skipped: $upn"
            Add-Action 'SetPrimary' $upn 'ChangeUpn' $newAddr 'Skipped' 'Synced user - change userPrincipalName in on-premises AD'
            continue
        }

        # Primary SMTP: keep every existing address, promote the new one to SMTP:
        try {
            $mbx = Get-Mailbox -Identity $upn -ErrorAction Stop
        } catch {
            $mbx = $null
        }
        if ($mbx) {
            if ($mbx.PrimarySmtpAddress -eq $newAddr) {
                Write-Detail "Primary SMTP already $newAddr"
                Add-Action 'SetPrimary' $upn 'SetPrimarySmtp' $newAddr 'AlreadyPresent'
            } elseif (-not $Apply) {
                Write-Detail "WHATIF primary SMTP $($mbx.PrimarySmtpAddress) -> $newAddr (old kept as alias)"
                Add-Action 'SetPrimary' $upn 'SetPrimarySmtp' $newAddr 'WhatIf'
            } else {
                try {
                    if ($mbx.EmailAddressPolicyEnabled) {
                        Set-Mailbox -Identity $upn -EmailAddressPolicyEnabled $false -ErrorAction Stop
                    }
                    $addresses = New-Object System.Collections.Generic.List[string]
                    $addresses.Add("SMTP:$newAddr")
                    foreach ($addr in $mbx.EmailAddresses) {
                        $a = [string]$addr
                        if ($a -ieq "SMTP:$newAddr") { continue }
                        if ($a -clike 'SMTP:*') { $a = 'smtp:' + $a.Substring(5) }
                        $addresses.Add($a)
                    }
                    Set-Mailbox -Identity $upn -EmailAddresses $addresses -ErrorAction Stop
                    Write-Ok "Primary SMTP set: $newAddr"
                    Add-Action 'SetPrimary' $upn 'SetPrimarySmtp' $newAddr 'Success'
                } catch {
                    Write-Fail "Primary SMTP failed for ${upn}: $($_.Exception.Message)"
                    Add-Action 'SetPrimary' $upn 'SetPrimarySmtp' $newAddr 'Failed' $_.Exception.Message
                    continue
                }
            }
        } else {
            Write-Warn "No mailbox for $upn; UPN will still be changed."
            Add-Action 'SetPrimary' $upn 'SetPrimarySmtp' $newAddr 'Skipped' 'No mailbox'
        }

        # UPN
        if (-not $Apply) {
            Write-Detail "WHATIF UPN $upn -> $newAddr"
            Add-Action 'SetPrimary' $upn 'ChangeUpn' $newAddr 'WhatIf'
            continue
        }
        try {
            Update-MgUser -UserId $u.Id -UserPrincipalName $newAddr -ErrorAction Stop
            Write-Ok "UPN changed: $upn -> $newAddr"
            Add-Action 'SetPrimary' $upn 'ChangeUpn' $newAddr 'Success'
        } catch {
            Write-Fail "UPN change failed for ${upn}: $($_.Exception.Message)"
            Add-Action 'SetPrimary' $upn 'ChangeUpn' $newAddr 'Failed' $_.Exception.Message
        }
        Start-Sleep -Milliseconds $ThrottleMs
    }
}

# =============================================================================
# Phase: Groups
# =============================================================================

function Invoke-GroupsPhase {
    Write-Banner 'PHASE: GROUPS (M365 groups, distribution groups, shared mailboxes)'
    Connect-ExchangeIfNeeded
    Write-Step "Apply=$($Apply.IsPresent)"

    $work = New-Object System.Collections.Generic.List[object]
    foreach ($g in Get-UnifiedGroup -ResultSize Unlimited) {
        if ($g.PrimarySmtpAddress -like "*@$OldDomain") { $work.Add([PSCustomObject]@{ Kind = 'UnifiedGroup'; Identity = $g.Identity; Address = [string]$g.PrimarySmtpAddress }) }
    }
    foreach ($d in Get-DistributionGroup -ResultSize Unlimited) {
        if ($d.PrimarySmtpAddress -like "*@$OldDomain") { $work.Add([PSCustomObject]@{ Kind = 'DistributionGroup'; Identity = $d.Identity; Address = [string]$d.PrimarySmtpAddress }) }
    }
    foreach ($s in Get-Mailbox -RecipientTypeDetails SharedMailbox, RoomMailbox, EquipmentMailbox -ResultSize Unlimited) {
        if ($s.PrimarySmtpAddress -like "*@$OldDomain") { $work.Add([PSCustomObject]@{ Kind = 'Mailbox'; Identity = $s.Identity; Address = [string]$s.PrimarySmtpAddress }) }
    }
    Write-Step "$($work.Count) objects in scope"

    foreach ($item in $work) {
        $newAddr = Get-NewAddress $item.Address
        if (-not $Apply) {
            Write-Detail "WHATIF $($item.Kind) $($item.Address) -> $newAddr"
            Add-Action 'Groups' $item.Address $item.Kind $newAddr 'WhatIf'
            continue
        }
        try {
            switch ($item.Kind) {
                'UnifiedGroup' {
                    Set-UnifiedGroup -Identity $item.Identity -EmailAddresses @{ add = $newAddr } -ErrorAction Stop
                    Set-UnifiedGroup -Identity $item.Identity -PrimarySmtpAddress $newAddr -ErrorAction Stop
                }
                'DistributionGroup' {
                    Set-DistributionGroup -Identity $item.Identity -EmailAddresses @{ add = $newAddr } -ErrorAction Stop
                    Set-DistributionGroup -Identity $item.Identity -PrimarySmtpAddress $newAddr -ErrorAction Stop
                }
                'Mailbox' {
                    Set-Mailbox -Identity $item.Identity -EmailAddresses @{ add = $newAddr } -ErrorAction Stop
                    Set-Mailbox -Identity $item.Identity -PrimarySmtpAddress $newAddr -ErrorAction Stop
                }
            }
            Write-Ok "$($item.Kind) repointed: $($item.Address) -> $newAddr"
            Add-Action 'Groups' $item.Address $item.Kind $newAddr 'Success'
        } catch {
            Write-Fail "$($item.Kind) failed for $($item.Address): $($_.Exception.Message)"
            Add-Action 'Groups' $item.Address $item.Kind $newAddr 'Failed' $_.Exception.Message
        }
        Start-Sleep -Milliseconds $ThrottleMs
    }
}

# =============================================================================
# Phase: Default
# =============================================================================

function Invoke-DefaultPhase {
    Write-Banner 'PHASE: DEFAULT (set tenant default domain)'
    Connect-GraphIfNeeded
    $new = Get-MgDomain -DomainId $NewDomain
    if ($new.IsDefault) {
        Write-Ok "$NewDomain is already the default domain"
        Add-Action 'Default' $NewDomain 'SetDefault' 'Already default' 'AlreadyPresent'
        return
    }
    if (-not $Apply) {
        Write-Detail "WHATIF set default domain -> $NewDomain"
        Add-Action 'Default' $NewDomain 'SetDefault' 'Set as tenant default' 'WhatIf'
        return
    }
    try {
        Update-MgDomain -DomainId $NewDomain -IsDefault -ErrorAction Stop
        Write-Ok "Default domain is now $NewDomain"
        Add-Action 'Default' $NewDomain 'SetDefault' 'Set as tenant default' 'Success'
    } catch {
        Write-Fail "Could not set default domain: $($_.Exception.Message)"
        Add-Action 'Default' $NewDomain 'SetDefault' 'Set as tenant default' 'Failed' $_.Exception.Message
    }
}

# =============================================================================
# Phase: Validate
# =============================================================================

function Invoke-ValidatePhase {
    Write-Banner 'PHASE: VALIDATE (post-change state)'
    Connect-GraphIfNeeded
    Connect-ExchangeIfNeeded

    $props = 'Id', 'DisplayName', 'UserPrincipalName', 'Mail', 'OnPremisesSyncEnabled', 'UserType'
    $users = if ($UserListCsv) {
        Get-TargetUsers
    } else {
        @(Get-MgUser -All -Property $props -ConsistencyLevel eventual | Where-Object { $_.UserType -ne 'Guest' })
    }

    $odbByPath = @{}
    if ($AdminUrl) {
        try {
            Connect-SpoIfNeeded
            foreach ($site in Get-SPOSite -IncludePersonalSite $true -Limit All -Filter "Url -like '-my.sharepoint.com/personal/'") {
                $odbByPath[$site.Url.ToLower()] = $site
            }
            Write-Ok "$($odbByPath.Count) OneDrive sites enumerated"
        } catch {
            Write-Warn "Could not enumerate OneDrive sites: $($_.Exception.Message)"
        }
    }

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($u in $users) {
        $mbx = $null
        try { $mbx = Get-Mailbox -Identity $u.UserPrincipalName -ErrorAction Stop } catch { }
        $expectedPath = "/personal/$(Get-OneDrivePathFromUpn $u.UserPrincipalName)"
        $match = $odbByPath.Keys | Where-Object { $_.EndsWith($expectedPath.ToLower()) } | Select-Object -First 1
        $rows.Add([PSCustomObject]@{
            DisplayName        = $u.DisplayName
            UserPrincipalName  = $u.UserPrincipalName
            UpnOnNewDomain     = ($u.UserPrincipalName -like "*@$NewDomain")
            PrimarySmtpAddress = $mbx.PrimarySmtpAddress
            SmtpOnNewDomain    = ($mbx -and $mbx.PrimarySmtpAddress -like "*@$NewDomain")
            OldAddressRetained = ($mbx -and (($mbx.EmailAddresses -join ';') -like "*@$OldDomain*"))
            OneDriveUrl        = $match
            OneDriveMatchesUpn = [bool]$match
            OnPremisesSynced   = [bool]$u.OnPremisesSyncEnabled
        })
    }
    Export-Report $rows 'Validate-PostChange.csv'

    $pending = @($rows | Where-Object { -not $_.UpnOnNewDomain })
    $odbLag  = @($rows | Where-Object { $_.UpnOnNewDomain -and -not $_.OneDriveMatchesUpn })
    Write-Step 'Summary'
    Write-Detail "UPN on ${NewDomain}: $($rows.Count - $pending.Count) of $($rows.Count)"
    Write-Detail "Still on old domain: $($pending.Count)"
    Write-Detail "OneDrive URL not yet rebuilt: $($odbLag.Count) (allow up to 24 hours)"
}

# =============================================================================
# Phase: ReShare
# =============================================================================

function Invoke-ReSharePhase {
    Write-Banner 'PHASE: RESHARE (OneDrive and Teams chat files needing re-share)'
    Connect-SpoIfNeeded
    Connect-PnPIfNeeded
    Write-Warn 'Report only. Sharing links that embed the old UPN cannot be rewritten in bulk; owners must re-share.'
    Write-Detail 'Teams channel files live in the team SharePoint site and are not affected by a UPN change.'

    $rows = New-Object System.Collections.Generic.List[object]
    $drives = @(Get-SPOSite -IncludePersonalSite $true -Limit All -Filter "Url -like '-my.sharepoint.com/personal/'")
    Write-Step "Scanning $($drives.Count) OneDrive sites (max $MaxItemsPerDrive items each)"

    $i = 0
    foreach ($drive in $drives) {
        $i++
        Write-Progress -Activity 'ReShare scan' -Status $drive.Url -PercentComplete (($i / [Math]::Max($drives.Count, 1)) * 100)
        try {
            Connect-PnPOnline -Url $drive.Url -Interactive -ClientId $(if ($ClientId) { $ClientId } else { $env:ENTRAID_APP_ID })
            $items = Get-PnPListItem -List 'Documents' -PageSize 500 -Fields 'FileRef', 'FileLeafRef', 'SharedWithUsers', 'FSObjType' |
                     Select-Object -First $MaxItemsPerDrive
            foreach ($item in $items) {
                $sharedWith = $item.FieldValues['SharedWithUsers']
                if (-not $sharedWith) { continue }
                $fileRef = [string]$item.FieldValues['FileRef']
                $principals = ($sharedWith | ForEach-Object { $_.Email }) -join ';'
                $rows.Add([PSCustomObject]@{
                    OneDriveUrl   = $drive.Url
                    Owner         = $drive.Owner
                    FileUrl       = $fileRef
                    FileName      = [string]$item.FieldValues['FileLeafRef']
                    IsTeamsChatFile = ($fileRef -like '*Microsoft Teams Chat Files*')
                    SharedWith    = $principals
                    Action        = 'Re-share after UPN change'
                })
            }
            Write-Detail "$($drive.Url): scanned"
        } catch {
            Write-Warn "$($drive.Url): $($_.Exception.Message)"
            Add-Action 'ReShare' $drive.Url 'Scan' 'Enumerate shared items' 'Failed' $_.Exception.Message
        }
    }
    Write-Progress -Activity 'ReShare scan' -Completed
    Export-Report $rows 'ReShare-SharedItems.csv'

    $chat = @($rows | Where-Object { $_.IsTeamsChatFile }).Count
    Write-Step 'Summary'
    Write-Detail "Shared items found: $($rows.Count) (Teams chat files: $chat)"
    Write-Detail 'Recommended: move business-critical content into a Teams/SharePoint site before cutover so its URL is UPN-independent.'
}

# =============================================================================
# Phase: AutoReShare
# =============================================================================

function Invoke-AutoReSharePhase {
    Write-Banner 'PHASE: AUTORESHARE (re-share OneDrive items via Microsoft Graph invite)'
    Write-Detail 'Requires an Entra ID app registration granted these Graph APPLICATION permissions with admin consent:'
    Write-Detail '  User.Read.All, Files.ReadWrite.All, Sites.ReadWrite.All'
    Connect-GraphAppOnlyIfNeeded
    Write-Detail 'Named-recipient shares (including Teams chat file shares) are re-invited automatically.'
    Write-Detail 'Anyone-in-org / anyone-with-link shares cannot be resolved to recipients and are reported as NeedsManualReshare.'
    Write-Warn "Apply=$($Apply.IsPresent) - each successful re-share sends a new-link email to the original recipients."

    $targets = Get-TargetUsers
    Write-Step "$($targets.Count) users in scope (post-cutover UPNs on $NewDomain expected)"

    $rows = New-Object System.Collections.Generic.List[object]
    $i = 0
    foreach ($u in $targets) {
        $i++
        Write-Progress -Activity 'AutoReShare' -Status $u.UserPrincipalName -PercentComplete (($i / [Math]::Max($targets.Count, 1)) * 100)

        $driveUri = "/v1.0/users/$($u.Id)/drive"
        try {
            $drive = Invoke-MgGraphRequest -Method GET -Uri $driveUri -ErrorAction Stop
        } catch {
            Write-Warn "No OneDrive for $($u.UserPrincipalName), skipped: $($_.Exception.Message)"
            Add-Action 'AutoReShare' $u.UserPrincipalName 'ResolveDrive' '' 'Skipped' $_.Exception.Message
            continue
        }
        $driveId = $drive.id

        # Note: $expand=permissions is not supported on driveItem/children by
        # Microsoft Graph (returns 400 BadRequest) - permissions must be
        # fetched per item via the dedicated /permissions endpoint instead.
        $childrenUri = "/v1.0/drives/$driveId/root/children?`$top=$MaxItemsPerDrive"
        $items = New-Object System.Collections.Generic.List[object]
        try {
            do {
                $page = Invoke-MgGraphRequest -Method GET -Uri $childrenUri -ErrorAction Stop
                foreach ($it in $page.value) { $items.Add($it) }
                $childrenUri = $page.'@odata.nextLink'
            } while ($childrenUri -and $items.Count -lt $MaxItemsPerDrive)
        } catch {
            Write-Warn "$($u.UserPrincipalName): could not enumerate drive items: $($_.Exception.Message)"
            Add-Action 'AutoReShare' $u.UserPrincipalName 'EnumerateItems' '' 'Failed' $_.Exception.Message
            continue
        }

        foreach ($item in $items) {
            $permissions = $null
            try {
                $permResp = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/drives/$driveId/items/$($item.id)/permissions" -ErrorAction Stop
                $permissions = @($permResp.value)
            } catch {
                Write-Warn "$($u.UserPrincipalName): $($item.name): could not list permissions: $($_.Exception.Message)"
                Add-Action 'AutoReShare' $u.UserPrincipalName 'ListPermissions' $item.name 'Failed' $_.Exception.Message
                continue
            }
            Start-Sleep -Milliseconds $ThrottleMs
            if (-not $permissions) { continue }
            foreach ($perm in $permissions) {
                $recipientEmails = @()
                if ($perm.grantedToV2.user.email) { $recipientEmails += $perm.grantedToV2.user.email }
                if ($perm.grantedToIdentitiesV2) {
                    foreach ($gid in $perm.grantedToIdentitiesV2) {
                        if ($gid.user.email) { $recipientEmails += $gid.user.email }
                    }
                }
                $isTeamsChat = [bool]($item.parentReference.path -and $item.parentReference.path -like '*Microsoft Teams Chat Files*')

                if ($perm.link -and -not $recipientEmails) {
                    # "Anyone in the org" or "Anyone with the link" - Graph does not expose who received it.
                    $rows.Add([PSCustomObject]@{
                        Owner           = $u.UserPrincipalName
                        FileName        = $item.name
                        FileWebUrl      = $item.webUrl
                        IsTeamsChatFile = $isTeamsChat
                        PermissionType  = "Link:$($perm.link.scope)"
                        Recipients      = ''
                        Status          = 'NeedsManualReshare'
                        Detail          = 'Anonymous/organization link - no recipient list to re-invite'
                    })
                    continue
                }
                if (-not $recipientEmails) { continue }

                # The owner permission on every item lists the file owner
                # themselves with roles=['owner']. Graph's invite action only
                # accepts read/write roles - passing 'owner' (or re-inviting
                # the owner to their own file) always returns 400 BadRequest.
                if ($perm.roles -contains 'owner') { continue }
                $recipientEmails = @($recipientEmails | Where-Object { $_ -ne $u.Mail -and $_ -ne $u.UserPrincipalName })
                if (-not $recipientEmails) { continue }

                $roles = @(if ($perm.roles) { @($perm.roles) } else { @('read') })
                $recipients = @($recipientEmails | ForEach-Object { @{ email = $_ } })

                if (-not $Apply) {
                    $rows.Add([PSCustomObject]@{
                        Owner           = $u.UserPrincipalName
                        FileName        = $item.name
                        FileWebUrl      = $item.webUrl
                        IsTeamsChatFile = $isTeamsChat
                        PermissionType  = "Named:$($roles -join ',')"
                        Recipients      = ($recipientEmails -join ';')
                        Status          = 'WhatIf'
                        Detail          = 'Would call driveItem invite for these recipients'
                    })
                    Add-Action 'AutoReShare' $u.UserPrincipalName 'Invite' $item.name 'WhatIf'
                    continue
                }

                $body = @{
                    recipients     = $recipients
                    message        = $ReshareMessage
                    requireSignIn  = $true
                    sendInvitation = $true
                    roles          = $roles
                }
                # Serialize to a JSON string ourselves and send it as raw text.
                # Passing the hashtable directly to -Body makes Invoke-MgGraphRequest
                # serialize it internally, which throws "Self referencing loop
                # detected for property 'Value' ... Path 'recipients[0].email.Chars'"
                # under Windows PowerShell 5.1 (its bundled Newtonsoft.Json walks
                # into the ETS Chars indexer that 5.1 attaches to [string]).
                $bodyJson = $body | ConvertTo-Json -Depth 10
                try {
                    Invoke-MgGraphRequest -Method POST -Uri "/v1.0/drives/$driveId/items/$($item.id)/invite" -Body $bodyJson -ContentType 'application/json' -ErrorAction Stop | Out-Null
                    Write-Ok "$($item.name): re-shared with $($recipientEmails -join ', ')"
                    $rows.Add([PSCustomObject]@{
                        Owner           = $u.UserPrincipalName
                        FileName        = $item.name
                        FileWebUrl      = $item.webUrl
                        IsTeamsChatFile = $isTeamsChat
                        PermissionType  = "Named:$($roles -join ',')"
                        Recipients      = ($recipientEmails -join ';')
                        Status          = 'Success'
                        Detail          = ''
                    })
                    Add-Action 'AutoReShare' $u.UserPrincipalName 'Invite' $item.name 'Success'
                } catch {
                    Write-Fail "$($item.name): re-share failed: $($_.Exception.Message)"
                    $rows.Add([PSCustomObject]@{
                        Owner           = $u.UserPrincipalName
                        FileName        = $item.name
                        FileWebUrl      = $item.webUrl
                        IsTeamsChatFile = $isTeamsChat
                        PermissionType  = "Named:$($roles -join ',')"
                        Recipients      = ($recipientEmails -join ';')
                        Status          = 'Failed'
                        Detail          = $_.Exception.Message
                    })
                    Add-Action 'AutoReShare' $u.UserPrincipalName 'Invite' $item.name 'Failed' $_.Exception.Message
                }
                Start-Sleep -Milliseconds $ThrottleMs
            }
        }
    }
    Write-Progress -Activity 'AutoReShare' -Completed
    Export-Report $rows 'AutoReShare-Results.csv'

    $ok     = @($rows | Where-Object { $_.Status -eq 'Success' }).Count
    $manual = @($rows | Where-Object { $_.Status -eq 'NeedsManualReshare' }).Count
    $failed = @($rows | Where-Object { $_.Status -eq 'Failed' }).Count
    Write-Step 'Summary'
    Write-Detail "Auto re-shared: $ok"
    Write-Detail "Needs manual re-share (anonymous/org links): $manual"
    Write-Detail "Failed: $failed"
}

# =============================================================================
# Phase: Rename
# =============================================================================

function Invoke-RenamePhase {
    Write-Banner 'PHASE: RENAME (SharePoint tenant domain name)'
    if (-not $NewSpoName) { throw '-NewSpoName is required for the Rename phase.' }
    Connect-GraphIfNeeded
    Connect-SpoIfNeeded

    Write-Warn "Target URLs become https://$NewSpoName.sharepoint.com and https://$NewSpoName-my.sharepoint.com."
    Write-Warn "They cannot be $NewDomain - vanity domains are not supported for SharePoint or OneDrive."
    Write-Detail 'Not supported for Multi-Geo tenants or GCC/GCC High/DoD.'
    Write-Detail 'Run this in its own change window, separate from the UPN cutover.'

    $required = "$NewSpoName.onmicrosoft.com"
    $domain = Get-MgDomain -All | Where-Object { $_.Id -eq $required }
    if (-not $domain -or -not $domain.IsVerified) {
        throw "$required must be added and verified in Entra ID before renaming."
    }
    Write-Ok "$required is verified"

    try {
        $existing = Get-SPOTenantRenameStatus -ErrorAction Stop
        if ($existing) {
            Write-Warn "An existing rename job was found: State=$($existing.State)"
            $existing | Format-List | Out-String | Write-Host
            return
        }
    } catch { }

    if (-not $Apply) {
        Write-Detail "WHATIF Start-SPOTenantRename -DomainName $NewSpoName -ScheduledDateTime $($RenameScheduledUtc.ToString('yyyy-MM-ddTHH:mm:ssZ'))"
        Add-Action 'Rename' $NewSpoName 'StartRename' $RenameScheduledUtc.ToString('u') 'WhatIf'
        return
    }
    try {
        Start-SPOTenantRename -DomainName $NewSpoName -ScheduledDateTime $RenameScheduledUtc.ToString('yyyy-MM-ddTHH:mm:ssZ') -ErrorAction Stop
        Write-Ok "Rename scheduled for $($RenameScheduledUtc.ToString('u')) UTC"
        Add-Action 'Rename' $NewSpoName 'StartRename' $RenameScheduledUtc.ToString('u') 'Success'
        Write-Detail 'Monitor with: Get-SPOTenantRenameStatus'
    } catch {
        Write-Fail "Rename could not be scheduled: $($_.Exception.Message)"
        Add-Action 'Rename' $NewSpoName 'StartRename' $RenameScheduledUtc.ToString('u') 'Failed' $_.Exception.Message
    }
}

# =============================================================================
# Main
# =============================================================================

if (-not (Test-Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
}

$phases = if ($Phase -contains 'All') { @('Precheck', 'Inventory', 'Validate', 'ReShare') } else { $Phase }

Write-Banner 'MICROSOFT 365 DOMAIN CUTOVER'
Write-Host "Old domain : $OldDomain"      -ForegroundColor White
Write-Host "New domain : $NewDomain"      -ForegroundColor White
Write-Host "Phases     : $($phases -join ', ')" -ForegroundColor White
Write-Host "Mode       : $(if ($Apply) { 'APPLY - changes will be made' } else { 'DRY RUN - no changes' })" -ForegroundColor $(if ($Apply) { 'Red' } else { 'Green' })
Write-Host "Output     : $OutputFolder"   -ForegroundColor White

if ($Apply -and ($phases | Where-Object { $_ -in 'SetPrimary', 'Groups', 'Default', 'AutoReShare', 'Rename' })) {
    $confirm = Read-Host "Type APPLY to confirm tenant-wide changes"
    if ($confirm -ne 'APPLY') { Write-Warn 'Aborted by operator.'; return }
}

try {
    foreach ($p in $phases) {
        switch ($p) {
            'Precheck'   { Invoke-PrecheckPhase }
            'Inventory'  { Invoke-InventoryPhase }
            'AddAlias'   { Invoke-AddAliasPhase }
            'SetPrimary' { Invoke-SetPrimaryPhase }
            'Groups'     { Invoke-GroupsPhase }
            'Default'    { Invoke-DefaultPhase }
            'Validate'   { Invoke-ValidatePhase }
            'ReShare'    { Invoke-ReSharePhase }
            'AutoReShare' { Invoke-AutoReSharePhase }
            'Rename'     { Invoke-RenamePhase }
        }
    }
} finally {
    if ($script:ActionLog.Count -gt 0) {
        Export-Report $script:ActionLog.ToArray() 'ActionLog.csv'
    }
    if ($script:Connected.Exchange) { try { Disconnect-ExchangeOnline -Confirm:$false } catch { } }
    if ($script:Connected.PnP)      { try { Disconnect-PnPOnline } catch { } }
    if ($script:Connected.Graph)    { try { Disconnect-MgGraph | Out-Null } catch { } }
    Write-Banner 'COMPLETE'
    Write-Host "Reports written to: $OutputFolder" -ForegroundColor Green
}
