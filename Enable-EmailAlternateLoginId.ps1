<#
.SYNOPSIS
    Manages Entra ID's "Email as an alternate login ID" (public preview)
    feature - letting users sign in with either their old domain email or
    their new UPN, while the UPN remains the default identity everywhere.

.DESCRIPTION
    This does NOT touch any user's UPN. It configures a tenant-level policy
    so the Entra ID sign-in page also matches against each user's
    ProxyAddresses (mail aliases) in addition to UserPrincipalName. For your
    migrated users, that means they can type either name@kaaratech.com or
    name@kaara.ai at sign-in and land on the same account - UPN (kaara.ai)
    stays what's shown everywhere else (Outlook, Teams, admin center, etc).

    Prerequisite: each user must still have their old-domain address as a
    ProxyAddress. Update-UserUpn.ps1 never removes aliases, so this should
    already be true - verify with Test-AlternateLoginIdReadiness.ps1 first.

    Two independent mechanisms, matching Microsoft's guidance to pilot
    before going tenant-wide:

    * Staged rollout (pilot)  - enables the feature only for members of one
      security group. Use -Phase EnablePilot / DisablePilot.
    * Org-wide (Home Realm Discovery policy) - enables the feature for the
      entire tenant. Use -Phase PromoteToOrgWide / DisableOrgWide.

    All write actions require -Apply; without it, the script only reports
    what it would do. -Phase Status is always read-only.

.PARAMETER Phase
    Status (default), EnablePilot, DisablePilot, PromoteToOrgWide, or
    DisableOrgWide.

.PARAMETER PilotGroupName
    Display name of an existing Entra ID security group to use for staged
    rollout. Required for EnablePilot / DisablePilot.

.PARAMETER Apply
    Actually make the change. Without it, EnablePilot/DisablePilot/
    PromoteToOrgWide/DisableOrgWide only report what they would do.

.EXAMPLE
    # See current state (staged rollout + org-wide policy)
    .\Enable-EmailAlternateLoginId.ps1

.EXAMPLE
    # Pilot with a test group first
    .\Enable-EmailAlternateLoginId.ps1 -Phase EnablePilot -PilotGroupName "Pilot-AltLoginId" -Apply

.EXAMPLE
    # Once confirmed working for the pilot group, roll out tenant-wide
    .\Enable-EmailAlternateLoginId.ps1 -Phase PromoteToOrgWide -Apply
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('Status', 'EnablePilot', 'DisablePilot', 'PromoteToOrgWide', 'DisableOrgWide')]
    [string]$Phase = 'Status',

    [Parameter(Mandatory = $false)]
    [string]$PilotGroupName,

    [Parameter(Mandatory = $false)]
    [switch]$Apply
)

$ErrorActionPreference = 'Stop'

function Write-Ok     { param([string]$Text) Write-Host "   OK   $Text" -ForegroundColor Green }
function Write-Warn   { param([string]$Text) Write-Host "   WARN $Text" -ForegroundColor Yellow }
function Write-Fail   { param([string]$Text) Write-Host "   FAIL $Text" -ForegroundColor Red }
function Write-Detail { param([string]$Text) Write-Host "        $Text" -ForegroundColor Gray }
function Write-Step   { param([string]$Text) Write-Host "-> $Text" -ForegroundColor Yellow }

function Get-FeatureRolloutPolicies {
    (Invoke-MgGraphRequest -Method GET -Uri '/beta/policies/featureRolloutPolicies').value |
        Where-Object { $_.feature -eq 'emailAsAlternateId' }
}

function Get-HomeRealmDiscoveryPolicies {
    (Invoke-MgGraphRequest -Method GET -Uri '/v1.0/policies/homeRealmDiscoveryPolicies').value
}

function Test-AlternateIdLoginEnabled {
    param($HrdPolicy)
    foreach ($d in $HrdPolicy.definition) {
        try {
            $obj = $d | ConvertFrom-Json
            if ($obj.HomeRealmDiscoveryPolicy.AlternateIdLogin.Enabled) { return $true }
        } catch { }
    }
    return $false
}

function Show-Status {
    Write-Step 'Staged rollout (pilot) policy for EmailAsAlternateId'
    $rollout = @(Get-FeatureRolloutPolicies)
    if ($rollout.Count -eq 0) {
        Write-Detail 'None configured.'
    } else {
        foreach ($p in $rollout) {
            Write-Detail "Id=$($p.id)  DisplayName='$($p.displayName)'  IsEnabled=$($p.isEnabled)"
            $applies = @((Invoke-MgGraphRequest -Method GET -Uri "/beta/policies/featureRolloutPolicies/$($p.id)/appliesTo").value)
            if ($applies.Count -eq 0) {
                Write-Detail '  (no groups assigned yet)'
            } else {
                foreach ($g in $applies) { Write-Detail "  -> group $($g.id)" }
            }
        }
    }

    Write-Step 'Org-wide Home Realm Discovery policy'
    $hrd = @(Get-HomeRealmDiscoveryPolicies)
    if ($hrd.Count -eq 0) {
        Write-Detail 'No HomeRealmDiscoveryPolicy exists - email-as-alternate-login-id is NOT enabled tenant-wide.'
    } else {
        foreach ($p in $hrd) {
            $enabled = Test-AlternateIdLoginEnabled -HrdPolicy $p
            Write-Detail "Id=$($p.id)  DisplayName='$($p.displayName)'  IsOrganizationDefault=$($p.isOrganizationDefault)  AlternateIdLogin.Enabled=$enabled"
        }
    }
}

function Resolve-PilotGroupId {
    if (-not $PilotGroupName) { throw '-PilotGroupName is required for this phase.' }
    $escaped = $PilotGroupName.Replace("'", "''")
    $groups = @((Invoke-MgGraphRequest -Method GET -Uri "/v1.0/groups?`$filter=displayName eq '$escaped'").value)
    if ($groups.Count -eq 0) { throw "Group not found: $PilotGroupName" }
    if ($groups.Count -gt 1) { throw "Multiple groups named '$PilotGroupName' found; rename to be unique." }
    return $groups[0].id
}

function Enable-Pilot {
    $groupId = Resolve-PilotGroupId
    $rollout = @(Get-FeatureRolloutPolicies)

    if ($rollout.Count -eq 0) {
        Write-Step 'Creating staged rollout policy for EmailAsAlternateId'
        $body = @{ feature = 'emailAsAlternateId'; displayName = 'EmailAsAlternateId Pilot Rollout'; isEnabled = $true } | ConvertTo-Json -Compress
        $policy = Invoke-MgGraphRequest -Method POST -Uri '/beta/policies/featureRolloutPolicies' -Body $body -ContentType 'application/json'
        Write-Ok "Created rollout policy $($policy.id)"
    } else {
        $policy = $rollout[0]
        if (-not $policy.isEnabled) {
            Invoke-MgGraphRequest -Method PATCH -Uri "/beta/policies/featureRolloutPolicies/$($policy.id)" -Body (@{ isEnabled = $true } | ConvertTo-Json -Compress) -ContentType 'application/json' | Out-Null
            Write-Ok 'Re-enabled existing rollout policy'
        } else {
            Write-Detail 'Rollout policy already exists and is enabled.'
        }
    }

    $refBody = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$groupId" } | ConvertTo-Json -Compress
    try {
        Invoke-MgGraphRequest -Method POST -Uri "/beta/policies/featureRolloutPolicies/$($policy.id)/appliesTo/`$ref" -Body $refBody -ContentType 'application/json' | Out-Null
        Write-Ok "Added group '$PilotGroupName' to the pilot rollout"
    } catch {
        if ($_.Exception.Message -match 'already exist|One or more added object references already exist') {
            Write-Detail "Group '$PilotGroupName' is already part of the pilot rollout."
        } else {
            throw
        }
    }
    Write-Detail 'May take up to 24 hours for newly-added group members to see the effect.'
}

function Disable-Pilot {
    $groupId = Resolve-PilotGroupId
    $rollout = @(Get-FeatureRolloutPolicies)
    if ($rollout.Count -eq 0) { Write-Warn 'No rollout policy exists - nothing to remove.'; return }
    $policy = $rollout[0]
    try {
        Invoke-MgGraphRequest -Method DELETE -Uri "/beta/policies/featureRolloutPolicies/$($policy.id)/appliesTo/$groupId/`$ref" | Out-Null
        Write-Ok "Removed group '$PilotGroupName' from the pilot rollout"
    } catch {
        Write-Warn "Could not remove group (it may not have been a member): $($_.Exception.Message)"
    }
}

function Set-OrgWide {
    param([bool]$Enable)
    $defObj = @{ HomeRealmDiscoveryPolicy = @{ AlternateIdLogin = @{ Enabled = $Enable } } }
    $defString = $defObj | ConvertTo-Json -Depth 5 -Compress
    $hrd = @(Get-HomeRealmDiscoveryPolicies)

    if ($hrd.Count -eq 0) {
        if (-not $Enable) { Write-Warn 'No HomeRealmDiscoveryPolicy exists - nothing to disable.'; return }
        $body = @{ displayName = 'AlternateLoginId-OrgWide'; definition = @($defString); isOrganizationDefault = $true } | ConvertTo-Json -Depth 5
        $created = Invoke-MgGraphRequest -Method POST -Uri '/v1.0/policies/homeRealmDiscoveryPolicies' -Body $body -ContentType 'application/json'
        Write-Ok "Created org-wide HomeRealmDiscoveryPolicy $($created.id) with AlternateIdLogin.Enabled=$Enable"
    } else {
        $policy = $hrd[0]
        $body = @{ displayName = $policy.displayName; definition = @($defString); isOrganizationDefault = $true } | ConvertTo-Json -Depth 5
        Invoke-MgGraphRequest -Method PATCH -Uri "/v1.0/policies/homeRealmDiscoveryPolicies/$($policy.id)" -Body $body -ContentType 'application/json' | Out-Null
        Write-Ok "Updated org-wide HomeRealmDiscoveryPolicy $($policy.id): AlternateIdLogin.Enabled=$Enable"
    }
    Write-Detail 'Can take up to 1 hour to propagate tenant-wide.'
}

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    throw "Required module 'Microsoft.Graph.Authentication' is not installed. Run: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
}

$connectedHere = $false
try {
    if (-not (Get-MgContext)) {
        Write-Step 'Connecting to Microsoft Graph'
        Connect-MgGraph -Scopes 'Policy.ReadWrite.ApplicationConfiguration', 'Directory.ReadWrite.All' -NoWelcome
        $connectedHere = $true
    } else {
        Write-Detail "Using existing Microsoft Graph session ($((Get-MgContext).Account))"
    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "EMAIL AS ALTERNATE LOGIN ID" -ForegroundColor Cyan
    Write-Host "Phase : $Phase" -ForegroundColor Cyan
    Write-Host "Mode  : $(if ($Apply) { 'APPLY - changes will be made' } else { 'DRY RUN (WhatIf)' })" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan

    switch ($Phase) {
        'Status' { Show-Status }

        'EnablePilot' {
            if (-not $Apply) {
                Write-Warn "Dry run: would create/enable a staged rollout policy and add group '$PilotGroupName'. Pass -Apply to make this change."
            } else {
                $confirm = Read-Host "Type APPLY to confirm enabling email-as-alternate-login-id for group '$PilotGroupName'"
                if ($confirm -ne 'APPLY') { Write-Warn 'Aborted by operator.'; return }
                Enable-Pilot
            }
        }

        'DisablePilot' {
            if (-not $Apply) {
                Write-Warn "Dry run: would remove group '$PilotGroupName' from the pilot rollout. Pass -Apply to make this change."
            } else {
                $confirm = Read-Host "Type APPLY to confirm removing group '$PilotGroupName' from the pilot rollout"
                if ($confirm -ne 'APPLY') { Write-Warn 'Aborted by operator.'; return }
                Disable-Pilot
            }
        }

        'PromoteToOrgWide' {
            if (-not $Apply) {
                Write-Warn 'Dry run: would enable email-as-alternate-login-id tenant-wide. Pass -Apply to make this change.'
            } else {
                $confirm = Read-Host 'Type APPLY to confirm enabling email-as-alternate-login-id for the ENTIRE tenant'
                if ($confirm -ne 'APPLY') { Write-Warn 'Aborted by operator.'; return }
                Set-OrgWide -Enable $true
            }
        }

        'DisableOrgWide' {
            if (-not $Apply) {
                Write-Warn 'Dry run: would disable email-as-alternate-login-id tenant-wide. Pass -Apply to make this change.'
            } else {
                $confirm = Read-Host 'Type APPLY to confirm disabling email-as-alternate-login-id for the ENTIRE tenant'
                if ($confirm -ne 'APPLY') { Write-Warn 'Aborted by operator.'; return }
                Set-OrgWide -Enable $false
            }
        }
    }
} finally {
    if ($connectedHere) { try { Disconnect-MgGraph | Out-Null } catch { } }
}
