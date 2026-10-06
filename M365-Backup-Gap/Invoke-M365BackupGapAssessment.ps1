<#
.SYNOPSIS
    Read-only Microsoft 365 "backup gap" assessment for ONE customer tenant.
.DESCRIPTION
    Uses Microsoft Graph (read-only) to collect user counts, data volumes, global admin
    count and any third-party backup apps registered in the tenant. Writes a JSON file
    (consumed by New-BackupGapOutreach.ps1) and an HTML customer report that maps the
    findings to three risks: employee deletion, ransomware on synced files, rogue admin.
.NOTES
    Requires: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
    App registration (customer consents once) with Application permissions:
      User.Read.All, Directory.Read.All, Application.Read.All, Reports.Read.All, AuditLog.Read.All
    Optional: run Invoke-M365PurviewChecks.ps1 first; its output/<Customer>.purview.json is merged in.
    Auth: pass -ClientId and set the env var named by -SecretEnvVar (unattended),
          or omit them to sign in interactively as a customer admin (delegated).
    Note: tenant report names may be concealed by the customer; totals are unaffected.
.EXAMPLE
    ./Invoke-M365BackupGapAssessment.ps1 -TenantId <guid> -ClientId <guid> -SecretEnvVar M365_SECRET_CONTOSO -PricePerUser 3
#>
[CmdletBinding()]
param(
    [string]$CustomerName,   # optional: defaults to the tenant's own name
    [string]$TenantId,
    [string]$ClientId,
    [string]$SecretEnvVar,
    [decimal]$PricePerUser = 3,            # GBP per user per month (your quote)
    [decimal]$LossCostPerUser = 500,       # GBP assumption: cost of losing a year of data per user
    [string]$OutputDir = (Join-Path $PSScriptRoot 'output')
)

$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

# ---- Connect ---------------------------------------------------------------
if ($ClientId -and $SecretEnvVar -and $TenantId) {
    $secret = [Environment]::GetEnvironmentVariable($SecretEnvVar)
    if (-not $secret) { throw "Env var '$SecretEnvVar' is empty." }
    $cred = [pscredential]::new($ClientId, (ConvertTo-SecureString $secret -AsPlainText -Force))
    Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $cred -NoWelcome
} else {
    $scopes = 'User.Read.All','Directory.Read.All','Application.Read.All','Reports.Read.All','AuditLog.Read.All'
    if ($TenantId) { Connect-MgGraph -TenantId $TenantId -Scopes $scopes -NoWelcome }
    else           { Connect-MgGraph -Scopes $scopes -NoWelcome }
}

function Get-GraphAll([string]$Uri) {
    $items = @()
    while ($Uri) {
        $r = Invoke-MgGraphRequest -Method GET -Uri $Uri -Headers @{ ConsistencyLevel = 'eventual' }
        $items += $r.value
        $Uri = $r.'@odata.nextLink'
    }
    $items
}

function Get-UsageTotalGB([string]$Report) {
    $tmp = [IO.Path]::GetTempFileName()
    try {
        Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/reports/$Report(period='D30')" -OutputFilePath $tmp | Out-Null
        $rows = Import-Csv $tmp | Where-Object { $_.'Is Deleted' -ne 'True' }
        $bytes = ($rows | ForEach-Object { [double]($_.'Storage Used (Byte)') } | Measure-Object -Sum).Sum
        [pscustomobject]@{ GB = [math]::Round(($bytes / 1GB), 1); Count = @($rows).Count }
    } catch {
        Write-Warning "Report $Report unavailable: $($_.Exception.Message)"
        [pscustomobject]@{ GB = $null; Count = $null }
    } finally { Remove-Item $tmp -ErrorAction SilentlyContinue }
}

# ---- Collect ---------------------------------------------------------------
$org = (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization').value[0]
if (-not $CustomerName) { $CustomerName = if ($org.displayName) { $org.displayName } elseif ($TenantId) { $TenantId } else { 'Tenant' } }
$safeName = ($CustomerName -replace '[^\w\-]', '_')
Write-Host "Assessing $CustomerName ..." -ForegroundColor Cyan

$licensed = (Invoke-MgGraphRequest -Method GET -Headers @{ ConsistencyLevel = 'eventual' } `
    -Uri 'https://graph.microsoft.com/v1.0/users?$filter=assignedLicenses/$count ne 0 and userType eq ''Member''&$count=true&$top=1').'@odata.count'

$guests = (Invoke-MgGraphRequest -Method GET -Headers @{ ConsistencyLevel = 'eventual' } `
    -Uri 'https://graph.microsoft.com/v1.0/users?$filter=userType eq ''Guest''&$count=true&$top=1').'@odata.count'

# Global Administrator role template id is fixed across tenants
$gaTemplate = '62e90394-69f5-4237-9190-012177145e10'
$globalAdmins = 0
try {
    $role = @(Get-GraphAll "https://graph.microsoft.com/v1.0/directoryRoles?`$filter=roleTemplateId eq '$gaTemplate'")
    if ($role) { $globalAdmins = @(Get-GraphAll "https://graph.microsoft.com/v1.0/directoryRoles/$($role[0].id)/members").Count }
} catch { Write-Warning "Global admin lookup failed: $($_.Exception.Message)" }

$vendors = 'Veeam','AvePoint','Afi','Keepit','Druva','Backupify','Spanning','Cohesity','Dropsuite','SkyKick',
           'Barracuda','Acronis','Datto','Commvault','Rubrik','HYCU','CloudAlly','Cove','N-able','Backupify','Redstor','Syncro','Altaro','Hornetsecurity','365 Total Backup'
$backupApps = @()
try {
    $sps = Get-GraphAll 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=displayName,appId&$top=999'
    $backupApps = @($sps | Where-Object { $n = $_.displayName; $vendors | Where-Object { $n -match [regex]::Escape($_) } } |
                    Select-Object -ExpandProperty displayName -Unique)
} catch { Write-Warning "Service principal lookup failed: $($_.Exception.Message)" }

$mail = Get-UsageTotalGB 'getMailboxUsageDetail'
$od   = Get-UsageTotalGB 'getOneDriveUsageAccountDetail'
$spo  = Get-UsageTotalGB 'getSharePointSiteUsageDetail'

# Entra directory audit: user/group deletions in last 30 days (admin deletion signal)
$since = (Get-Date).AddDays(-30).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
$entraDeletes = $null
try {
    $aud = Get-GraphAll "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits?`$filter=activityDateTime ge $since&`$top=999"
    $entraDeletes = @($aud | Where-Object { $_.activityDisplayName -match '^(Delete|Hard delete) (user|group)' }).Count
} catch { Write-Warning "Directory audit lookup failed: $($_.Exception.Message)" }

$purviewPath = Join-Path $OutputDir "$safeName.purview.json"
$pv = if (Test-Path $purviewPath) { Get-Content $purviewPath -Raw | ConvertFrom-Json } else { $null }

# ---- Score -----------------------------------------------------------------
$hasBackup = $backupApps.Count -gt 0
$findings = @(
    [pscustomobject]@{ Risk='Employee deletes a mailbox or SharePoint site'
        Level = $(if ($hasBackup) {'Low'} else {'High'})
        Detail = 'Native recovery is limited: Exchange deleted items default to 14 days (max 30); SharePoint/OneDrive recycle bins keep items 93 days; deleted Microsoft 365 groups 30 days.' },
    [pscustomobject]@{ Risk='Ransomware encrypting synced files'
        Level = $(if ($hasBackup) {'Low'} elseif ($od.GB -gt 0) {'High'} else {'Medium'})
        Detail = "OneDrive holds $($od.GB) GB and SharePoint $($spo.GB) GB that sync to devices. Version history is not a point-in-time backup." },
    [pscustomobject]@{ Risk='Departing admin wipes data'
        Level = $(if ($hasBackup -and $globalAdmins -le 4) {'Low'} elseif ($globalAdmins -gt 4) {'High'} else {'Medium'})
        Detail = "$globalAdmins Global Administrators. An admin can hard-delete data and the recycle bins; a separate backup copy is outside their tenant." }
)

if ($pv) {
    if ($pv.PSObject.Properties['AuditIngestionEnabled']) {
        $findings += [pscustomobject]@{ Risk='Audit logging (evidence after an incident)'
            Level = $(if ($pv.AuditIngestionEnabled) {'Low'} else {'High'})
            Detail = $(if ($pv.AuditIngestionEnabled) {'Unified audit log is on. Standard retention is 180 days, so older activity cannot be investigated.'} else {'Unified audit log is OFF: deletions and admin actions are not recorded.'}) }
    }
    if ($pv.PSObject.Properties['RetentionCoverage']) {
        $cov = $pv.RetentionCoverage
        $missing = @('Exchange','SharePoint','OneDrive','Teams') | Where-Object { -not $cov.$_ }
        $delPol = @($pv.DeleteRulePolicies)
        $detail = if ($missing.Count) { "No enabled retention policy covers: $($missing -join ', ')." } else { 'Enabled retention policies cover all main workloads.' }
        if ($delPol.Count) { $detail += " Policies that delete data after a period: $($delPol -join ', ')." }
        $detail += ' Retention policies are not a backup: they cannot restore after a tenant-level wipe, and an admin can edit or remove them.'
        $findings += [pscustomobject]@{ Risk='Purview retention policies'
            Level = $(if ($missing.Count -ge 2) {'High'} elseif ($missing.Count -eq 1 -or $delPol.Count) {'Medium'} else {'Low'}); Detail = $detail }
    }
    if ($pv.PSObject.Properties['Mailboxes']) {
        $m = $pv.Mailboxes
        $findings += [pscustomobject]@{ Risk='Mailbox holds and recoverable items'
            Level = $(if ($m.Total -gt 0 -and $m.WithAnyHold -eq 0) {'High'} elseif ($m.WithAnyHold -lt $m.Total) {'Medium'} else {'Low'})
            Detail = "$($m.WithAnyHold) of $($m.Total) mailboxes on any hold; single-item recovery off on $($m.SingleItemRecoveryOff)." }
    }
}
if ($entraDeletes -gt 0 -or ($pv -and $pv.PSObject.Properties['Deletions'])) {
    $d = if ($pv -and $pv.PSObject.Properties['Deletions']) { "$($pv.Deletions.Events) deletion events in $($pv.Deletions.LookbackDays) days ($($pv.Deletions.SiteDeletes) site deletes). " } else { '' }
    $findings += [pscustomobject]@{ Risk='Recent deletion activity'
        Level = $(if (($entraDeletes -gt 20) -or ($pv -and $pv.Deletions.Events -gt 1000)) {'Medium'} else {'Low'})
        Detail = "$d$entraDeletes user/group deletions in Entra in the last 30 days. Review top deleters before treating as normal." }
}

$users = [int]$licensed
$result = [ordered]@{
    CustomerName    = $CustomerName
    TenantName      = $org.displayName
    TenantId        = $org.id
    AssessedOn      = (Get-Date -Format 'yyyy-MM-dd')
    LicensedUsers   = $users
    GuestUsers      = [int]$guests
    GlobalAdmins    = $globalAdmins
    MailboxGB       = $mail.GB
    OneDriveGB      = $od.GB
    SharePointGB    = $spo.GB
    BackupAppsFound = $backupApps
    PricePerUser    = $PricePerUser
    AnnualBackupCost   = [math]::Round($users * $PricePerUser * 12, 0)
    EstimatedLossCost  = [math]::Round($users * $LossCostPerUser, 0)
    EntraDeletes30d = $entraDeletes
    Purview         = $pv
    Findings        = $findings
}

$monthlyCost = [math]::Round($users * $PricePerUser, 0)
$steps = @()
$steps += $(if ($hasBackup) {
    "<b>Verify your backup vendor.</b> Detected: $($backupApps -join ', '). Confirm it covers Exchange, SharePoint, OneDrive and Teams, keeps copies outside your tenant, and that a restore has been tested in the last 6 months."
} else {
    "<b>Choose a third-party Microsoft 365 backup vendor.</b> Look for Exchange, SharePoint, OneDrive and Teams coverage, storage outside your tenant (ideally immutable), daily backups, point-in-time restore, and a restore test during onboarding. Budget about &pound;$monthlyCost per month for $users users."
})
$steps += $(if ($globalAdmins -gt 4) {
    "<b>Reduce Global Administrators from $globalAdmins to 2-4.</b> Move day-to-day admins to narrower roles (Exchange, SharePoint, Backup Admin), use separate admin accounts with phishing-resistant MFA, make elevation time-limited (Privileged Identity Management, Entra ID P2) and keep two monitored break-glass accounts. Remove admin rights on the day someone leaves."
} else {
    "<b>Keep Global Administrators at $globalAdmins.</b> Maintain 2-4, use separate admin accounts with phishing-resistant MFA, keep two monitored break-glass accounts, and remove admin rights on the day someone leaves."
})
$steps += $(if ($pv -and $pv.PSObject.Properties['AuditIngestionEnabled'] -and -not $pv.AuditIngestionEnabled) {
    "<b>Turn on the unified audit log now</b> (Purview > Audit). Without it, deletions and admin actions are not recorded. Then extend retention as below."
} else {
    "<b>Lengthen audit log retention.</b> Standard keeps audit data for 180 days; Audit (Premium, included in Microsoft 365 E5) keeps 1 year, with up to 10 years as an add-on. Where that is not affordable, export audit logs regularly to storage or a SIEM you control, so an incident found late can still be investigated."
})
$stepsHtml = ($steps | ForEach-Object { "<li>$_</li>" }) -join "`n"

$jsonPath = Join-Path $OutputDir "$safeName.json"
$result.NextSteps = @($steps -replace '<[^>]+>','' -replace '&pound;','GBP ')
$result | ConvertTo-Json -Depth 5 | Set-Content -Path $jsonPath -Encoding UTF8

# ---- HTML report -----------------------------------------------------------
$rows = ($findings | ForEach-Object {
    $c = switch ($_.Level) { 'High' {'#c0392b'} 'Medium' {'#e67e22'} default {'#27ae60'} }
    "<tr><td>$($_.Risk)</td><td style='color:$c;font-weight:bold'>$($_.Level)</td><td>$($_.Detail)</td></tr>"
}) -join "`n"
$backupTxt = if ($hasBackup) { "Possible backup tooling detected: $($backupApps -join ', ') (verify it covers Exchange, SharePoint, OneDrive, Teams)." } else { 'No third-party backup application detected in the tenant.' }
$html = @"
<html><head><meta charset='utf-8'><title>$CustomerName - Microsoft 365 backup gap</title>
<style>body{font-family:Segoe UI,Arial;margin:2em;max-width:900px}td,th{border:1px solid #ccc;padding:6px;vertical-align:top}table{border-collapse:collapse;width:100%}</style></head><body>
<h1>Microsoft 365 backup gap - $CustomerName</h1>
<p>Microsoft keeps the <b>service</b> running. It does not protect you from deletion, ransomware or a rogue admin.</p>
<p>$backupTxt</p>
<ul><li>Licensed users: $users (guests: $guests)</li><li>Mailbox: $($mail.GB) GB | OneDrive: $($od.GB) GB | SharePoint: $($spo.GB) GB</li><li>Global admins: $globalAdmins</li></ul>
<table><tr><th>Risk</th><th>Level</th><th>Why</th></tr>$rows</table>
<h2>Cost comparison</h2>
<p>Third-party backup: about <b>&pound;$($result.AnnualBackupCost)</b> per year ($users users x &pound;$PricePerUser x 12).<br>
Losing a year of data (assumed &pound;$LossCostPerUser per user): about <b>&pound;$($result.EstimatedLossCost)</b>.</p>
<h2>Recommended next steps</h2>
<ol>$stepsHtml</ol>
<p style='color:#777;font-size:small'>Assessed $($result.AssessedOn) with read-only Microsoft Graph access. $(if ($pv) {'Purview data from Exchange/Compliance PowerShell.'} else {'Purview retention policies were not assessed (run Invoke-M365PurviewChecks.ps1); confirm manually.'}) Audit log history is limited to the tenant audit retention period.</p>
</body></html>
"@
$htmlPath = Join-Path $OutputDir "$safeName.html"
$html | Set-Content -Path $htmlPath -Encoding UTF8

Disconnect-MgGraph | Out-Null
Write-Host "Done: $jsonPath , $htmlPath" -ForegroundColor Green
