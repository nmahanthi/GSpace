<#
.SYNOPSIS
    Read-only Purview / Exchange retention and audit checks for ONE customer tenant.
.DESCRIPTION
    Writes output/<Customer>.purview.json, which Invoke-M365BackupGapAssessment.ps1 merges into
    its report if present (run this script first). Checks:
      - Unified audit log ingestion on/off
      - Retention policies and rules (locations, keep vs delete, duration)
      - Mailbox holds / Recoverable Items settings (litigation hold, single item recovery, deleted item window)
      - Deletion activity in the last 30 days from the unified audit log (top deleters)
.NOTES
    Requires: Install-Module ExchangeOnlineManagement -Scope CurrentUser   (v3.2+)
    These cmdlets are not available in Microsoft Graph, so this uses Exchange Online and
    Security & Compliance PowerShell.
    App-only setup (one-time, per customer):
      1. App registration with API permission Office 365 Exchange Online > Application > Exchange.ManageAsApp (admin consent).
      2. Upload a certificate to the app; export it as .pfx and base64 it into a secret.
      3. Assign the app's service principal the Entra role "Global Reader".
         If audit search returns nothing, also add it to the Purview role group "View-Only Audit Logs".
    Interactive alternative: omit -PfxEnvVar and sign in as a customer admin.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$CustomerName,
    [string]$Organization,          # e.g. contoso.onmicrosoft.com
    [string]$ClientId,
    [string]$PfxEnvVar,             # env var holding base64 .pfx
    [string]$PfxPasswordEnvVar,     # env var holding .pfx password (optional)
    [int]$LookbackDays = 30,
    [string]$OutputDir = (Join-Path $PSScriptRoot 'output')
)
$ErrorActionPreference = 'Stop'
$safe = ($CustomerName -replace '[^\w\-]', '_')
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

# ---- Connect ---------------------------------------------------------------
if ($PfxEnvVar -and $ClientId -and $Organization) {
    $b64 = [Environment]::GetEnvironmentVariable($PfxEnvVar)
    if (-not $b64) { throw "Env var '$PfxEnvVar' is empty." }
    $pw = if ($PfxPasswordEnvVar) { [Environment]::GetEnvironmentVariable($PfxPasswordEnvVar) } else { $null }
    $cert = if ($pw) { [Security.Cryptography.X509Certificates.X509Certificate2]::new([Convert]::FromBase64String($b64), $pw) }
            else     { [Security.Cryptography.X509Certificates.X509Certificate2]::new([Convert]::FromBase64String($b64)) }
    Connect-ExchangeOnline -Certificate $cert -AppId $ClientId -Organization $Organization -ShowBanner:$false
    Connect-IPPSSession    -Certificate $cert -AppId $ClientId -Organization $Organization -ShowBanner:$false
} else {
    Connect-ExchangeOnline -ShowBanner:$false
    Connect-IPPSSession -ShowBanner:$false
}

$out = [ordered]@{ CustomerName = $CustomerName; CheckedOn = (Get-Date -Format 'yyyy-MM-dd'); Errors = @() }
function Invoke-Check([string]$Name, [scriptblock]$Block) {
    try { & $Block } catch { $out.Errors += "${Name}: $($_.Exception.Message)"; Write-Warning "${Name}: $($_.Exception.Message)" }
}

# ---- Audit logging ---------------------------------------------------------
Invoke-Check 'AuditConfig' {
    $out.AuditIngestionEnabled = [bool](Get-AdminAuditLogConfig).UnifiedAuditLogIngestionEnabled
}

# ---- Retention policies ----------------------------------------------------
Invoke-Check 'RetentionPolicies' {
    $policies = foreach ($p in Get-RetentionCompliancePolicy -DistributionDetail) {
        $rules = @(Get-RetentionComplianceRule -Policy $p.Name)
        [pscustomobject]@{
            Name       = $p.Name
            Enabled    = [bool]$p.Enabled
            Exchange   = [bool]@($p.ExchangeLocation).Count
            SharePoint = [bool]@($p.SharePointLocation).Count
            OneDrive   = [bool]@($p.OneDriveLocation).Count
            Groups     = [bool]@($p.ModernGroupLocation).Count
            Teams      = [bool]@($p.TeamsChannelLocation).Count
            Rules      = @($rules | ForEach-Object { [pscustomobject]@{
                            Action = "$($_.RetentionComplianceAction)"; Days = $_.RetentionDuration } })
        }
    }
    $out.RetentionPolicies = @($policies)
    $enabled = @($policies | Where-Object Enabled)
    $out.RetentionCoverage = [ordered]@{
        Exchange   = [bool]($enabled | Where-Object Exchange)
        SharePoint = [bool]($enabled | Where-Object SharePoint)
        OneDrive   = [bool]($enabled | Where-Object OneDrive)
        Groups     = [bool]($enabled | Where-Object Groups)
        Teams      = [bool]($enabled | Where-Object Teams)
    }
    # Delete-type rules actively remove data; these are not protection
    $out.DeleteRulePolicies = @($enabled | Where-Object { $_.Rules | Where-Object { $_.Action -match 'Delete' } } | ForEach-Object Name)
}

# ---- Mailbox holds and recoverable items ----------------------------------
Invoke-Check 'Mailboxes' {
    $mbx = @(Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox `
        -Properties LitigationHoldEnabled, InPlaceHolds, SingleItemRecoveryEnabled, RetainDeletedItemsFor)
    $out.Mailboxes = [ordered]@{
        Total                  = $mbx.Count
        WithLitigationHold     = @($mbx | Where-Object LitigationHoldEnabled).Count
        WithAnyHold            = @($mbx | Where-Object { $_.LitigationHoldEnabled -or @($_.InPlaceHolds).Count }).Count
        SingleItemRecoveryOff  = @($mbx | Where-Object { -not $_.SingleItemRecoveryEnabled }).Count
        DeletedItemWindowDays  = @($mbx | Group-Object { [int]$_.RetainDeletedItemsFor.TotalDays } |
                                    ForEach-Object { [pscustomobject]@{ Days = [int]$_.Name; Mailboxes = $_.Count } })
    }
}

# ---- Deletion activity -----------------------------------------------------
Invoke-Check 'DeletionActivity' {
    $ops = 'HardDelete','SoftDelete','FileDeleted','FileDeletedSecondStageRecycleBin','FolderDeleted','SiteDeleted','Remove-Mailbox'
    $rec = @(Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-$LookbackDays) -EndDate (Get-Date) `
                -Operations $ops -ResultSize 5000)
    $out.Deletions = [ordered]@{
        LookbackDays = $LookbackDays
        Events       = $rec.Count
        CappedAt5000 = ($rec.Count -ge 5000)
        TopDeleters  = @($rec | Group-Object UserIds | Sort-Object Count -Descending | Select-Object -First 5 |
                          ForEach-Object { [pscustomobject]@{ User = $_.Name; Events = $_.Count } })
        SiteDeletes  = @($rec | Where-Object Operations -eq 'SiteDeleted').Count
    }
}

Disconnect-ExchangeOnline -Confirm:$false | Out-Null
$path = Join-Path $OutputDir "$safe.purview.json"
$out | ConvertTo-Json -Depth 6 | Set-Content -Path $path -Encoding UTF8
Write-Host "Done: $path" -ForegroundColor Green
