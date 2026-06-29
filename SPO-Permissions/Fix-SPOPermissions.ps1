<#
.SYNOPSIS
    Remediates SharePoint Online permission discrepancies found by Compare-Permissions.ps1.
.DESCRIPTION
    Reads the differences CSV, connects to each affected SPO site, and:
      - MISSING  -> Grants the expected permission level to the user/group.
      - MISMATCH -> Removes the wrong level and grants the expected one.
      - EXTRA    -> Optionally removes the unrecognised permission (controlled by Config.psd1 Fix.RemoveExtraPermissions).
    Every action is written to a timestamped fix-log CSV.
    Set Fix.WhatIf = $true in Config.psd1 to simulate without making changes.
.PARAMETER ConfigPath
    Path to Config.psd1.  Defaults to .\Config.psd1
.PARAMETER DifferencesFile
    Path to the differences CSV from Compare-Permissions.ps1.
    If omitted the newest matching file in Output.Directory is used.
.PARAMETER OutputFile
    Overrides the fix-log CSV path.
.EXAMPLE
    # Dry-run first (WhatIf = $true in Config.psd1)
    .\Fix-SPOPermissions.ps1
    # Then set WhatIf = $false and re-run to apply changes
#>
[CmdletBinding()]
param(
    [string]$ConfigPath       = ".\Config.psd1",
    [string]$DifferencesFile  = "",
    [string]$OutputFile       = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not (Get-Module -ListAvailable -Name "PnP.PowerShell")) {
    throw "PnP.PowerShell module is not installed. Run: Install-Module PnP.PowerShell -Scope CurrentUser"
}
Import-Module PnP.PowerShell -ErrorAction Stop

# ---------------------------------------------------------------------------
# Load config
# ---------------------------------------------------------------------------
$cfg    = Import-PowerShellDataFile -Path $ConfigPath
$date   = Get-Date -Format "yyyyMMdd_HHmmss"
$dir    = $cfg.Output.Directory
$whatIf = [bool]$cfg.Fix.WhatIf
$removeExtra = [bool]$cfg.Fix.RemoveExtraPermissions

if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }

function Resolve-LatestCsv { param([string]$Pattern)
    Get-ChildItem -Path $dir -Filter $Pattern | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName
}
if (-not $DifferencesFile) { $DifferencesFile = Resolve-LatestCsv "$($cfg.Output.DifferencesFile)*.csv" }
if (-not $OutputFile)      { $OutputFile = Join-Path $dir "$($cfg.Output.FixLogFile)_$date.csv" }

Write-Host "`n=== Fix-SPOPermissions.ps1 ===" -ForegroundColor Cyan
Write-Host "WhatIf mode       : $whatIf"
Write-Host "Remove extra perms: $removeExtra"
Write-Host "Differences source: $DifferencesFile`n"

$diffs  = Import-Csv $DifferencesFile
$fixLog = [System.Collections.Generic.List[object]]::new()

# Cache PnP connections per site to avoid repeated reconnects
$connectedSites = [System.Collections.Generic.HashSet[string]]::new()

function Connect-SPOSite { param([string]$SiteUrl)
    if (-not $connectedSites.Contains($SiteUrl)) {
        Connect-PnPOnline -Url $SiteUrl `
            -ClientId           $cfg.SharePoint.ClientId `
            -CertificatePath    $cfg.SharePoint.CertificatePath `
            -CertificatePassword (ConvertTo-SecureString $cfg.SharePoint.CertificatePassword -AsPlainText -Force) `
            -Tenant             $cfg.SharePoint.TenantId
        $connectedSites.Add($SiteUrl) | Out-Null
    }
}

function Write-Log {
    param([string]$SiteUrl, [string]$Email, [string]$Action, [string]$Detail, [string]$Status, [string]$Error = "")
    $fixLog.Add([PSCustomObject]@{
        Timestamp     = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        SPOSiteUrl    = $SiteUrl
        PrincipalEmail= $Email
        Action        = $Action
        Detail        = $Detail
        Status        = $Status
        WhatIf        = $whatIf
        Error         = $Error
    })
    $color = if ($Status -eq "SUCCESS") { "Green" } elseif ($Status -eq "WHATIF") { "Cyan" } else { "Red" }
    Write-Host "  [$Status] $Action | $Email | $Detail" -ForegroundColor $color
}

# ---------------------------------------------------------------------------
# Process each difference
# ---------------------------------------------------------------------------
foreach ($diff in $diffs) {
    $siteUrl = $diff.SPOSiteUrl.TrimEnd('/')
    $email   = $diff.PrincipalEmail.Trim()
    $type    = $diff.DifferenceType

    if (-not $siteUrl -or -not $email) { continue }

    Write-Host "Processing [$type] $email @ $siteUrl" -ForegroundColor DarkCyan

    try {
        Connect-SPOSite -SiteUrl $siteUrl

        switch ($type) {

            "MISSING" {
                $level = $diff.ExpectedSPOLevel
                if ($whatIf) {
                    Write-Log $siteUrl $email "GRANT" "Would grant '$level'" "WHATIF"
                } else {
                    Set-PnPWebPermission -User $email -AddRole $level
                    Write-Log $siteUrl $email "GRANT" "Granted '$level'" "SUCCESS"
                }
            }

            "MISMATCH" {
                $addLevel    = $diff.ExpectedSPOLevel
                $removeLevels = $diff.ActualSPOLevel -split ";\s*"
                if ($whatIf) {
                    Write-Log $siteUrl $email "FIX_LEVEL" "Would remove '$($diff.ActualSPOLevel)' and grant '$addLevel'" "WHATIF"
                } else {
                    foreach ($rl in $removeLevels) {
                        if ($rl -and $rl -ne $addLevel) {
                            try { Set-PnPWebPermission -User $email -RemoveRole $rl } catch { <# ignore if already gone #> }
                        }
                    }
                    Set-PnPWebPermission -User $email -AddRole $addLevel
                    Write-Log $siteUrl $email "FIX_LEVEL" "Replaced '$($diff.ActualSPOLevel)' with '$addLevel'" "SUCCESS"
                }
            }

            "EXTRA" {
                if ($removeExtra) {
                    $removeLevels = $diff.ActualSPOLevel -split ";\s*"
                    if ($whatIf) {
                        Write-Log $siteUrl $email "REMOVE" "Would remove extra permission '$($diff.ActualSPOLevel)'" "WHATIF"
                    } else {
                        foreach ($rl in $removeLevels) {
                            if ($rl) { Set-PnPWebPermission -User $email -RemoveRole $rl }
                        }
                        Write-Log $siteUrl $email "REMOVE" "Removed extra permission '$($diff.ActualSPOLevel)'" "SUCCESS"
                    }
                } else {
                    Write-Log $siteUrl $email "SKIP_EXTRA" "Extra permission kept (RemoveExtraPermissions=false)" "SKIPPED"
                }
            }
        }
    } catch {
        Write-Log $siteUrl $email "ERROR" $diff.DifferenceType "FAILED" $_.Exception.Message
    }
}

$fixLog | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8

$success = ($fixLog | Where-Object Status -eq "SUCCESS").Count
$failed  = ($fixLog | Where-Object Status -eq "FAILED").Count
$skipped = ($fixLog | Where-Object Status -in @("SKIPPED","WHATIF")).Count

Write-Host "`n--- Fix Summary ---" -ForegroundColor Cyan
Write-Host "  Successful actions : $success" -ForegroundColor Green
Write-Host "  Failed actions     : $failed"  -ForegroundColor Red
Write-Host "  Skipped / WhatIf   : $skipped" -ForegroundColor Yellow
Write-Host "Fix log written -> $OutputFile`n" -ForegroundColor Green
return $OutputFile
