<#
.SYNOPSIS
    Generates a tailored outreach email (.txt) for each customer in a CSV.
.DESCRIPTION
    Uses assessment JSON from Invoke-M365BackupGapAssessment.ps1 when it exists
    (output/<Customer>.json); otherwise falls back to the Users column in the CSV.
.EXAMPLE
    ./New-BackupGapOutreach.ps1 -CustomersCsv ./customers.csv -PricePerUser 3
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$CustomersCsv,
    [decimal]$PricePerUser = 3,
    [decimal]$LossCostPerUser = 500,
    [string]$YourName = 'Your Name',
    [string]$OutputDir = (Join-Path $PSScriptRoot 'output')
)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

foreach ($c in Import-Csv $CustomersCsv) {
    $safe = ($c.CustomerName -replace '[^\w\-]', '_')
    $jsonPath = Join-Path $OutputDir "$safe.json"
    $a = if (Test-Path $jsonPath) { Get-Content $jsonPath -Raw | ConvertFrom-Json } else { $null }

    $users = if ($a) { [int]$a.LicensedUsers } elseif ($c.Users) { [int]$c.Users } else { 0 }
    $annual = [math]::Round($users * $PricePerUser * 12, 0)
    $loss   = [math]::Round($users * $LossCostPerUser, 0)
    $monthly = [math]::Round($users * $PricePerUser, 0)
    $name = if ($c.ContactName) { ($c.ContactName -split ' ')[0] } else { 'there' }

    $own = ''
    if ($a) {
        $backup = if ($a.BackupAppsFound -and @($a.BackupAppsFound).Count) { "we saw possible backup tooling ($(@($a.BackupAppsFound) -join ', ')) that is worth verifying" } else { 'we found no third-party backup in your tenant' }
        $own = "`nWhat we found in $($a.TenantName): $($a.LicensedUsers) licensed users, $($a.OneDriveGB) GB in OneDrive, $($a.SharePointGB) GB in SharePoint, $($a.GlobalAdmins) global admins; $backup.`n"
    }
    $costLine = if ($users) { "For $users users that is about £$monthly a month (£$annual a year). Losing a year of email and files is estimated at £$loss." } else { 'It costs a few pounds per user per month; losing a year of email costs far more.' }

    $body = @"
To: $($c.ContactEmail)
Subject: Is your Microsoft 365 data really backed up, $($c.CustomerName)?

Hi $name,

"Microsoft backs up my 365 data." Not quite.

Microsoft keeps the service running. It does not protect you from:
- An employee deleting a mailbox or SharePoint site
- Ransomware encrypting synced files
- A departing admin wiping data on the way out

Retention windows are short, and recovery isn't guaranteed.
$own
A third-party backup costs a few pounds per user per month. $costLine

Happy to share a short report or set this up for you. Worth a 15-minute call?

$YourName
"@
    $path = Join-Path $OutputDir "$safe-email.txt"
    $body | Set-Content -Path $path -Encoding UTF8
    Write-Host "Wrote $path"
}
