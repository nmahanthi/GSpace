#Requires -Version 5.1
<#
.SYNOPSIS
    Backs up each user's OneDrive (personal site) content into a per-user
    folder on a target SharePoint Online site, for a list of users in a CSV.

.DESCRIPTION
    Reads a CSV of users (column: UserPrincipalName), and for every user:
      1. Resolves their OneDrive (personal site) via Microsoft Graph.
      2. Creates a folder named after the user under the destination SPO
         site's document library (optionally nested under -DestinationBaseFolder).
      3. Recursively copies every file from the user's OneDrive into that
         folder, preserving the original sub-folder structure.

    Copying is done server-side via the Microsoft Graph driveItem "copy"
    action (no file content passes through this machine), and each copy
    job is polled to completion before moving to the next file.

    Runs in two phases: (1) discovers every file for every user first, so
    the totals are known up front, then (2) copies them while showing a
    status bar with % complete, files pending, elapsed time and an ETA
    (also logged to the console every 25 files).

    AUTHENTICATION
    Reading another user's OneDrive (/users/{id}/drive) is blocked under
    delegated auth (403 Forbidden), even for a signed-in Global
    Administrator. This script therefore requires app-only (client
    credentials) Graph auth. Register an Entra ID app and grant it these
    Graph APPLICATION permissions with admin consent:
        User.Read.All, Files.ReadWrite.All, Sites.ReadWrite.All
    Requires the Microsoft.Graph.Authentication module, v2.15.0 or later
    (needed for -ResponseHeadersVariable on Invoke-MgGraphRequest):
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser

    Defaults to a dry run (reports what would be copied). Pass -Apply to
    perform the actual copies.

.PARAMETER UserListCsv
    Path to a CSV with a UserPrincipalName column. An optional FolderName
    column can be included to override the computed destination folder
    name for a given row (default: the user's UPN, sanitized).

.PARAMETER DestinationSiteUrl
    Full URL of the target SharePoint Online site, e.g.
    https://contoso.sharepoint.com/sites/OneDriveBackup

.PARAMETER DestinationLibrary
    Name of the document library on the destination site to back up into.
    Default: "Documents".

.PARAMETER DestinationBaseFolder
    Optional sub-folder path under the library to nest all per-user folders
    under, e.g. "2026-Backup". Default: "" (user folders sit directly under
    the library root).

.PARAMETER TenantId
    Entra ID tenant ID (GUID or verified domain). Required.

.PARAMETER ClientId
    Entra ID app registration (client) ID with the application permissions
    listed above, admin-consented. Required.

.PARAMETER ClientSecret
    Client secret for the app registration, as a SecureString. If omitted,
    the script prompts for it interactively.

.PARAMETER MaxItemsPerUser
    Cap on the number of files copied per user; 0 = unlimited. Default: 0.

.PARAMETER ThrottleMs
    Delay between Graph calls to avoid throttling. Default: 250ms.

.PARAMETER MonitorTimeoutSec
    Max seconds to wait for each server-side copy job to complete before
    treating it as failed. Default: 180.

.PARAMETER GraphClientTimeoutSec
    Seconds to wait for the Graph/Azure AD token request before giving up
    (Connect-MgGraph -ClientTimeout). Raise this on slow/high-latency links.
    Default: 100 (the Microsoft.Graph.Authentication default).

.PARAMETER Apply
    Actually perform the copies. Without this switch, the script only
    reports what it would copy.

.PARAMETER OutputCsv
    Path for the result report CSV. Defaults to
    .\OneDriveBackupReport_<timestamp>.csv

.EXAMPLE
    # Dry run - see what would be copied
    .\Backup-OneDriveToSPO.ps1 -UserListCsv .\Users.csv `
        -DestinationSiteUrl "https://contoso.sharepoint.com/sites/OneDriveBackup" `
        -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

.EXAMPLE
    # Real run, nested under a dated folder
    .\Backup-OneDriveToSPO.ps1 -UserListCsv .\Users.csv `
        -DestinationSiteUrl "https://contoso.sharepoint.com/sites/OneDriveBackup" `
        -DestinationBaseFolder "2026-08-26" `
        -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -Apply
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$UserListCsv,
    [Parameter(Mandatory)][string]$DestinationSiteUrl,
    [string]$DestinationLibrary     = "Documents",
    [string]$DestinationBaseFolder  = "",
    [Parameter(Mandatory)][string]$TenantId,
    [Parameter(Mandatory)][string]$ClientId,
    [System.Security.SecureString]$ClientSecret,
    [int]$MaxItemsPerUser    = 0,
    [int]$ThrottleMs         = 250,
    [int]$MonitorTimeoutSec  = 180,
    [int]$GraphClientTimeoutSec = 100,
    [switch]$Apply,
    [string]$OutputCsv = (Join-Path $PSScriptRoot "OneDriveBackupReport_$(Get-Date -f 'yyyyMMdd_HHmmss').csv")
)

# Note: deliberately no Set-StrictMode here - Graph JSON responses are loosely
# typed (e.g. '@odata.nextLink', 'folder'/'file' facets, response headers) and
# only appear on some objects; strict mode would throw on routine absence.
$ErrorActionPreference = "Stop"

# ── Logging helper ───────────────────────────────────────────────────────────
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts    = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) { "SUCCESS"{"Green"} "WARN"{"Yellow"} "ERROR"{"Red"} default{"Cyan"} }
    Write-Host "[$ts][$Level] $Message" -ForegroundColor $color
}

function Assert-Module {
    param([string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        throw "Required module '$Name' is not installed. Run: Install-Module $Name -Scope CurrentUser"
    }
    Import-Module $Name -ErrorAction Stop
}

# ── Graph app-only connection ────────────────────────────────────────────────
function Connect-GraphAppOnly {
    Assert-Module "Microsoft.Graph.Authentication"
    $secret = $ClientSecret
    if (-not $secret) {
        $secret = Read-Host -Prompt "Client secret for app $ClientId" -AsSecureString
    }
    $credential = New-Object System.Management.Automation.PSCredential ($ClientId, $secret)
    Write-Log "Connecting to Microsoft Graph (app-only, Tenant: $TenantId, ClientId: $ClientId) ..."
    try {
        Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $credential -ClientTimeout $GraphClientTimeoutSec -NoWelcome -ErrorAction Stop
    } catch {
        $full = $_.Exception.ToString()
        if ($_.Exception.InnerException) { $full += "`n--- Inner ---`n" + $_.Exception.InnerException.ToString() }
        Write-Log "Full error detail:`n$full" "ERROR"
        if ($full -match "request_timeout|TaskCanceledException|HttpClient.Timeout") {
            Write-Log @"
This is a NETWORK-LEVEL timeout reaching Azure AD (login.microsoftonline.com) - not a wrong secret/ClientId. The token request never got a response at all (StatusCode: 0). Likely causes on this machine/network:
  - A corporate proxy or firewall is blocking/inspecting outbound HTTPS to login.microsoftonline.com and graph.microsoft.com. Ask network/security to allow those hosts.
  - VPN or SSL-inspection software silently drops the TLS handshake - try from a different network (e.g. mobile hotspot) to confirm.
  - If this machine needs a proxy to reach the internet, .NET's HttpClient needs it configured system-wide, e.g.: netsh winhttp import proxy source=ie, or set the HTTPS_PROXY environment variable in a NEW PowerShell window before running this script.
  - Confirm basic connectivity first: Test-NetConnection login.microsoftonline.com -Port 443
  - You can raise the wait (helps on slow/high-latency links, will not fix a real block) by re-running with -GraphClientTimeoutSec 300.
"@ "WARN"
        } else {
            Write-Log @"
Common causes for 'ClientSecretCredential authentication failed':
  - AADSTS7000215 / AADSTS7000222: secret value is wrong, expired, or you pasted the Secret ID instead of the Secret Value (Entra ID > App registrations > your app > Certificates & secrets - the VALUE is only shown once at creation).
  - AADSTS700016: ClientId does not match an app registration in this tenant, or TenantId is wrong.
  - AADSTS90002: TenantId not found - use the tenant GUID or a verified domain (e.g. contoso.onmicrosoft.com), not the friendly display name.
  - AADSTS65001 / permissions not consented: the app's API permissions (User.Read.All, Files.ReadWrite.All, Sites.ReadWrite.All - Application type) need admin consent granted (Grant admin consent button in the Azure Portal).
  - Module too old: Update-Module Microsoft.Graph.Authentication
"@ "WARN"
        }
        throw
    }
    Write-Log "Connected." "SUCCESS"
}

# ── CSV of users ──────────────────────────────────────────────────────────────
function Get-BackupTargets {
    if (-not (Test-Path $UserListCsv)) { throw "UserListCsv not found: $UserListCsv" }
    $rows = Import-Csv -Path $UserListCsv
    if (-not ($rows | Get-Member -Name UserPrincipalName)) {
        throw "UserListCsv must contain a UserPrincipalName column."
    }
    $hasFolderCol = [bool]($rows | Get-Member -Name FolderName)
    $rows | Where-Object { $_.UserPrincipalName -and $_.UserPrincipalName.Trim() } | ForEach-Object {
        $upn = $_.UserPrincipalName.Trim()
        $folder = if ($hasFolderCol -and $_.FolderName -and $_.FolderName.Trim()) { $_.FolderName.Trim() } else { $upn }
        [PSCustomObject]@{
            UserPrincipalName = $upn
            FolderName        = (ConvertTo-SafeFolderName $folder)
        }
    }
}

function ConvertTo-SafeFolderName {
    param([string]$Name)
    # SharePoint-invalid folder/file name characters: " * : < > ? / \ | # %
    $clean = ($Name -replace '["\*:<>\?/\\\|#%]', '_').Trim()
    $clean = $clean.TrimEnd('.', ' ')
    if (-not $clean) { $clean = "Unnamed" }
    return $clean
}

function Join-RelPath {
    param([string]$Base, [string]$Leaf)
    if (-not $Base) { return $Leaf }
    if (-not $Leaf) { return $Base }
    return "$Base/$Leaf"
}

# ── Destination site / drive resolution ──────────────────────────────────────
function Get-DestinationDriveId {
    $uri = [Uri]$DestinationSiteUrl
    $hostName = $uri.Host
    $sitePath = $uri.AbsolutePath.Trim('/')
    $site = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/sites/${hostName}:/${sitePath}" -ErrorAction Stop
    $drivesResp = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/sites/$($site.id)/drives" -ErrorAction Stop
    $drive = $drivesResp.value | Where-Object { $_.name -eq $DestinationLibrary }
    if (-not $drive) {
        $available = ($drivesResp.value | ForEach-Object { $_.name }) -join ', '
        throw "Document library '$DestinationLibrary' not found on $DestinationSiteUrl. Available libraries: $available"
    }
    return $drive.id
}

# ── Ensure a (possibly nested) folder path exists in a drive, return its item id ──
function Resolve-DriveFolder {
    param([string]$DriveId, [string[]]$PathSegments)
    $parentId = "root"
    foreach ($seg in $PathSegments) {
        if (-not $seg) { continue }
        $existingId = $null
        $childrenUri = "/v1.0/drives/$DriveId/items/$parentId/children?`$top=500&`$select=id,name,folder"
        while ($childrenUri) {
            $page = Invoke-MgGraphRequest -Method GET -Uri $childrenUri -ErrorAction Stop
            $match = $page.value | Where-Object { $_.folder -and $_.name -eq $seg } | Select-Object -First 1
            if ($match) { $existingId = $match.id; break }
            $childrenUri = $page.'@odata.nextLink'
        }
        if ($existingId) {
            $parentId = $existingId
        } else {
            $body = @{ name = $seg; folder = @{}; "@microsoft.graph.conflictBehavior" = "fail" } | ConvertTo-Json
            try {
                $created = Invoke-MgGraphRequest -Method POST -Uri "/v1.0/drives/$DriveId/items/$parentId/children" -Body $body -ContentType "application/json" -ErrorAction Stop
                $parentId = $created.id
            } catch {
                # Race: created by a concurrent run between our check and create - re-fetch it.
                $page = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/drives/$DriveId/items/$parentId/children?`$top=500&`$select=id,name,folder" -ErrorAction Stop
                $match = $page.value | Where-Object { $_.folder -and $_.name -eq $seg } | Select-Object -First 1
                if (-not $match) { throw }
                $parentId = $match.id
            }
        }
    }
    return $parentId
}

# ── Recursively enumerate every file in a user's OneDrive ────────────────────
function Get-DriveFilesRecursive {
    param([string]$DriveId, [string]$ItemId = "root", [string]$RelativeFolder = "", [ref]$Count)
    $results = New-Object System.Collections.Generic.List[object]
    $uri = "/v1.0/drives/$DriveId/items/$ItemId/children?`$top=200"
    while ($uri) {
        if ($MaxItemsPerUser -gt 0 -and $Count.Value -ge $MaxItemsPerUser) { break }
        $page = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        foreach ($item in $page.value) {
            if ($MaxItemsPerUser -gt 0 -and $Count.Value -ge $MaxItemsPerUser) { break }
            if ($item.folder) {
                $sub = Get-DriveFilesRecursive -DriveId $DriveId -ItemId $item.id -RelativeFolder (Join-RelPath $RelativeFolder $item.name) -Count $Count
                foreach ($s in $sub) { $results.Add($s) }
            } elseif ($item.file) {
                $results.Add([PSCustomObject]@{ Id = $item.id; Name = $item.name; RelativeFolder = $RelativeFolder })
                $Count.Value++
            }
        }
        $uri = $page.'@odata.nextLink'
    }
    return $results
}

# ── Fire a server-side copy and wait for it to complete ──────────────────────
function Copy-DriveItemAndWait {
    param([string]$SourceDriveId, [string]$ItemId, [string]$DestDriveId, [string]$DestFolderId, [string]$Name)

    $body = @{ parentReference = @{ driveId = $DestDriveId; id = $DestFolderId }; name = $Name } | ConvertTo-Json
    $headers = $null
    Invoke-MgGraphRequest -Method POST -Uri "/v1.0/drives/$SourceDriveId/items/$ItemId/copy" `
        -Body $body -ContentType "application/json" -ResponseHeadersVariable headers -ErrorAction Stop | Out-Null

    $location = $headers.Location
    if ($location -is [array]) { $location = $location[0] }
    if (-not $location) { return }  # Some tenants complete the copy synchronously with no monitor URL.

    $sw = [Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $MonitorTimeoutSec) {
        $status = $null
        try { $status = Invoke-RestMethod -Uri $location -Method GET -ErrorAction Stop }
        catch { $status = Invoke-MgGraphRequest -Method GET -Uri $location -ErrorAction Stop }
        if ($status.status -eq "completed") { return }
        if ($status.status -in @("failed", "cannotConvert", "malwareDetected")) {
            throw "Copy job status '$($status.status)': $($status.statusDescription)"
        }
        Start-Sleep -Milliseconds 1000
    }
    throw "Copy job timed out after $MonitorTimeoutSec sec."
}

# ── Format a TimeSpan as d.hh:mm:ss (days omitted when zero) ─────────────────
function Format-Duration {
    param([TimeSpan]$Span)
    if ($Span.TotalDays -ge 1) { return $Span.ToString("d\.hh\:mm\:ss") }
    return $Span.ToString("hh\:mm\:ss")
}

# =============================================================================
# MAIN
# =============================================================================
Write-Log "Mode: $(if ($Apply) { 'APPLY (real copy)' } else { 'DRY RUN (no changes - pass -Apply to copy)' })" "WARN"

Connect-GraphAppOnly

Write-Log "Resolving destination library '$DestinationLibrary' on $DestinationSiteUrl ..."
$destDriveId = Get-DestinationDriveId
Write-Log "Destination drive resolved." "SUCCESS"

$targets = @(Get-BackupTargets)
Write-Log "$($targets.Count) users in scope from $UserListCsv"

$results = New-Object System.Collections.Generic.List[object]
$baseSegments = @($DestinationBaseFolder -split '/' | Where-Object { $_ })

# ── Phase 1: discovery - enumerate every file for every user first, so the
#    overall %, pending count and ETA in Phase 2 are accurate from the start.
#    No destination writes happen in this phase. ──────────────────────────────
Write-Log "Phase 1/2: discovering files across $($targets.Count) user(s) ..."
$work = New-Object System.Collections.Generic.List[object]
$u = 0
foreach ($target in $targets) {
    $u++
    $upn = $target.UserPrincipalName
    Write-Progress -Id 1 -Activity "Discovering OneDrive content" -Status "$upn ($u of $($targets.Count))" -PercentComplete (($u / [Math]::Max($targets.Count, 1)) * 100)

    try {
        $drive = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/users/$upn/drive" -ErrorAction Stop
    } catch {
        Write-Log "[$upn] No OneDrive found, skipped: $($_.Exception.Message)" "WARN"
        $results.Add([PSCustomObject]@{
            UserPrincipalName = $upn; FolderName = $target.FolderName
            File = ""; RelativeFolder = ""; Status = "Skipped"; Detail = "No OneDrive: $($_.Exception.Message)"
        })
        continue
    }

    $count = 0
    try {
        $files = Get-DriveFilesRecursive -DriveId $drive.id -Count ([ref]$count)
    } catch {
        Write-Log "[$upn] Could not enumerate OneDrive files: $($_.Exception.Message)" "ERROR"
        $results.Add([PSCustomObject]@{
            UserPrincipalName = $upn; FolderName = $target.FolderName
            File = ""; RelativeFolder = ""; Status = "Failed"; Detail = "Enumerate: $($_.Exception.Message)"
        })
        continue
    }
    Write-Log "[$upn] $($files.Count) file(s) found."
    foreach ($f in $files) {
        $work.Add([PSCustomObject]@{
            UserPrincipalName = $upn; FolderName = $target.FolderName; SourceDriveId = $drive.id
            FileId = $f.Id; FileName = $f.Name; RelativeFolder = $f.RelativeFolder
        })
    }
}
Write-Progress -Id 1 -Activity "Discovering OneDrive content" -Completed

$totalWork = $work.Count
Write-Log "Phase 1/2 complete: $totalWork file(s) to process across $($targets.Count) user(s)." "SUCCESS"

# ── Phase 2: copy (or report, in dry run) each file, with an overall status
#    bar showing % complete, files pending and an ETA based on the running
#    average time per file. Folder lookups are cached per user/sub-folder so
#    repeat files in the same folder don't re-walk the destination tree. ─────
$userFolderCache = @{}
$subFolderCache  = @{}
$sw = [Diagnostics.Stopwatch]::StartNew()
$done = 0; $ok = 0; $skip = 0; $fail = 0
$perUserOk = @{}; $perUserSkip = @{}; $perUserFail = @{}

foreach ($item in $work) {
    $done++
    $pending = $totalWork - $done
    $pct = if ($totalWork -gt 0) { [Math]::Round(($done / $totalWork) * 100, 1) } else { 100 }
    $avgPerItemSec = if ($done -gt 0) { $sw.Elapsed.TotalSeconds / $done } else { 0 }
    $etaStr = if ($done -lt 3) { "calculating..." } else { Format-Duration ([TimeSpan]::FromSeconds([Math]::Max(0, $avgPerItemSec * $pending))) }
    $status = "$done/$totalWork done ($pct%) | Pending: $pending | Elapsed: $(Format-Duration $sw.Elapsed) | ETA: $etaStr"
    Write-Progress -Id 2 -Activity "OneDrive Backup" -Status $status -PercentComplete $pct -CurrentOperation "$($item.UserPrincipalName): $($item.FileName)"
    if ($done % 25 -eq 0 -or $done -eq $totalWork) {
        Write-Log "Progress: $status"
    }

    $upn = $item.UserPrincipalName

    if (-not $Apply) {
        $results.Add([PSCustomObject]@{
            UserPrincipalName = $upn; FolderName = $item.FolderName
            File = $item.FileName; RelativeFolder = $item.RelativeFolder
            Status = "WhatIf"; Detail = "Would copy to $DestinationLibrary/$($item.FolderName)/$($item.RelativeFolder)"
        })
        continue
    }

    try {
        if (-not $userFolderCache.ContainsKey($item.FolderName)) {
            $userFolderCache[$item.FolderName] = Resolve-DriveFolder -DriveId $destDriveId -PathSegments ($baseSegments + $item.FolderName)
        }
        $targetFolderId = $userFolderCache[$item.FolderName]
        if ($item.RelativeFolder) {
            $subKey = "$($item.FolderName)|$($item.RelativeFolder)"
            if (-not $subFolderCache.ContainsKey($subKey)) {
                $subFolderCache[$subKey] = Resolve-DriveFolder -DriveId $destDriveId -PathSegments ($baseSegments + $item.FolderName + ($item.RelativeFolder -split '/' | Where-Object { $_ }))
            }
            $targetFolderId = $subFolderCache[$subKey]
        }
    } catch {
        $fail++; $perUserFail[$upn] = ($perUserFail[$upn] + 1)
        $results.Add([PSCustomObject]@{
            UserPrincipalName = $upn; FolderName = $item.FolderName
            File = $item.FileName; RelativeFolder = $item.RelativeFolder
            Status = "Failed"; Detail = "CreateFolder: $($_.Exception.Message)"
        })
        continue
    }

    try {
        Copy-DriveItemAndWait -SourceDriveId $item.SourceDriveId -ItemId $item.FileId -DestDriveId $destDriveId -DestFolderId $targetFolderId -Name $item.FileName
        $ok++; $perUserOk[$upn] = ($perUserOk[$upn] + 1)
        $results.Add([PSCustomObject]@{
            UserPrincipalName = $upn; FolderName = $item.FolderName
            File = $item.FileName; RelativeFolder = $item.RelativeFolder
            Status = "Success"; Detail = ""
        })
    } catch {
        $msg = $_.Exception.Message
        if ($msg -like "*nameAlreadyExists*") {
            $skip++; $perUserSkip[$upn] = ($perUserSkip[$upn] + 1)
            $results.Add([PSCustomObject]@{
                UserPrincipalName = $upn; FolderName = $item.FolderName
                File = $item.FileName; RelativeFolder = $item.RelativeFolder
                Status = "Skipped"; Detail = "Already exists at destination"
            })
        } else {
            $fail++; $perUserFail[$upn] = ($perUserFail[$upn] + 1)
            Write-Log "[$upn] $($item.FileName): copy failed: $msg" "ERROR"
            $results.Add([PSCustomObject]@{
                UserPrincipalName = $upn; FolderName = $item.FolderName
                File = $item.FileName; RelativeFolder = $item.RelativeFolder
                Status = "Failed"; Detail = $msg
            })
        }
    }
    Start-Sleep -Milliseconds $ThrottleMs
}
Write-Progress -Id 2 -Activity "OneDrive Backup" -Completed

$results | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
Write-Log "Report written: $OutputCsv" "SUCCESS"

$total   = $results.Count
$success = @($results | Where-Object { $_.Status -eq "Success" }).Count
$skipped = @($results | Where-Object { $_.Status -eq "Skipped" }).Count
$failed  = @($results | Where-Object { $_.Status -eq "Failed" }).Count
$whatif  = @($results | Where-Object { $_.Status -eq "WhatIf" }).Count
Write-Log "Summary: Total=$total Success=$success Skipped=$skipped Failed=$failed WhatIf=$whatif | Total time: $(Format-Duration $sw.Elapsed)" "SUCCESS"
if (-not $Apply) { Write-Log "This was a dry run. Re-run with -Apply to perform the copies." "WARN" }
