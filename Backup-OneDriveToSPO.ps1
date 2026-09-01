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
    action (no file content passes through this machine). Copies run in a
    sliding window of up to -BatchSize concurrent jobs: as soon as any one
    finishes, its slot is immediately refilled with the next queued file, so
    the server-side wait time for many files overlaps instead of being paid
    one file at a time, and a single large straggler file never blocks the
    rest of the queue behind it.

    Runs in two phases: (1) discovers every file for every user first, so
    the totals (including total data size) are known up front, then (2)
    copies them while showing a status bar with % complete, GB copied, files
    pending, elapsed time and an ETA
    (also logged to the console every 25 files).

    KEEP-ALIVE / LONG RUNS
    Every Graph call is bounded by a hard per-call timeout and run through a
    retry-with-backoff wrapper, and the Graph connection is proactively
    refreshed every -TokenRefreshIntervalMin minutes - so a multi-hour backup
    survives network blips, VPN reconnects, laptop sleep/wake, or a silently
    dropped idle connection instead of hanging forever on one stuck call.

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
    Delay between copy-job dispatches, to avoid throttling. Since the actual
    file transfer happens entirely server-side, this is the main cost of
    dispatching each job - lower it (even to 0) for large/many-user backups
    if you are not seeing 429 (throttling) errors in the log. Default: 100ms.

.PARAMETER BatchSize
    Max number of copy jobs kept in flight at once, in a sliding window: as
    soon as any one finishes, its slot is immediately refilled with the next
    queued file, rather than waiting for a whole fixed batch to finish. This
    is the main lever for speeding up large (multi-GB per user) backups,
    since it overlaps the server-side copy wait time across many files
    instead of paying it one file at a time, and a single large straggler no
    longer blocks the rest of the queue behind it. The copy itself is a
    lightweight async dispatch (no data flows through this machine), so this
    can usually be raised well above the default for large backups (e.g.
    50-100) - raise it if you are not seeing 429 (throttling) errors; lower
    it if you are. Default: 10.

.PARAMETER MonitorTimeoutSec
    Max seconds to wait for each server-side copy job to complete before
    treating it as failed. Large files can genuinely take much longer than
    small ones to copy server-side - raise this for users with multi-GB
    files (e.g. 1800 for very large files). Default: 300.

.PARAMETER GraphClientTimeoutSec
    Seconds to wait for the Graph/Azure AD token request before giving up
    (Connect-MgGraph -ClientTimeout). Raise this on slow/high-latency links.
    Default: 100 (the Microsoft.Graph.Authentication default).

.PARAMETER RequestTimeoutSec
    Hard cap, in seconds, on any single Graph HTTP call. On long-running
    backups (hours), a network blip, sleep/wake or VPN reconnect can leave a
    call hung forever with no error and no progress - this bounds every call
    so the script always recovers (aborts the stuck call, reconnects, and
    retries) instead of hanging indefinitely. Default: 120.

.PARAMETER TokenRefreshIntervalMin
    Proactively reconnect to Microsoft Graph (fresh token + fresh HTTP
    connection) every N minutes of wall-clock time, independent of errors,
    to keep the session alive across long runs. Default: 45.

.PARAMETER MaxGraphRetries
    Max attempts per Graph call before giving up on that call (with
    exponential backoff between attempts, and a forced reconnect if the
    failure looks auth/connection-related). Default: 5.

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
    [int]$ThrottleMs         = 100,
    [int]$BatchSize          = 20,
    [int]$MonitorTimeoutSec  = 300,
    [int]$GraphClientTimeoutSec = 100,
    [int]$RequestTimeoutSec = 120,
    [int]$TokenRefreshIntervalMin = 45,
    [int]$MaxGraphRetries = 5,
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

# ── Keep-alive: bounded-time Graph calls on a disposable background runspace ──
# On multi-hour runs a single Invoke-MgGraphRequest call can hang forever with
# no error and no progress after a network blip, laptop sleep/wake, VPN
# reconnect, or a silently-dropped idle TCP connection. Every Graph call is
# routed through a background runspace with a hard wall-clock timeout; if it
# doesn't return in time we abort it, throw the stuck runspace away, and let
# the retry loop reconnect and try again - so the script always keeps moving
# instead of hanging.
$script:LastConnectTime = $null
$script:GraphRunspace   = $null

function New-GraphRunspace {
    if ($script:GraphRunspace) {
        try { $script:GraphRunspace.Close() } catch {}
        try { $script:GraphRunspace.Dispose() } catch {}
    }
    $rs = [runspacefactory]::CreateRunspace()
    $rs.Open()
    $init = [PowerShell]::Create()
    $init.Runspace = $rs
    [void]$init.AddScript('Import-Module Microsoft.Graph.Authentication -ErrorAction Stop')
    $init.Invoke() | Out-Null
    $init.Dispose()
    $script:GraphRunspace = $rs
}

function Invoke-WithHardTimeout {
    param([scriptblock]$ScriptBlock, [hashtable]$Params, [int]$TimeoutSec)
    if (-not $script:GraphRunspace -or $script:GraphRunspace.RunspaceStateInfo.State -ne 'Opened') {
        New-GraphRunspace
    }
    $ps = [PowerShell]::Create()
    $ps.Runspace = $script:GraphRunspace
    [void]$ps.AddScript($ScriptBlock).AddParameters($Params)
    $async = $ps.BeginInvoke()
    if (-not $async.AsyncWaitHandle.WaitOne([TimeSpan]::FromSeconds($TimeoutSec))) {
        try { $ps.Stop() } catch {}
        $ps.Dispose()
        New-GraphRunspace   # the old runspace may be wedged - discard it, start clean
        throw "Timed out after $TimeoutSec sec waiting for a response."
    }
    try {
        $out = $ps.EndInvoke($async)
        if ($ps.HadErrors -and $ps.Streams.Error.Count -gt 0) { throw $ps.Streams.Error[0].Exception }
        return $out
    } finally {
        $ps.Dispose()
    }
}

# Runs a Graph call with: proactive token/connection refresh, a hard per-call
# timeout, and retry-with-backoff (reconnecting first if the failure looks
# auth/connection-related). Returns @{ Result = ...; Headers = ... }.
function Invoke-GraphRequestSafe {
    param(
        [string]$Method = "GET",
        [string]$Uri,
        [string]$Body = $null,
        [string]$ContentType = "application/json",
        [switch]$NeedHeaders
    )
    if (-not $script:LastConnectTime -or ((Get-Date) - $script:LastConnectTime).TotalMinutes -ge $TokenRefreshIntervalMin) {
        Write-Log "Keep-alive: refreshing Graph connection (every $TokenRefreshIntervalMin min) ..."
        Connect-GraphAppOnly
    }

    $scriptBlock = {
        param($Method, $Uri, $Body, $ContentType, $NeedHeaders)
        $rh = $null
        if ($NeedHeaders) {
            $r = Invoke-MgGraphRequest -Method $Method -Uri $Uri -Body $Body -ContentType $ContentType -ResponseHeadersVariable rh -ErrorAction Stop
        } else {
            $r = Invoke-MgGraphRequest -Method $Method -Uri $Uri -Body $Body -ContentType $ContentType -ErrorAction Stop
        }
        [PSCustomObject]@{ Result = $r; Headers = $rh }
    }
    $params = @{ Method = $Method; Uri = $Uri; Body = $Body; ContentType = $ContentType; NeedHeaders = [bool]$NeedHeaders }

    $attempt = 0
    $delay = 2
    while ($true) {
        $attempt++
        try {
            $out = Invoke-WithHardTimeout -ScriptBlock $scriptBlock -Params $params -TimeoutSec $RequestTimeoutSec
            return $out[0]
        } catch {
            $msg = $_.Exception.Message
            # Permanent/expected errors - retrying wastes time (e.g. 30s of backoff)
            # for an outcome that will never change, most commonly hit on re-runs
            # where the file already exists at the destination. Fail fast instead.
            if ($msg -match "nameAlreadyExists|invalidRequest|accessDenied|Forbidden|itemNotFound|resourceNotFound|malwareDetected|BadRequest|notAllowed") {
                throw
            }
            if ($attempt -ge $MaxGraphRetries) { throw }
            Write-Log "Graph call attempt $attempt/$MaxGraphRetries failed ($Method $Uri): $msg. Retrying in $delay sec ..." "WARN"
            if ($msg -match "Timed out|Unauthorized|401|expired|token|Could not establish|SSL|connection") {
                try { Connect-GraphAppOnly } catch { Write-Log "Reconnect attempt failed: $($_.Exception.Message)" "WARN" }
            }
            Start-Sleep -Seconds $delay
            $delay = [Math]::Min($delay * 2, 60)
        }
    }
}

# Polls a copy-job monitor URL with the same hard-timeout protection. The
# monitor URL is short-lived/unauthenticated per Microsoft's docs, but some
# tenants require the bearer token anyway - falls back to Invoke-MgGraphRequest.
function Get-CopyJobStatus {
    param([string]$MonitorUrl)
    $scriptBlock = {
        param($MonitorUrl)
        try { Invoke-RestMethod -Uri $MonitorUrl -Method GET -ErrorAction Stop }
        catch { Invoke-MgGraphRequest -Method GET -Uri $MonitorUrl -ErrorAction Stop }
    }
    $out = Invoke-WithHardTimeout -ScriptBlock $scriptBlock -Params @{ MonitorUrl = $MonitorUrl } -TimeoutSec ([Math]::Min(30, $RequestTimeoutSec))
    return $out[0]
}

# Fallback completion check used when the copy-job monitor URL itself can't
# be reached (some tenants return a Location header on the tenant's own
# *-my.sharepoint.com/*.sharepoint.com host, which many corporate networks
# don't allow/resolve even though graph.microsoft.com works fine). Checks
# directly via Graph whether the file now exists at the destination.
function Test-DriveItemExistsByName {
    param([string]$DriveId, [string]$ParentId, [string]$Name)
    $encoded = [Uri]::EscapeDataString($Name)
    try {
        (Invoke-GraphRequestSafe -Method GET -Uri "/v1.0/drives/$DriveId/items/${ParentId}:/$encoded") | Out-Null
        return $true
    } catch {
        if ($_.Exception.Message -match "itemNotFound|404|NotFound|Not Found") { return $false }
        throw
    }
}

# ── Graph app-only connection ────────────────────────────────────────────────
function Connect-GraphAppOnly {
    Assert-Module "Microsoft.Graph.Authentication"
    # Prompt for the secret at most once per run; cache it so the periodic
    # keep-alive refresh and any error-triggered reconnect reuse it instead
    # of prompting again every time.
    if (-not $script:CachedClientSecret) {
        $script:CachedClientSecret = $ClientSecret
        if (-not $script:CachedClientSecret) {
            $script:CachedClientSecret = Read-Host -Prompt "Client secret for app $ClientId" -AsSecureString
        }
    }
    $credential = New-Object System.Management.Automation.PSCredential ($ClientId, $script:CachedClientSecret)
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
    $script:LastConnectTime = Get-Date
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
    $site = (Invoke-GraphRequestSafe -Method GET -Uri "/v1.0/sites/${hostName}:/${sitePath}").Result
    $drivesResp = (Invoke-GraphRequestSafe -Method GET -Uri "/v1.0/sites/$($site.id)/drives").Result
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
            $page = (Invoke-GraphRequestSafe -Method GET -Uri $childrenUri).Result
            $match = $page.value | Where-Object { $_.folder -and $_.name -eq $seg } | Select-Object -First 1
            if ($match) { $existingId = $match.id; break }
            $childrenUri = $page.'@odata.nextLink'
        }
        if ($existingId) {
            $parentId = $existingId
        } else {
            $body = @{ name = $seg; folder = @{}; "@microsoft.graph.conflictBehavior" = "fail" } | ConvertTo-Json
            try {
                $created = (Invoke-GraphRequestSafe -Method POST -Uri "/v1.0/drives/$DriveId/items/$parentId/children" -Body $body -ContentType "application/json").Result
                $parentId = $created.id
            } catch {
                # Race: created by a concurrent run between our check and create - re-fetch it.
                $page = (Invoke-GraphRequestSafe -Method GET -Uri "/v1.0/drives/$DriveId/items/$parentId/children?`$top=500&`$select=id,name,folder").Result
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
    $uri = "/v1.0/drives/$DriveId/items/$ItemId/children?`$top=999"
    while ($uri) {
        if ($MaxItemsPerUser -gt 0 -and $Count.Value -ge $MaxItemsPerUser) { break }
        $page = (Invoke-GraphRequestSafe -Method GET -Uri $uri).Result
        foreach ($item in $page.value) {
            if ($MaxItemsPerUser -gt 0 -and $Count.Value -ge $MaxItemsPerUser) { break }
            if ($item.folder) {
                $sub = Get-DriveFilesRecursive -DriveId $DriveId -ItemId $item.id -RelativeFolder (Join-RelPath $RelativeFolder $item.name) -Count $Count
                foreach ($s in $sub) { $results.Add($s) }
            } elseif ($item.file) {
                $results.Add([PSCustomObject]@{ Id = $item.id; Name = $item.name; RelativeFolder = $RelativeFolder; Size = [int64]$item.size })
                $Count.Value++
            }
        }
        $uri = $page.'@odata.nextLink'
    }
    return $results
}

# ── Fire a server-side copy; returns a monitor URL to poll, or $null if the
#    copy already completed synchronously (some tenants do this for small
#    files). Does NOT wait for completion - callers pipeline many of these
#    in a batch, then poll them all together (see Phase 2 below). ───────────
function Start-DriveItemCopy {
    param([string]$SourceDriveId, [string]$ItemId, [string]$DestDriveId, [string]$DestFolderId, [string]$Name)

    $body = @{ parentReference = @{ driveId = $DestDriveId; id = $DestFolderId }; name = $Name } | ConvertTo-Json
    $resp = Invoke-GraphRequestSafe -Method POST -Uri "/v1.0/drives/$SourceDriveId/items/$ItemId/copy" `
        -Body $body -ContentType "application/json" -NeedHeaders

    $location = $resp.Headers.Location
    if ($location -is [array]) { $location = $location[0] }
    return $location
}

# ── Format a TimeSpan as d.hh:mm:ss (days omitted when zero) ─────────────────
function Format-Duration {
    param([TimeSpan]$Span)
    if ($Span.TotalDays -ge 1) { return $Span.ToString("d\.hh\:mm\:ss") }
    return $Span.ToString("hh\:mm\:ss")
}

# ── Format a byte count as GB/MB for human-readable progress on large backups ─
function Format-Bytes {
    param([long]$Bytes)
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N1} MB" -f ($Bytes / 1MB) }
    return "$Bytes B"
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
        $drive = (Invoke-GraphRequestSafe -Method GET -Uri "/v1.0/users/$upn/drive").Result
    } catch {
        Write-Log "[$upn] No OneDrive found, skipped: $($_.Exception.Message)" "WARN"
        $results.Add([PSCustomObject]@{
            UserPrincipalName = $upn; FolderName = $target.FolderName
            File = ""; RelativeFolder = ""; Size = 0; Status = "Skipped"; Detail = "No OneDrive: $($_.Exception.Message)"
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
            File = ""; RelativeFolder = ""; Size = 0; Status = "Failed"; Detail = "Enumerate: $($_.Exception.Message)"
        })
        continue
    }
    Write-Log "[$upn] $($files.Count) file(s) found."
    foreach ($f in $files) {
        $work.Add([PSCustomObject]@{
            UserPrincipalName = $upn; FolderName = $target.FolderName; SourceDriveId = $drive.id
            FileId = $f.Id; FileName = $f.Name; RelativeFolder = $f.RelativeFolder; Size = $f.Size
        })
    }
}
Write-Progress -Id 1 -Activity "Discovering OneDrive content" -Completed

$totalWork  = $work.Count
$totalBytes = ($work | Measure-Object -Property Size -Sum).Sum
if (-not $totalBytes) { $totalBytes = 0 }
Write-Log "Phase 1/2 complete: $totalWork file(s), $(Format-Bytes $totalBytes) total, to process across $($targets.Count) user(s)." "SUCCESS"

# ── Phase 2: copy (or report, in dry run) files using a sliding window of up
#    to -BatchSize concurrent copy jobs. As soon as ANY job finishes, its slot
#    is immediately refilled with the next queued file - unlike a fixed batch
#    (fire N, wait for ALL N, then fire the next N), a single large straggler
#    file no longer blocks dispatch of the rest of the queue behind it. This
#    is the main lever for large (multi-GB/user) backups. An overall status
#    bar shows % complete, GB copied, files pending, elapsed and ETA (by
#    bytes, not just file count, since a few huge files can dominate the
#    total far more than the item count suggests).
#    Folder lookups are cached per user/sub-folder so repeat files in the
#    same folder don't re-walk the destination tree. ──────────────────────────
$userFolderCache = @{}
$subFolderCache  = @{}
$sw = [Diagnostics.Stopwatch]::StartNew()
$done = 0; $ok = 0; $skip = 0; $fail = 0; $doneBytes = 0
$perUserOk = @{}; $perUserSkip = @{}; $perUserFail = @{}

function Write-BackupProgress {
    param([string]$CurrentLabel = "")
    $pending = $totalWork - $done
    $pendingBytes = [Math]::Max(0, $totalBytes - $doneBytes)
    $pct = if ($totalWork -gt 0) { [Math]::Round(($done / $totalWork) * 100, 1) } else { 100 }
    $avgBytesPerSec = if ($doneBytes -gt 0) { $doneBytes / $sw.Elapsed.TotalSeconds } else { 0 }
    $etaStr = if ($done -lt 3 -or $avgBytesPerSec -le 0) { "calculating..." } else { Format-Duration ([TimeSpan]::FromSeconds([Math]::Max(0, $pendingBytes / $avgBytesPerSec))) }
    $status = "$done/$totalWork done ($pct%) | $(Format-Bytes $doneBytes)/$(Format-Bytes $totalBytes) | Pending: $pending | Elapsed: $(Format-Duration $sw.Elapsed) | ETA: $etaStr"
    Write-Progress -Id 2 -Activity "OneDrive Backup" -Status $status -PercentComplete $pct -CurrentOperation $CurrentLabel
    if ($done % 25 -eq 0 -or $done -eq $totalWork) { Write-Log "Progress: $status" }
}

function Add-BackupResult {
    param($Item, [string]$Status, [string]$Detail)
    $upn = $Item.UserPrincipalName
    $done++
    if ($Status -in @("Success", "Skipped")) { $doneBytes += $Item.Size }
    switch ($Status) {
        "Success" { $ok++;   $perUserOk[$upn]   = ($perUserOk[$upn] + 1) }
        "Skipped" { $skip++; $perUserSkip[$upn] = ($perUserSkip[$upn] + 1) }
        "Failed"  { $fail++; $perUserFail[$upn] = ($perUserFail[$upn] + 1) }
    }
    $results.Add([PSCustomObject]@{
        UserPrincipalName = $upn; FolderName = $Item.FolderName
        File = $Item.FileName; RelativeFolder = $Item.RelativeFolder; Size = $Item.Size
        Status = $Status; Detail = $Detail
    })
    Write-BackupProgress -CurrentLabel "$upn`: $($Item.FileName)"
}

$nextIndex = 0
$inFlight  = New-Object System.Collections.Generic.List[object]

# Sliding window: keep dispatching new copy jobs as long as there is room in
# the window and work left, then poll everything currently in flight. Any
# entry that finishes frees its slot immediately on the next iteration -
# jobs are never held up waiting for the rest of a fixed batch to finish.
while ($nextIndex -lt $totalWork -or $inFlight.Count -gt 0) {
    while ($inFlight.Count -lt $BatchSize -and $nextIndex -lt $totalWork) {
        $item = $work[$nextIndex]
        $nextIndex++
        $upn  = $item.UserPrincipalName

        if (-not $Apply) {
            Add-BackupResult -Item $item -Status "WhatIf" -Detail "Would copy to $DestinationLibrary/$($item.FolderName)/$($item.RelativeFolder)"
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
            Add-BackupResult -Item $item -Status "Failed" -Detail "CreateFolder: $($_.Exception.Message)"
            continue
        }

        try {
            $monitorUrl = Start-DriveItemCopy -SourceDriveId $item.SourceDriveId -ItemId $item.FileId -DestDriveId $destDriveId -DestFolderId $targetFolderId -Name $item.FileName
            if (-not $monitorUrl) {
                Add-BackupResult -Item $item -Status "Success" -Detail ""
            } else {
                $inFlight.Add([PSCustomObject]@{ Item = $item; MonitorUrl = $monitorUrl; DestDriveId = $destDriveId; DestFolderId = $targetFolderId; Started = [Diagnostics.Stopwatch]::StartNew(); LastLogSec = 0; PollErrors = 0; UseExistenceCheck = $false })
            }
        } catch {
            $msg = $_.Exception.Message
            if ($msg -like "*nameAlreadyExists*") {
                Add-BackupResult -Item $item -Status "Skipped" -Detail "Already exists at destination"
            } else {
                Write-Log "[$upn] $($item.FileName): copy failed: $msg" "ERROR"
                Add-BackupResult -Item $item -Status "Failed" -Detail $msg
            }
        }

        if ($ThrottleMs -gt 0) { Start-Sleep -Milliseconds $ThrottleMs }
    }

    if ($inFlight.Count -eq 0) { continue }

    # ── Poll phase: check every in-flight copy job together, so the wait is
    #    shared across the whole window instead of paid one file at a time. ──
    Start-Sleep -Milliseconds 1000
    $stillPending = New-Object System.Collections.Generic.List[object]
    foreach ($entry in $inFlight) {
        $item = $entry.Item
        if ($entry.Started.Elapsed.TotalSeconds -ge $MonitorTimeoutSec) {
            Write-Log "[$($item.UserPrincipalName)] $($item.FileName): copy job timed out after $MonitorTimeoutSec sec" "ERROR"
            Add-BackupResult -Item $item -Status "Failed" -Detail "Copy job timed out after $MonitorTimeoutSec sec"
            continue
        }
        if ($entry.UseExistenceCheck) {
            # Monitor URL host is unreachable from this network - fall back to
            # asking Graph directly whether the file now exists at the
            # destination (same graph.microsoft.com endpoint as everything else).
            try {
                if (Test-DriveItemExistsByName -DriveId $entry.DestDriveId -ParentId $entry.DestFolderId -Name $item.FileName) {
                    Add-BackupResult -Item $item -Status "Success" -Detail ""
                } else {
                    $stillPending.Add($entry)
                }
            } catch {
                $entry.PollErrors++
                if ($entry.PollErrors -eq 1 -or $entry.PollErrors % 10 -eq 0) {
                    Write-Log "[$($item.UserPrincipalName)] $($item.FileName): existence check failed: $($_.Exception.Message)" "WARN"
                }
                $stillPending.Add($entry)
            }
            continue
        }

        $status = $null
        try { $status = Get-CopyJobStatus -MonitorUrl $entry.MonitorUrl }
        catch {
            $entry.PollErrors++
            # Log the actual error periodically instead of silently retrying
            # forever - a copy that only ever fails to POLL (vs. one that is
            # genuinely still running on Graph's side) looks identical in the
            # final "timed out" message otherwise, and this is the difference
            # between a slow copy and a broken monitor URL/permission issue.
            if ($entry.PollErrors -eq 1 -or $entry.PollErrors % 10 -eq 0) {
                Write-Log "[$($item.UserPrincipalName)] $($item.FileName): poll attempt $($entry.PollErrors) failed: $($_.Exception.Message)" "WARN"
            }
            # A handful of failures in a row on the SAME host almost always
            # means the monitor URL's host is unreachable from this network
            # (DNS/firewall), not a transient blip - switch to the existence
            # check instead of retrying an endpoint that will never answer.
            if ($entry.PollErrors -eq 3) {
                Write-Log "[$($item.UserPrincipalName)] $($item.FileName): monitor URL unreachable after 3 attempts, switching to destination existence check" "WARN"
                $entry.UseExistenceCheck = $true
            }
            $stillPending.Add($entry); continue
        }
        if ($status.status -eq "completed") {
            Add-BackupResult -Item $item -Status "Success" -Detail ""
        } elseif ($status.status -in @("failed", "cannotConvert", "malwareDetected")) {
            Write-Log "[$($item.UserPrincipalName)] $($item.FileName): copy failed: $($status.status) $($status.statusDescription)" "ERROR"
            Add-BackupResult -Item $item -Status "Failed" -Detail "$($status.status): $($status.statusDescription)"
        } else {
            $elapsedSec = [int]$entry.Started.Elapsed.TotalSeconds
            if ($elapsedSec -ge $entry.LastLogSec + 30) {
                $entry.LastLogSec = $elapsedSec
                Write-Log "[$($item.UserPrincipalName)] $($item.FileName): copy still in progress after ${elapsedSec}s (status: $($status.status), $($status.percentageComplete)% complete)" "INFO"
            }
            $stillPending.Add($entry)
        }
    }
    $inFlight = $stillPending
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
