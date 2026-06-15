<#
.SYNOPSIS
  Enumerate Google Workspace Tasks per user via GAM and correlate to the
  originating surface (Docs / Chat Space) to derive creator + assignee.

.DESCRIPTION
  The Google Tasks API does not expose a 'creator' or 'assignee' user field on
  a Task. This script derives them as follows:
    - Assignee       = the user whose Tasks list holds the task (GAM context).
    - Creator (Docs) = owner of the Drive file referenced by assignmentInfo.
    - Creator (Chat) = not exposed by the API; the Chat space name/link is
                       reported instead and Creator fields are left blank.
    - Self-created   = no assignmentInfo => creator == assignee.

.PARAMETER Users
  One or more primaryEmail addresses, OR the literal string 'all' to enumerate
  every active user via GAM, OR a path to a CSV containing a 'primaryEmail'
  column.

.PARAMETER OutputCsv
  Output CSV path. Default: .\GoogleTasks_WithCreator.csv

.PARAMETER IncludeCompleted
.PARAMETER IncludeHidden
.PARAMETER IncludeDeleted
  Pass-through flags to GAM 'print tasks'. Default: only needsAction tasks.

.PARAMETER GamPath
  Path to the gam executable. Default: 'gam' (must be on PATH).

.EXAMPLE
  .\Get-GoogleTasksWithCreator.ps1 -Users user1@contoso.com,user2@contoso.com

.EXAMPLE
  .\Get-GoogleTasksWithCreator.ps1 -Users all -IncludeCompleted -OutputCsv out.csv
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string[]]$Users,
    [string]$OutputCsv = (Join-Path (Get-Location) 'GoogleTasks_WithCreator.csv'),
    [string]$SummaryCsv,
    [string]$UsersFile,
    [int]$MaxUsers = 0,
    [string]$CheckpointCsv,
    [int]$RestParallel = 8,
    [switch]$IncludeCompleted,
    [switch]$IncludeHidden,
    [switch]$IncludeDeleted,
    [switch]$SkipTenantSpaceScan,
    [switch]$SkipDocCommentScan,
    [switch]$ForceDocCommentScan,
    [switch]$IncludeResolvedDocComments,
    [string]$GamPath = 'gam',
    [string]$KeyFile = ''
)
if (-not $SummaryCsv) {
    $SummaryCsv = [IO.Path]::ChangeExtension($OutputCsv, $null).TrimEnd('.') + '_Summary.csv'
}
# -UsersFile overrides -Users when provided (plain-text: one email per line,
# or CSV with primaryEmail column). Keeps the Mandatory -Users requirement
# satisfied for safety while letting bulk runs pass a file.
if ($UsersFile) {
    if (-not (Test-Path $UsersFile)) { throw "UsersFile not found: $UsersFile" }
    $Users = @($UsersFile)
}

# Native executables (gam.exe) emit progress to stderr; under 'Stop' that is
# treated as a terminating error by PowerShell 5.x. Use 'Continue' at script
# scope and rely on explicit $LASTEXITCODE checks for GAM failures.
$ErrorActionPreference = 'Continue'

# --- Pre-flight: verify GAM executable is reachable ---
$resolvedGam = $null
try {
    $cmd = Get-Command -Name $GamPath -ErrorAction Stop
    $resolvedGam = $cmd.Source
}
catch {
    Write-Error @"
Cannot find GAM executable '$GamPath'.
Either:
  1. Install GAM7 (https://github.com/GAM-team/GAM/wiki/How-to-Install-Advanced-GAM) and ensure 'gam.exe' is on PATH, OR
  2. Re-run this script with -GamPath pointing at your existing install, e.g.
       .\Get-GoogleTasksWithCreator.ps1 -Users ... -GamPath 'C:\GAM7\gam.exe'
"@
    exit 2
}
$GamPath = $resolvedGam
Write-Host ("Using GAM: {0}" -f $GamPath) -ForegroundColor DarkGray

# --- Locate pwsh (PowerShell 7+) for the REST fallback used to fetch
#     Docs/Chat-assigned tasks (needs RSA.ImportFromPem + showAssigned=true).
$PwshPath = $null
$PwshHelper = Join-Path $PSScriptRoot '_Get-AssignedTasks.ps1'
try {
    $cmd = Get-Command pwsh -ErrorAction Stop
    $PwshPath = $cmd.Source
    if (-not (Test-Path $PwshHelper)) { $PwshPath = $null }
}
catch { $PwshPath = $null }
if ($PwshPath) {
    Write-Host ("REST fallback enabled via: {0}" -f $PwshPath) -ForegroundColor DarkGray
}
else {
    Write-Warning "pwsh (PowerShell 7+) or _Get-AssignedTasks.ps1 not available; Docs/Chat-assigned tasks will not be fetched."
}

function Invoke-Gam {
    param(
        [Parameter(Mandatory)][string[]]$GamArgs,
        [int[]]$SuppressExitCodes = @(60)
    )
    $errFile = [IO.Path]::GetTempFileName()
    $prevEAP = $ErrorActionPreference
    $ErrorActionPreference = 'SilentlyContinue'
    try {
        $out = & $GamPath @GamArgs 2>$errFile
        $code = $LASTEXITCODE
        if ($code -ne 0) {
            if ($SuppressExitCodes -contains $code) { return $null }
            $err = (Get-Content -Raw -LiteralPath $errFile -ErrorAction SilentlyContinue)
            Write-Warning ("GAM failed ({0}): {1} :: {2}" -f $code, ($GamArgs -join ' '), $err)
            return $null
        }
        return $out
    }
    finally {
        $ErrorActionPreference = $prevEAP
        Remove-Item $errFile -Force -ErrorAction SilentlyContinue
    }
}

function Resolve-UserList {
    param([string[]]$In)
    if ($In.Count -eq 1) {
        $only = $In[0]
        if ($only -eq 'all') {
            $tmp = [IO.Path]::GetTempFileName()
            $out = Invoke-Gam @('print', 'users', 'fields', 'primaryEmail,suspended')
            if (-not $out) { return @() }
            $out | Out-File -FilePath $tmp -Encoding UTF8
            $rows = Import-Csv $tmp
            Remove-Item $tmp -Force
            return $rows | Where-Object { $_.suspended -ne 'True' } | Select-Object -ExpandProperty primaryEmail
        }
        if (Test-Path $only) {
            $head = Get-Content -LiteralPath $only -TotalCount 1 -ErrorAction SilentlyContinue
            if ($head -and $head -match '(?i)primaryEmail') {
                return (Import-Csv $only).primaryEmail | Where-Object { $_ }
            }
            return Get-Content -LiteralPath $only | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '@' }
        }
    }
    return $In
}

$driveOwnerCache = @{}
function Get-DriveFileOwner {
    param([string]$User, [string]$FileId)
    if (-not $FileId) { return $null }
    if ($driveOwnerCache.ContainsKey($FileId)) { return $driveOwnerCache[$FileId] }
    $raw = Invoke-Gam @('user', $User, 'show', 'fileinfo', $FileId, 'fields', 'owners,name')
    $ownerEmail = $null; $ownerName = $null; $fileName = $null
    if ($raw) {
        foreach ($line in $raw) {
            if ($line -match '^\s*name:\s*(.+)$') { $fileName = $Matches[1].Trim() }
            if ($line -match '^\s*emailAddress:\s*(.+)$') { if (-not $ownerEmail) { $ownerEmail = $Matches[1].Trim() } }
            if ($line -match '^\s*displayName:\s*(.+)$') { if (-not $ownerName) { $ownerName = $Matches[1].Trim() } }
        }
    }
    $result = [pscustomobject]@{ OwnerEmail = $ownerEmail; OwnerName = $ownerName; FileName = $fileName }
    $driveOwnerCache[$FileId] = $result
    return $result
}

$spaceCache = @{}
function Get-ChatSpaceInfo {
    param([string]$User, [string]$Space)
    if (-not $Space) { return $null }
    if ($spaceCache.ContainsKey($Space)) { return $spaceCache[$Space] }
    $raw = Invoke-Gam @('user', $User, 'info', 'chatspace', $Space)
    $displayName = $null
    if ($raw) {
        foreach ($line in $raw) {
            if ($line -match '^\s*displayName:\s*(.+)$') { $displayName = $Matches[1].Trim(); break }
        }
    }
    $result = [pscustomobject]@{ SpaceName = $Space; DisplayName = $displayName }
    $spaceCache[$Space] = $result
    return $result
}

$chatMessagesCache = @{}
function Get-SpaceTaskCreator {
    param(
        [string]$User,
        [string]$Space,
        [string]$TaskUpdatedIso,
        [int]$ToleranceSeconds = 10
    )
    if (-not $Space -or -not $TaskUpdatedIso) { return $null }

    if (-not $chatMessagesCache.ContainsKey($Space)) {
        $raw = Invoke-Gam @('user', $User, 'print', 'chatmessages', 'space', $Space)
        $msgs = @()
        if ($raw) {
            $tmp = [IO.Path]::GetTempFileName()
            try {
                $raw | Out-File -FilePath $tmp -Encoding UTF8
                $rows = Import-Csv -Path $tmp -ErrorAction SilentlyContinue
                foreach ($m in $rows) {
                    if ($m.'sender.type' -ne 'HUMAN') { continue }
                    if ($m.argumentText -notmatch '^(Created|Assigned) a task') { continue }
                    $msgs += $m
                }
            }
            finally { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
        }
        $chatMessagesCache[$Space] = $msgs
    }
    $msgs = $chatMessagesCache[$Space]
    if (-not $msgs -or $msgs.Count -eq 0) { return $null }

    $target = $null
    try { $target = [DateTimeOffset]::Parse($TaskUpdatedIso).UtcDateTime } catch { return $null }

    $bestEmail = $null
    $bestDelta = [double]::MaxValue
    foreach ($m in $msgs) {
        $t = $null
        try { $t = [DateTimeOffset]::Parse($m.'createTime').UtcDateTime } catch { continue }
        $delta = [Math]::Abs(($t - $target).TotalSeconds)
        if ($delta -le $ToleranceSeconds -and $delta -lt $bestDelta) {
            $bestDelta = $delta
            $bestEmail = $m.'sender.email'
        }
    }
    return $bestEmail
}

function ConvertFrom-GamCsv {
    param([string[]]$Raw)
    if (-not $Raw) { return @() }
    $tmp = [IO.Path]::GetTempFileName()
    try {
        $Raw | Out-File -FilePath $tmp -Encoding UTF8
        return , (Import-Csv -Path $tmp -ErrorAction SilentlyContinue)
    }
    finally { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
}

function Get-TenantChatSpaces {
    param([string]$ImpersonationUser)
    if (-not $ImpersonationUser) { return @() }
    $raw = Invoke-Gam @('user', $ImpersonationUser, 'print', 'chatspaces')
    $rows = ConvertFrom-GamCsv -Raw $raw
    $out = @()
    foreach ($r in $rows) {
        $id = $r.name
        if ($id) {
            $out += [pscustomobject]@{
                Id          = $id
                DisplayName = $r.displayName
                MemberCount = $r.'membershipCount.joinedDirectHumanUserCount'
            }
        }
    }
    return $out
}

$spaceMembersCache = @{}
function Get-SpaceHumanMember {
    param([string]$ImpersonationUser, [string]$Space)
    if (-not $Space) { return $null }
    if ($spaceMembersCache.ContainsKey($Space)) { return $spaceMembersCache[$Space] }
    $raw = Invoke-Gam @('user', $ImpersonationUser, 'print', 'chatmembers', 'space', $Space)
    $rows = ConvertFrom-GamCsv -Raw $raw
    $first = $null
    foreach ($r in $rows) {
        if ($r.'member.type' -eq 'HUMAN' -and $r.'member.email') { $first = $r.'member.email'; break }
    }
    $spaceMembersCache[$Space] = $first
    return $first
}

function Get-SpaceAllMessages {
    param([string]$ImpersonationUser, [string]$Space)
    if (-not $Space) { return @() }
    if ($chatMessagesCache.ContainsKey("__ALL__::$Space")) { return $chatMessagesCache["__ALL__::$Space"] }
    $raw = Invoke-Gam @('user', $ImpersonationUser, 'print', 'chatmessages', 'space', $Space)
    $rows = ConvertFrom-GamCsv -Raw $raw
    $msgs = @()
    foreach ($r in $rows) {
        if ($r.'sender.type' -ne 'HUMAN') { continue }
        if ($r.argumentText -notmatch '(Created|Assigned|Unassigned|Completed|Re-opened|Deleted) a task') { continue }
        $msgs += $r
    }
    $chatMessagesCache["__ALL__::$Space"] = $msgs
    return $msgs
}

function Build-SpaceTaskThreads {
    # Returns a map: thread.name -> @{ Creator; CreateTime; Assignee; Deleted; LastEventTime }
    param([object[]]$Messages)
    $threads = @{}
    foreach ($m in $Messages) {
        $tn = $m.'thread.name'
        if (-not $tn) { continue }
        if (-not $threads.ContainsKey($tn)) {
            $threads[$tn] = [ordered]@{
                Creator         = $null
                CreateTime      = $null
                Assignee        = $null
                AssigneeMention = $null
                Deleted         = $false
                LastEventTime   = $null
                LastEventText   = $null
            }
        }
        $th = $threads[$tn]
        $th.LastEventTime = $m.createTime
        $th.LastEventText = $m.argumentText

        if ($m.argumentText -match '^Created a task( for @(.+?))? \(via Tasks\)') {
            if (-not $th.Creator) {
                $th.Creator = $m.'sender.email'
                $th.CreateTime = $m.createTime
            }
            if ($Matches[2]) { $th.AssigneeMention = $Matches[2].Trim(); $th.Assignee = $m.'annotations.0.userMention.user.name' }
        }
        elseif ($m.argumentText -match '^Assigned a task to @(.+?) \(via Tasks\)') {
            $th.AssigneeMention = $Matches[1].Trim()
            $th.Assignee = $m.'annotations.0.userMention.user.name'
        }
        elseif ($m.argumentText -match '^Unassigned a task') {
            $th.AssigneeMention = $null
            $th.Assignee = $null
        }
        elseif ($m.argumentText -match '^Deleted a task') {
            $th.Deleted = $true
        }
    }
    return $threads
}

$userDocsCache = @{}
function Get-UserOwnedDocs {
    param([string]$User)
    if ($userDocsCache.ContainsKey($User)) { return $userDocsCache[$User] }
    $q = "mimeType = 'application/vnd.google-apps.document' or mimeType = 'application/vnd.google-apps.spreadsheet' or mimeType = 'application/vnd.google-apps.presentation'"
    $raw = Invoke-Gam @('user', $User, 'print', 'filelist', 'query', $q, 'fields', 'id,name,mimeType')
    $rows = ConvertFrom-GamCsv -Raw $raw
    $out = @()
    foreach ($r in $rows) {
        if ($r.id) {
            $out += [pscustomobject]@{ Id = $r.id; Name = $r.name; MimeType = $r.mimeType }
        }
    }
    $userDocsCache[$User] = $out
    return $out
}

function Get-FileAssignedComments {
    param([string]$User, [string]$FileId)
    if (-not $FileId) { return @() }
    $raw = Invoke-Gam @('user', $User, 'print', 'filecomments', $FileId)
    $rows = ConvertFrom-GamCsv -Raw $raw
    $out = @()
    foreach ($r in $rows) {
        if (-not $r.assigneeEmailAddress) { continue }
        if ($r.replyId) { continue }
        if ($r.deleted -eq 'True') { continue }
        $out += $r
    }
    return $out
}

# Tenant-wide bulk prefetch. One GAM process handles many users via 'file' user
# entity; GAM internally parallelises across its worker pool. The per-user
# helpers below read from these maps instead of launching a GAM process each.
$script:tenantTasksByUser = $null      # @{ email-lower -> List[row] }
$script:tenantTasklistsByUser = $null  # @{ email-lower -> @{ listId -> title } }
$script:tenantUserFiles = @()

function Get-TenantUserEntityArgs {
    param([string[]]$Users)
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("gamusers_" + [guid]::NewGuid().ToString('N').Substring(0, 10) + ".txt")
    ($Users -join "`r`n") | Out-File -FilePath $tmp -Encoding ASCII
    $script:tenantUserFiles += $tmp
    return @('file', $tmp)
}

function Get-TenantTasksJsonBulk {
    param([Parameter(Mandatory)][string[]]$Users)
    if (-not $Users -or $Users.Count -eq 0) { return @{} }
    $entity = Get-TenantUserEntityArgs -Users $Users
    $gargs = $entity + @('print', 'tasks', 'formatjson')
    if ($IncludeCompleted) { $gargs += 'showcompleted' }
    if ($IncludeHidden) { $gargs += 'showhidden' }
    if ($IncludeDeleted) { $gargs += 'showdeleted' }
    $raw = Invoke-Gam -GamArgs $gargs
    $rows = ConvertFrom-GamCsv -Raw $raw
    $map = @{}
    foreach ($r in $rows) {
        if (-not $r.User) { continue }
        $u = $r.User.ToLowerInvariant()
        if (-not $map.ContainsKey($u)) { $map[$u] = New-Object System.Collections.Generic.List[object] }
        [void]$map[$u].Add($r)
    }
    return $map
}

function Get-TenantTasklistsBulk {
    param([Parameter(Mandatory)][string[]]$Users)
    if (-not $Users -or $Users.Count -eq 0) { return @{} }
    $entity = Get-TenantUserEntityArgs -Users $Users
    $gargs = $entity + @('print', 'tasklists')
    $raw = Invoke-Gam -GamArgs $gargs
    $rows = ConvertFrom-GamCsv -Raw $raw
    $map = @{}
    foreach ($r in $rows) {
        if (-not $r.User) { continue }
        $u = $r.User.ToLowerInvariant()
        if (-not $map.ContainsKey($u)) { $map[$u] = @{} }
        if ($r.id) { $map[$u][$r.id] = $r.title }
    }
    return $map
}

$tasklistTitleCache = @{}
function Get-UserTasklistMap {
    param([string]$User)
    if ($script:tenantTasklistsByUser) {
        $k = $User.ToLowerInvariant()
        if ($script:tenantTasklistsByUser.ContainsKey($k)) { return $script:tenantTasklistsByUser[$k] }
        return @{}
    }
    if ($tasklistTitleCache.ContainsKey($User)) { return $tasklistTitleCache[$User] }
    $raw = Invoke-Gam @('user', $User, 'print', 'tasklists')
    $rows = ConvertFrom-GamCsv -Raw $raw
    $map = @{}
    foreach ($r in $rows) {
        if ($r.id) { $map[$r.id] = $r.title }
    }
    $tasklistTitleCache[$User] = $map
    return $map
}

function Get-UserTasksJson {
    param([string]$User)
    if ($script:tenantTasksByUser) {
        $k = $User.ToLowerInvariant()
        if ($script:tenantTasksByUser.ContainsKey($k)) { return $script:tenantTasksByUser[$k] }
        return @()
    }
    $gargs = @('user', $User, 'print', 'tasks', 'formatjson')
    if ($IncludeCompleted) { $gargs += 'showcompleted' }
    if ($IncludeHidden) { $gargs += 'showhidden' }
    if ($IncludeDeleted) { $gargs += 'showdeleted' }
    $tmp = [IO.Path]::GetTempFileName()
    $errFile = [IO.Path]::GetTempFileName()
    $prevEAP = $ErrorActionPreference
    $ErrorActionPreference = 'SilentlyContinue'
    try {
        $raw = & $GamPath @gargs 2>$errFile
        $code = $LASTEXITCODE
        if ($code -ne 0) {
            if ($code -ne 60) {
                $err = (Get-Content -Raw -LiteralPath $errFile -ErrorAction SilentlyContinue)
                Write-Warning ("GAM print tasks failed for {0}: {1}" -f $User, $err)
            }
            return @()
        }
        $raw | Out-File -FilePath $tmp -Encoding UTF8
        return Import-Csv -Path $tmp -ErrorAction SilentlyContinue
    }
    finally {
        $ErrorActionPreference = $prevEAP
        Remove-Item $tmp     -Force -ErrorAction SilentlyContinue
        Remove-Item $errFile -Force -ErrorAction SilentlyContinue
    }
}

$script:tenantAssignedByUser = $null  # @{ email-lower -> [entry] }
$restAssignedCache = @{}

function Invoke-TenantAssignedTasksBulk {
    # One pwsh 7 process; feeds the whole user list to the REST helper which
    # runs requests in parallel and emits NDJSON (one JSON object per user).
    param([Parameter(Mandatory)][string[]]$Users)
    if (-not $PwshPath) { return @{} }
    if (-not $Users -or $Users.Count -eq 0) { return @{} }

    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("restusers_" + [guid]::NewGuid().ToString('N').Substring(0, 10) + ".txt")
    ($Users -join "`r`n") | Out-File -FilePath $tmp -Encoding ASCII
    $hargs = @('-NoProfile', '-NoLogo', '-File', $PwshHelper, '-UsersFile', $tmp, '-Parallel', [string]$RestParallel)
    if ($KeyFile) { $hargs += '-KeyFile'; $hargs += $KeyFile }
    if ($IncludeCompleted) { $hargs += '-IncludeCompleted' }
    if ($IncludeHidden) { $hargs += '-IncludeHidden' }
    if ($IncludeDeleted) { $hargs += '-IncludeDeleted' }

    $errFile = [IO.Path]::GetTempFileName()
    $prevEAP = $ErrorActionPreference
    $ErrorActionPreference = 'SilentlyContinue'
    $map = @{}
    try {
        $out = & $PwshPath @hargs 2>$errFile
        if ($LASTEXITCODE -ne 0) {
            $err = (Get-Content -Raw -LiteralPath $errFile -ErrorAction SilentlyContinue)
            Write-Warning ("Bulk REST fallback failed: {0}" -f $err)
            return $map
        }
        foreach ($line in @($out)) {
            if (-not $line -or [string]::IsNullOrWhiteSpace($line)) { continue }
            $obj = $null
            try { $obj = $line | ConvertFrom-Json -ErrorAction Stop } catch { continue }
            if (-not $obj -or -not $obj.user) { continue }
            $k = $obj.user.ToLowerInvariant()
            $map[$k] = if ($obj.entries) { @($obj.entries) } else { @() }
        }
        return $map
    }
    finally {
        $ErrorActionPreference = $prevEAP
        Remove-Item $errFile -Force -ErrorAction SilentlyContinue
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    }
}

function Get-AssignedTasksViaRest {
    param([string]$User)
    if (-not $User) { return @() }
    if ($script:tenantAssignedByUser) {
        $k = $User.ToLowerInvariant()
        if ($script:tenantAssignedByUser.ContainsKey($k)) { return $script:tenantAssignedByUser[$k] }
        return @()
    }
    if ($restAssignedCache.ContainsKey($User)) { return $restAssignedCache[$User] }
    if (-not $PwshPath) { $restAssignedCache[$User] = @(); return @() }
    $hargs = @('-NoProfile', '-NoLogo', '-File', $PwshHelper, '-User', $User)
    if ($KeyFile) { $hargs += '-KeyFile'; $hargs += $KeyFile }
    if ($IncludeCompleted) { $hargs += '-IncludeCompleted' }
    if ($IncludeHidden) { $hargs += '-IncludeHidden' }
    if ($IncludeDeleted) { $hargs += '-IncludeDeleted' }
    $errFile = [IO.Path]::GetTempFileName()
    $prevEAP = $ErrorActionPreference
    $ErrorActionPreference = 'SilentlyContinue'
    try {
        $out = & $PwshPath @hargs 2>$errFile
        if ($LASTEXITCODE -ne 0) {
            $err = (Get-Content -Raw -LiteralPath $errFile -ErrorAction SilentlyContinue)
            Write-Warning ("REST fallback failed for {0}: {1}" -f $User, $err)
            $restAssignedCache[$User] = @()
            return @()
        }
        $joined = ($out -join '')
        if ([string]::IsNullOrWhiteSpace($joined)) { $restAssignedCache[$User] = @(); return @() }
        $obj = $joined | ConvertFrom-Json -ErrorAction SilentlyContinue
        if ($null -eq $obj) { $restAssignedCache[$User] = @(); return @() }
        $arr = if ($obj -is [System.Array]) { $obj } else { @($obj) }
        $restAssignedCache[$User] = $arr
        return $arr
    }
    finally {
        $ErrorActionPreference = $prevEAP
        Remove-Item $errFile -Force -ErrorAction SilentlyContinue
    }
}

function Get-AssigneeDefaultTasklist {
    # Returns the assignee's default ("My Tasks") tasklist as
    # @{ TaskListId; TaskListTitle } if known, else $null. Prefers a list
    # literally named 'My Tasks'; falls back to the first list returned.
    param([string]$Assignee)
    if (-not $Assignee) { return $null }
    $tlMap = Get-UserTasklistMap -User $Assignee
    if (-not $tlMap -or $tlMap.Count -eq 0) { return $null }
    foreach ($id in $tlMap.Keys) {
        if ($tlMap[$id] -eq 'My Tasks') {
            return @{ TaskListId = $id; TaskListTitle = $tlMap[$id] }
        }
    }
    $firstId = ($tlMap.Keys | Select-Object -First 1)
    return @{ TaskListId = $firstId; TaskListTitle = $tlMap[$firstId] }
}

function Get-AssigneeTasklistForDoc {
    # Returns @{ TaskListId; TaskListTitle } for the assignee's view of a
    # Doc-comment assignment. Resolution order:
    #   1. A real Tasks entry whose assignmentInfo/links point at this Doc.
    #   2. The assignee's default tasklist (where a mirrored task would land).
    #   3. $null (caller falls back to the '(Google Docs comment)' label).
    param([string]$Assignee, [string]$DocId)
    if (-not $Assignee -or -not $DocId) { return $null }
    $entries = Get-AssignedTasksViaRest -User $Assignee
    foreach ($e in $entries) {
        $t = $e.task; if (-not $t) { continue }
        $link = $t.assignmentInfo.linkToTask
        if ($link -and $link -like "*$DocId*") {
            return @{ TaskListId = $e.tasklistId; TaskListTitle = $e.tasklistTitle }
        }
        foreach ($l in @($t.links)) {
            if ($l.link -and $l.link -like "*$DocId*") {
                return @{ TaskListId = $e.tasklistId; TaskListTitle = $e.tasklistTitle }
            }
        }
    }
    return Get-AssigneeDefaultTasklist -Assignee $Assignee
}

function Add-TaskRow {
    param(
        [Parameter(Mandatory)][string]$User,
        [Parameter(Mandatory)]$Task,
        [string]$TaskListId,
        [string]$TaskListTitle,
        [string]$Origin
    )
    $ai = $Task.assignmentInfo
    $surface = if ($ai) { $ai.surfaceType } else { 'SELF' }

    $creatorEmail = $null; $creatorName = $null
    $sourceRef = $null; $sourceLink = $null; $notes = $null

    if ($ai) {
        $sourceLink = $ai.linkToTask
        switch ($surface) {
            'DOCUMENT' {
                $fid = $ai.driveResourceInfo.driveFileId
                $sourceRef = $fid
                $info = Get-DriveFileOwner -User $User -FileId $fid
                if ($info) {
                    $creatorEmail = $info.OwnerEmail
                    $creatorName = $info.OwnerName
                    if ($info.FileName) { $notes = "DocName: $($info.FileName)" }
                }
            }
            'SPACE' {
                $sp = $ai.spaceInfo.space
                $sourceRef = $sp
                $info = Get-ChatSpaceInfo -User $User -Space $sp
                $correlated = Get-SpaceTaskCreator -User $User -Space $sp -TaskUpdatedIso $Task.updated
                if ($correlated) { $creatorEmail = $correlated }
                $parts = @()
                if ($info -and $info.DisplayName) { $parts += "SpaceName: $($info.DisplayName)" }
                if ($correlated) {
                    $parts += 'Creator resolved via Chat message correlation.'
                }
                else {
                    $parts += 'Creator identity not exposed by Tasks API for Chat-assigned tasks.'
                }
                $notes = $parts -join ' | '
            }
            default { $notes = "Unhandled surfaceType: $surface" }
        }
    }
    else { $creatorEmail = $User }

    $isShared = [bool]($creatorEmail -and $User -and ($creatorEmail -ne $User))

    $script:results.Add([pscustomobject]@{
            AssigneeEmail = $User
            TaskListId    = $TaskListId
            TaskListTitle = $TaskListTitle
            TaskId        = $Task.id
            Title         = $Task.title
            Status        = $Task.status
            Due           = $Task.due
            Completed     = $Task.completed
            Updated       = $Task.updated
            WebViewLink   = $Task.webViewLink
            SurfaceType   = $surface
            SourceRef     = $sourceRef
            SourceLink    = $sourceLink
            CreatorEmail  = $creatorEmail
            CreatorName   = $creatorName
            Shared        = $isShared
            Origin        = $Origin
            Notes         = $notes
        })
}

$results = New-Object System.Collections.Generic.List[object]
$resolvedUsers = Resolve-UserList -In $Users
if ($MaxUsers -gt 0 -and $resolvedUsers.Count -gt $MaxUsers) {
    Write-Host ("MaxUsers cap in effect: truncating {0} -> {1}" -f $resolvedUsers.Count, $MaxUsers) -ForegroundColor Yellow
    $resolvedUsers = $resolvedUsers | Select-Object -First $MaxUsers
}

# Auto-disable the Drive comment scan on large tenant runs unless explicitly
# forced. Doc-comment scanning invokes one GAM 'print filecomments' per owned
# Drive file, which is the single heaviest path and typically not needed for
# routine tenant-wide audits.
$isAllMode = ($Users.Count -eq 1 -and $Users[0] -eq 'all')
if ($isAllMode -and -not $ForceDocCommentScan -and -not $SkipDocCommentScan) {
    Write-Host "Auto-disabling -DocCommentScan for -Users all run (use -ForceDocCommentScan to override)." -ForegroundColor Yellow
    $SkipDocCommentScan = $true
}

Write-Host ("Processing {0} user(s)..." -f $resolvedUsers.Count) -ForegroundColor Cyan

# Bulk prefetch: one GAM process for tasklists, one for tasks, one pwsh-7
# process for the parallel REST assigned-tasks pass. Replaces N * 3 external
# calls with 3 total, scaling the per-user loop to a pure in-memory walk.
if ($resolvedUsers.Count -gt 0) {
    Write-Host "Prefetching tasklists (bulk)..." -ForegroundColor Cyan
    $t0 = Get-Date
    $script:tenantTasklistsByUser = Get-TenantTasklistsBulk -Users $resolvedUsers
    Write-Host ("  done in {0:N1}s ({1} user maps)" -f ((Get-Date) - $t0).TotalSeconds, $script:tenantTasklistsByUser.Count) -ForegroundColor DarkGray

    Write-Host "Prefetching tasks (bulk, formatjson)..." -ForegroundColor Cyan
    $t0 = Get-Date
    $script:tenantTasksByUser = Get-TenantTasksJsonBulk -Users $resolvedUsers
    Write-Host ("  done in {0:N1}s ({1} users with tasks)" -f ((Get-Date) - $t0).TotalSeconds, $script:tenantTasksByUser.Count) -ForegroundColor DarkGray

    if ($PwshPath) {
        Write-Host ("Prefetching assigned tasks via REST (parallel={0})..." -f $RestParallel) -ForegroundColor Cyan
        $t0 = Get-Date
        $script:tenantAssignedByUser = Invoke-TenantAssignedTasksBulk -Users $resolvedUsers
        Write-Host ("  done in {0:N1}s ({1} users with assigned entries)" -f ((Get-Date) - $t0).TotalSeconds, $script:tenantAssignedByUser.Count) -ForegroundColor DarkGray
    }
}

# Initialise checkpoint CSV with a header row so append writes are well-formed.
if ($CheckpointCsv) {
    'AssigneeEmail,TaskListId,TaskListTitle,TaskId,Title,Status,Due,Completed,Updated,WebViewLink,SurfaceType,SourceRef,SourceLink,CreatorEmail,CreatorName,Shared,Origin,Notes' |
    Out-File -FilePath $CheckpointCsv -Encoding UTF8
    Write-Host ("Checkpoint CSV: {0}" -f $CheckpointCsv) -ForegroundColor DarkGray
}

foreach ($u in $resolvedUsers) {
    Write-Host ("  -> {0}" -f $u) -ForegroundColor Gray
    $seenIds = New-Object 'System.Collections.Generic.HashSet[string]'
    $preCount = $results.Count
    $tlMap = Get-UserTasklistMap -User $u

    # 1) Tasks via GAM CLI (user-created tasks; no assigned-from-Docs/Chat).
    $rows = Get-UserTasksJson -User $u
    if ($rows) {
        foreach ($row in $rows) {
            $jsonCol = $row.PSObject.Properties | Where-Object { $_.Name -match '^JSON' } | Select-Object -First 1
            if (-not $jsonCol) { continue }
            $task = $null
            try { $task = $jsonCol.Value | ConvertFrom-Json -ErrorAction Stop } catch { continue }
            if (-not $task) { continue }
            if ($task.id -and -not $seenIds.Add([string]$task.id)) { continue }
            $tlTitle = if ($row.tasklistId -and $tlMap.ContainsKey($row.tasklistId)) { $tlMap[$row.tasklistId] } else { $null }
            Add-TaskRow -User $u -Task $task -TaskListId $row.tasklistId -TaskListTitle $tlTitle -Origin 'GAM'
        }
    }

    # 2) Assigned-from-Docs/Chat tasks via direct REST (showAssigned=true).
    $assigned = Get-AssignedTasksViaRest -User $u
    foreach ($entry in $assigned) {
        $task = $entry.task
        if (-not $task) { continue }
        if ($task.id -and -not $seenIds.Add([string]$task.id)) { continue }
        $tlTitle = $entry.tasklistTitle
        if (-not $tlTitle -and $entry.tasklistId -and $tlMap.ContainsKey($entry.tasklistId)) { $tlTitle = $tlMap[$entry.tasklistId] }
        Add-TaskRow -User $u -Task $task -TaskListId $entry.tasklistId -TaskListTitle $tlTitle -Origin 'REST'
    }

    # 3) Drive comment scan: Docs/Sheets/Slides where this user is the file
    #    owner and a comment has assigneeEmailAddress set (Docs action items).
    if (-not $SkipDocCommentScan) {
        $apiDocKeys = @{}
        foreach ($r in $results) {
            if ($r.SurfaceType -eq 'DOCUMENT' -and $r.Origin -ne 'DOCSCAN' -and $r.SourceRef -and $r.AssigneeEmail) {
                $apiDocKeys[("{0}|{1}" -f $r.SourceRef, $r.AssigneeEmail)] = $true
            }
        }

        $docs = Get-UserOwnedDocs -User $u
        foreach ($doc in $docs) {
            $comments = Get-FileAssignedComments -User $u -FileId $doc.Id
            foreach ($c in $comments) {
                if (-not $IncludeResolvedDocComments -and $c.resolved -eq 'True') { continue }
                $dupKey = "{0}|{1}" -f $doc.Id, $c.assigneeEmailAddress
                if ($apiDocKeys.ContainsKey($dupKey)) { continue }
                $cid = '{0}/{1}' -f $doc.Id, $c.commentId
                if (-not $seenIds.Add($cid)) { continue }

                $creator = if ($c.'author.me' -eq 'True') { $u } else { $null }
                $creatorName = $c.'author.displayName'
                $title = if ($c.'quotedFileContent.value') { $c.'quotedFileContent.value' } else { $c.content }
                $status = if ($c.resolved -eq 'True') { 'completed' } else { 'needsAction' }
                $isShared = [bool]($creator -and $c.assigneeEmailAddress -and ($creator -ne $c.assigneeEmailAddress))
                $noteParts = @()
                if ($doc.Name) { $noteParts += "DocName: $($doc.Name)" }
                $noteParts += 'Assigned via Docs comment (discovered by Drive comment scan).'
                if (-not $creator) { $noteParts += "AuthorName: $creatorName (email not exposed by Drive API)" }
                if ($c.content) { $noteParts += "CommentText: $($c.content)" }

                $link = "https://docs.google.com/document/d/$($doc.Id)/edit?disco=$($c.commentId)"

                $aTlId = $null; $aTlTitle = $null
                $mirror = Get-AssigneeTasklistForDoc -Assignee $c.assigneeEmailAddress -DocId $doc.Id
                if ($mirror) { $aTlId = $mirror.TaskListId; $aTlTitle = $mirror.TaskListTitle }
                else { $aTlTitle = '(Google Docs comment)' }

                $results.Add([pscustomobject]@{
                        AssigneeEmail = $c.assigneeEmailAddress
                        TaskListId    = $aTlId
                        TaskListTitle = $aTlTitle
                        TaskId        = $cid
                        Title         = $title
                        Status        = $status
                        Due           = $null
                        Completed     = if ($c.resolved -eq 'True') { $c.modifiedTime } else { $null }
                        Updated       = $c.modifiedTime
                        WebViewLink   = $link
                        SurfaceType   = 'DOCUMENT'
                        SourceRef     = $doc.Id
                        SourceLink    = $link
                        CreatorEmail  = $creator
                        CreatorName   = $creatorName
                        Shared        = $isShared
                        Origin        = 'DOCSCAN'
                        Notes         = ($noteParts -join ' | ')
                    })
            }
        }
    }

    if ($CheckpointCsv -and $results.Count -gt $preCount) {
        $new = for ($i = $preCount; $i -lt $results.Count; $i++) { $results[$i] }
        $new | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 |
        Out-File -FilePath $CheckpointCsv -Encoding UTF8 -Append
    }
}

# 3) Tenant-wide Chat space scan: capture orphaned / unassigned tasks visible
#    only as Chat system messages (e.g. created in a space but never assigned).
if (-not $SkipTenantSpaceScan -and $resolvedUsers.Count -gt 0) {
    $impUser = $resolvedUsers | Where-Object { $_ -eq 'utsab@rocheua.com' } | Select-Object -First 1
    if (-not $impUser) { $impUser = $resolvedUsers[0] }
    Write-Host ("Scanning tenant Chat spaces via {0}..." -f $impUser) -ForegroundColor Cyan

    $spaces = Get-TenantChatSpaces -ImpersonationUser $impUser
    Write-Host ("  Found {0} space(s)" -f $spaces.Count) -ForegroundColor Gray

    # Index of already-captured tasks: key = "space|roundedSeconds".
    $existingIdx = @{}
    foreach ($r in $results) {
        if ($r.SurfaceType -ne 'SPACE' -or -not $r.SourceRef -or -not $r.Updated) { continue }
        try {
            $t = [DateTimeOffset]::Parse($r.Updated).UtcDateTime
            $key = "{0}|{1}" -f $r.SourceRef, [int64]($t - [datetime]'1970-01-01').TotalSeconds
            $existingIdx[$key] = $true
        }
        catch {}
    }

    foreach ($sp in $spaces) {
        $member = Get-SpaceHumanMember -ImpersonationUser $impUser -Space $sp.Id
        if (-not $member) { continue }
        $msgs = Get-SpaceAllMessages -ImpersonationUser $member -Space $sp.Id
        if (-not $msgs -or $msgs.Count -eq 0) { continue }
        $threads = Build-SpaceTaskThreads -Messages $msgs

        foreach ($tn in $threads.Keys) {
            $th = $threads[$tn]
            if ($th.Deleted) { continue }
            if (-not $th.Creator -or -not $th.CreateTime) { continue }

            $dup = $false
            try {
                $t0 = [DateTimeOffset]::Parse($th.CreateTime).UtcDateTime
                for ($d = -10; $d -le 10; $d++) {
                    $k = "{0}|{1}" -f $sp.Id, ([int64]($t0 - [datetime]'1970-01-01').TotalSeconds + $d)
                    if ($existingIdx.ContainsKey($k)) { $dup = $true; break }
                }
            }
            catch {}
            if ($dup) { continue }

            $assignee = $th.Assignee
            $isShared = [bool]($assignee -and $th.Creator -and ($assignee -ne $th.Creator))
            $noteParts = @()
            if ($sp.DisplayName) { $noteParts += "SpaceName: $($sp.DisplayName)" }
            if (-not $assignee) { $noteParts += 'Orphaned/unassigned task discovered via Chat space scan.' }
            else { $noteParts += 'Task discovered via Chat space scan (not returned by Tasks API).' }
            if ($th.AssigneeMention -and -not $assignee) { $noteParts += "MentionText: @$($th.AssigneeMention)" }

            $results.Add([pscustomobject]@{
                    AssigneeEmail = $assignee
                    TaskListId    = $null
                    TaskListTitle = $null
                    TaskId        = $tn
                    Title         = $th.LastEventText
                    Status        = if ($assignee) { 'needsAction' } else { 'unassigned' }
                    Due           = $null
                    Completed     = $null
                    Updated       = $th.CreateTime
                    WebViewLink   = $null
                    SurfaceType   = 'SPACE'
                    SourceRef     = $sp.Id
                    SourceLink    = $null
                    CreatorEmail  = $th.Creator
                    CreatorName   = $null
                    Shared        = $isShared
                    Origin        = 'SPACESCAN'
                    Notes         = ($noteParts -join ' | ')
                })
        }
    }
}

$results | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
Write-Host ("Wrote {0} task row(s) to {1}" -f $results.Count, $OutputCsv) -ForegroundColor Green

# 4) Sharer-centric summary CSV (grouped by creator, shared rows only).
$sharedRows = $results | Where-Object { $_.Shared -eq $true -and $_.CreatorEmail }
if ($sharedRows) {
    $summary = $sharedRows | Group-Object CreatorEmail | ForEach-Object {
        [pscustomobject]@{
            CreatorEmail   = $_.Name
            SharedCount    = $_.Count
            AssigneeEmails = (($_.Group | ForEach-Object AssigneeEmail | Where-Object { $_ } | Sort-Object -Unique) -join '; ')
            Titles         = (($_.Group | ForEach-Object Title | Where-Object { $_ } | Sort-Object -Unique) -join '; ')
            SurfaceTypes   = (($_.Group | ForEach-Object SurfaceType | Sort-Object -Unique) -join '; ')
        }
    } | Sort-Object SharedCount -Descending
    $summary | Export-Csv -Path $SummaryCsv -NoTypeInformation -Encoding UTF8
    Write-Host ("Wrote {0} sharer row(s) to {1}" -f @($summary).Count, $SummaryCsv) -ForegroundColor Green
}
else {
    Write-Host "No shared rows detected; summary CSV not written." -ForegroundColor Yellow
}

foreach ($tf in $script:tenantUserFiles) { Remove-Item $tf -Force -ErrorAction SilentlyContinue }
