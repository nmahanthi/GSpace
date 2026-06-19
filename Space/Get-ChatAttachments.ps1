# =============================================================================
# Get-ChatAttachments.ps1
# Audits Google Chat attachment count + size per Space / Group Chat / DM
#
# USAGE EXAMPLES
#   # All users (full tenant)
#   .\Get-ChatAttachments.ps1
#
#   # One specific user
#   .\Get-ChatAttachments.ps1 -User john@domain.com
#
#   # Several users inline
#   .\Get-ChatAttachments.ps1 -User "john@domain.com","jane@domain.com"
#
#   # From a file
#   .\Get-ChatAttachments.ps1 -Mode TargetUsers
#
#   # Spaces only, no DMs, fastest
#   .\Get-ChatAttachments.ps1 -Mode AdminOnly
#
#   # Narrow to last 30 days, flag anything > 5 MB
#   .\Get-ChatAttachments.ps1 -User john@domain.com -Days 30 -LargeMB 5
#
# PARAMETERS
#   -User           One or more email addresses.  Overrides Mode -> TargetUsers.
#   -Mode           AllUsers | TargetUsers | AdminOnly  (default: AllUsers)
#   -Days           How many days back to scan messages (default: 90)
#   -LargeMB        Flag threshold in MB for "large" attachments (default: 10)
#   -NoDMs          Switch: skip Direct Message spaces
#   -MetadataOnly   Switch: size uploaded files via Drive metadata ONLY.
#                   Chat-uploaded files live in the sender's Drive "Chat Files"
#                   folder; we find them by name+type without touching the
#                   attachment download URL at all.  If a file is not found in
#                   Drive (rare), size stays 0 rather than doing a Range GET.
#                   Omit this flag to allow a 1-byte Range GET fallback for
#                   files that aren't found via Drive search.
# =============================================================================
[CmdletBinding()]
param(
    [string[]] $User,
    [ValidateSet("AllUsers","TargetUsers","AdminOnly")]
    [string]   $Mode,
    [int]      $Days,
    [int]      $LargeMB,
    [switch]   $NoDMs,
    [switch]   $MetadataOnly,
    [int]      $Throttle,      # parallel workers per phase (default 5; requires PS7+)
    [int]      $MaxDMUsers     # cap DM-user scan in AllUsers mode; 0 = no cap
)

# ── CONFIGURATION ─────────────────────────────────────────────────────────────
$AdminEmail          = "admin-narendra@rocheua.com"
$GamPath             = Join-Path (Split-Path $PSScriptRoot) "gam7\gam.exe"
$OutputDir           = $PSScriptRoot
$Timestamp           = Get-Date -Format "yyyyMMdd_HHmmss"
$TimeWindowDays      = 90          # how far back to scan messages
$LargeFileMB         = 10         # flag threshold for "large" attachments
$IncludeGroupChats   = $true      # include GROUP_CHAT spaces
$IncludeDMs          = $true      # include Direct Message spaces (needs per-user calls)
$HeadTimeoutSec      = 10         # timeout per HEAD request for upload sizing
$MaxHeadRequests     = 500        # cap to avoid rate-limit on large tenants
$ThrottleLimit       = 5          # parallel workers (Phase 1 DM, Phase 2, Phase 3a) – PS7+ only
$MaxDMUsersCount     = 0          # 0 = scan every user for DMs; set > 0 to cap (AllUsers mode)

# ── RUN MODE ──────────────────────────────────────────────────────────────────
#   AllUsers    : all named Spaces (asadmin) + every user's DMs (per-user)
#   TargetUsers : only spaces/DMs for users listed in $TargetUsersFile
#   AdminOnly   : original - only spaces visible asadmin, no DMs
$RunMode             = "AllUsers"    # AllUsers | TargetUsers | AdminOnly

# Used only when RunMode = "TargetUsers"
# One email per line (plain .txt)  OR  CSV with column User / Email / primaryEmail
$TargetUsersFile     = Join-Path $PSScriptRoot "TargetUsers.txt"

# ── APPLY COMMAND-LINE PARAM OVERRIDES ────────────────────────────────────────
# -User  supplied  -> switch to TargetUsers mode, use those emails directly
if ($User -and $User.Count -gt 0) {
    $RunMode = "TargetUsers"
    # Store inline list; Load-TargetUsers will return this instead of reading a file
    $InlineTargetUsers = @($User | ForEach-Object { $_.Trim() } | Where-Object { $_ -match "@" })
} else {
    $InlineTargetUsers = $null
}
if ($Mode)         { $RunMode        = $Mode    }
if ($Days)         { $TimeWindowDays = $Days    }
if ($LargeMB)      { $LargeFileMB   = $LargeMB }
if ($NoDMs)        { $IncludeDMs    = $false   }
if ($MetadataOnly) { $UseMetadataOnly = $true  } else { $UseMetadataOnly = $false }
if ($Throttle)     { $ThrottleLimit   = $Throttle   }
if ($PSBoundParameters.ContainsKey('MaxDMUsers')) { $MaxDMUsersCount = $MaxDMUsers }

# Detect PowerShell version — ForEach-Object -Parallel requires PS 7+
$RunParallel = ($PSVersionTable.PSVersion.Major -ge 7)
if (-not $RunParallel) {
    Write-Warning "PS7+ required for parallel speedup. Running sequentially. Upgrade: winget install Microsoft.PowerShell"
    $ThrottleLimit = 1
}

# ── EXISTING ADMIN SPACE LIST (for AllUsers / AdminOnly modes) ─────────────────
# chat_spaces_target.csv from the BOT/CHAT run gives User + name + spaceType.
# Script re-pulls automatically if the path does not exist.
$ExistingSpaceCsv    = Join-Path (Split-Path $PSScriptRoot) `
    "gam_exports\Bot&Chat App related_usingGAM\RUN_BOTCHAT_20260615_154743Z\chat_spaces_target.csv"

# ── OUTPUT PATHS ──────────────────────────────────────────────────────────────
$TmpMsgCsv           = Join-Path $OutputDir "tmp_chatmsgs_$Timestamp.csv"
$DetailCsv           = Join-Path $OutputDir "CHAT_ATTACH_DETAIL_$Timestamp.csv"
$PerSpaceCsv         = Join-Path $OutputDir "CHAT_ATTACH_PER_SPACE_$Timestamp.csv"
$TopCsv              = Join-Path $OutputDir "CHAT_ATTACH_TOP_LARGEST_$Timestamp.csv"

Write-Host ("=" * 60) -ForegroundColor Cyan
Write-Host "  Google Chat Attachment Audit (GAM7 Hybrid)" -ForegroundColor Cyan
Write-Host ("=" * 60) -ForegroundColor Cyan
Write-Host "Admin       : $AdminEmail"
Write-Host "Mode        : $RunMode"
if ($InlineTargetUsers) {
    Write-Host "Users       : $($InlineTargetUsers -join ', ')" -ForegroundColor Green
}
Write-Host "Window      : Last $TimeWindowDays days"
Write-Host "Include DMs : $IncludeDMs"
Write-Host "Sizing mode : $(if($UseMetadataOnly){'MetadataOnly (Drive search, no Range GET)'}else{'Drive search + Range GET fallback'})"
Write-Host "Large flag  : > $LargeFileMB MB"
Write-Host "Parallel    : $(if($RunParallel){"PS7 - ThrottleLimit $ThrottleLimit  (-Throttle N to change)"}else{"PS5 - sequential"})"
if ($MaxDMUsersCount -gt 0) {
    Write-Host "DM user cap : $MaxDMUsersCount users max (-MaxDMUsers 0 to disable cap)" -ForegroundColor DarkYellow
}
Write-Host "Timestamp   : $Timestamp"
Write-Host ""

# ── HELPER: wait for a file to appear ─────────────────────────────────────────
function Wait-ForFile($path, $timeoutSec = 300) {
    $elapsed = 0
    Write-Host "   Waiting for GAM output" -NoNewline
    while (-not (Test-Path $path) -or (Get-Item $path).Length -eq 0) {
        if ($elapsed -ge $timeoutSec) {
            Write-Host "`n   ERROR: timed out after $timeoutSec s" -ForegroundColor Red; exit 1
        }
        Start-Sleep -Seconds 3; $elapsed += 3; Write-Host "." -NoNewline
    }
    Write-Host " OK" -ForegroundColor Green
}

# ── HELPER: safe column reader ─────────────────────────────────────────────────
function Get-Col($row, [string[]]$names) {
    foreach ($n in $names) {
        $v = $row.PSObject.Properties[$n]
        if ($v -and "$($v.Value)".Trim() -ne "") { return "$($v.Value)".Trim() }
    }
    return ""
}

# Capture Get-Col source so it can be injected into parallel runspaces via Invoke-Expression
$GetColSrc = @'
function Get-Col { param($row, [string[]]$names)
    foreach ($n in $names) {
        $v = $row.PSObject.Properties[$n]
        if ($v -and "$($v.Value)".Trim() -ne "") { return "$($v.Value)".Trim() }
    }; return "" }
'@

# ── HELPER: pull chatspaces for a single user, return rows ────────────────────
function Get-UserSpaces($userEmail, [string[]]$types) {
    $typeArg = $types -join ","
    $tmp     = Join-Path $OutputDir "tmp_uspaces_${Timestamp}.csv"
    if (Test-Path $tmp) { Remove-Item $tmp -Force }
    & $GamPath redirect csv $tmp user $userEmail print chatspaces `
        types $typeArg `
        fields "name,displayname,spacetype,membershipcount,createtime,lastactivetime,spaceuri" 2>$null
    $rows = @()
    if (Test-Path $tmp) {
        $rows = @(Import-Csv $tmp)
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    }
    return $rows
}

# ── HELPER: merge rows into $spacemap (deduplicates by space name) ────────────
function Add-ToSpaceMap($rows, $runnerHint) {
    foreach ($r in $rows) {
        $sname = (Get-Col $r "name").Trim()
        $stype = (Get-Col $r "spaceType","spacetype").Trim().ToUpper()
        if (-not $sname -or $script:spacemap.ContainsKey($sname)) { continue }
        if ($stype -notin $script:allowedTypes) { continue }
        $runner = if ($runnerHint) { $runnerHint } else {
            $v = Get-Col $r "User","TargetUser"
            if ($v) { $v } else { $script:AdminEmail }
        }
        $script:spacemap[$sname] = [ordered]@{
            Runner      = $runner
            DisplayName = Get-Col $r "displayName","displayname"
            SpaceType   = $stype
            MemberCount = Get-Col $r "membershipCount.joinedDirectHumanUserCount","membershipcount"
            SpaceUri    = Get-Col $r "spaceUri","spaceuri"
            CreateTime  = Get-Col $r "createTime","createtime"
            LastActive  = Get-Col $r "lastActiveTime","lastactivetime"
        }
    }
}

# =============================================================================
# PHASE 1 – Build unique space list  (mode-aware)
#
#   AdminOnly   : asadmin CSV -> SPACE + GROUP_CHAT only, no DMs
#   AllUsers    : asadmin CSV for SPACE/GROUP_CHAT  +  per-user pull for DMs
#   TargetUsers : per-user pull for each target (SPACE + GROUP_CHAT + DMs)
# =============================================================================
Write-Host "[1/3] Building space list  (mode: $RunMode)..." -ForegroundColor Yellow

$allowedTypes = [System.Collections.Generic.List[string]]@("SPACE")
if ($IncludeGroupChats) { $allowedTypes.Add("GROUP_CHAT") }
if ($IncludeDMs -and $RunMode -ne "AdminOnly") { $allowedTypes.Add("DIRECT_MESSAGE") }

$spacemap = @{}   # spaces/XXX -> { Runner; DisplayName; SpaceType; ... }

# ── Sub-function: load asadmin space CSV (SPACE + GROUP_CHAT) ─────────────────
function Load-AdminSpaces {
    $rows = $null
    if (Test-Path $script:ExistingSpaceCsv) {
        $rows = @(Import-Csv $script:ExistingSpaceCsv)
        Write-Host "   Loaded admin space CSV  : $($rows.Count) rows"
    } else {
        Write-Host "   Admin CSV not found - re-pulling asadmin..." -ForegroundColor DarkYellow
        $tmp = Join-Path $script:OutputDir "tmp_adminspaces_$($script:Timestamp).csv"
        & $script:GamPath redirect csv $tmp user $script:AdminEmail print chatspaces asadmin `
            fields "name,displayname,spacetype,membershipcount,createtime,lastactivetime,spaceuri" 2>$null
        Wait-ForFile $tmp
        $rows = @(Import-Csv $tmp)
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        Write-Host "   Pulled $($rows.Count) admin space rows"
    }
    Add-ToSpaceMap $rows $null
}

# ── Sub-function: load target user list (inline -User param OR file) ──────────
function Load-TargetUsers {
    # Inline -User param takes priority over file
    if ($script:InlineTargetUsers -and $script:InlineTargetUsers.Count -gt 0) {
        Write-Host "   Using inline -User list: $($script:InlineTargetUsers -join ', ')" -ForegroundColor Green
        return $script:InlineTargetUsers
    }
    if (-not (Test-Path $script:TargetUsersFile)) {
        Write-Host "   ERROR: TargetUsersFile not found: $($script:TargetUsersFile)" -ForegroundColor Red
        Write-Host "   Tip: pass emails directly with  -User user@domain.com" -ForegroundColor Yellow
        exit 1
    }
    $ext = [System.IO.Path]::GetExtension($script:TargetUsersFile).ToLower()
    if ($ext -eq ".csv") {
        $rows = Import-Csv $script:TargetUsersFile
        $col  = $rows[0].PSObject.Properties.Name |
                Where-Object { $_ -match "user|email|primaryemail" } |
                Select-Object -First 1
        return @($rows | Select-Object -ExpandProperty $col | Where-Object { $_ -match "@" })
    } else {
        return @(Get-Content $script:TargetUsersFile |
                 Where-Object { $_.Trim() -match "@" } |
                 ForEach-Object { $_.Trim() })
    }
}

# ── BRANCH BY RUN MODE ────────────────────────────────────────────────────────
switch ($RunMode) {

    "AdminOnly" {
        # SPACE + GROUP_CHAT only via asadmin — original behaviour
        Load-AdminSpaces
    }

    "AllUsers" {
        # Step A: SPACE + GROUP_CHAT from asadmin (fast, comprehensive)
        Load-AdminSpaces

        # Step B: DMs — must be pulled per-user (admin is not a DM participant)
        if ($IncludeDMs) {
            Write-Host "   Pulling all-user list for DM scan..." -ForegroundColor DarkYellow
            $tmpUsers = Join-Path $OutputDir "tmp_allusers_$Timestamp.csv"
            & $GamPath redirect csv $tmpUsers print users fields primaryEmail 2>$null
            Wait-ForFile $tmpUsers
            $allUsers = @(Import-Csv $tmpUsers | Select-Object -ExpandProperty primaryEmail |
                          Where-Object { $_ -match "@" })
            Remove-Item $tmpUsers -Force -ErrorAction SilentlyContinue
            Write-Host "   Users found : $($allUsers.Count)"

            # Cap DM scan if requested (AllUsers with 100k users can take hours without a cap)
            $dmUserList = if ($MaxDMUsersCount -gt 0 -and $allUsers.Count -gt $MaxDMUsersCount) {
                Write-Host ("   Capping DM scan to {0} of {1} users  (raise with -MaxDMUsers N)" `
                    -f $MaxDMUsersCount, $allUsers.Count) -ForegroundColor DarkYellow
                @($allUsers | Select-Object -First $MaxDMUsersCount)
            } else { $allUsers }

            Write-Host ("   Scanning DMs for {0} users  (parallel={1}, workers={2})..." `
                -f $dmUserList.Count, $RunParallel, $ThrottleLimit) -ForegroundColor DarkGray

            if ($RunParallel) {
                # Each worker pulls DM spaces for one user and outputs flat hashtables
                $dmRaw = $dmUserList | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
                    Invoke-Expression $using:GetColSrc
                    $u       = $_
                    $gamPath = $using:GamPath
                    $outDir  = $using:OutputDir
                    $tmp = Join-Path $outDir "tmp_dm_$(New-Guid).csv"
                    & $gamPath redirect csv $tmp user $u print chatspaces types directmessage `
                        fields "name,displayname,spacetype,membershipcount,createtime,lastactivetime,spaceuri" 2>$null
                    if (Test-Path $tmp) {
                        $rows = @(Import-Csv $tmp)
                        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
                        foreach ($r in $rows) {
                            @{
                                name            = Get-Col $r "name"
                                displayname     = Get-Col $r "displayName","displayname"
                                spacetype       = Get-Col $r "spaceType","spacetype"
                                membershipcount = Get-Col $r "membershipCount.joinedDirectHumanUserCount","membershipcount"
                                spaceuri        = Get-Col $r "spaceUri","spaceuri"
                                createtime      = Get-Col $r "createTime","createtime"
                                lastactivetime  = Get-Col $r "lastActiveTime","lastactivetime"
                                Runner          = $u
                            }
                        }
                    }
                }
                # Merge into spacemap on the main thread (safe to mutate here)
                foreach ($ht in @($dmRaw)) {
                    if (-not $ht -or -not $ht.name) { continue }
                    $sn = "$($ht.name)".Trim()
                    if (-not $sn -or $spacemap.ContainsKey($sn)) { continue }
                    if ("DIRECT_MESSAGE" -notin $allowedTypes) { continue }
                    $spacemap[$sn] = [ordered]@{
                        Runner      = $ht.Runner
                        DisplayName = $ht.displayname
                        SpaceType   = "DIRECT_MESSAGE"
                        MemberCount = $ht.membershipcount
                        SpaceUri    = $ht.spaceuri
                        CreateTime  = $ht.createtime
                        LastActive  = $ht.lastactivetime
                    }
                }
            } else {
                # PS5 sequential fallback
                $ui = 0
                foreach ($u in $dmUserList) {
                    $ui++
                    if ($ui % 25 -eq 0) {
                        Write-Host ("   DM scan: {0}/{1} users ({2} DMs so far)" `
                            -f $ui, $dmUserList.Count, `
                            ($spacemap.Values | Where-Object { $_.SpaceType -eq "DIRECT_MESSAGE" }).Count) `
                            -ForegroundColor DarkGray
                    }
                    $dmRows = Get-UserSpaces $u @("directmessage")
                    Add-ToSpaceMap $dmRows $u
                }
            }
            Write-Host ("   DM spaces added : {0}" `
                -f ($spacemap.Values | Where-Object { $_.SpaceType -eq "DIRECT_MESSAGE" }).Count) `
                -ForegroundColor Green
        }
    }

    "TargetUsers" {
        # Pull every space type per target user (scopes to their actual memberships)
        $targetUsers = Load-TargetUsers
        Write-Host "   Target users loaded : $($targetUsers.Count)"

        $typeArgs = [System.Collections.Generic.List[string]]@("space","groupchat")
        if ($IncludeDMs) { $typeArgs.Add("directmessage") }

        Write-Host ("   Pulling spaces for {0} target user(s)  (parallel={1}, workers={2})..." `
            -f $targetUsers.Count, $RunParallel, $ThrottleLimit) -ForegroundColor DarkGray

        if ($RunParallel) {
            $typeArgStr = $typeArgs -join ","
            $tuRaw = $targetUsers | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
                Invoke-Expression $using:GetColSrc
                $u          = $_
                $gamPath    = $using:GamPath
                $outDir     = $using:OutputDir
                $typeStr    = $using:typeArgStr
                $tmp = Join-Path $outDir "tmp_tu_$(New-Guid).csv"
                & $gamPath redirect csv $tmp user $u print chatspaces types $typeStr `
                    fields "name,displayname,spacetype,membershipcount,createtime,lastactivetime,spaceuri" 2>$null
                if (Test-Path $tmp) {
                    $rows = @(Import-Csv $tmp)
                    Remove-Item $tmp -Force -ErrorAction SilentlyContinue
                    foreach ($r in $rows) {
                        @{
                            name            = Get-Col $r "name"
                            displayname     = Get-Col $r "displayName","displayname"
                            spacetype       = (Get-Col $r "spaceType","spacetype").ToUpper()
                            membershipcount = Get-Col $r "membershipCount.joinedDirectHumanUserCount","membershipcount"
                            spaceuri        = Get-Col $r "spaceUri","spaceuri"
                            createtime      = Get-Col $r "createTime","createtime"
                            lastactivetime  = Get-Col $r "lastActiveTime","lastactivetime"
                            Runner          = $u
                        }
                    }
                }
            }
            foreach ($ht in @($tuRaw)) {
                if (-not $ht -or -not $ht.name) { continue }
                $sn = "$($ht.name)".Trim()
                if (-not $sn -or $spacemap.ContainsKey($sn)) { continue }
                if ($ht.spacetype -notin $allowedTypes) { continue }
                $spacemap[$sn] = [ordered]@{
                    Runner      = $ht.Runner
                    DisplayName = $ht.displayname
                    SpaceType   = $ht.spacetype
                    MemberCount = $ht.membershipcount
                    SpaceUri    = $ht.spaceuri
                    CreateTime  = $ht.createtime
                    LastActive  = $ht.lastactivetime
                }
            }
        } else {
            $ti = 0
            foreach ($u in $targetUsers) {
                $ti++
                Write-Host ("  [{0,3}/{1}] {2}" -f $ti, $targetUsers.Count, $u) -ForegroundColor DarkCyan
                $rows = Get-UserSpaces $u $typeArgs
                Add-ToSpaceMap $rows $u
            }
        }
    }

    default {
        Write-Host "   ERROR: Unknown RunMode '$RunMode'. Use AllUsers, TargetUsers, or AdminOnly." `
            -ForegroundColor Red
        exit 1
    }
}

$uniqueSpaces = @($spacemap.Keys | Sort-Object)
$dmCount      = ($spacemap.Values | Where-Object { $_.SpaceType -eq "DIRECT_MESSAGE" }).Count
$spaceCount   = ($spacemap.Values | Where-Object { $_.SpaceType -eq "SPACE" }).Count
$gcCount      = ($spacemap.Values | Where-Object { $_.SpaceType -eq "GROUP_CHAT" }).Count

Write-Host ""
Write-Host ("   Unique spaces to scan : {0}  (SPACE:{1}  GROUP_CHAT:{2}  DM:{3})" `
    -f $uniqueSpaces.Count, $spaceCount, $gcCount, $dmCount) -ForegroundColor Green

# =============================================================================
# PHASE 2 – Per-space: pull messages, parse attachment columns
# =============================================================================
Write-Host ""
Write-Host ("[2/3] Pulling messages per space  (parallel={0}, workers={1})..." `
    -f $RunParallel, $ThrottleLimit) -ForegroundColor Yellow

# Time filter: ISO-8601, UTC
$filterTime = (Get-Date).AddDays(-$TimeWindowDays).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
$msgFilter  = "createTime > `"$filterTime`""

# ── PARALLEL PATH (PS7+) ──────────────────────────────────────────────────────
if ($RunParallel) {
    # Snapshot all work items before entering parallel (avoids $using: on a live hashtable)
    $spaceWork = @($uniqueSpaces | ForEach-Object {
        $meta = $spacemap[$_]
        [PSCustomObject]@{
            SpaceName   = $_
            Runner      = $meta.Runner
            DisplayName = if ($meta.DisplayName) { $meta.DisplayName } else { $_ }
            SpaceType   = $meta.SpaceType
            MemberCount = $meta.MemberCount
        }
    })

    # Each worker pulls one space and emits flat hashtables — one per attachment
    $phase2Rows = $spaceWork | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
        Invoke-Expression $using:GetColSrc

        $item      = $_
        $gamPath   = $using:GamPath
        $outDir    = $using:OutputDir
        $flt       = $using:msgFilter

        $tmpFile = Join-Path $outDir "tmp_p2_$(New-Guid).csv"
        $gArgs   = @("redirect","csv",$tmpFile,"user",$item.Runner,"print","chatmessages",
                     $item.SpaceName,"filter",$flt,"fields",
                     "name,createtime,sender,attachment,attachedgifs")
        & $gamPath @gArgs 2>$null

        if (-not (Test-Path $tmpFile) -or (Get-Item $tmpFile).Length -lt 10) { return }
        $msgs = @(Import-Csv $tmpFile)
        Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue
        if (-not $msgs -or $msgs.Count -eq 0) { return }

        $sampleCols = $msgs[0].PSObject.Properties.Name
        $attachIdxs = $sampleCols |
            Where-Object { $_ -match '^attachment\.(\d+)\.' } |
            ForEach-Object { [int]($_ -replace '^attachment\.(\d+)\..*','$1') } |
            Sort-Object -Unique

        foreach ($msg in $msgs) {
            $msgName    = Get-Col $msg "name"
            $createTime = Get-Col $msg "createTime","createtime"
            $sender     = Get-Col $msg "sender.name","sender.email","sender"
            $senderDisp = Get-Col $msg "sender.displayName","sender.displayname"

            $gifCols  = $msg.PSObject.Properties.Name | Where-Object { $_ -match '^attachedGifs\.' }
            $gifCount = ($gifCols | ForEach-Object { $_ -replace '^attachedGifs\.(\d+)\..*','$1' } |
                         Sort-Object -Unique).Count
            if ($gifCount -gt 0) {
                @{  SpaceID="$($item.SpaceName)"; SpaceName="$($item.DisplayName)"
                    SpaceType="$($item.SpaceType)"; MemberCount="$($item.MemberCount)"
                    MessageID="$msgName"; CreateTime="$createTime"
                    Sender="$senderDisp"; SenderID="$sender"
                    AttachIndex="gifs"; AttachName="Animated GIF(s)"
                    ContentType="image/gif"; Source="GIPHY_TENOR"
                    IsInlineImage=$true; DriveFileID=""; DownloadUri=""
                    Bytes=0; SizeMB=0; IsLarge=$false
                    Note="$gifCount gif(s) - URL only, no size" }
            }

            foreach ($idx in $attachIdxs) {
                $prefix  = "attachment.$idx"
                $aName   = Get-Col $msg "$prefix.name"
                if (-not $aName) { continue }

                $cName   = Get-Col $msg "$prefix.contentName"
                $cType   = Get-Col $msg "$prefix.contentType"
                $src     = Get-Col $msg "$prefix.source"
                $driveId = Get-Col $msg "$prefix.driveDataRef.driveFileId"
                $dlUri   = Get-Col $msg "$prefix.downloadUri"
                $isImg   = $cType -match "^image/"

                @{  SpaceID="$($item.SpaceName)"; SpaceName="$($item.DisplayName)"
                    SpaceType="$($item.SpaceType)"; MemberCount="$($item.MemberCount)"
                    MessageID="$msgName"; CreateTime="$createTime"
                    Sender="$senderDisp"; SenderID="$sender"
                    AttachIndex=$idx; AttachName="$cName"
                    ContentType="$cType"; Source="$src"
                    IsInlineImage=$isImg; DriveFileID="$driveId"
                    DownloadUri="$dlUri"; Bytes=0; SizeMB=0; IsLarge=$false; Note="" }
            }
        }
    }

    # Collect flat hashtables into typed lists (main thread preserves mutable references for Phase 3)
    $detailRows   = [System.Collections.Generic.List[PSCustomObject]]::new()
    $driveIds     = [System.Collections.Generic.List[PSCustomObject]]::new()
    $uploadedRows = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($ht in @($phase2Rows)) {
        if (-not $ht) { continue }
        $row = [PSCustomObject]$ht
        $detailRows.Add($row)
        if ($ht.Source -eq "DRIVE_FILE" -and $ht.DriveFileID) {
            $driveIds.Add([PSCustomObject]@{ DriveFileID = $ht.DriveFileID; Row = $row })
        } elseif ($ht.Source -eq "UPLOADED_CONTENT" -and $ht.DownloadUri) {
            $uploadedRows.Add([PSCustomObject]@{ Uri = $ht.DownloadUri; Row = $row })
        }
    }

} else {
    # ── SEQUENTIAL FALLBACK (PS5) ─────────────────────────────────────────────
    $detailRows   = [System.Collections.Generic.List[PSCustomObject]]::new()
    $driveIds     = [System.Collections.Generic.List[PSCustomObject]]::new()
    $uploadedRows = [System.Collections.Generic.List[PSCustomObject]]::new()

    $spaceIdx = 0
    foreach ($sname in $uniqueSpaces) {
        $spaceIdx++
        $meta   = $spacemap[$sname]
        $runner = $meta.Runner
        $disp   = if ($meta.DisplayName) { $meta.DisplayName } else { $sname }

        Write-Host ("  [{0,3}/{1}] {2} ({3}) as {4}" -f `
            $spaceIdx, $uniqueSpaces.Count, $disp, $meta.SpaceType, $runner) -ForegroundColor DarkCyan

        if (Test-Path $TmpMsgCsv) { Remove-Item $TmpMsgCsv -Force }
        $gamArgs = @(
            "redirect","csv",$TmpMsgCsv,"user",$runner,
            "print","chatmessages",$sname,"filter",$msgFilter,
            "fields","name,createtime,sender,attachment,attachedgifs"
        )
        & $GamPath @gamArgs 2>$null

        if (-not (Test-Path $TmpMsgCsv) -or (Get-Item $TmpMsgCsv).Length -lt 10) {
            Write-Host "      (no messages or access denied)" -ForegroundColor DarkGray
            continue
        }
        $msgs = Import-Csv $TmpMsgCsv
        if (-not $msgs -or $msgs.Count -eq 0) { continue }

        $sampleCols = $msgs[0].PSObject.Properties.Name
        $attachIdxs = $sampleCols |
            Where-Object { $_ -match '^attachment\.(\d+)\.' } |
            ForEach-Object { [int]($_ -replace '^attachment\.(\d+)\..*','$1') } |
            Sort-Object -Unique

        foreach ($msg in $msgs) {
            $msgName    = Get-Col $msg "name"
            $createTime = Get-Col $msg "createTime","createtime"
            $sender     = Get-Col $msg "sender.name","sender.email","sender"
            $senderDisp = Get-Col $msg "sender.displayName","sender.displayname"

            $gifCols  = $msg.PSObject.Properties.Name | Where-Object { $_ -match '^attachedGifs\.' }
            $gifCount = ($gifCols | ForEach-Object { $_ -replace '^attachedGifs\.(\d+)\..*','$1' } |
                         Sort-Object -Unique).Count
            if ($gifCount -gt 0) {
                $detailRows.Add([PSCustomObject]@{
                    SpaceID="$sname"; SpaceName="$disp"; SpaceType=$meta.SpaceType
                    MemberCount=$meta.MemberCount; MessageID=$msgName; CreateTime=$createTime
                    Sender=$senderDisp; SenderID=$sender; AttachIndex="gifs"
                    AttachName="Animated GIF(s)"; ContentType="image/gif"; Source="GIPHY_TENOR"
                    IsInlineImage=$true; DriveFileID=""; DownloadUri=""
                    Bytes=0; SizeMB=0; IsLarge=$false
                    Note="$gifCount gif(s) - URL only, no size"
                })
            }

            foreach ($idx in $attachIdxs) {
                $prefix      = "attachment.$idx"
                $aName       = Get-Col $msg "$prefix.name"
                if (-not $aName) { continue }

                $contentName = Get-Col $msg "$prefix.contentName"
                $contentType = Get-Col $msg "$prefix.contentType"
                $source      = Get-Col $msg "$prefix.source"
                $driveId     = Get-Col $msg "$prefix.driveDataRef.driveFileId"
                $dlUri       = Get-Col $msg "$prefix.downloadUri"
                $isImage     = $contentType -match "^image/"

                $detRow = [PSCustomObject]@{
                    SpaceID=$sname; SpaceName=$disp; SpaceType=$meta.SpaceType
                    MemberCount=$meta.MemberCount; MessageID=$msgName; CreateTime=$createTime
                    Sender=$senderDisp; SenderID=$sender; AttachIndex=$idx
                    AttachName=$contentName; ContentType=$contentType; Source=$source
                    IsInlineImage=$isImage; DriveFileID=$driveId; DownloadUri=$dlUri
                    Bytes=0; SizeMB=0; IsLarge=$false; Note=""
                }
                $detailRows.Add($detRow)

                if ($source -eq "DRIVE_FILE" -and $driveId) {
                    $driveIds.Add([PSCustomObject]@{ DriveFileID = $driveId; Row = $detRow })
                } elseif ($source -eq "UPLOADED_CONTENT" -and $dlUri) {
                    $uploadedRows.Add([PSCustomObject]@{ Uri = $dlUri; Row = $detRow })
                }
            }
        }
        Remove-Item $TmpMsgCsv -Force -ErrorAction SilentlyContinue
    }
}

Write-Host ""
Write-Host "   Total attachment detail rows : $($detailRows.Count)" -ForegroundColor Green
Write-Host "   Drive-source attachments     : $($driveIds.Count)"
Write-Host "   Uploaded/inline attachments  : $($uploadedRows.Count)"

# =============================================================================
# PHASE 3a – Enrich Drive-source attachments with real byte size
#            gam user <runner> print filelist select drivefileid <id> fields size
# =============================================================================
Write-Host ""
Write-Host ("[3a/3] Enriching Drive attachment sizes  (parallel={0}, workers={1})..." `
    -f $RunParallel, $ThrottleLimit) -ForegroundColor Yellow

if ($driveIds.Count -gt 0) {
    $uniqueDriveIds = @($driveIds | Select-Object -ExpandProperty DriveFileID -Unique)
    Write-Host "   Unique Drive file IDs to query : $($uniqueDriveIds.Count)"

    $driveSizeMap = @{}   # fileId -> bytes

    # Pre-compute the best runner per file ID (admin first, then space runner as fallback)
    $driveWork = @($uniqueDriveIds | ForEach-Object {
        $id       = $_
        $ref      = $driveIds | Where-Object { $_.DriveFileID -eq $id } | Select-Object -First 1
        $fbRunner = $AdminEmail
        if ($ref -and $ref.Row.SpaceID -and $spacemap.ContainsKey($ref.Row.SpaceID)) {
            $sr = $spacemap[$ref.Row.SpaceID].Runner
            if ($sr -and $sr -ne $AdminEmail) { $fbRunner = $sr }
        }
        [PSCustomObject]@{ Id = $id; AdminEmail = $AdminEmail; FallbackRunner = $fbRunner }
    })

    if ($RunParallel) {
        # ── Parallel: each worker fetches one file ID and returns @{Id;Bytes}
        $driveResults = $driveWork | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
            $item    = $_
            $gamPath = $using:GamPath
            $outDir  = $using:OutputDir

            $bytes = [long]0
            $runnersToTry = @($item.AdminEmail)
            if ($item.FallbackRunner -ne $item.AdminEmail) { $runnersToTry += $item.FallbackRunner }

            foreach ($runner in $runnersToTry) {
                $tmpF = Join-Path $outDir "tmp_drv_$(New-Guid).csv"
                & $gamPath redirect csv $tmpF user $runner print filelist `
                    select drivefileid $item.Id fields "id,name,size,mimetype" 2>$null
                if (Test-Path $tmpF) {
                    $rows = @(Import-Csv $tmpF)
                    Remove-Item $tmpF -Force -ErrorAction SilentlyContinue
                    foreach ($r in $rows) {
                        $rawProp = $r.PSObject.Properties['size']
                        if ($rawProp -and $rawProp.Value -match '^\d+$') {
                            $bytes = [long]$rawProp.Value; break
                        }
                    }
                }
                if ($bytes -gt 0) { break }
            }
            [PSCustomObject]@{ Id = $item.Id; Bytes = $bytes }
        }

        foreach ($r in @($driveResults)) {
            if ($r) { $driveSizeMap[$r.Id] = $r.Bytes }
        }

    } else {
        # ── Sequential fallback (PS5) ─────────────────────────────────────────
        $tmpDriveCsv = Join-Path $OutputDir "tmp_driveinfo_$Timestamp.csv"
        $fi = 0
        foreach ($item in $driveWork) {
            $fi++
            if ($fi % 10 -eq 0 -or $fi -eq 1) {
                Write-Host ("   Drive lookup {0}/{1}..." -f $fi, $driveWork.Count) -ForegroundColor DarkGray
            }
            if (Test-Path $tmpDriveCsv) { Remove-Item $tmpDriveCsv -Force }

            $runnersToTry = @($item.AdminEmail)
            if ($item.FallbackRunner -ne $item.AdminEmail) { $runnersToTry += $item.FallbackRunner }

            foreach ($runner in $runnersToTry) {
                & $GamPath redirect csv $tmpDriveCsv user $runner print filelist `
                    select drivefileid $item.Id fields "id,name,size,mimetype" 2>$null
                if (Test-Path $tmpDriveCsv) {
                    $fileRows = @(Import-Csv $tmpDriveCsv)
                    foreach ($dr in $fileRows) {
                        $fid  = Get-Col $dr "id","Owner.id"
                        if (-not $fid) { $fid = $item.Id }
                        $raw  = Get-Col $dr "size"
                        $driveSizeMap[$fid] = if ($raw -match '^\d+$') { [long]$raw } else { [long]0 }
                    }
                    Remove-Item $tmpDriveCsv -Force -ErrorAction SilentlyContinue
                }
                if ($driveSizeMap.ContainsKey($item.Id)) { break }
            }
        }
        if (Test-Path $tmpDriveCsv) { Remove-Item $tmpDriveCsv -Force -ErrorAction SilentlyContinue }
    }

    # Write sizes back to detail rows (single-threaded — safe mutable access)
    $notFound = 0
    foreach ($entry in $driveIds) {
        $fid = $entry.DriveFileID
        if ($driveSizeMap.ContainsKey($fid)) {
            $bytes = $driveSizeMap[$fid]
            $entry.Row.Bytes  = $bytes
            $entry.Row.SizeMB = [Math]::Round($bytes / 1MB, 2)
            $entry.Row.IsLarge= ($bytes -gt ($LargeFileMB * 1MB))
            if (-not $driveSizeMap[$fid]) { $entry.Row.Note = "size=0 - may be Google Workspace native file" }
        } else {
            $notFound++
            $entry.Row.Note = "Drive file not accessible by admin"
        }
    }

    Write-Host "   Drive sizes resolved : $($driveSizeMap.Count) / $($uniqueDriveIds.Count)" -ForegroundColor Green
    if ($notFound -gt 0) {
        Write-Host "   Not accessible (Shared Drive / permission) : $notFound" -ForegroundColor DarkYellow
    }
} else {
    Write-Host "   No Drive attachments found - skipping." -ForegroundColor DarkGray
}

# =============================================================================
# PHASE 3b – Size uploaded/inline attachments
#
# HOW SIZE IS OBTAINED (no full file download in either path):
#
#   Stage i  — Drive metadata search  (zero bytes from attachment URL)
#     Chat-uploaded files are stored in the sender's Google Drive under the
#     "Chat Files" folder.  We search by filename + MIME type and read the
#     Drive 'size' field — pure API metadata, nothing downloaded.
#
#   Stage ii — Range GET fallback  (1 byte in memory, never saved to disk)
#     For files not found via Drive search (e.g. already deleted from Drive,
#     or name collision).  "GET Range: bytes=0-0" triggers HTTP 206 +
#     Content-Range: bytes 0-0/<total> from which we extract the real size.
#     Pass -MetadataOnly to skip this stage entirely.
# =============================================================================
Write-Host ""
Write-Host "[3b/3] Sizing uploaded/inline attachments..." -ForegroundColor Yellow

$driveDone = 0; $headDone = 0; $headFailed = 0; $headSkipped = 0

if ($uploadedRows.Count -gt 0) {
    $toProcess = $uploadedRows
    if ($toProcess.Count -gt $MaxHeadRequests) {
        Write-Host ("   Capping at {0} (of {1} total) to avoid rate-limit" `
            -f $MaxHeadRequests, $toProcess.Count) -ForegroundColor DarkYellow
        $toProcess = $toProcess | Select-Object -First $MaxHeadRequests
        $headSkipped = $uploadedRows.Count - $MaxHeadRequests
    }

    # ── Stage i: Drive metadata search ──────────────────────────────────────
    Write-Host "   Stage i : Drive metadata search (no download)..." -ForegroundColor DarkGray

    $tmpSearch   = Join-Path $OutputDir "tmp_uploadsearch_$Timestamp.csv"
    $needRangeGet = [System.Collections.Generic.List[object]]::new()

    foreach ($entry in $toProcess) {
        $fname  = $entry.Row.AttachName
        $ftype  = $entry.Row.ContentType
        $runner = if ($spacemap.ContainsKey($entry.Row.SpaceID)) {
                      $spacemap[$entry.Row.SpaceID].Runner } else { $AdminEmail }

        if (-not $fname -or -not $ftype) { $needRangeGet.Add($entry); continue }

        # Drive query: name + MIME type match, not trashed.
        # Escape single quotes for Drive query syntax ( ' -> \' )
        $safeName  = $fname -replace "'", "\'"
        $safeType  = $ftype -replace "'", "\'"
        $driveQuery = "name = '$safeName' and mimeType = '$safeType' and trashed = false"

        if (Test-Path $tmpSearch) { Remove-Item $tmpSearch -Force }
        & $GamPath redirect csv $tmpSearch user $runner print filelist `
            query $driveQuery fields "id,name,size,mimetype,createdTime" 2>$null

        $matched = $false
        if (Test-Path $tmpSearch) {
            $hits = @(Import-Csv $tmpSearch)
            Remove-Item $tmpSearch -Force -ErrorAction SilentlyContinue

            if ($hits.Count -ge 1) {
                # If multiple name+type matches pick the one closest to message time
                $best = if ($hits.Count -eq 1) { $hits[0] } else {
                    $msgTime = try { [datetime]$entry.Row.CreateTime } catch { [datetime]::UtcNow }
                    $hits | Sort-Object {
                        $ct = try { [datetime](Get-Col $_ "createdTime") } catch { [datetime]::MinValue }
                        [Math]::Abs(($ct - $msgTime).TotalSeconds)
                    } | Select-Object -First 1
                }
                $raw = Get-Col $best "size"
                if ($raw -match '^\d+$') {
                    $bytes = [long]$raw
                    $entry.Row.Bytes   = $bytes
                    $entry.Row.SizeMB  = [Math]::Round($bytes / 1MB, 2)
                    $entry.Row.IsLarge = ($bytes -gt ($LargeFileMB * 1MB))
                    $entry.Row.Note    = "Size via Drive metadata"
                    $driveDone++
                    $matched = $true
                }
            }
        }
        if (-not $matched) { $needRangeGet.Add($entry) }
    }
    if (Test-Path $tmpSearch) { Remove-Item $tmpSearch -Force -ErrorAction SilentlyContinue }
    Write-Host ("   Drive search resolved : {0} / {1}" -f $driveDone, $toProcess.Count) `
        -ForegroundColor Green

    # ── Stage ii: Range GET fallback (skip when -MetadataOnly) ──────────────
    if ($needRangeGet.Count -gt 0 -and -not $UseMetadataOnly) {
        Write-Host ("   Stage ii: Range GET fallback for {0} file(s) not found in Drive..." `
            -f $needRangeGet.Count) -ForegroundColor DarkGray

        Add-Type -AssemblyName System.Net.Http
        $handler = [System.Net.Http.HttpClientHandler]::new()
        $handler.AllowAutoRedirect = $true
        $client  = [System.Net.Http.HttpClient]::new($handler)
        $client.Timeout = [TimeSpan]::FromSeconds($HeadTimeoutSec)

        foreach ($entry in $needRangeGet) {
            try {
                $req = [System.Net.Http.HttpRequestMessage]::new(
                            [System.Net.Http.HttpMethod]::Get, $entry.Uri)
                $req.Headers.TryAddWithoutValidation("Range", "bytes=0-0") | Out-Null
                $resp = $client.SendAsync($req,
                            [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead).GetAwaiter().GetResult()

                $bytes = [long]0
                if ([int]$resp.StatusCode -eq 206) {
                    $cr = $resp.Content.Headers.ContentRange
                    if ($cr -ne $null -and $cr.Length.HasValue) { $bytes = [long]$cr.Length.Value }
                }
                if ($bytes -eq 0) {
                    $cl = $resp.Content.Headers.ContentLength
                    if ($cl -and $cl -gt 0) { $bytes = [long]$cl }
                }
                # Discard the at-most-1-byte body to free the connection
                try { $resp.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult() | Out-Null } catch {}
                $resp.Dispose()

                if ($bytes -gt 0) {
                    $entry.Row.Bytes   = $bytes
                    $entry.Row.SizeMB  = [Math]::Round($bytes / 1MB, 2)
                    $entry.Row.IsLarge = ($bytes -gt ($LargeFileMB * 1MB))
                    $entry.Row.Note    = "Size via Range GET"
                    $headDone++
                } else {
                    $entry.Row.Note = "Size unknown: no Content-Range/Length from Range GET"
                    $headFailed++
                }
            } catch {
                $entry.Row.Note = "Range GET failed: $($_.Exception.Message -replace '`n',' ')"
                $headFailed++
            }
        }
        $client.Dispose()

        Write-Host ("   Range GET resolved  : {0}  failed : {1}" -f $headDone, $headFailed) `
            -ForegroundColor $(if($headFailed -gt 0){"DarkYellow"}else{"Green"})

    } elseif ($needRangeGet.Count -gt 0 -and $UseMetadataOnly) {
        Write-Host ("   {0} file(s) not found in Drive - size stays 0 (-MetadataOnly is set)." `
            -f $needRangeGet.Count) -ForegroundColor DarkYellow
        foreach ($entry in $needRangeGet) {
            $entry.Row.Note = "Size unknown: not found in Drive (MetadataOnly mode)"
        }
    }

    if ($headSkipped -gt 0) {
        Write-Host "   Skipped (cap) : $headSkipped - re-run with higher MaxHeadRequests" -ForegroundColor DarkYellow
    }
} else {
    Write-Host "   No uploaded attachments found - skipping." -ForegroundColor DarkGray
}

# =============================================================================
# OUTPUT – Export CSVs
# =============================================================================
Write-Host ""
Write-Host "Exporting reports..." -ForegroundColor Yellow

# 1. Full detail (one row per attachment)
$detailRows | Sort-Object Bytes -Descending |
    Export-Csv $DetailCsv -NoTypeInformation
Write-Host "   Detail CSV       -> $DetailCsv"

# 2. Per-Space aggregated summary
$perSpaceRows = $detailRows |
    Where-Object { $_.Source -ne "GIPHY_TENOR" } |   # exclude pure-gif rows from byte totals
    Group-Object -Property SpaceID |
    ForEach-Object {
        $grp   = $_.Group
        $first = $grp[0]
        $totalBytes    = ($grp | Measure-Object Bytes -Sum).Sum
        $driveBytes    = ($grp | Where-Object { $_.Source -eq "DRIVE_FILE"       } | Measure-Object Bytes -Sum).Sum
        $uploadBytes   = ($grp | Where-Object { $_.Source -eq "UPLOADED_CONTENT" } | Measure-Object Bytes -Sum).Sum
        $driveCount    = ($grp | Where-Object { $_.Source -eq "DRIVE_FILE" }).Count
        $uploadCount   = ($grp | Where-Object { $_.Source -eq "UPLOADED_CONTENT" }).Count
        # Image count + bytes: all images regardless of source (uploaded PNG/JPG/GIF + Drive images)
        $imageRows     = $grp | Where-Object { $_.IsInlineImage -eq $true }
        $imageCount    = $imageRows.Count
        $imageBytes    = ($imageRows | Measure-Object Bytes -Sum).Sum
        $gifCount      = ($detailRows | Where-Object { $_.SpaceID -eq $first.SpaceID -and $_.Source -eq "GIPHY_TENOR" } |
                          Select-Object -ExpandProperty Note | ForEach-Object {
                              if ($_ -match "^(\d+) gif") { [int]$Matches[1] } else { 0 }
                          } | Measure-Object -Sum).Sum
        $largeCount    = ($grp | Where-Object { $_.IsLarge -eq $true }).Count
        $largest       = $grp | Sort-Object Bytes -Descending | Select-Object -First 1

        [PSCustomObject]@{
            SpaceID                  = $first.SpaceID
            SpaceName                = $first.SpaceName
            SpaceType                = $first.SpaceType
            MemberCount              = $first.MemberCount
            SpaceUri                 = $spacemap[$first.SpaceID].SpaceUri
            TotalAttachmentCount     = $grp.Count
            Drive_AttachCount        = $driveCount
            Uploaded_AttachCount     = $uploadCount
            Image_Count              = $imageCount
            Image_TotalBytes         = $imageBytes
            Image_TotalMB            = [Math]::Round($imageBytes / 1MB, 2)
            AnimatedGif_Count        = [int]$gifCount
            TotalBytes               = $totalBytes
            TotalSizeMB              = [Math]::Round($totalBytes / 1MB, 2)
            Drive_TotalBytes         = $driveBytes
            Drive_TotalMB            = [Math]::Round($driveBytes / 1MB, 2)
            Uploaded_TotalBytes      = $uploadBytes
            Uploaded_TotalMB         = [Math]::Round($uploadBytes / 1MB, 2)
            LargeAttachments_Count   = $largeCount
            Largest_AttachName       = $largest.AttachName
            Largest_Bytes            = $largest.Bytes
            Largest_SizeMB           = $largest.SizeMB
            Largest_Source           = $largest.Source
            Largest_MessageID        = $largest.MessageID
        }
    } | Sort-Object TotalBytes -Descending

$perSpaceRows | Export-Csv $PerSpaceCsv -NoTypeInformation
Write-Host "   Per-Space CSV    -> $PerSpaceCsv"

# 3. Top 100 largest individual attachments
$detailRows |
    Where-Object { $_.Bytes -gt 0 } |
    Sort-Object Bytes -Descending |
    Select-Object -First 100 |
    Export-Csv $TopCsv -NoTypeInformation
Write-Host "   Top Largest CSV  -> $TopCsv"

# =============================================================================
# CONSOLE SUMMARY
# =============================================================================
$grandTotalBytes  = ($detailRows | Measure-Object Bytes -Sum).Sum
$grandTotalCount  = ($detailRows | Where-Object { $_.Source -ne "GIPHY_TENOR" }).Count
$driveTotal       = ($detailRows | Where-Object { $_.Source -eq "DRIVE_FILE"       } | Measure-Object Bytes -Sum).Sum
$uploadTotal      = ($detailRows | Where-Object { $_.Source -eq "UPLOADED_CONTENT" } | Measure-Object Bytes -Sum).Sum
$imageRows_all    = $detailRows  | Where-Object { $_.IsInlineImage -eq $true -and $_.Source -ne "GIPHY_TENOR" }
$imageTotal       = ($imageRows_all | Measure-Object Bytes -Sum).Sum
$imageTotalCount  = $imageRows_all.Count
$largeTotal       = ($detailRows | Where-Object { $_.IsLarge -eq $true }).Count
$spacesWithAttach = ($detailRows | Select-Object -ExpandProperty SpaceID -Unique).Count
$top5             = $detailRows | Where-Object { $_.Bytes -gt 0 } | Sort-Object Bytes -Descending | Select-Object -First 5

Write-Host ""
Write-Host ("=" * 60) -ForegroundColor Cyan
Write-Host "              ATTACHMENT AUDIT COMPLETE" -ForegroundColor Green
Write-Host ("=" * 60) -ForegroundColor Cyan
Write-Host ("Spaces scanned          : {0}"   -f $uniqueSpaces.Count)
Write-Host ("Spaces with attachments : {0}"   -f $spacesWithAttach)
Write-Host ("Total attachments found : {0}"   -f $grandTotalCount)
Write-Host ("  Drive-source          : {0}"   -f ($driveIds.Count))
Write-Host ("  Uploaded/inline       : {0}"   -f ($uploadedRows.Count))
Write-Host ("  Images (inline/Drive) : {0}"   -f $imageTotalCount)
Write-Host ("  Flagged large (>{0} MB): {1}"  -f $LargeFileMB, $largeTotal) `
    -ForegroundColor $(if($largeTotal -gt 0){"Yellow"}else{"Green"})
Write-Host ""
Write-Host ("Total size (Drive)      : {0:N2} MB"  -f ($driveTotal  / 1MB))
Write-Host ("Total size (Uploaded)   : {0:N2} MB"  -f ($uploadTotal / 1MB))
Write-Host ("  Sized via Drive meta  : {0}"         -f $driveDone)
if (-not $UseMetadataOnly) {
Write-Host ("  Sized via Range GET   : {0}"         -f $headDone) }
Write-Host ("Total size (Images)     : {0:N2} MB"  -f ($imageTotal  / 1MB)) `
    -ForegroundColor $(if($imageTotal -gt 0){"Cyan"}else{"Gray"})
Write-Host ("Grand total sized       : {0:N2} MB"  -f ($grandTotalBytes / 1MB))
Write-Host ""
Write-Host "TOP 5 LARGEST ATTACHMENTS:" -ForegroundColor Cyan
$rank = 1
foreach ($t in $top5) {
    Write-Host ("  #{0}  {1:N1} MB  [{2}]  {3}  in {4}" -f `
        $rank++,
        ($t.Bytes / 1MB),
        $t.Source,
        $(if($t.AttachName){"'$($t.AttachName)'"}else{"(unnamed)"}),
        $t.SpaceName
    )
}
Write-Host ""
Write-Host "Reports:" -ForegroundColor Cyan
Write-Host "  Detail     -> $DetailCsv"
Write-Host "  Per-Space  -> $PerSpaceCsv"
Write-Host "  Top 100    -> $TopCsv"
Write-Host ("=" * 60) -ForegroundColor Cyan
