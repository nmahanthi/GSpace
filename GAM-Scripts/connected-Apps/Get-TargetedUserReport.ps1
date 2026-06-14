#Requires -Version 5.1
<#
.SYNOPSIS
    Targeted Connected Apps & Chat Audit - scan only a selected list of users.

.DESCRIPTION
    Designed for large Google Workspace tenants (100k+ users) where you do NOT
    want a full-domain sweep every time. Supply 1-N user emails and the script:

      [1] OAuth tokens  - third-party apps each target user has authorised
      [2] App summary   - distinct apps across the target set, with user counts
      [3] Chat Spaces   - spaces the target users belong to       (-IncludeChatSpaces)
      [4] Chat Bots     - bot/app activity involving target users  (-IncludeChatBots)

    Outputs: per-section CSVs  +  a self-contained HTML report.

.PARAMETER Users
    One or more primary email addresses.
    Example: -Users "alice@corp.com","bob@corp.com"

.PARAMETER UsersFile
    Path to a CSV file with a 'primaryEmail' column  OR  a plain-text file
    with one email per line. Use instead of, or in addition to, -Users.

.PARAMETER GamPath
    Full path to gam / gam.exe. Auto-detected from common locations if omitted.

.PARAMETER OutputDir
    Folder for all output files. Defaults to .\TargetedReport_<timestamp>

.PARAMETER IncludeChatSpaces
    Enumerate Chat Spaces each target user belongs to.
    Requires GAM7 / GAMADV-XTD3 and the Google Chat API admin scope.

.PARAMETER IncludeChatBots
    Pull the tenant-wide Chat audit log (last -ChatDaysAgo days) and filter
    for bot/app events where the actor or target is a selected user.
    Requires GAM7 / GAMADV-XTD3 and the Reports API scope.

.PARAMETER ChatDaysAgo
    How many days back to look in the Chat audit log.  Default: 30.

.PARAMETER DwdTimeoutSeconds
    Per-strategy timeout (seconds) for background GAM jobs.  Default: 90.

.EXAMPLE
    # Scan 3 users - connected apps only
    .\Get-TargetedUserReport.ps1 -Users "alice@corp.com","bob@corp.com","carol@corp.com"

.EXAMPLE
    # Load from a CSV file, include Chat data, look back 60 days
    .\Get-TargetedUserReport.ps1 -UsersFile ".\targets.csv" -IncludeChatSpaces -IncludeChatBots -ChatDaysAgo 60

.EXAMPLE
    # Specify a custom GAM path and output folder
    .\Get-TargetedUserReport.ps1 -Users "user@domain.com" -GamPath "C:\GAM7\gam.exe" -OutputDir "C:\Reports\Targeted"
#>

[CmdletBinding()]
param(
    [string[]]$Users          = @(),
    [string]  $UsersFile      = "",
    [string]  $GamPath        = "",
    [string]  $OutputDir      = "",
    [switch]  $IncludeChatSpaces,
    [switch]  $IncludeChatBots,
    [int]     $ChatDaysAgo    = 30,
    [int]     $DwdTimeoutSeconds = 90
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

# ======================================================
#  HELPERS
# ======================================================
function Write-Banner { param([string]$t); $l = "=" * 64; Write-Host "`n$l`n  $t`n$l" -ForegroundColor Cyan }
function Write-Step   { param([string]$t); Write-Host "  [$(Get-Date -f 'HH:mm:ss')] $t" -ForegroundColor Yellow }
function Write-OK     { param([string]$t); Write-Host "  [OK]   $t" -ForegroundColor Green }
function Write-Warn   { param([string]$t); Write-Host "  [WARN] $t" -ForegroundColor DarkYellow }
function Write-Fail   { param([string]$t); Write-Host "  [FAIL] $t" -ForegroundColor Red }

function SafeCsv {
    param([string]$Path)
    if (Test-Path $Path) {
        $raw = Get-Content $Path -Raw -ErrorAction SilentlyContinue
        if ($raw -and $raw.Trim().Length -gt 5) {
            try { return Import-Csv $Path } catch {}
        }
    }
    return @()
}

function Invoke-GamTimeout {
    param([string[]]$GamArgs, [int]$TimeoutSec = 90)
    $bin = $script:GAM
    $job = Start-Job -ScriptBlock { param($b,$a); & $b @a 2>&1 } -ArgumentList $bin, (, $GamArgs)
    $done = Wait-Job $job -Timeout $TimeoutSec
    if ($done) { $out = Receive-Job $job; Remove-Job $job -Force; return $out }
    Stop-Job $job; Remove-Job $job -Force; return $null
}

# ======================================================
#  FIND GAM
# ======================================================
Clear-Host
Write-Host @"

  +----------------------------------------------------------+
  |  Google Workspace - Targeted Connected Apps & Chat Audit |
  |  Powered by GAM / GAMADV-XTD3 / GAM7                    |
  +----------------------------------------------------------+

"@ -ForegroundColor Cyan

Write-Banner "0. Locating GAM"

$candidates = @(
    $GamPath,
    ".\gam.exe", ".\gam",
    "$PSScriptRoot\gam.exe", "$PSScriptRoot\gam",
    "$PWD\gam.exe", "$PWD\gam",
    "gam",
    "$env:USERPROFILE\AppData\Local\GAM7\gam.exe",
    "$env:USERPROFILE\AppData\Local\GAMADV-XTD3\gam.exe",
    "C:\GAM7\gam.exe", "C:\GAMADV-XTD3\gam.exe",
    "C:\GAM6\gam.exe", "C:\GAM\gam.exe"
)

$script:GAM = $null
foreach ($c in $candidates) {
    if (-not $c) { continue }
    $found = Get-Command $c -ErrorAction SilentlyContinue
    if ($found) { $script:GAM = $found.Source; break }
    if (Test-Path $c) { $script:GAM = (Resolve-Path $c).Path; break }
}

if (-not $script:GAM) {
    Write-Fail "GAM not found. Pass -GamPath or run from inside your GAM folder."
    exit 1
}

$verLines   = & $script:GAM version 2>&1 | Select-Object -First 5
$isGAM7     = @($verLines | Where-Object { $_ -match "GAM7|GAMADV" }).Count -gt 0
$verString  = ($verLines | Where-Object { $_ -match "\d+\.\d+" } | Select-Object -First 1) -replace "^\s+", ""
Write-OK "GAM binary : $script:GAM"
Write-OK "Version    : $verString  $(if($isGAM7){'[GAM7/GAMADV-XTD3]'}else{'[Standard GAM]'})"

if (($IncludeChatSpaces -or $IncludeChatBots) -and -not $isGAM7) {
    Write-Warn "Chat features require GAM7/GAMADV-XTD3 - disabling -IncludeChatSpaces / -IncludeChatBots."
    $IncludeChatSpaces = [switch]$false
    $IncludeChatBots   = [switch]$false
}

if (-not $OutputDir) {
    $OutputDir = Join-Path $PWD "TargetedReport_$(Get-Date -f 'yyyyMMdd_HHmmss')"
}
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
Write-OK "Output dir : $OutputDir"
$T0 = Get-Date

# ======================================================
#  SECTION 0 - LOAD TARGET USERS
# ======================================================
Write-Banner "0. Loading Target Users"

$targetEmails = [System.Collections.Generic.List[string]]::new()

# --- From -Users parameter ---
foreach ($e in $Users) {
    $e = $e.Trim()
    if ($e -and $e -match "@") { $targetEmails.Add($e.ToLower()) }
}

# --- From -UsersFile (CSV with primaryEmail, or plain text one-per-line) ---
if ($UsersFile -and (Test-Path $UsersFile)) {
    $ext = [IO.Path]::GetExtension($UsersFile).ToLower()
    if ($ext -eq ".csv") {
        $fileRows = Import-Csv $UsersFile -ErrorAction SilentlyContinue
        if ($fileRows) {
            # Accept 'primaryEmail' OR 'email' OR first column
            $col = $fileRows[0].PSObject.Properties.Name |
                   Where-Object { $_ -match "primaryEmail|email" } |
                   Select-Object -First 1
            if (-not $col) { $col = $fileRows[0].PSObject.Properties.Name | Select-Object -First 1 }
            foreach ($r in $fileRows) {
                $e = "$($r.$col)".Trim().ToLower()
                if ($e -and $e -match "@") { $targetEmails.Add($e) }
            }
        }
    } else {
        # Plain text - one email per line
        Get-Content $UsersFile | ForEach-Object {
            $e = $_.Trim().ToLower()
            if ($e -and $e -match "@") { $targetEmails.Add($e) }
        }
    }
    Write-OK "Loaded users from file: $UsersFile"
}

# Deduplicate
$targetEmails = @($targetEmails | Select-Object -Unique)

if ($targetEmails.Count -eq 0) {
    Write-Fail "No valid email addresses found. Use -Users or -UsersFile."
    Write-Host ""
    Write-Host "  Examples:" -ForegroundColor Yellow
    Write-Host "    .\Get-TargetedUserReport.ps1 -Users 'alice@corp.com','bob@corp.com'" -ForegroundColor White
    Write-Host "    .\Get-TargetedUserReport.ps1 -UsersFile '.\targets.csv' -IncludeChatSpaces" -ForegroundColor White
    exit 1
}

Write-OK "$($targetEmails.Count) target user(s) loaded:"
$targetEmails | ForEach-Object { Write-Host "    $_" -ForegroundColor Gray }

# Write target users to a temp CSV for GAM multiprocess commands
$targetCsv = Join-Path $OutputDir "_target_users.csv"
$targetEmails | ForEach-Object { [PSCustomObject]@{ primaryEmail = $_ } } |
    Export-Csv $targetCsv -NoTypeInformation

# ======================================================
#  SECTION 1 - OAUTH TOKENS (CONNECTED APPS)
# ======================================================
Write-Banner "1. OAuth Tokens (Connected Apps) for Target Users"

$tokensFile = Join-Path $OutputDir "tokens_target_users.csv"
Write-Step "Fetching OAuth tokens for $($targetEmails.Count) user(s) via GAM multiprocess..."

& $script:GAM redirect csv $tokensFile multiprocess csv $targetCsv `
    gam user "~primaryEmail" print tokens 2>$null | Out-Null

$tokenData = SafeCsv $tokensFile

# Auto-detect column names (GAM version differences)
$colUser = "user"; $colApp = "displayText"
if ($tokenData.Count -gt 0) {
    $sp = $tokenData[0].PSObject.Properties.Name
    if ($sp -contains "userEmail") { $colUser = "userEmail" }
    elseif ($sp -contains "user") { $colUser = "user" }
    if ($sp -contains "displayText") { $colApp = "displayText" }
    elseif ($sp -contains "appName") { $colApp = "appName" }
    Write-OK "Token column names detected: user='$colUser'  app='$colApp'"
}
Write-OK "$($tokenData.Count) token record(s) retrieved."

# --- Per-user token summary ---
$perUserSummary = @(if ($tokenData.Count -gt 0) {
    $tokenData | Group-Object -Property { $_.$colUser } | ForEach-Object {
        [PSCustomObject]@{
            User     = $_.Name
            AppCount = $_.Count
            Apps     = (@($_.Group | ForEach-Object { $_.$colApp } | Select-Object -Unique) -join " | ")
        }
    } | Sort-Object AppCount -Descending
})

# --- App aggregation across target users ---
$appSummary = @(if ($tokenData.Count -gt 0) {
    $tokenData |
    Where-Object { $_.$colApp -and $_.$colApp.Trim() -ne "" } |
    Group-Object -Property { $_.$colApp } |
    ForEach-Object {
        $grp      = $_.Group
        $first    = $grp | Select-Object -First 1
        $cid      = if ($first.PSObject.Properties['clientId'])  { $first.clientId }  else { "N/A" }
        $scopes   = if ($first.PSObject.Properties['scopes'])    { $first.scopes }    else { "N/A" }
        $usrList  = @($grp | ForEach-Object { $_.$colUser } | Select-Object -Unique | Sort-Object)
        [PSCustomObject]@{
            AppName    = $_.Name
            ClientId   = $cid
            UserCount  = $usrList.Count
            Users      = $usrList -join "; "
            Scopes     = $scopes
        }
    } | Sort-Object UserCount -Descending
})

$perUserSummary | Export-Csv (Join-Path $OutputDir "tokens_per_user.csv") -NoTypeInformation -Force
$appSummary     | Export-Csv (Join-Path $OutputDir "apps_aggregated.csv") -NoTypeInformation -Force
Write-OK "$($appSummary.Count) distinct connected app(s) across all target users."

# ======================================================
#  SECTION 2 - CHAT SPACES (optional)
# ======================================================
$chatSpaceData  = @()
$chatMemberData = @()
Write-Banner "2. Chat Spaces$(if(-not $IncludeChatSpaces){' [SKIPPED - use -IncludeChatSpaces]'})"

if ($IncludeChatSpaces) {
    $chatSpacesFile  = Join-Path $OutputDir "chat_spaces_target.csv"
    $chatMembersFile = Join-Path $OutputDir "chat_members_target.csv"

    Write-Step "Fetching Chat Spaces for each target user..."

    # Collect spaces per user and merge
    $allSpaceRows = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($email in $targetEmails) {
        Write-Step "  -> $email"
        $tmpFile = Join-Path $OutputDir "_spaces_$($email -replace '@','_at_').csv"
        & $script:GAM redirect csv $tmpFile user $email print chatspaces 2>$null | Out-Null
        $rows = SafeCsv $tmpFile
        foreach ($r in $rows) {
            # Tag which target user this came from
            $r | Add-Member -NotePropertyName "TargetUser" -NotePropertyValue $email -Force
            $allSpaceRows.Add($r)
        }
        Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue
    }
    $chatSpaceData = @($allSpaceRows)
    if ($chatSpaceData.Count -gt 0) {
        $chatSpaceData | Export-Csv $chatSpacesFile -NoTypeInformation -Force
    }
    Write-OK "$($chatSpaceData.Count) Chat Space membership record(s) across target users."

    # Fetch members of each unique space found
    $uniqueSpaceNames = @($chatSpaceData | Where-Object { $_.name } |
                          Select-Object -ExpandProperty name -Unique)
    if ($uniqueSpaceNames.Count -gt 0) {
        Write-Step "Fetching members for $($uniqueSpaceNames.Count) unique space(s)..."
        $allMemberRows = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach ($spaceName in $uniqueSpaceNames) {
            $tmpMem = Join-Path $OutputDir "_members_tmp.csv"
            & $script:GAM redirect csv $tmpMem print chatmembers $spaceName asadmin 2>$null | Out-Null
            $mrows = SafeCsv $tmpMem
            foreach ($mr in $mrows) {
                $mr | Add-Member -NotePropertyName "SpaceResourceName" -NotePropertyValue $spaceName -Force
                $allMemberRows.Add($mr)
            }
            Remove-Item $tmpMem -Force -ErrorAction SilentlyContinue
        }
        $chatMemberData = @($allMemberRows)
        if ($chatMemberData.Count -gt 0) {
            $chatMemberData | Export-Csv $chatMembersFile -NoTypeInformation -Force
        }
        Write-OK "$($chatMemberData.Count) Chat membership record(s) retrieved."
    }
} else {
    Write-Warn "Skipped. Re-run with -IncludeChatSpaces to enumerate Chat Spaces."
}

# ======================================================
#  SECTION 3 - CHAT BOT ACTIVITY (optional)
# ======================================================
$botEventData    = @()
$botSummaryData  = @()
Write-Banner "3. Chat Bot Activity$(if(-not $IncludeChatBots){' [SKIPPED - use -IncludeChatBots]'})"

if ($IncludeChatBots) {
    $chatRawFile    = Join-Path $OutputDir "chat_report_raw.csv"
    $botEventsFile  = Join-Path $OutputDir "chat_bot_events_target.csv"
    $botSummaryFile = Join-Path $OutputDir "chat_bot_summary_target.csv"

    Write-Step "Pulling Chat audit report (last $ChatDaysAgo days) via Reports API..."
    $chatRaw = Invoke-GamTimeout @("redirect","csv",$chatRawFile,"report","chat","daysago","$ChatDaysAgo") ($DwdTimeoutSeconds * 2)

    if ($null -eq $chatRaw) {
        Write-Warn "Chat report timed out. Try increasing -DwdTimeoutSeconds."
    } else {
        $allChatRows = SafeCsv $chatRawFile
        Write-OK "$($allChatRows.Count) total Chat report rows retrieved."

        if ($allChatRows.Count -gt 0) {
            # Build a fast lookup set of target emails (lower-case)
            $targetSet = @{}
            $targetEmails | ForEach-Object { $targetSet[$_] = $true }

            # Helper - safe column reader
            function Get-ChatCol($row, [string[]]$names) {
                foreach ($n in $names) {
                    $v = $row.PSObject.Properties[$n]
                    if ($v -and "$($v.Value)".Trim() -ne "") { return "$($v.Value)".Trim() }
                }
                return ""
            }

            $filteredEvents = [System.Collections.Generic.List[PSCustomObject]]::new()

            foreach ($row in $allChatRows) {
                $actorEmail = Get-ChatCol $row "actor.email","actor"
                $targetUser = Get-ChatCol $row "resourceDetails.1.ownerDetails.ownerIdentity.0.userIdentity.userEmail","target_users"
                $eventName  = Get-ChatCol $row "events.name","event.name","name"
                $eventTime  = Get-ChatCol $row "id.time","time","eventTime"
                $spaceId    = Get-ChatCol $row "resourceDetails.0.id","room_id"
                $spaceName  = Get-ChatCol $row "resourceDetails.0.title","room_name"
                $res1Type   = Get-ChatCol $row "resourceDetails.1.type"
                $res1Title  = Get-ChatCol $row "resourceDetails.1.title"
                $res1Id     = Get-ChatCol $row "resourceDetails.1.id"

                $appName    = if ($res1Type -eq "APPLICATION") { $res1Title } else { "" }
                $appId      = if ($res1Type -eq "APPLICATION") { $res1Id }   else { "" }
                $isBotEvent = ($eventName -match "app|bot") -or ($res1Type -eq "APPLICATION")

                # Only keep rows involving a target user
                $actorLow  = $actorEmail.ToLower()
                $targetLow = $targetUser.ToLower()
                $involved  = $targetSet.ContainsKey($actorLow) -or $targetSet.ContainsKey($targetLow)

                if ($involved -and $isBotEvent) {
                    $filteredEvents.Add([PSCustomObject]@{
                        EventTime        = $eventTime
                        EventName        = $eventName
                        TargetUserInvolved = if ($targetSet.ContainsKey($actorLow)) { $actorEmail } else { $targetUser }
                        ActorEmail       = $actorEmail
                        AppName          = $appName
                        AppId            = $appId
                        SpaceID          = $spaceId
                        SpaceName        = $spaceName
                        ResourceType     = $res1Type
                        IsBotEvent       = $isBotEvent
                    })
                }
            }

            $botEventData = @($filteredEvents)
            Write-OK "$($botEventData.Count) bot event(s) involving target users."

            if ($botEventData.Count -gt 0) {
                $botEventData | Export-Csv $botEventsFile -NoTypeInformation -Force

                # Per-bot summary for target users
                $botSummaryData = @(
                    $botEventData |
                    Where-Object { $_.AppId -ne "" } |
                    Group-Object -Property AppId |
                    ForEach-Object {
                        $g = $_.Group
                        [PSCustomObject]@{
                            AppName          = $g[0].AppName
                            AppId            = $g[0].AppId
                            TotalEvents      = $g.Count
                            EventTypes       = (($g | Select-Object -ExpandProperty EventName -Unique) -join " | ")
                            TargetUsersActive= (($g | Select-Object -ExpandProperty TargetUserInvolved -Unique | Sort-Object) -join "; ")
                            UniqueSpaces     = ($g | Where-Object { $_.SpaceID } | Select-Object -ExpandProperty SpaceID -Unique).Count
                            SpaceNames       = (($g | Where-Object { $_.SpaceName } | Select-Object -ExpandProperty SpaceName -Unique) -join " | ")
                            FirstSeen        = ($g | Select-Object -ExpandProperty EventTime | Sort-Object | Select-Object -First 1)
                            LastSeen         = ($g | Select-Object -ExpandProperty EventTime | Sort-Object | Select-Object -Last 1)
                        }
                    } | Sort-Object TotalEvents -Descending
                )
                $botSummaryData | Export-Csv $botSummaryFile -NoTypeInformation -Force
                Write-OK "$($botSummaryData.Count) unique bot(s) active for target users."
            }
        }
        Remove-Item $chatRawFile -Force -ErrorAction SilentlyContinue
    }
} else {
    Write-Warn "Skipped. Re-run with -IncludeChatBots to include Chat bot/app activity."
}

# ======================================================
#  BUILD HTML REPORT
# ======================================================
Write-Banner "4. Building HTML Report"

Add-Type -AssemblyName System.Web

function To-HtmlTable {
    param([object[]]$Data, [int]$Limit = 500, [string]$Empty = "No data available.")
    if (-not $Data -or $Data.Count -eq 0) { return "<p class='empty'>$Empty</p>" }
    $rows  = if ($Data.Count -gt $Limit) { $Data | Select-Object -First $Limit } else { $Data }
    $props = $rows[0].PSObject.Properties.Name
    $sb    = [System.Text.StringBuilder]::new()
    [void]$sb.Append("<table><thead><tr>")
    foreach ($p in $props) { [void]$sb.Append("<th>$([System.Web.HttpUtility]::HtmlEncode($p))</th>") }
    [void]$sb.Append("</tr></thead><tbody>")
    foreach ($row in $rows) {
        [void]$sb.Append("<tr>")
        foreach ($p in $props) {
            $v = "$($row.$p)"
            if ($v.Length -gt 350) { $v = $v.Substring(0, 347) + "..." }
            [void]$sb.Append("<td>$([System.Web.HttpUtility]::HtmlEncode($v))</td>")
        }
        [void]$sb.Append("</tr>")
    }
    [void]$sb.Append("</tbody></table>")
    if ($Data.Count -gt $Limit) {
        [void]$sb.Append("<p class='note'>Showing first $Limit of $($Data.Count) rows. See CSV for full data.</p>")
    }
    return $sb.ToString()
}

# Pre-compute HTML fragments
$htmlTargetList = ($targetEmails | ForEach-Object { "<li>$([System.Web.HttpUtility]::HtmlEncode($_))</li>" }) -join "`n"

$htmlPerUser    = To-HtmlTable $perUserSummary  -Limit 200  -Empty "No OAuth tokens found for any target user."
$htmlApps       = To-HtmlTable $appSummary      -Limit 500  -Empty "No connected apps found for target users."
$htmlRawTokens  = To-HtmlTable $tokenData       -Limit 500  -Empty "No OAuth token records."

if ($IncludeChatSpaces -and $chatSpaceData.Count -gt 0) {
    $htmlSpaces = To-HtmlTable $chatSpaceData -Limit 300
} elseif ($IncludeChatSpaces) {
    $htmlSpaces = "<div class='warn'>Chat Spaces requested but no data returned. Check that the Google Chat API admin scope is authorised via <code>gam oauth update</code>.</div>"
} else {
    $htmlSpaces = "<div class='warn'>Not collected. Re-run with <strong>-IncludeChatSpaces</strong> to enumerate Chat Spaces for target users.</div>"
}

if ($IncludeChatBots -and $botEventData.Count -gt 0) {
    $htmlBotSummary = To-HtmlTable $botSummaryData -Limit 200
    $htmlBotEvents  = To-HtmlTable $botEventData   -Limit 500
} elseif ($IncludeChatBots) {
    $htmlBotSummary = "<div class='warn'>No bot events found involving target users in the last $ChatDaysAgo days.</div>"
    $htmlBotEvents  = $htmlBotSummary
} else {
    $htmlBotSummary = "<div class='warn'>Not collected. Re-run with <strong>-IncludeChatBots</strong> to include Chat bot activity.</div>"
    $htmlBotEvents  = $htmlBotSummary
}

# Stats
$reportDate   = Get-Date -Format "dddd dd MMM yyyy, HH:mm:ss"
$elapsed      = [math]::Round(((Get-Date) - $T0).TotalSeconds, 1)
$totalTargets = $targetEmails.Count
$totalApps    = $appSummary.Count
$totalTokens  = $tokenData.Count
$totalSpaces  = $chatSpaceData.Count
$totalBotEvts = $botEventData.Count
$totalBots    = $botSummaryData.Count

$htmlCss = "<style>`n" +
":root{background-color:#f0f4f8}`n" +
"body{font-family:'Segoe UI',Roboto,Arial,sans-serif;background:#f0f4f8;color:#202124;font-size:14px;margin:0}`n" +
"header{background:linear-gradient(135deg,#1a73e8,#0d47a1);color:#fff;padding:28px 36px}`n" +
"header h1{font-size:1.6rem;font-weight:600}`n" +
"header p{margin-top:5px;opacity:.8;font-size:.88rem}`n" +
".wrap{max-width:1400px;margin:0 auto;padding:24px 20px}`n" +
"*{box-sizing:border-box}`n" +
".cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:14px;margin-bottom:22px}`n" +
".card{background:#fff;border-radius:10px;padding:18px 14px;text-align:center;box-shadow:0 1px 3px rgba(0,0,0,.1);border-top:4px solid #1a73e8}`n" +
".card.g{border-top-color:#34a853}.card.y{border-top-color:#f9ab00}.card.r{border-top-color:#ea4335}.card.p{border-top-color:#9334e6}`n" +
".card-num{font-size:2rem;font-weight:700;color:#1a73e8}`n" +
".card.g .card-num{color:#34a853}.card.y .card-num{color:#e37400}.card.r .card-num{color:#ea4335}.card.p .card-num{color:#9334e6}`n" +
".card-label{font-size:.72rem;color:#5f6368;margin-top:3px;text-transform:uppercase;letter-spacing:.5px}`n" +
".sec{background:#fff;border-radius:10px;padding:22px;margin-bottom:20px;box-shadow:0 1px 3px rgba(0,0,0,.07)}`n" +
".sec h2{font-size:1rem;font-weight:600;color:#1a73e8;border-bottom:2px solid #dadce0;padding-bottom:9px;margin-bottom:16px;display:flex;align-items:center;gap:8px}`n" +
".badge{font-size:.72rem;background:#1a73e8;color:#fff;padding:2px 8px;border-radius:10px}`n" +
".badge.r{background:#ea4335}.badge.g{background:#34a853}.badge.y{background:#f9ab00;color:#333}`n" +
".info{background:#e8f0fe;border-left:4px solid #1a73e8;border-radius:4px;padding:10px 14px;margin-bottom:14px;font-size:.83rem;color:#174ea6}`n" +
".warn{background:#fef7e0;border-left:4px solid #f9ab00;border-radius:4px;padding:10px 14px;margin-bottom:14px;font-size:.83rem;color:#7a4f00}`n" +
".ok{background:#e6f4ea;border-left:4px solid #34a853;border-radius:4px;padding:10px 14px;margin-bottom:14px;font-size:.83rem;color:#1e4620}`n" +
".tabs{display:flex;flex-wrap:wrap;gap:2px;border-bottom:2px solid #dadce0;margin-bottom:0}`n" +
".tab{padding:7px 16px;cursor:pointer;border:none;background:none;font-size:.83rem;color:#5f6368;border-bottom:3px solid transparent;margin-bottom:-2px;border-radius:4px 4px 0 0}`n" +
".tab:hover{background:#f1f3f4;color:#202124}.tab.on{color:#1a73e8;border-bottom-color:#1a73e8;font-weight:600}`n" +
".pane{display:none;padding-top:18px}.pane.on{display:block}`n" +
".tbl{overflow-x:auto}`n" +
"table{width:100%;border-collapse:collapse;font-size:.81rem}`n" +
"thead th{background:#f8f9fa;color:#5f6368;font-weight:600;text-transform:uppercase;font-size:.71rem;padding:9px 11px;text-align:left;border-bottom:2px solid #dadce0;white-space:nowrap}`n" +
"tbody tr:hover{background:#f1f3f4}`n" +
"tbody td{padding:8px 11px;border-bottom:1px solid #dadce0;vertical-align:top;max-width:380px;word-break:break-word}`n" +
"tbody tr:last-child td{border-bottom:none}`n" +
".empty,.note{color:#5f6368;font-style:italic;padding:10px 0;font-size:.83rem}`n" +
"code{background:#f1f3f4;padding:1px 5px;border-radius:3px;font-size:.82rem}`n" +
"ul.userlist{margin:0;padding-left:20px;columns:3;-webkit-columns:3;column-gap:20px}`n" +
"ul.userlist li{font-size:.82rem;color:#202124;margin-bottom:3px}`n" +
"footer{text-align:center;padding:20px;color:#5f6368;font-size:.78rem}`n" +
"</style>"

$reportFile = Join-Path $OutputDir "TargetedReport.html"

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Targeted User Report</title>
$htmlCss
</head>
<body>
<header>
  <h1>&#128269; Targeted Connected Apps &amp; Chat Audit</h1>
  <p>Google Workspace &nbsp;&middot;&nbsp; $reportDate &nbsp;&middot;&nbsp; Runtime: ${elapsed}s &nbsp;&middot;&nbsp; Target users: $totalTargets</p>
</header>
<div class="wrap">

<!-- STAT CARDS -->
<div class="cards">
  <div class="card">  <div class="card-num">$totalTargets</div> <div class="card-label">Target Users</div></div>
  <div class="card y"><div class="card-num">$totalApps</div>    <div class="card-label">OAuth Apps</div></div>
  <div class="card">  <div class="card-num">$totalTokens</div>  <div class="card-label">Token Records</div></div>
  <div class="card g"><div class="card-num">$totalSpaces</div>  <div class="card-label">Space Records</div></div>
  <div class="card p"><div class="card-num">$totalBots</div>    <div class="card-label">Unique Bots</div></div>
  <div class="card r"><div class="card-num">$totalBotEvts</div> <div class="card-label">Bot Events</div></div>
</div>

<!-- TARGET USERS -->
<div class="sec">
  <h2>&#128101; Target Users Scanned <span class="badge">$totalTargets</span></h2>
  <ul class="userlist">
$htmlTargetList
  </ul>
</div>

<!-- CONNECTED APPS -->
<div class="sec">
  <h2>&#128241; Connected Apps (OAuth Tokens) <span class="badge y">$totalApps apps</span></h2>
  <div class="info">
    Shows every third-party app authorised by the target users via OAuth.
    <strong>UserCount</strong> = how many of the target users have authorised this app.
    <strong>Scopes</strong> = permissions the app holds.
  </div>
  <div class="tabs">
    <button class="tab on" onclick="tab(event,'t-apps')">By App</button>
    <button class="tab"    onclick="tab(event,'t-user')">By User</button>
    <button class="tab"    onclick="tab(event,'t-raw')">Raw Tokens</button>
  </div>
  <div id="t-apps" class="pane on"><div class="tbl">$htmlApps</div></div>
  <div id="t-user" class="pane"><div class="tbl">$htmlPerUser</div></div>
  <div id="t-raw"  class="pane">
    <div class="info">One row per user per app. Up to 500 rows shown. See tokens_target_users.csv for full data.</div>
    <div class="tbl">$htmlRawTokens</div>
  </div>
</div>

<!-- CHAT SPACES -->
<div class="sec">
  <h2>&#128172; Chat Spaces (Target Users) <span class="badge g">$totalSpaces records</span></h2>
  $htmlSpaces
</div>

<!-- CHAT BOTS -->
<div class="sec">
  <h2>&#129302; Chat Bot Activity (Target Users) <span class="badge p">$totalBots bots &nbsp;|&nbsp; $totalBotEvts events</span></h2>
  <div class="tabs">
    <button class="tab on" onclick="tab(event,'b-sum')">Bot Summary</button>
    <button class="tab"    onclick="tab(event,'b-evt')">All Bot Events</button>
  </div>
  <div id="b-sum" class="pane on"><div class="tbl">$htmlBotSummary</div></div>
  <div id="b-evt" class="pane"><div class="tbl">$htmlBotEvents</div></div>
</div>

<!-- OUTPUT FILES -->
<div class="sec">
  <h2>&#128193; Output Files</h2>
  <table>
    <thead><tr><th>File</th><th>Description</th><th>Records</th></tr></thead>
    <tbody>
      <tr><td>tokens_target_users.csv</td><td>Raw OAuth token records for target users</td><td>$totalTokens</td></tr>
      <tr><td>tokens_per_user.csv</td><td>Per-user app count and app list</td><td>$($perUserSummary.Count)</td></tr>
      <tr><td>apps_aggregated.csv</td><td>Apps aggregated with user counts and scopes</td><td>$totalApps</td></tr>
      <tr><td>chat_spaces_target.csv</td><td>Chat Spaces the target users belong to</td><td>$totalSpaces</td></tr>
      <tr><td>chat_members_target.csv</td><td>Members of spaces found for target users</td><td>$($chatMemberData.Count)</td></tr>
      <tr><td>chat_bot_events_target.csv</td><td>Bot events involving target users</td><td>$totalBotEvts</td></tr>
      <tr><td>chat_bot_summary_target.csv</td><td>Per-bot summary for target users</td><td>$totalBots</td></tr>
      <tr><td>TargetedReport.html</td><td>This interactive HTML report</td><td>-</td></tr>
    </tbody>
  </table>
</div>

</div>
<footer>Generated by Get-TargetedUserReport.ps1 &nbsp;&middot;&nbsp; $reportDate</footer>
<script>
function tab(e, id) {
  var sec = e.currentTarget.closest('.sec');
  sec.querySelectorAll('.tab').forEach(function(t){t.classList.remove('on');});
  sec.querySelectorAll('.pane').forEach(function(p){p.classList.remove('on');});
  e.currentTarget.classList.add('on');
  document.getElementById(id).classList.add('on');
}
</script>
</body>
</html>
"@

$html | Out-File $reportFile -Encoding UTF8
Write-OK "HTML report written: $reportFile"

# ======================================================
#  FINAL SUMMARY
# ======================================================
Write-Banner "Done"
$elapsed2 = [math]::Round(((Get-Date) - $T0).TotalSeconds, 1)

Write-Host ""
Write-Host "  Target Users Scanned   : $totalTargets" -ForegroundColor White
Write-Host "  OAuth Connected Apps   : $totalApps  ($totalTokens token records)" -ForegroundColor White
if ($IncludeChatSpaces) { Write-Host "  Chat Space Records     : $totalSpaces" -ForegroundColor White }
if ($IncludeChatBots)   { Write-Host "  Chat Bot Events        : $totalBotEvts  ($totalBots unique bot(s))" -ForegroundColor White }
Write-Host ""
Write-Host "  Output folder : $OutputDir" -ForegroundColor Cyan
Write-Host "  HTML report   : $reportFile" -ForegroundColor Cyan
Write-Host ""
Write-OK "Completed in ${elapsed2}s"
Write-Host ""

$open = Read-Host "  Open HTML report in browser? (y/n)"
if ($open -match '^y') { Start-Process $reportFile }

