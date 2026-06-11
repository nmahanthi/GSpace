# ---------------------------------------------------------------
# Get-ChatBots.ps1
# Pulls Google Chat Bot activity directly via GAM Reports API
# then enriches with Space details from GAM Chat API
#
# GAM command used: gam report chat
# This pulls from the same data source as Admin Console Reports
# ---------------------------------------------------------------

# Self-unblock: removes the Windows Zone Identifier (internet download tag)
# This prevents the PowerShell security warning on first run
$selfPath = $MyInvocation.MyCommand.Path
if ($selfPath -and (Test-Path $selfPath)) {
    $zone = Get-Item $selfPath -Stream "Zone.Identifier" -ErrorAction SilentlyContinue
    if ($zone) {
        Unblock-File -Path $selfPath -ErrorAction SilentlyContinue
    }
}

# --- CONFIGURATION ---
$AdminEmail  = "admin-narendra@rocheua.com"
$GamPath     = Join-Path $PSScriptRoot "gam.exe"
$CurrentDir  = Get-Location
$Timestamp   = Get-Date -Format "yyyyMMdd_HHmmss"

# Date range for bot activity report (default: last 30 days)
# Change these if you want a specific range e.g. "2026-01-01" "2026-04-10"
$StartDate   = "-30d"   # relative: 30 days ago
$EndDate     = "today"

# --- Temp files ---
$RawChatReportCSV = "$CurrentDir\raw_chat_report.csv"
$RawSpacesCSV     = "$CurrentDir\raw_spaces_bots.csv"

# --- Output files ---
$BotRawCSV        = "$CurrentDir\BOTS_RAW_ACTIVITY_$Timestamp.csv"
$EnrichedCSV      = "$CurrentDir\BOTS_ENRICHED_$Timestamp.csv"
$PerBotCSV        = "$CurrentDir\BOTS_PER_BOT_$Timestamp.csv"
$PerSpaceCSV      = "$CurrentDir\BOTS_PER_SPACE_$Timestamp.csv"
$LastModifiedCSV  = "$CurrentDir\BOTS_LAST_MODIFIED_$Timestamp.csv"

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "   Google Chat Bot Activity Audit (GAM7)   " -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "Timestamp  : $Timestamp"
Write-Host "Admin      : $AdminEmail"
Write-Host "Date Range : $StartDate -> $EndDate"
Write-Host ""

# ---------------------------------------------------------------
# Helper: Wait for file with timeout
# ---------------------------------------------------------------
function Wait-ForFile($filePath, $timeoutSec = 180) {
    $elapsed = 0
    Write-Host "   Waiting for GAM output..." -NoNewline
    while (-not (Test-Path $filePath)) {
        if ($elapsed -ge $timeoutSec) {
            Write-Host "`nERROR: Timed out after $timeoutSec seconds." -ForegroundColor Red
            exit 1
        }
        Start-Sleep -Seconds 2; $elapsed += 2
        Write-Host "." -NoNewline
    }
    Write-Host " Done." -ForegroundColor Green
}

# ---------------------------------------------------------------
# STEP 1: Pull Chat Activity Report via GAM Reports API
# gam report chat = same data as Admin Console > Reports > Chat
# This includes bot interactions, message events, space events
# ---------------------------------------------------------------
Write-Host "[1/3] Pulling Chat Activity report from Admin Reports API..." -ForegroundColor Yellow
Write-Host "      (covers bot invocations, messages, space events)" -ForegroundColor DarkGray

$gamReportArgs = @(
    "redirect", "csv", $RawChatReportCSV,
    "report", "chat",
    "start", $StartDate,
    "end",   $EndDate
)
& $GamPath @gamReportArgs
Wait-ForFile $RawChatReportCSV

$rawReport = Import-Csv $RawChatReportCSV
if (-not $rawReport -or $rawReport.Count -eq 0) {
    Write-Host "ERROR: No chat activity data returned." -ForegroundColor Red
    Write-Host "       Check that Reports API is enabled for your domain." -ForegroundColor Yellow
    exit 1
}

Write-Host "   Total activity records pulled: $($rawReport.Count)"

# ---------------------------------------------------------------
# Auto-detect actual GAM column names for app/bot/id fields
# GAM column names vary by version - this makes script resilient
# ---------------------------------------------------------------
$cols = $rawReport[0].PSObject.Properties.Name

Write-Host ""
Write-Host "   --- Column Discovery ---" -ForegroundColor Cyan
Write-Host "   All columns returned by GAM:" -ForegroundColor DarkGray
$cols | ForEach-Object { Write-Host "     $_" -ForegroundColor DarkGray }

# Find actual column names that relate to app/bot identity
$appNameCol   = $cols | Where-Object { $_ -match "app_name|appname|app\.name|botname|bot_name" }    | Select-Object -First 1
$appIdCol     = $cols | Where-Object { $_ -match "app_id|appid|app\.id|botid|bot_id" }              | Select-Object -First 1
$robotEmailCol= $cols | Where-Object { $_ -match "robot_email|robotemail|bot_email|botemail|app_email" } | Select-Object -First 1
$msgLenCol    = $cols | Where-Object { $_ -match "message_length|msglength|msg_length|messagesize" } | Select-Object -First 1
$msgIdCol     = $cols | Where-Object { $_ -match "message_id|messageid|msg_id" }                    | Select-Object -First 1
$userEmailCol = $cols | Where-Object { $_ -match "user_email|useremail" }                           | Select-Object -First 1

Write-Host ""
Write-Host "   --- Auto-detected bot/app columns ---" -ForegroundColor Cyan
Write-Host "   AppName    column : $(if ($appNameCol)    { $appNameCol }    else { 'NOT FOUND - will try fallbacks' })" -ForegroundColor $(if ($appNameCol) {"Green"} else {"Yellow"})
Write-Host "   AppId      column : $(if ($appIdCol)      { $appIdCol }      else { 'NOT FOUND - will try fallbacks' })" -ForegroundColor $(if ($appIdCol) {"Green"} else {"Yellow"})
Write-Host "   RobotEmail column : $(if ($robotEmailCol) { $robotEmailCol } else { 'NOT FOUND - will try fallbacks' })" -ForegroundColor $(if ($robotEmailCol) {"Green"} else {"Yellow"})
Write-Host "   MsgLength  column : $(if ($msgLenCol)     { $msgLenCol }     else { 'NOT FOUND' })" -ForegroundColor $(if ($msgLenCol) {"Green"} else {"Yellow"})
Write-Host "   MsgId      column : $(if ($msgIdCol)      { $msgIdCol }      else { 'NOT FOUND' })" -ForegroundColor $(if ($msgIdCol) {"Green"} else {"Yellow"})
Write-Host "   UserEmail  column : $(if ($userEmailCol)  { $userEmailCol }  else { 'NOT FOUND' })" -ForegroundColor $(if ($userEmailCol) {"Green"} else {"Yellow"})

# Show a sample of bot-related rows so we can see actual values
Write-Host ""
Write-Host "   --- Sample rows where actor or event looks bot-related ---" -ForegroundColor Cyan
$sampleBotRows = $rawReport | Where-Object {
    ($_.PSObject.Properties.Value -join " ") -match "bot|app|APPLICATION|robot"
} | Select-Object -First 3
if ($sampleBotRows) {
    $sampleBotRows | ForEach-Object {
        $row = $_
        Write-Host "   ---- Row ----" -ForegroundColor DarkGray
        $row.PSObject.Properties | Where-Object { $_.Value -ne "" } | ForEach-Object {
            Write-Host "     $($_.Name) = $($_.Value)" -ForegroundColor DarkGray
        }
    }
} else {
    Write-Host "   No obvious bot rows found in sample - showing first 2 rows for reference:" -ForegroundColor Yellow
    $rawReport | Select-Object -First 2 | ForEach-Object {
        $row = $_
        $row.PSObject.Properties | Where-Object { $_.Value -ne "" } | ForEach-Object {
            Write-Host "     $($_.Name) = $($_.Value)" -ForegroundColor DarkGray
        }
    }
}
Write-Host ""

# Save raw report as-is for reference
$rawReport | Export-Csv $BotRawCSV -NoTypeInformation
Write-Host "   Raw report saved: $BotRawCSV" -ForegroundColor DarkGray

# ---------------------------------------------------------------
# STEP 2: Pull ALL spaces from GAM to enrich with names
# ---------------------------------------------------------------
Write-Host "[2/3] Fetching all Chat Spaces for name enrichment..." -ForegroundColor Yellow

$gamSpaceArgs = @(
    "redirect", "csv", $RawSpacesCSV,
    "user", $AdminEmail,
    "print", "chatspaces",
    "asadmin",
    "fields", "name,displayname,spacetype,membershipcount"
)
& $GamPath @gamSpaceArgs
Wait-ForFile $RawSpacesCSV

$allSpaces = Import-Csv $RawSpacesCSV
Write-Host "   Spaces loaded: $($allSpaces.Count)"

# Build SpaceID -> details lookup
$spaceLookup = @{}
foreach ($sp in $allSpaces) {
    $spaceLookup[$sp.name] = @{
        DisplayName  = $sp.displayName
        SpaceType    = $sp.spaceType
        MemberCount  = $sp.'membershipCount.joinedDirectHumanUserCount'
    }
}
# Build SpaceName -> SpaceID reverse lookup
$spaceNameLookup = @{}
foreach ($sp in $allSpaces) {
    if ($sp.displayName) {
        $spaceNameLookup[$sp.displayName.Trim().ToLower()] = $sp.name
    }
}
Remove-Item $RawSpacesCSV -Force -ErrorAction SilentlyContinue

# ---------------------------------------------------------------
# STEP 3: Filter for Bot-related events and enrich with space info
# Chat report events include: bot_added, bot_removed, message_posted etc.
# We look for any event involving a bot (app) actor or bot event type
# ---------------------------------------------------------------
Write-Host "[3/3] Processing bot-related events and enriching..." -ForegroundColor Yellow

# Helper to safely get column value
function Get-Col($row, [string[]]$names) {
    foreach ($n in $names) {
        $v = $row.PSObject.Properties[$n]
        if ($v -and "$($v.Value)".Trim() -ne "") { return "$($v.Value)".Trim() }
    }
    return ""
}

$enriched  = [System.Collections.Generic.List[PSCustomObject]]::new()
$botEvents = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($row in $rawReport) {
    $eventName  = Get-Col $row "events.name","event.name","name"
    $actorEmail = Get-Col $row "actor.email","actor"
    $actorType  = Get-Col $row "actor.callerType","actor.callerType"
    $ipAddress  = Get-Col $row "ipAddress","ip_address"
    $eventTime  = Get-Col $row "id.time","time","eventTime"

    # Extract Chat-specific parameters
    # GAM stores data in resourceDetails.0 (space) and resourceDetails.1 (app/message)
    $spaceId    = Get-Col $row "resourceDetails.0.id","room_id"
    $spaceName  = Get-Col $row "resourceDetails.0.title","room_name"
    $msgType    = Get-Col $row "conversation_type"
    $msgId      = Get-Col $row "message_id"
    $msgLengthInt = 0   # GAM chat report does not provide message length

    # User who triggered the action
    $userEmail  = Get-Col $row "resourceDetails.1.ownerDetails.ownerIdentity.0.userIdentity.userEmail","target_users"

    # Bot identity: resourceDetails.1 is APPLICATION when a bot is involved
    $res1Type   = Get-Col $row "resourceDetails.1.type"
    $res1Title  = Get-Col $row "resourceDetails.1.title"
    $res1Id     = Get-Col $row "resourceDetails.1.id"

    if ($res1Type -eq "APPLICATION") {
        $appName    = $res1Title      # e.g. "Polly"
        $appId      = $res1Id         # e.g. "gcp/1077762947798"
        $robotEmail = $res1Id         # GAM does not expose bot email; GCP resource ID used instead
    } else {
        $appName    = ""
        $appId      = ""
        $robotEmail = ""
    }

    # Classify event type - bot event if name contains app/bot OR resourceDetails.1 is APPLICATION
    $isBotEvent = $eventName -match "app|bot" -or $res1Type -eq "APPLICATION"

    if ($isBotEvent) {
        # Try to resolve space name from GAM spaces
        $resolvedSpaceName = ""
        $resolvedSpaceType = ""
        $resolvedMemberCount = ""

        if ($spaceId -and $spaceLookup.ContainsKey($spaceId)) {
            $resolvedSpaceName   = $spaceLookup[$spaceId].DisplayName
            $resolvedSpaceType   = $spaceLookup[$spaceId].SpaceType
            $resolvedMemberCount = $spaceLookup[$spaceId].MemberCount
        } elseif ($spaceName) {
            $resolvedSpaceName = $spaceName
            $key = $spaceName.ToLower()
            if ($spaceNameLookup.ContainsKey($key)) {
                $sid = $spaceNameLookup[$key]
                $resolvedSpaceType   = $spaceLookup[$sid].SpaceType
                $resolvedMemberCount = $spaceLookup[$sid].MemberCount
            }
        }

        $botEvents.Add([PSCustomObject]@{
            EventTime        = $eventTime
            EventName        = $eventName
            ActorEmail       = $actorEmail
            ActorType        = $actorType
            AppName          = $appName
            AppId            = $appId
            RobotEmail       = $robotEmail
            UserEmail        = $userEmail
            SpaceID          = $spaceId
            SpaceName        = if ($resolvedSpaceName) { $resolvedSpaceName } else { $spaceName }
            SpaceType        = $resolvedSpaceType
            MemberCount      = $resolvedMemberCount
            ConversationType = $msgType
            MessageId        = $msgId
            MessageLength    = $msgLengthInt
            IPAddress        = $ipAddress
        })
    }

    # Also add all records to enriched (full picture)
    $resolvedSpaceName2  = ""
    $resolvedSpaceType2  = ""
    if ($spaceId -and $spaceLookup.ContainsKey($spaceId)) {
        $resolvedSpaceName2 = $spaceLookup[$spaceId].DisplayName
        $resolvedSpaceType2 = $spaceLookup[$spaceId].SpaceType
    }

    $enriched.Add([PSCustomObject]@{
        EventTime        = $eventTime
        EventName        = $eventName
        ActorEmail       = $actorEmail
        ActorType        = $actorType
        AppName          = $appName
        AppId            = $appId
        SpaceID          = $spaceId
        SpaceName        = if ($resolvedSpaceName2) { $resolvedSpaceName2 } else { $spaceName }
        SpaceType        = $resolvedSpaceType2
        ConversationType = $msgType
        IsBotEvent       = $isBotEvent
    })
}

Write-Host "   Total events processed   : $($rawReport.Count)"
Write-Host "   Bot-related events found : $($botEvents.Count)" -ForegroundColor $(if ($botEvents.Count -gt 0) {"Green"} else {"DarkYellow"})

# ---------------------------------------------------------------
# Per-Bot summary (group by AppId/AppName)
# ---------------------------------------------------------------
$perBotRows = $botEvents |
    Group-Object -Property AppId |
    ForEach-Object {
        $g = $_.Group

        # app_added events tell us when/where the bot was added and by whom
        $addedEvents   = $g | Where-Object { $_.EventName -eq "app_added" }
        $addedToDate   = ($addedEvents | Select-Object -ExpandProperty EventTime | Sort-Object | Select-Object -First 1)
        $addedBy       = ($addedEvents | Where-Object { $_.ActorEmail } | Select-Object -ExpandProperty ActorEmail -Unique | Select-Object -First 1)
        $addedToSpaces = ($addedEvents | Where-Object { $_.SpaceName } | Select-Object -ExpandProperty SpaceName -Unique)
        $addedToSpaceIDs = ($addedEvents | Where-Object { $_.SpaceID } | Select-Object -ExpandProperty SpaceID -Unique)

        # All spaces the bot has been active in
        $allSpaceNames = ($g | Where-Object { $_.SpaceName -ne "" } | Select-Object -ExpandProperty SpaceName -Unique)
        $allSpaceIDs   = ($g | Where-Object { $_.SpaceID   -ne "" } | Select-Object -ExpandProperty SpaceID   -Unique)

        # Space types and member counts for spaces bot is in
        $spaceDetails  = $allSpaceIDs | ForEach-Object {
            $sid = $_
            $sname = ($g | Where-Object { $_.SpaceID -eq $sid } | Select-Object -ExpandProperty SpaceName -First 1)
            $stype = ($g | Where-Object { $_.SpaceID -eq $sid } | Select-Object -ExpandProperty SpaceType -First 1)
            $smembers = ($g | Where-Object { $_.SpaceID -eq $sid } | Select-Object -ExpandProperty MemberCount -First 1)
            "$sname (Type:$stype Members:$smembers)"
        }

        [PSCustomObject]@{
            AppName          = $g[0].AppName
            AppId            = $g[0].AppId
            GcpResourceId    = ($g | Where-Object { $_.RobotEmail } | Select-Object -ExpandProperty RobotEmail -Unique | Select-Object -First 1)
            AddedToDate      = if ($addedToDate)   { $addedToDate }   else { "N/A" }
            AddedBy          = if ($addedBy)        { $addedBy }       else { "N/A" }
            AddedToSpaces    = if ($addedToSpaces)  { $addedToSpaces -join " | " } else { "N/A" }
            LastActivityDate = ($g | Select-Object -ExpandProperty EventTime | Sort-Object | Select-Object -Last 1)
            TotalEvents      = $g.Count
            UniqueSpaces     = $allSpaceIDs.Count
            ActiveInSpaces   = if ($allSpaceNames)  { $allSpaceNames -join " | " }  else { "N/A" }
            SpaceDetails     = if ($spaceDetails)   { $spaceDetails  -join " || " } else { "N/A" }
            EventTypes       = (($g | Select-Object -ExpandProperty EventName -Unique) -join " | ")
            FirstSeen        = ($g | Select-Object -ExpandProperty EventTime | Sort-Object | Select-Object -First 1)
        }
    } | Sort-Object TotalEvents -Descending

# Per-Space summary
$perSpaceRows = $botEvents |
    Where-Object { $_.SpaceName -ne "" } |
    Group-Object -Property SpaceName |
    ForEach-Object {
        $g = $_.Group
        [PSCustomObject]@{
            SpaceName    = $g[0].SpaceName
            SpaceID      = $g[0].SpaceID
            SpaceType    = $g[0].SpaceType
            MemberCount  = $g[0].MemberCount
            TotalEvents  = $g.Count
            UniqueBots   = ($g | Where-Object { $_.AppId } | Select-Object -ExpandProperty AppId -Unique).Count
            BotNames     = (($g | Where-Object { $_.AppName } | Select-Object -ExpandProperty AppName -Unique) -join " | ")
        }
    } | Sort-Object TotalEvents -Descending

# ---------------------------------------------------------------
# Export
# ---------------------------------------------------------------
$enriched    | Export-Csv $EnrichedCSV -NoTypeInformation
$perBotRows  | Export-Csv $PerBotCSV  -NoTypeInformation
$perSpaceRows| Export-Csv $PerSpaceCSV -NoTypeInformation

# Export Last Modified Per Bot
$lastModifiedPerBot = $perBotRows | Select-Object `
    AppName, AppId, GcpResourceId,
    AddedToDate, AddedBy, AddedToSpaces,
    LastActivityDate, FirstSeen,
    TotalEvents, UniqueSpaces, ActiveInSpaces, SpaceDetails,
    EventTypes
$lastModifiedPerBot | Export-Csv $LastModifiedCSV -NoTypeInformation

# Cleanup
Remove-Item $RawChatReportCSV -Force -ErrorAction SilentlyContinue

# ---------------------------------------------------------------
# Per-Bot Detail Table (console)
# ---------------------------------------------------------------
Write-Host ""
Write-Host "================================================================================================================================" -ForegroundColor Cyan
Write-Host "   BOT ACTIVITY DETAIL                                                                                                          " -ForegroundColor Cyan
Write-Host "================================================================================================================================" -ForegroundColor Cyan

if ($perBotRows.Count -eq 0) {
    Write-Host "  No bot activity found in the selected date range." -ForegroundColor DarkYellow
    Write-Host "  Check the column discovery output above - GAM may use different column names." -ForegroundColor Yellow
} else {
    $perBotRows | ForEach-Object {
        $nameDisplay  = if ($_.AppName)         { $_.AppName }         else { "(unknown)" }
        $idDisplay    = if ($_.AppId)           { $_.AppId }           else { "(unknown)" }
        $addedToDate  = if ($_.AddedToDate)     { $_.AddedToDate }     else { "N/A" }
        $addedBy      = if ($_.AddedBy)         { $_.AddedBy }         else { "N/A" }
        $addedSpaces  = if ($_.AddedToSpaces)   { $_.AddedToSpaces }   else { "N/A" }
        $lastActive   = if ($_.LastActivityDate){ $_.LastActivityDate } else { "N/A" }
        $activeSpaces = if ($_.ActiveInSpaces)  { $_.ActiveInSpaces }  else { "N/A" }

        Write-Host ""
        Write-Host "  Bot Name       : $nameDisplay"        -ForegroundColor White
        Write-Host "  App ID         : $idDisplay"          -ForegroundColor DarkGray
        Write-Host "  Added On       : $addedToDate"        -ForegroundColor Cyan
        Write-Host "  Added By       : $addedBy"            -ForegroundColor Cyan
        Write-Host "  Added To Space : $addedSpaces"        -ForegroundColor Cyan
        Write-Host "  Last Active    : $lastActive"         -ForegroundColor Green
        Write-Host "  Total Events   : $($_.TotalEvents)"   -ForegroundColor White
        Write-Host "  Active In      : $activeSpaces"       -ForegroundColor Yellow
        Write-Host "  Space Details  : $($_.SpaceDetails)"  -ForegroundColor DarkGray
        Write-Host "  Event Types    : $($_.EventTypes)"    -ForegroundColor DarkGray
        Write-Host "  -----------------------------------------------" -ForegroundColor DarkGray
    }
}
Write-Host ""

# ---------------------------------------------------------------
# FINAL OUTPUT
# ---------------------------------------------------------------
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "           AUDIT COMPLETE                  " -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "Total Chat Events        : $($rawReport.Count)"
Write-Host "Bot-Related Events       : $($botEvents.Count)"
Write-Host "Unique Bots/Apps         : $(($botEvents | Select-Object -ExpandProperty AppId -Unique).Count)"
Write-Host "Unique Spaces with Bots  : $(($botEvents | Where-Object { $_.SpaceID } | Select-Object -ExpandProperty SpaceID -Unique).Count)"
Write-Host ""
Write-Host "Reports saved:" -ForegroundColor Cyan
Write-Host "  Raw Activity   -> $BotRawCSV"
Write-Host "  Enriched Detail-> $EnrichedCSV"
Write-Host "  Per Bot        -> $PerBotCSV"
Write-Host "  Per Space      -> $PerSpaceCSV"
Write-Host "  Last Modified  -> $LastModifiedCSV"

if ($botEvents.Count -eq 0) {
    Write-Host ""
    Write-Host "NOTE: No bot events found. This could mean:" -ForegroundColor DarkYellow
    Write-Host "  1. No bots were active in the last 30 days" -ForegroundColor DarkYellow
    Write-Host "  2. Check raw report columns above - event field names may differ" -ForegroundColor DarkYellow
    Write-Host "  3. Raw report saved to $BotRawCSV - open it to inspect actual column names" -ForegroundColor DarkYellow
}