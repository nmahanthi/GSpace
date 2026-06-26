<#
.SYNOPSIS
    Microsoft 365 Assessment Report — Groups, Teams & Planner
.DESCRIPTION
    Collects governance and usage data for all M365 Groups, Teams and Planner
    plans in the tenant and produces a self-contained HTML report plus optional
    CSV exports.
.PARAMETER TenantId
    Azure AD Tenant ID. Leave blank to use interactive (delegated) login.
.PARAMETER ClientId
    App-Registration Client ID for app-only (unattended) auth.
.PARAMETER ClientSecret
    App-Registration Client Secret for app-only auth.
.PARAMETER OutputPath
    Folder where report files are written. Created automatically.
.PARAMETER ExportCSV
    Also write each dataset to a CSV file in the output folder.
.PARAMETER SkipPlannerDetail
    Skip per-task enumeration (faster for large tenants).
.PARAMETER MaxGroups
    Limit the number of M365 Groups processed (0 = unlimited).
.EXAMPLE
    .\M365-Assessment-Report.ps1
    .\M365-Assessment-Report.ps1 -TenantId "xxxx" -ClientId "yyyy" -ClientSecret "zzzz" -ExportCSV
.NOTES
    Modules auto-installed if missing:
        Microsoft.Graph.Authentication, Microsoft.Graph.Groups,
        Microsoft.Graph.Teams, Microsoft.Graph.Planner, Microsoft.Graph.Users
    Permissions required (delegated or application):
        Group.Read.All, Team.ReadBasic.All, TeamMember.Read.All,
        Channel.ReadBasic.All, Tasks.Read, User.Read.All,
        Directory.Read.All, GroupMember.Read.All
#>

[CmdletBinding()]
param (
    [string]$TenantId     = "",
    [string]$ClientId     = "",
    [string]$ClientSecret = "",
    [string]$OutputPath   = ".\M365_Assessment_$(Get-Date -Format 'yyyyMMdd_HHmmss')",
    [switch]$ExportCSV,
    [switch]$SkipPlannerDetail,
    [int]   $MaxGroups    = 0
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$WarningPreference     = "SilentlyContinue"
$script:MaxGroups      = $MaxGroups

#region HELPERS
function Write-Log {
    param([string]$Msg, [string]$Level = "INFO")
    $color = switch ($Level) {
        "SUCCESS" { "Green"  } "WARN" { "Yellow" } "ERROR" { "Red" } default { "Cyan" }
    }
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')][$Level] $Msg" -ForegroundColor $color
}

function Install-RequiredModules {
    $mods = @(
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.Groups",
        "Microsoft.Graph.Teams",
        "Microsoft.Graph.Planner",
        "Microsoft.Graph.Users"
    )
    foreach ($m in $mods) {
        if (-not (Get-Module -ListAvailable -Name $m -ErrorAction SilentlyContinue)) {
            Write-Log "Installing module: $m ..." "WARN"
            Install-Module $m -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        }
        Import-Module $m -Force -ErrorAction Stop
    }
    Write-Log "All Microsoft.Graph modules are ready." "SUCCESS"
}

function Connect-ToGraph {
    param([string]$TenantId, [string]$ClientId, [string]$ClientSecret)
    $scopes = @(
        "Group.Read.All","Team.ReadBasic.All","TeamMember.Read.All",
        "Channel.ReadBasic.All","Tasks.Read","User.Read.All",
        "Directory.Read.All","GroupMember.Read.All"
    )
    if ($ClientId -and $ClientSecret -and $TenantId) {
        Write-Log "Connecting via App-Only (Client Credentials)..." "INFO"
        $sec  = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
        $cred = New-Object System.Management.Automation.PSCredential($ClientId, $sec)
        Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $cred -NoWelcome
    } else {
        Write-Log "Connecting via Device Code flow..." "INFO"
        Write-Log "A URL and code will appear below. Open the URL in any browser and enter the code." "WARN"
        $tenantId = if ($TenantId) { $TenantId } else { "organizations" }
        Connect-MgGraph -Scopes $scopes -UseDeviceAuthentication -NoWelcome -TenantId $tenantId
    }
    $ctx = Get-MgContext
    if (-not $ctx) { throw "Authentication failed. No Graph context established." }
    Write-Log "Connected as: $($ctx.Account) | Tenant: $($ctx.TenantId)" "SUCCESS"
}

function Invoke-SafeGraph {
    param([scriptblock]$Block)
    try   { & $Block 2>$null }
    catch {
        # Extract only the first line so 403 HTTP dumps don't flood the log
        $msg = ($_.Exception.Message -split "`n")[0].Trim()
        Write-Log "Graph call skipped: $msg" "WARN"
        return $null
    }
}
#endregion

#region M365 GROUPS
function Get-M365GroupsAssessment {
    Write-Log "=== Collecting M365 Groups ===" "INFO"
    $props  = "id,displayName,description,mail,visibility,createdDateTime,resourceProvisioningOptions,groupTypes"
    $groups = Get-MgGroup -Filter "groupTypes/any(c:c eq 'Unified')" -Property $props -All -PageSize 999
    if ($script:MaxGroups -gt 0) { $groups = $groups | Select-Object -First $script:MaxGroups }
    Write-Log "Found $($groups.Count) M365 Groups. Collecting member/owner details..." "INFO"

    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    $i = 0
    foreach ($g in $groups) {
        $i++
        Write-Progress -Activity "Analyzing M365 Groups" `
                       -Status "$i / $($groups.Count) : $($g.DisplayName)" `
                       -PercentComplete (($i / $groups.Count) * 100)

        $owners  = @(Invoke-SafeGraph { Get-MgGroupOwner  -GroupId $g.Id -All })
        $members = @(Invoke-SafeGraph { Get-MgGroupMember -GroupId $g.Id -All })
        $isTeam  = $g.ResourceProvisioningOptions -contains "Team"

        $results.Add([PSCustomObject]@{
            GroupId          = $g.Id
            DisplayName      = $g.DisplayName
            Description      = if ($g.Description) { $g.Description } else { "" }
            Email            = $g.Mail
            Visibility       = if ($g.Visibility) { $g.Visibility } else { "Private" }
            CreatedDate      = if ($g.CreatedDateTime) { $g.CreatedDateTime.ToString("yyyy-MM-dd") } else { "N/A" }
            OwnersCount      = $owners.Count
            MembersCount     = $members.Count
            IsTeamsConnected = $isTeam
            HasPlanner       = $false
            Owners           = if ($owners.Count -gt 0) { ($owners | ForEach-Object { $_.AdditionalProperties["displayName"] }) -join "; " } else { "None" }
        })
    }
    Write-Progress -Activity "Analyzing M365 Groups" -Completed
    Write-Log "Groups complete: $($results.Count) records." "SUCCESS"
    return $results
}
#endregion

#region MICROSOFT TEAMS
function Get-TeamsAssessment {
    Write-Log "=== Collecting Microsoft Teams ===" "INFO"
    $allTeams = Invoke-SafeGraph { Get-MgTeam -All -PageSize 999 }
    if (-not $allTeams) { Write-Log "No Teams returned (check permissions)." "WARN"; return @() }
    Write-Log "Found $($allTeams.Count) Teams. Collecting channel/member details..." "INFO"

    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    $i = 0
    foreach ($t in $allTeams) {
        $i++
        Write-Progress -Activity "Analyzing Teams" `
                       -Status "$i / $($allTeams.Count) : $($t.DisplayName)" `
                       -PercentComplete (($i / $allTeams.Count) * 100)

        $detail   = Invoke-SafeGraph { Get-MgTeam -TeamId $t.Id }
        $channels = @(Invoke-SafeGraph { Get-MgTeamChannel -TeamId $t.Id -All })
        $members  = @(Invoke-SafeGraph { Get-MgTeamMember  -TeamId $t.Id -All })

        # Fallback to basic info from the enumerated object when Get-MgTeam is Forbidden
        $teamDesc    = if ($detail -and $detail.Description) { $detail.Description } else { "" }
        $teamVis     = if ($detail -and $detail.Visibility)  { $detail.Visibility  } else { if ($t.Visibility) { $t.Visibility } else { "Unknown" } }
        $teamArchive = if ($detail) { $detail.IsArchived } else { $false }

        $owners  = @(); $guests = @(); $regular = @()
        if ($members) {
            $owners  = @($members | Where-Object { $_.Roles -contains "owner" })
            $guests  = @($members | Where-Object { $_.AdditionalProperties["userType"] -eq "Guest" })
            $regular = @($members | Where-Object { ($_.Roles -notcontains "owner") -and ($_.AdditionalProperties["userType"] -ne "Guest") })
        }

        $stdCh  = @($channels | Where-Object { $_.MembershipType -eq "standard" }).Count
        $privCh = @($channels | Where-Object { $_.MembershipType -eq "private"  }).Count
        $shrdCh = @($channels | Where-Object { $_.MembershipType -eq "shared"   }).Count

        $results.Add([PSCustomObject]@{
            TeamId            = $t.Id
            DisplayName       = $t.DisplayName
            Description       = $teamDesc
            Visibility        = $teamVis
            IsArchived        = $teamArchive
            OwnersCount       = $owners.Count
            MembersCount      = $regular.Count
            GuestsCount       = $guests.Count
            TotalMembersCount = $members.Count
            TotalChannels     = $channels.Count
            StandardChannels  = $stdCh
            PrivateChannels   = $privCh
            SharedChannels    = $shrdCh
            Owners            = ($owners | ForEach-Object { $_.DisplayName }) -join "; "
        })
    }
    Write-Progress -Activity "Analyzing Teams" -Completed
    Write-Log "Teams complete: $($results.Count) records." "SUCCESS"
    return $results
}
#endregion

#region PLANNER
function Get-PlannerAssessment {
    param([System.Collections.Generic.List[PSCustomObject]]$GroupsData)
    Write-Log "=== Collecting Planner Plans ===" "INFO"
    $planResults = [System.Collections.Generic.List[PSCustomObject]]::new()
    $i = 0

    foreach ($grp in $GroupsData) {
        $i++
        Write-Progress -Activity "Analyzing Planner" `
                       -Status "$i / $($GroupsData.Count) : $($grp.DisplayName)" `
                       -PercentComplete (($i / $GroupsData.Count) * 100)

        $plans = Invoke-SafeGraph { Get-MgGroupPlannerPlan -GroupId $grp.GroupId -All }
        if (-not $plans) { continue }
        $grp.HasPlanner = $true

        foreach ($plan in $plans) {
            $buckets    = Invoke-SafeGraph { Get-MgPlannerPlanBucket -PlannerPlanId $plan.Id -All }
            $notStarted = 0; $inProgress = 0; $completed = 0; $totalTasks = 0

            if (-not $SkipPlannerDetail) {
                $tasks = Invoke-SafeGraph { Get-MgPlannerPlanTask -PlannerPlanId $plan.Id -All }
                if ($tasks) {
                    $tasks      = @($tasks)
                    $totalTasks = $tasks.Count
                    $notStarted = @($tasks | Where-Object { $_.PercentComplete -eq 0 }).Count
                    $inProgress = @($tasks | Where-Object { $_.PercentComplete -gt 0 -and $_.PercentComplete -lt 100 }).Count
                    $completed  = @($tasks | Where-Object { $_.PercentComplete -eq 100 }).Count
                }
            }

            $creatorName = ""
            if ($plan.CreatedBy -and $plan.CreatedBy.User -and $plan.CreatedBy.User.Id) {
                $u = Invoke-SafeGraph { Get-MgUser -UserId $plan.CreatedBy.User.Id -Property "displayName" }
                if ($u) { $creatorName = $u.DisplayName }
            }

            $planResults.Add([PSCustomObject]@{
                PlanId       = $plan.Id
                PlanTitle    = $plan.Title
                GroupId      = $grp.GroupId
                GroupName    = $grp.DisplayName
                CreatedBy    = $creatorName
                CreatedDate  = if ($plan.CreatedDateTime) { $plan.CreatedDateTime.ToString("yyyy-MM-dd") } else { "N/A" }
                BucketsCount = if ($buckets) { @($buckets).Count } else { 0 }
                TotalTasks   = $totalTasks
                NotStarted   = $notStarted
                InProgress   = $inProgress
                Completed    = $completed
                Buckets      = if ($buckets) { ($buckets | ForEach-Object { $_.Name }) -join "; " } else { "" }
            })
        }
    }
    Write-Progress -Activity "Analyzing Planner" -Completed
    Write-Log "Planner complete: $($planResults.Count) plans found." "SUCCESS"
    return $planResults
}
#endregion

#region HTML REPORT
function ConvertTo-HtmlTable {
    param([object[]]$Data, [string]$TableId)
    if (-not $Data -or $Data.Count -eq 0) { return "<p class='no-data'>No data available.</p>" }
    $cols = $Data[0].PSObject.Properties.Name
    $sb   = [System.Text.StringBuilder]::new()
    [void]$sb.Append("<div class='table-wrap'><table id='$TableId'>")
    [void]$sb.Append("<thead><tr>")
    foreach ($c in $cols) { [void]$sb.Append("<th>$c</th>") }
    [void]$sb.Append("</tr></thead><tbody>")
    foreach ($row in $Data) {
        [void]$sb.Append("<tr>")
        foreach ($c in $cols) {
            $val  = $row.$c
            $cell = if ($null -eq $val) { "" } else { [System.Web.HttpUtility]::HtmlEncode($val.ToString()) }
            if      ($val -eq $true)      { [void]$sb.Append("<td><span class='badge badge-yes'>Yes</span></td>") }
            elseif  ($val -eq $false)     { [void]$sb.Append("<td><span class='badge badge-no'>No</span></td>") }
            elseif  ($val -eq "Public")   { [void]$sb.Append("<td><span class='badge badge-pub'>$cell</span></td>") }
            elseif  ($val -eq "Private")  { [void]$sb.Append("<td><span class='badge badge-prv'>$cell</span></td>") }
            else                          { [void]$sb.Append("<td title='$cell'>$cell</td>") }
        }
        [void]$sb.Append("</tr>")
    }
    [void]$sb.Append("</tbody></table></div>")
    return $sb.ToString()
}

function Export-HTMLReport {
    param(
        [string]$OutputFile,
        [System.Collections.Generic.List[PSCustomObject]]$GroupsData,
        [System.Collections.Generic.List[PSCustomObject]]$TeamsData,
        [System.Collections.Generic.List[PSCustomObject]]$PlannerData,
        [string]$TenantInfo,
        [datetime]$RunDate
    )
    Add-Type -AssemblyName System.Web

    $totalGroups   = $GroupsData.Count
    $teamsLinked   = @($GroupsData | Where-Object { $_.IsTeamsConnected }).Count
    $plannerLinked = @($GroupsData | Where-Object { $_.HasPlanner }).Count
    $noOwner       = @($GroupsData | Where-Object { $_.OwnersCount -eq 0 }).Count
    $pubGroups     = @($GroupsData | Where-Object { $_.Visibility -eq "Public" }).Count
    $totalTeams    = $TeamsData.Count
    $archivedTeams = @($TeamsData  | Where-Object { $_.IsArchived }).Count
    $totalGuests   = [int](($TeamsData   | Measure-Object GuestsCount -Sum).Sum)
    $totalPlans    = $PlannerData.Count
    $totalTasks    = [int](($PlannerData | Measure-Object TotalTasks  -Sum).Sum)
    $doneTasks     = [int](($PlannerData | Measure-Object Completed   -Sum).Sum)
    $openTasks     = [int](($PlannerData | Measure-Object NotStarted  -Sum).Sum)

    $grpCols = $GroupsData  | Select-Object DisplayName,Email,Visibility,CreatedDate,OwnersCount,MembersCount,IsTeamsConnected,HasPlanner,Owners
    $tmsCols = $TeamsData   | Select-Object DisplayName,Visibility,IsArchived,OwnersCount,MembersCount,GuestsCount,TotalChannels,StandardChannels,PrivateChannels,SharedChannels,Owners
    $plnCols = $PlannerData | Select-Object PlanTitle,GroupName,CreatedBy,CreatedDate,BucketsCount,TotalTasks,NotStarted,InProgress,Completed,Buckets

    $grpTable = ConvertTo-HtmlTable -Data $grpCols -TableId "tblGroups"
    $tmsTable = ConvertTo-HtmlTable -Data $tmsCols -TableId "tblTeams"
    $plnTable = ConvertTo-HtmlTable -Data $plnCols -TableId "tblPlanner"

    $warnNoOwner = if ($noOwner -gt 0) { "warn" } else { "ok" }
    $warnArch    = if ($archivedTeams -gt 0) { "warn" } else { "ok" }
    $reportDate  = $RunDate.ToString("dddd, dd MMM yyyy HH:mm")

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>M365 Assessment Report</title>
<style>
:root{--bg:#f0f4f8;--card:#ffffff;--accent:#0078d4;--accent2:#00b4d8;--text:#1a1a2e;
      --sub:#6b7280;--border:#e2e8f0;--yes:#107c10;--no:#d13438;--pub:#0078d4;--prv:#8764b8;}
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--text);font-size:14px;}
header{background:linear-gradient(135deg,#0078d4 0%,#00b4d8 100%);color:#fff;padding:30px 40px;
       box-shadow:0 2px 8px rgba(0,120,212,.3);}
header h1{font-size:1.75rem;font-weight:700;display:flex;align-items:center;gap:12px;}
header p{opacity:.9;margin-top:6px;font-size:.88rem;}
.summary{display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:14px;padding:24px 40px;}
.card{background:var(--card);border-radius:10px;padding:18px 20px;box-shadow:0 1px 3px rgba(0,0,0,.08);
      border-left:4px solid var(--accent);transition:transform .15s;}
.card:hover{transform:translateY(-2px);box-shadow:0 4px 12px rgba(0,0,0,.12);}
.card .val{font-size:2rem;font-weight:700;color:var(--accent);}
.card .lbl{color:var(--sub);font-size:.75rem;margin-top:4px;text-transform:uppercase;letter-spacing:.6px;}
.card.warn{border-color:#d13438;}.card.warn .val{color:#d13438;}
.card.ok  {border-color:#107c10;}.card.ok   .val{color:#107c10;}
nav{display:flex;gap:2px;padding:0 40px;background:var(--card);
    border-bottom:1px solid var(--border);box-shadow:0 1px 3px rgba(0,0,0,.06);}
nav button{background:none;border:none;padding:14px 22px;cursor:pointer;font-size:.92rem;
           color:var(--sub);border-bottom:3px solid transparent;transition:.2s;font-weight:500;}
nav button.active{color:var(--accent);border-bottom-color:var(--accent);font-weight:600;}
nav button:hover:not(.active){color:var(--accent);background:#f8fbff;}
.tab{display:none;padding:24px 40px;}
.tab.active{display:block;}
.section-hdr{font-size:1.05rem;font-weight:600;margin-bottom:14px;color:var(--text);}
.search-bar{margin-bottom:10px;}
.search-bar input{padding:8px 14px;border:1px solid var(--border);border-radius:8px;
                  width:300px;font-size:.88rem;outline:none;transition:.2s;}
.search-bar input:focus{border-color:var(--accent);box-shadow:0 0 0 3px rgba(0,120,212,.15);}
.table-wrap{overflow-x:auto;border-radius:10px;box-shadow:0 1px 4px rgba(0,0,0,.07);}
table{width:100%;border-collapse:collapse;background:var(--card);}
th{background:var(--accent);color:#fff;padding:11px 14px;text-align:left;
   font-weight:600;font-size:.8rem;white-space:nowrap;letter-spacing:.3px;}
th:first-child{border-radius:10px 0 0 0;}
th:last-child {border-radius:0 10px 0 0;}
td{padding:9px 14px;border-bottom:1px solid var(--border);font-size:.84rem;
   max-width:250px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
tr:nth-child(even) td{background:#f8fafc;}
tr:hover td{background:#ebf5ff;}
.badge{display:inline-block;padding:2px 10px;border-radius:20px;font-size:.76rem;font-weight:600;color:#fff;}
.badge-yes{background:var(--yes);}
.badge-no {background:var(--no);}
.badge-pub{background:var(--pub);}
.badge-prv{background:var(--prv);}
.no-data{color:var(--sub);padding:40px;text-align:center;font-size:.95rem;}
footer{text-align:center;padding:20px 40px;color:var(--sub);font-size:.78rem;
       border-top:1px solid var(--border);margin-top:20px;}
@media(max-width:768px){header,.summary,nav,.tab{padding-left:16px;padding-right:16px;}
  .summary{grid-template-columns:repeat(2,1fr);}
  .search-bar input{width:100%;}}
</style>
</head>
<body>
<header>
  <h1>&#128202; Microsoft 365 Assessment Report</h1>
  <p>Tenant: <strong>$TenantInfo</strong> &nbsp;&bull;&nbsp; Generated: <strong>$reportDate</strong></p>
</header>

<div class="summary">
  <div class="card"><div class="val">$totalGroups</div><div class="lbl">Total M365 Groups</div></div>
  <div class="card"><div class="val">$teamsLinked</div><div class="lbl">Teams-Connected Groups</div></div>
  <div class="card"><div class="val">$plannerLinked</div><div class="lbl">Groups with Planner</div></div>
  <div class="card $warnNoOwner"><div class="val">$noOwner</div><div class="lbl">Groups with No Owner</div></div>
  <div class="card"><div class="val">$pubGroups</div><div class="lbl">Public Groups</div></div>
  <div class="card"><div class="val">$totalTeams</div><div class="lbl">Total Teams</div></div>
  <div class="card $warnArch"><div class="val">$archivedTeams</div><div class="lbl">Archived Teams</div></div>
  <div class="card"><div class="val">$totalGuests</div><div class="lbl">Total Guest Users (Teams)</div></div>
  <div class="card"><div class="val">$totalPlans</div><div class="lbl">Planner Plans</div></div>
  <div class="card ok"><div class="val">$doneTasks</div><div class="lbl">Completed Tasks</div></div>
  <div class="card warn"><div class="val">$openTasks</div><div class="lbl">Not-Started Tasks</div></div>
  <div class="card"><div class="val">$totalTasks</div><div class="lbl">Total Planner Tasks</div></div>
</div>

<nav>
  <button class="active" onclick="showTab('groups',this)">&#128101; M365 Groups ($totalGroups)</button>
  <button onclick="showTab('teams',this)">&#129309; Teams ($totalTeams)</button>
  <button onclick="showTab('planner',this)">&#128203; Planner ($totalPlans)</button>
</nav>

<div id="groups" class="tab active">
  <p class="section-hdr">Microsoft 365 Unified Groups — Ownership, Membership &amp; Connectivity</p>
  <div class="search-bar"><input type="text" id="srchG" onkeyup="filterTable('srchG','tblGroups')" placeholder="&#128269; Search groups..."/></div>
  $grpTable
</div>
<div id="teams" class="tab">
  <p class="section-hdr">Microsoft Teams — Channels, Members &amp; Guest Access</p>
  <div class="search-bar"><input type="text" id="srchT" onkeyup="filterTable('srchT','tblTeams')" placeholder="&#128269; Search teams..."/></div>
  $tmsTable
</div>
<div id="planner" class="tab">
  <p class="section-hdr">Planner Plans — Buckets &amp; Task Progress</p>
  <div class="search-bar"><input type="text" id="srchP" onkeyup="filterTable('srchP','tblPlanner')" placeholder="&#128269; Search plans..."/></div>
  $plnTable
</div>

<footer>Microsoft 365 Assessment Report &mdash; Generated by M365-Assessment-Report.ps1 &mdash; $reportDate</footer>

<script>
function showTab(id, btn) {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('nav button').forEach(b => b.classList.remove('active'));
    document.getElementById(id).classList.add('active');
    btn.classList.add('active');
}
function filterTable(inputId, tableId) {
    var q = document.getElementById(inputId).value.toLowerCase();
    var rows = document.getElementById(tableId).getElementsByTagName('tr');
    for (var i = 1; i < rows.length; i++) {
        rows[i].style.display = rows[i].innerText.toLowerCase().includes(q) ? '' : 'none';
    }
}
</script>
</body>
</html>
"@
    $html | Out-File -FilePath $OutputFile -Encoding UTF8 -Force
    Write-Log "HTML report saved: $OutputFile" "SUCCESS"
}
#endregion

#region MAIN
try {
    Write-Log "============================================================" "INFO"
    Write-Log "  M365 Assessment Report — Starting" "INFO"
    Write-Log "============================================================" "INFO"

    # 1. Ensure modules
    Install-RequiredModules

    # 2. Connect to Graph
    Connect-ToGraph -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
    $ctx        = Get-MgContext
    $tenantInfo = if ($ctx.Account) { "$($ctx.Account)  ($($ctx.TenantId))" } else { $ctx.TenantId }

    # 3. Create output folder
    $folder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
    Write-Log "Output folder: $folder" "INFO"

    # 4. Collect data
    $groupsData  = Get-M365GroupsAssessment
    $teamsData   = Get-TeamsAssessment
    $plannerData = Get-PlannerAssessment -GroupsData $groupsData

    # 5. Optional CSV exports
    if ($ExportCSV) {
        $groupsData  | Export-Csv -Path (Join-Path $folder "M365_Groups.csv")  -NoTypeInformation -Encoding UTF8
        $teamsData   | Export-Csv -Path (Join-Path $folder "M365_Teams.csv")   -NoTypeInformation -Encoding UTF8
        $plannerData | Export-Csv -Path (Join-Path $folder "M365_Planner.csv") -NoTypeInformation -Encoding UTF8
        Write-Log "CSV files written to: $folder" "SUCCESS"
    }

    # 6. Generate HTML report
    $htmlFile = Join-Path $folder "M365_Assessment_Report.html"
    Export-HTMLReport -OutputFile   $htmlFile    `
                      -GroupsData   $groupsData  `
                      -TeamsData    $teamsData   `
                      -PlannerData  $plannerData `
                      -TenantInfo   $tenantInfo  `
                      -RunDate      (Get-Date)

    # 7. Disconnect & open report
    Disconnect-MgGraph | Out-Null
    Write-Log "Disconnected from Microsoft Graph." "INFO"
    Write-Log "============================================================" "SUCCESS"
    Write-Log "  Assessment complete! Report: $htmlFile" "SUCCESS"
    Write-Log "============================================================" "SUCCESS"
    Start-Process $htmlFile

} catch {
    Write-Log "FATAL ERROR: $_" "ERROR"
    Write-Log $_.ScriptStackTrace "ERROR"
    exit 1
}
#endregion
