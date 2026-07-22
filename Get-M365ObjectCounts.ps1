#Requires -Modules Microsoft.Graph, PnP.PowerShell, MicrosoftTeams

<#
.SYNOPSIS
    Extracts Microsoft 365 Object Counts and Volumes
.DESCRIPTION
    This script retrieves various Microsoft 365 metrics including Teams, SharePoint, and other object counts
.NOTES
    Prerequisites:
    - Install-Module Microsoft.Graph -Scope CurrentUser
    - Install-Module PnP.PowerShell -Scope CurrentUser
    - Install-Module MicrosoftTeams -Scope CurrentUser

    - Since Sept 9, 2024, PnP PowerShell requires your own Entra ID App
      Registration for -Interactive login. Register one (one-time) with:
        Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP.PowerShell" -Tenant yourtenant.onmicrosoft.com
      Then pass the resulting Application (client) ID via -ClientId below,
      or set an ENTRAID_APP_ID / ENTRAID_CLIENT_ID / AZURE_CLIENT_ID
      environment variable so you don't have to pass it every time.
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$TenantUrl = "https://yourtenant.sharepoint.com",

    [Parameter(Mandatory=$false)]
    [string]$ClientId,

    [Parameter(Mandatory=$false)]
    [string]$OutputPath = ".\M365ObjectCounts_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Helper: connect (or reconnect) to a SharePoint site/admin URL via PnP PowerShell
function Connect-ToPnP {
    param([Parameter(Mandatory=$true)][string]$Url)

    if ($ClientId) {
        Connect-PnPOnline -Url $Url -Interactive -ClientId $ClientId
    } elseif ($env:ENTRAID_APP_ID -or $env:ENTRAID_CLIENT_ID -or $env:AZURE_CLIENT_ID) {
        Connect-PnPOnline -Url $Url -Interactive
    } else {
        throw "No ClientId supplied and no ENTRAID_APP_ID/ENTRAID_CLIENT_ID/AZURE_CLIENT_ID environment variable set. Since Sept 9, 2024, PnP PowerShell requires your own Entra ID App Registration for -Interactive login. Run Register-PnPEntraIDAppForInteractiveLogin once, then pass -ClientId <appId> to this script."
    }
}

# Connect to Microsoft Graph
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "Group.Read.All", "Team.ReadBasic.All", "Sites.Read.All", "User.Read.All"

# Connect to Teams
Write-Host "Connecting to Microsoft Teams..." -ForegroundColor Cyan
Connect-MicrosoftTeams

# Connect to SharePoint Online
Write-Host "Connecting to SharePoint Online..." -ForegroundColor Cyan
Connect-ToPnP -Url $TenantUrl

# Initialize results array
$results = @()

Write-Host "`n=== TEAMS METRICS ===" -ForegroundColor Green

# Get Teams sites count
Write-Host "Getting Teams sites count..." -ForegroundColor Yellow
$teamsCount = (Get-MgGroup -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -All -CountVariable teamsSiteCount).Count
$results += [PSCustomObject]@{
    Category = "Teams"
    ObjectType = "Teams sites"
    Count = $teamsCount
}
Write-Host "Teams sites: $teamsCount" -ForegroundColor White

# Get Groups with Dynamic membership
Write-Host "Getting groups with dynamic membership..." -ForegroundColor Yellow
$dynamicGroups = (Get-MgGroup -Filter "membershipRuleProcessingState eq 'On'" -All).Count
$results += [PSCustomObject]@{
    Category = "Teams"
    ObjectType = "Groups with Dynamic membership"
    Count = $dynamicGroups
}
Write-Host "Groups with Dynamic membership: $dynamicGroups" -ForegroundColor White

# Get Chats count (requires Teams PowerShell)
Write-Host "Getting chats count..." -ForegroundColor Yellow
try {
    $allUsers = Get-MgUser -All -Property Id
    $totalChats = 0
    foreach ($user in $allUsers) {
        $chats = Get-MgUserChat -UserId $user.Id -All
        $totalChats += $chats.Count
    }
    $results += [PSCustomObject]@{
        Category = "Teams"
        ObjectType = "Chats"
        Count = $totalChats
    }
    Write-Host "Chats: $totalChats" -ForegroundColor White
} catch {
    Write-Host "Unable to retrieve chats count: $_" -ForegroundColor Red
}

# Get Guests count
Write-Host "Getting guests count..." -ForegroundColor Yellow
$guestsCount = (Get-MgUser -Filter "userType eq 'Guest'" -All).Count
$results += [PSCustomObject]@{
    Category = "Teams"
    ObjectType = "Guests"
    Count = $guestsCount
}
Write-Host "Guests: $guestsCount" -ForegroundColor White

Write-Host "`n=== SHAREPOINT ONLINE METRICS ===" -ForegroundColor Green

# Get all site collections
Write-Host "Getting all SharePoint sites..." -ForegroundColor Yellow
$allSites = Get-PnPTenantSite -Detailed

# Not connected to O365 Groups or Teams
$notConnectedSites = ($allSites | Where-Object { $_.GroupId -eq "00000000-0000-0000-0000-000000000000" }).Count
$results += [PSCustomObject]@{
    Category = "SharePoint Online"
    ObjectType = "Not connected to O365 Groups or Teams"
    Count = $notConnectedSites
}
Write-Host "Not connected to O365 Groups or Teams: $notConnectedSites" -ForegroundColor White

# Connected to O365 Groups (but not Teams)
$connectedToGroups = ($allSites | Where-Object { 
    $_.GroupId -ne "00000000-0000-0000-0000-000000000000" -and 
    $_.Template -notlike "*TEAMCHANNEL*" 
}).Count
$results += [PSCustomObject]@{
    Category = "SharePoint Online"
    ObjectType = "Connected to O365 Groups"
    Count = $connectedToGroups
}
Write-Host "Connected to O365 Groups: $connectedToGroups" -ForegroundColor White

# Connected to O365 Groups and Microsoft Teams
$connectedToTeams = ($allSites | Where-Object { 
    $_.Template -like "*TEAMCHANNEL*" -or 
    ($_.GroupId -ne "00000000-0000-0000-0000-000000000000" -and $_.Template -eq "GROUP#0")
}).Count
$results += [PSCustomObject]@{
    Category = "SharePoint Online"
    ObjectType = "Connected to O365 Groups and Microsoft Teams"
    Count = $connectedToTeams
}
Write-Host "Connected to O365 Groups and Microsoft Teams: $connectedToTeams" -ForegroundColor White

Write-Host "`n=== OTHERS (DETAILED SHAREPOINT METRICS) ===" -ForegroundColor Green

# Large site collection by size (>100GB)
Write-Host "Analyzing large site collections by size..." -ForegroundColor Yellow
$largeSitesBySize = ($allSites | Where-Object { $_.StorageUsageCurrent -gt 102400 }).Count
$results += [PSCustomObject]@{
    Category = "SharePoint Online - Others"
    ObjectType = "Large site collection by size (>100GB)"
    Count = $largeSitesBySize
}
Write-Host "Large site collection by size: $largeSitesBySize" -ForegroundColor White

# Large site collection by numbers (>100K items)
Write-Host "Analyzing large site collections by item count..." -ForegroundColor Yellow
$largeSitesByNumber = 0
foreach ($site in $allSites) {
    try {
        Connect-ToPnP -Url $site.Url
        $lists = Get-PnPList
        $totalItems = ($lists | Measure-Object -Property ItemCount -Sum).Sum
        if ($totalItems -gt 100000) {
            $largeSitesByNumber++
        }
    } catch {
        Write-Host "Unable to access site: $($site.Url)" -ForegroundColor DarkYellow
    }
}
$results += [PSCustomObject]@{
    Category = "SharePoint Online - Others"
    ObjectType = "Large site collection by numbers (>100K items)"
    Count = $largeSitesByNumber
}
Write-Host "Large site collection by numbers: $largeSitesByNumber" -ForegroundColor White

# Workflow 2010
Write-Host "Checking Workflow 2010..." -ForegroundColor Yellow
$workflow2010Count = 0
foreach ($site in $allSites) {
    try {
        Connect-ToPnP -Url $site.Url
        $lists = Get-PnPList
        foreach ($list in $lists) {
            $workflows = Get-PnPWorkflowDefinition -List $list.Title
            $workflow2010Count += ($workflows | Where-Object { $_.RestrictToType -eq "List" }).Count
        }
    } catch {
        # Skip if unable to access
    }
}
$results += [PSCustomObject]@{
    Category = "SharePoint Online - Others"
    ObjectType = "Workflow 2010"
    Count = $workflow2010Count
}
Write-Host "Workflow 2010: $workflow2010Count" -ForegroundColor White

# Check-out files in use
Write-Host "Checking files in check-out..." -ForegroundColor Yellow
$checkedOutFiles = 0
foreach ($site in $allSites) {
    try {
        Connect-ToPnP -Url $site.Url
        $lists = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 } # Document libraries
        foreach ($list in $lists) {
            $items = Get-PnPListItem -List $list -PageSize 1000
            $checkedOutFiles += ($items | Where-Object { $_["CheckoutUser"] -ne $null }).Count
        }
    } catch {
        # Skip if unable to access
    }
}
$results += [PSCustomObject]@{
    Category = "SharePoint Online - Others"
    ObjectType = "Check-out file in use"
    Count = $checkedOutFiles
}
Write-Host "Check-out files in use: $checkedOutFiles" -ForegroundColor White

# Lists With Lookup Threshold Exceeded
Write-Host "Checking lists with lookup threshold exceeded..." -ForegroundColor Yellow
$listsWithLookupThreshold = 0
foreach ($site in $allSites) {
    try {
        Connect-ToPnP -Url $site.Url
        $lists = Get-PnPList
        foreach ($list in $lists) {
            $fields = Get-PnPField -List $list
            $lookupFields = $fields | Where-Object { $_.TypeAsString -eq "Lookup" }
            if ($lookupFields.Count -gt 12) {
                $listsWithLookupThreshold++
            }
        }
    } catch {
        # Skip if unable to access
    }
}
$results += [PSCustomObject]@{
    Category = "SharePoint Online - Others"
    ObjectType = "Lists With Lookup Threshold Exceeded"
    Count = $listsWithLookupThreshold
}
Write-Host "Lists With Lookup Threshold Exceeded: $listsWithLookupThreshold" -ForegroundColor White

# Sandbox Solutions
Write-Host "Checking sandbox solutions..." -ForegroundColor Yellow
$sandboxSolutions = 0
foreach ($site in $allSites) {
    try {
        Connect-ToPnP -Url $site.Url
        $solutions = Get-PnPSolution
        $sandboxSolutions += $solutions.Count
    } catch {
        # Skip if unable to access
    }
}
$results += [PSCustomObject]@{
    Category = "SharePoint Online - Others"
    ObjectType = "Sandbox Solutions"
    Count = $sandboxSolutions
}
Write-Host "Sandbox Solutions: $sandboxSolutions" -ForegroundColor White

# Custom Features
Write-Host "Checking custom features..." -ForegroundColor Yellow
$customFeatures = 0
foreach ($site in $allSites) {
    try {
        Connect-ToPnP -Url $site.Url
        $features = Get-PnPFeature -Scope Site
        $customFeatures += ($features | Where-Object {
            $_.DefinitionId -notlike "00000000-0000-0000-0000-*"
        }).Count
    } catch {
        # Skip if unable to access
    }
}
$results += [PSCustomObject]@{
    Category = "SharePoint Online - Others"
    ObjectType = "Custom Features"
    Count = $customFeatures
}
Write-Host "Custom Features: $customFeatures" -ForegroundColor White

# Custom List Templates
Write-Host "Checking custom list templates..." -ForegroundColor Yellow
$customListTemplates = 0
foreach ($site in $allSites) {
    try {
        Connect-ToPnP -Url $site.Url
        $templates = Get-PnPListItem -List "List Template Gallery"
        $customListTemplates += $templates.Count
    } catch {
        # Skip if unable to access
    }
}
$results += [PSCustomObject]@{
    Category = "SharePoint Online - Others"
    ObjectType = "Custom List Templates"
    Count = $customListTemplates
}
Write-Host "Custom List Templates: $customListTemplates" -ForegroundColor White

# Custom Site Templates
Write-Host "Checking custom site templates..." -ForegroundColor Yellow
$customSiteTemplates = 0
foreach ($site in $allSites) {
    try {
        Connect-ToPnP -Url $site.Url
        $templates = Get-PnPListItem -List "Solution Gallery"
        $customSiteTemplates += $templates.Count
    } catch {
        # Skip if unable to access
    }
}
$results += [PSCustomObject]@{
    Category = "SharePoint Online - Others"
    ObjectType = "Custom Site Templates"
    Count = $customSiteTemplates
}
Write-Host "Custom Site Templates: $customSiteTemplates" -ForegroundColor White

# Customised Pages
Write-Host "Checking customised pages..." -ForegroundColor Yellow
$customisedPages = 0
foreach ($site in $allSites) {
    try {
        Connect-ToPnP -Url $site.Url
        $web = Get-PnPWeb
        $files = Get-PnPFile -AsListItem
        $customisedPages += ($files | Where-Object {
            $_["vti_x005f_level"] -eq "2"
        }).Count
    } catch {
        # Skip if unable to access
    }
}
$results += [PSCustomObject]@{
    Category = "SharePoint Online - Others"
    ObjectType = "Customised Pages"
    Count = $customisedPages
}
Write-Host "Customised Pages: $customisedPages" -ForegroundColor White

# Export results
Write-Host "`n=== EXPORTING RESULTS ===" -ForegroundColor Green
$results | Export-Csv -Path $OutputPath -NoTypeInformation
Write-Host "Results exported to: $OutputPath" -ForegroundColor Cyan

# Display summary table
Write-Host "`n=== SUMMARY ===" -ForegroundColor Green
$results | Format-Table -AutoSize

# Disconnect sessions
Write-Host "`nDisconnecting sessions..." -ForegroundColor Cyan
Disconnect-MgGraph
Disconnect-MicrosoftTeams
Disconnect-PnPOnline

Write-Host "`nScript completed successfully!" -ForegroundColor Green
