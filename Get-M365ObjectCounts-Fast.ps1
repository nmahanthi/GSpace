#Requires -Modules Microsoft.Graph, PnP.PowerShell

<#
.SYNOPSIS
    Quick extraction of Microsoft 365 Object Counts and Volumes
.DESCRIPTION
    Optimized script to retrieve Microsoft 365 metrics using Graph API and SharePoint Online
.NOTES
    Prerequisites:
    - Install-Module Microsoft.Graph -Scope CurrentUser
    - Install-Module PnP.PowerShell -Scope CurrentUser

    - Since Sept 9, 2024, PnP PowerShell requires your own Entra ID App
      Registration for -Interactive login. Register one (one-time) with:
        Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP.PowerShell" -Tenant yourtenant.onmicrosoft.com
      Then pass the resulting Application (client) ID via -ClientId below,
      or set an ENTRAID_APP_ID / ENTRAID_CLIENT_ID / AZURE_CLIENT_ID
      environment variable so you don't have to pass it every time.

    Required Permissions:
    - Group.Read.All
    - Team.ReadBasic.All
    - Sites.Read.All
    - User.Read.All
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$AdminUrl = "https://yourtenant-admin.sharepoint.com",

    [Parameter(Mandatory=$false)]
    [string]$ClientId,

    [Parameter(Mandatory=$false)]
    [string]$OutputPath = ".\M365ObjectCounts_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Initialize results
$results = @()

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Microsoft 365 Object Counts Extraction" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# Connect to Microsoft Graph
Write-Host "[1/2] Connecting to Microsoft Graph..." -ForegroundColor Yellow
try {
    Connect-MgGraph -Scopes "Group.Read.All", "Team.ReadBasic.All", "Sites.Read.All", "User.Read.All", "Chat.Read.All" -NoWelcome
    Write-Host "✓ Connected to Microsoft Graph" -ForegroundColor Green
} catch {
    Write-Host "✗ Failed to connect to Microsoft Graph: $_" -ForegroundColor Red
    exit
}

# Connect to SharePoint Online Admin
Write-Host "[2/2] Connecting to SharePoint Online Admin..." -ForegroundColor Yellow
try {
    if ($ClientId) {
        Connect-PnPOnline -Url $AdminUrl -Interactive -ClientId $ClientId
    } elseif ($env:ENTRAID_APP_ID -or $env:ENTRAID_CLIENT_ID -or $env:AZURE_CLIENT_ID) {
        Connect-PnPOnline -Url $AdminUrl -Interactive
    } else {
        throw "No ClientId supplied and no ENTRAID_APP_ID/ENTRAID_CLIENT_ID/AZURE_CLIENT_ID environment variable set. Since Sept 9, 2024, PnP PowerShell requires your own Entra ID App Registration for -Interactive login. Run Register-PnPEntraIDAppForInteractiveLogin once, then pass -ClientId <appId> to this script."
    }
    Write-Host "✓ Connected to SharePoint Online`n" -ForegroundColor Green
} catch {
    Write-Host "✗ Failed to connect to SharePoint Online: $_" -ForegroundColor Red
    exit
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "TEAMS METRICS" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# Teams sites (Groups with Teams provisioned)
Write-Host "→ Getting Teams sites count..." -ForegroundColor Yellow
$teamsGroups = Get-MgGroup -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -All -CountVariable teamsCount
$results += [PSCustomObject]@{
    Category = "Teams"
    ObjectType = "Teams sites"
    Count = $teamsGroups.Count
}
Write-Host "  Teams sites: $($teamsGroups.Count)`n" -ForegroundColor White

# Groups with Dynamic membership
Write-Host "→ Getting groups with dynamic membership..." -ForegroundColor Yellow
$dynamicGroups = Get-MgGroup -Filter "groupTypes/any(c:c eq 'DynamicMembership')" -All
$results += [PSCustomObject]@{
    Category = "Teams"
    ObjectType = "Groups with Dynamic membership"
    Count = $dynamicGroups.Count
}
Write-Host "  Groups with Dynamic membership: $($dynamicGroups.Count)`n" -ForegroundColor White

# Chats count (sample approach - full count would be very slow)
Write-Host "→ Estimating chats count (sampling approach)..." -ForegroundColor Yellow
try {
    $sampleUsers = Get-MgUser -Top 100 -Property Id
    $totalChats = 0
    $processedUsers = 0
    
    foreach ($user in $sampleUsers) {
        try {
            $chats = Get-MgUserChat -UserId $user.Id -All -ErrorAction SilentlyContinue
            $totalChats += $chats.Count
            $processedUsers++
        } catch {
            # Skip users without chat access
        }
    }
    
    # Extrapolate from sample
    $allUsersCount = (Get-MgUser -CountVariable userCount -ConsistencyLevel eventual).Count
    $estimatedChats = if ($processedUsers -gt 0) { 
        [math]::Round(($totalChats / $processedUsers) * $allUsersCount) 
    } else { 
        0 
    }
    
    $results += [PSCustomObject]@{
        Category = "Teams"
        ObjectType = "Chats (estimated)"
        Count = $estimatedChats
    }
    Write-Host "  Chats (estimated): $estimatedChats`n" -ForegroundColor White
} catch {
    Write-Host "  Unable to retrieve chats: $_`n" -ForegroundColor DarkYellow
}

# Guests count
Write-Host "→ Getting guests count..." -ForegroundColor Yellow
$guests = Get-MgUser -Filter "userType eq 'Guest'" -All -CountVariable guestsCount
$results += [PSCustomObject]@{
    Category = "Teams"
    ObjectType = "Guests"
    Count = $guests.Count
}
Write-Host "  Guests: $($guests.Count)`n" -ForegroundColor White

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "SHAREPOINT ONLINE METRICS" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# Get all site collections
Write-Host "→ Retrieving all SharePoint sites (this may take a while)..." -ForegroundColor Yellow
$allSites = Get-PnPTenantSite -IncludeOneDriveSites -Detailed
Write-Host "  Total sites retrieved: $($allSites.Count)`n" -ForegroundColor White

# Not connected to O365 Groups or Teams
Write-Host "→ Analyzing site connections..." -ForegroundColor Yellow
$notConnectedSites = ($allSites | Where-Object { 
    $_.GroupId -eq "00000000-0000-0000-0000-000000000000" -or 
    [string]::IsNullOrEmpty($_.GroupId) 
}).Count
$results += [PSCustomObject]@{
    Category = "SharePoint Online"
    ObjectType = "Not connected to O365 Groups or Teams"
    Count = $notConnectedSites
}
Write-Host "  Not connected to O365 Groups or Teams: $notConnectedSites" -ForegroundColor White

# Connected to O365 Groups (but not Teams)
$connectedToGroups = ($allSites | Where-Object {
    $_.GroupId -ne "00000000-0000-0000-0000-000000000000" -and
    -not [string]::IsNullOrEmpty($_.GroupId) -and
    $_.Template -notmatch "TEAMCHANNEL|STS#3"
}).Count
$results += [PSCustomObject]@{
    Category = "SharePoint Online"
    ObjectType = "Connected to O365 Groups"
    Count = $connectedToGroups
}
Write-Host "  Connected to O365 Groups: $connectedToGroups" -ForegroundColor White

# Connected to O365 Groups and Microsoft Teams
$connectedToTeams = ($allSites | Where-Object {
    $_.Template -match "TEAMCHANNEL" -or
    ($_.GroupId -ne "00000000-0000-0000-0000-000000000000" -and $_.Template -eq "GROUP#0")
}).Count
$results += [PSCustomObject]@{
    Category = "SharePoint Online"
    ObjectType = "Connected to O365 Groups and Microsoft Teams"
    Count = $connectedToTeams
}
Write-Host "  Connected to O365 Groups and Microsoft Teams: $connectedToTeams`n" -ForegroundColor White

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "SHAREPOINT ONLINE - OTHERS" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# Large site collection by size (>100GB = 102400 MB)
Write-Host "→ Analyzing large site collections by size..." -ForegroundColor Yellow
$largeSitesBySize = ($allSites | Where-Object { $_.StorageUsageCurrent -gt 102400 }).Count
$results += [PSCustomObject]@{
    Category = "SharePoint Online - Others"
    ObjectType = "Large site collection by size (>100GB)"
    Count = $largeSitesBySize
}
Write-Host "  Large site collections by size (>100GB): $largeSitesBySize`n" -ForegroundColor White

# Large site collection by item numbers (requires detailed analysis)
Write-Host "→ Analyzing large site collections by item count (top 50 sites)..." -ForegroundColor Yellow
$largeSitesByNumber = 0
$topSites = $allSites | Sort-Object -Property StorageUsageCurrent -Descending | Select-Object -First 50

foreach ($site in $topSites) {
    try {
        $itemCount = Get-PnPTenantSite -Identity $site.Url | Select-Object -ExpandProperty ItemCount
        if ($itemCount -gt 100000) {
            $largeSitesByNumber++
        }
    } catch {
        # Skip if unable to get item count
    }
}
$results += [PSCustomObject]@{
    Category = "SharePoint Online - Others"
    ObjectType = "Large site collection by numbers (>100K items - sample)"
    Count = $largeSitesByNumber
}
Write-Host "  Large site collections by numbers (>100K items): $largeSitesByNumber`n" -ForegroundColor White

# Workflow 2010 (requires site-by-site check - sample top sites)
Write-Host "→ Checking for Workflow 2010 (sampling approach)..." -ForegroundColor Yellow
$workflow2010Count = 0
Write-Host "  Workflow 2010 detection requires individual site inspection" -ForegroundColor DarkYellow
Write-Host "  Skipping detailed check for performance. Manual inspection recommended.`n" -ForegroundColor DarkYellow
$results += [PSCustomObject]@{
    Category = "SharePoint Online - Others"
    ObjectType = "Workflow 2010 (manual check required)"
    Count = "N/A"
}

# Checked-out files (requires detailed scan)
Write-Host "→ Files in check-out status..." -ForegroundColor Yellow
Write-Host "  Checked-out files require individual library inspection" -ForegroundColor DarkYellow
Write-Host "  Skipping for performance. Use separate script for detailed scan.`n" -ForegroundColor DarkYellow
$results += [PSCustomObject]@{
    Category = "SharePoint Online - Others"
    ObjectType = "Check-out file in use (manual check required)"
    Count = "N/A"
}

# Lists with Lookup Threshold Exceeded
Write-Host "→ Lists with lookup threshold exceeded..." -ForegroundColor Yellow
Write-Host "  Lookup threshold check requires individual list inspection" -ForegroundColor DarkYellow
Write-Host "  Skipping for performance.`n" -ForegroundColor DarkYellow
$results += [PSCustomObject]@{
    Category = "SharePoint Online - Others"
    ObjectType = "Lists With Lookup Threshold Exceeded (manual check required)"
    Count = "N/A"
}

# Sandbox Solutions
Write-Host "→ Sandbox solutions..." -ForegroundColor Yellow
Write-Host "  Sandbox solutions are deprecated in SharePoint Online" -ForegroundColor DarkYellow
Write-Host "  Modern SharePoint Online doesn't support sandbox solutions.`n" -ForegroundColor DarkYellow
$results += [PSCustomObject]@{
    Category = "SharePoint Online - Others"
    ObjectType = "Sandbox Solutions"
    Count = 0
}

# Custom Features, Templates, and Customized Pages
Write-Host "→ Custom features, templates, and pages..." -ForegroundColor Yellow
Write-Host "  These require individual site collection inspection" -ForegroundColor DarkYellow
Write-Host "  Skipping for performance.`n" -ForegroundColor DarkYellow

$results += [PSCustomObject]@{
    Category = "SharePoint Online - Others"
    ObjectType = "Custom Features (manual check required)"
    Count = "N/A"
}
$results += [PSCustomObject]@{
    Category = "SharePoint Online - Others"
    ObjectType = "Custom List Templates (manual check required)"
    Count = "N/A"
}
$results += [PSCustomObject]@{
    Category = "SharePoint Online - Others"
    ObjectType = "Custom Site Templates (manual check required)"
    Count = "N/A"
}
$results += [PSCustomObject]@{
    Category = "SharePoint Online - Others"
    ObjectType = "Customised Pages (manual check required)"
    Count = "N/A"
}

# Export results
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "EXPORTING RESULTS" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "✓ Results exported to: $OutputPath`n" -ForegroundColor Green

# Display summary table
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "SUMMARY" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan
$results | Format-Table -AutoSize

# Disconnect sessions
Write-Host "`nDisconnecting sessions..." -ForegroundColor Yellow
Disconnect-MgGraph -ErrorAction SilentlyContinue
Disconnect-PnPOnline -ErrorAction SilentlyContinue

Write-Host "✓ Script completed successfully!`n" -ForegroundColor Green
