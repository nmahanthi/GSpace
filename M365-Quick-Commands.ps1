# Microsoft 365 Quick Commands for Object Counts
# Run these commands individually to get specific metrics

# ============================================
# PREREQUISITES - Install Modules (run once)
# ============================================

# Install Microsoft Graph PowerShell
Install-Module Microsoft.Graph -Scope CurrentUser -Force

# Install PnP PowerShell
Install-Module PnP.PowerShell -Scope CurrentUser -Force

# Install Teams PowerShell
Install-Module MicrosoftTeams -Scope CurrentUser -Force

# ============================================
# ONE-TIME SETUP - Register your own PnP Entra ID App
# ============================================
# Since Sept 9, 2024, PnP PowerShell requires your own Entra ID App
# Registration for -Interactive login (the old shared PnP app was retired).
# Run this ONCE per tenant (requires rights to create App Registrations):

# Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP.PowerShell" -Tenant "yourtenant.onmicrosoft.com"
# Note the "Application (client) ID" it outputs, then either:
#   a) Pass it via -ClientId on every Connect-PnPOnline call below, OR
#   b) Set it once as an environment variable so you never have to pass it:
#      [Environment]::SetEnvironmentVariable("ENTRAID_APP_ID", "<your-app-id>", "User")

# ============================================
# CONNECT TO SERVICES
# ============================================

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Group.Read.All", "Team.ReadBasic.All", "Sites.Read.All", "User.Read.All", "Chat.Read.All"

# Connect to SharePoint Admin Center (replace with your tenant and ClientId)
Connect-PnPOnline -Url "https://yourtenant-admin.sharepoint.com" -Interactive -ClientId "<your-app-id>"

# Connect to Teams (if needed)
Connect-MicrosoftTeams

# ============================================
# TEAMS METRICS
# ============================================

# Teams sites count
$teamsCount = (Get-MgGroup -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -All).Count
Write-Host "Teams sites: $teamsCount" -ForegroundColor Green

# Groups with Dynamic membership
$dynamicGroups = (Get-MgGroup -Filter "groupTypes/any(c:c eq 'DynamicMembership')" -All).Count
Write-Host "Groups with Dynamic membership: $dynamicGroups" -ForegroundColor Green

# Guest users count
$guestsCount = (Get-MgUser -Filter "userType eq 'Guest'" -All).Count
Write-Host "Guests: $guestsCount" -ForegroundColor Green

# Total chats (sample - first 10 users)
$sampleUsers = Get-MgUser -Top 10
$totalChats = 0
foreach ($user in $sampleUsers) {
    try {
        $chats = Get-MgUserChat -UserId $user.Id -All -ErrorAction SilentlyContinue
        $totalChats += $chats.Count
    } catch { }
}
Write-Host "Chats (sample from 10 users): $totalChats" -ForegroundColor Yellow

# ============================================
# SHAREPOINT ONLINE METRICS
# ============================================

# Get all sites
$allSites = Get-PnPTenantSite -Detailed
Write-Host "Total SharePoint sites: $($allSites.Count)" -ForegroundColor Cyan

# Not connected to O365 Groups or Teams
$notConnected = ($allSites | Where-Object { 
    $_.GroupId -eq "00000000-0000-0000-0000-000000000000" -or 
    [string]::IsNullOrEmpty($_.GroupId) 
}).Count
Write-Host "Not connected to O365 Groups or Teams: $notConnected" -ForegroundColor Green

# Connected to O365 Groups (not Teams)
$connectedToGroups = ($allSites | Where-Object { 
    $_.GroupId -ne "00000000-0000-0000-0000-000000000000" -and
    -not [string]::IsNullOrEmpty($_.GroupId) -and
    $_.Template -notmatch "TEAMCHANNEL"
}).Count
Write-Host "Connected to O365 Groups: $connectedToGroups" -ForegroundColor Green

# Connected to Teams
$connectedToTeams = ($allSites | Where-Object { 
    $_.Template -match "TEAMCHANNEL" -or 
    $_.Template -eq "GROUP#0"
}).Count
Write-Host "Connected to O365 Groups and Microsoft Teams: $connectedToTeams" -ForegroundColor Green

# ============================================
# SHAREPOINT ONLINE - OTHERS
# ============================================

# Large site collections by size (>100GB = 102400 MB)
$largeSitesBySize = ($allSites | Where-Object { $_.StorageUsageCurrent -gt 102400 }).Count
Write-Host "Large site collection by size (>100GB): $largeSitesBySize" -ForegroundColor Green

# Large site collections by storage (>50GB for visibility)
$largeSites50GB = ($allSites | Where-Object { $_.StorageUsageCurrent -gt 51200 })
Write-Host "`nSites over 50GB:" -ForegroundColor Yellow
$largeSites50GB | Select-Object Title, Url, @{N='SizeGB';E={[math]::Round($_.StorageUsageCurrent/1024, 2)}} | Format-Table

# Sites by template type distribution
Write-Host "`nSite Template Distribution:" -ForegroundColor Yellow
$allSites | Group-Object Template | Select-Object Name, Count | Sort-Object Count -Descending | Format-Table

# ============================================
# DETAILED CHECKS (run on specific sites)
# ============================================

# To check a specific site for detailed metrics, connect to it:
# Connect-PnPOnline -Url "https://yourtenant.sharepoint.com/sites/yoursite" -Interactive -ClientId "<your-app-id>"

# Check for Workflow 2010
# Get-PnPWorkflowDefinition

# Check for checked-out files in a library
# Get-PnPListItem -List "Documents" | Where-Object { $_["CheckoutUser"] -ne $null }

# Check lookup columns in a list
# $fields = Get-PnPField -List "YourList"
# ($fields | Where-Object { $_.TypeAsString -eq "Lookup" }).Count

# Check for custom features
# Get-PnPFeature -Scope Site

# Check for customized pages
# Get-PnPFile -AsListItem | Where-Object { $_["vti_x005f_level"] -eq "2" }

# ============================================
# EXPORT ALL DATA TO CSV
# ============================================

# Create results array
$results = @(
    [PSCustomObject]@{Category="Teams"; ObjectType="Teams sites"; Count=$teamsCount}
    [PSCustomObject]@{Category="Teams"; ObjectType="Groups with Dynamic membership"; Count=$dynamicGroups}
    [PSCustomObject]@{Category="Teams"; ObjectType="Guests"; Count=$guestsCount}
    [PSCustomObject]@{Category="SharePoint Online"; ObjectType="Not connected to O365 Groups or Teams"; Count=$notConnected}
    [PSCustomObject]@{Category="SharePoint Online"; ObjectType="Connected to O365 Groups"; Count=$connectedToGroups}
    [PSCustomObject]@{Category="SharePoint Online"; ObjectType="Connected to O365 Groups and Microsoft Teams"; Count=$connectedToTeams}
    [PSCustomObject]@{Category="SharePoint Online - Others"; ObjectType="Large site collection by size (>100GB)"; Count=$largeSitesBySize}
)

# Export to CSV
$outputPath = ".\M365ObjectCounts_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$results | Export-Csv -Path $outputPath -NoTypeInformation
Write-Host "`nResults exported to: $outputPath" -ForegroundColor Cyan

# Display results
$results | Format-Table -AutoSize

# ============================================
# DISCONNECT
# ============================================

# Disconnect-MgGraph
# Disconnect-PnPOnline
# Disconnect-MicrosoftTeams
