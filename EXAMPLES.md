# 📚 Usage Examples - Microsoft 365 Object Counts Scripts

## Example 1: Basic Quick Assessment

**Scenario:** You need a quick count of Teams, SharePoint sites, and guests.

```powershell
# Install modules (one-time)
Install-Module Microsoft.Graph -Scope CurrentUser -Force
Install-Module PnP.PowerShell -Scope CurrentUser -Force

# Register your own PnP Entra ID App (one-time, required since Sept 9, 2024)
Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP.PowerShell" -Tenant "contoso.onmicrosoft.com"

# Run the fast script (pass the Client ID from the registration above)
.\Get-M365ObjectCounts-Fast.ps1 -AdminUrl "https://contoso-admin.sharepoint.com" -ClientId "<your-app-id>"

# Output:
# - Console display with all metrics
# - CSV file: M365ObjectCounts_20260722_083700.csv
```

**Time:** 5-15 minutes  
**Best for:** Regular monthly checks, initial assessments

---

## Example 2: Generate Beautiful HTML Report

**Scenario:** You need to present results to management in a visual format.

```powershell
# Step 1: Run the extraction
.\Get-M365ObjectCounts-Fast.ps1 -AdminUrl "https://contoso-admin.sharepoint.com" -ClientId "<your-app-id>"

# Step 2: Generate HTML report from CSV
.\Generate-HTML-Report.ps1 -CsvPath ".\M365ObjectCounts_20260722_083700.csv"

# Opens automatically in browser!
# Output: M365ObjectCounts_Report.html
```

**Result:** Professional HTML report with charts and color-coded tables

---

## Example 3: Custom Output Location

**Scenario:** You want to save reports to a specific network location.

```powershell
# Run with custom output path
.\Get-M365ObjectCounts-Fast.ps1 `
    -AdminUrl "https://contoso-admin.sharepoint.com" `
    -ClientId "<your-app-id>" `
    -OutputPath "\\networkshare\reports\M365Report_$(Get-Date -Format 'yyyy-MM').csv"

# Generate HTML in same location
.\Generate-HTML-Report.ps1 `
    -CsvPath "\\networkshare\reports\M365Report_2026-07.csv" `
    -OutputPath "\\networkshare\reports\M365Report_2026-07.html"
```

**Use case:** Centralized reporting, team access, historical tracking

---

## Example 4: Detailed Site Analysis

**Scenario:** You need comprehensive data including site-level custom features and workflows.

```powershell
# Run the detailed script (takes longer)
.\Get-M365ObjectCounts.ps1 -TenantUrl "https://contoso.sharepoint.com"

# This will:
# - Check each site individually
# - Count custom features, templates, workflows
# - Take 1-4 hours depending on tenant size
```

**Warning:** Only use this for small tenants or when absolutely needed!

---

## Example 5: Manual Command-by-Command Approach

**Scenario:** You want to learn or run specific checks only.

```powershell
# Connect to services
Connect-MgGraph -Scopes "Group.Read.All", "Sites.Read.All", "User.Read.All"
Connect-PnPOnline -Url "https://contoso-admin.sharepoint.com" -Interactive -ClientId "<your-app-id>"

# Get Teams count
$teamsCount = (Get-MgGroup -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -All).Count
Write-Host "Teams sites: $teamsCount"

# Get Guests count
$guestsCount = (Get-MgUser -Filter "userType eq 'Guest'" -All).Count
Write-Host "Guests: $guestsCount"

# Get all SharePoint sites
$allSites = Get-PnPTenantSite -Detailed

# Not connected to Groups
$notConnected = ($allSites | Where-Object { 
    $_.GroupId -eq "00000000-0000-0000-0000-000000000000" 
}).Count
Write-Host "Sites not connected to Groups: $notConnected"

# Disconnect
Disconnect-MgGraph
Disconnect-PnPOnline
```

**Best for:** Learning, debugging, specific metric extraction

---

## Example 6: Scheduled Monthly Report

**Scenario:** Automate monthly reporting with Task Scheduler.

### Create a wrapper script: `Monthly-M365-Report.ps1`

```powershell
# Monthly-M365-Report.ps1
$reportDate = Get-Date -Format "yyyy-MM"
$outputFolder = "C:\M365Reports"

# Ensure folder exists
if (-not (Test-Path $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder
}

# Run extraction
.\Get-M365ObjectCounts-Fast.ps1 `
    -AdminUrl "https://contoso-admin.sharepoint.com" `
    -ClientId "<your-app-id>" `
    -OutputPath "$outputFolder\M365Report_$reportDate.csv"

# Generate HTML
$csvPath = "$outputFolder\M365Report_$reportDate.csv"
.\Generate-HTML-Report.ps1 `
    -CsvPath $csvPath `
    -OutputPath "$outputFolder\M365Report_$reportDate.html"

# Email report (optional)
Send-MailMessage `
    -From "reports@contoso.com" `
    -To "admin@contoso.com" `
    -Subject "M365 Monthly Report - $reportDate" `
    -Body "See attached monthly M365 object counts report." `
    -Attachments "$outputFolder\M365Report_$reportDate.html" `
    -SmtpServer "smtp.contoso.com"
```

### Schedule with Task Scheduler:

```powershell
# Create scheduled task
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" `
    -Argument "-File C:\Scripts\Monthly-M365-Report.ps1"

$trigger = New-ScheduledTaskTrigger -Monthly -DaysOfMonth 1 -At 2am

Register-ScheduledTask -TaskName "M365 Monthly Report" `
    -Action $action `
    -Trigger $trigger `
    -Description "Automated monthly M365 object counts report"
```

---

## Example 7: Compare Two Time Periods

**Scenario:** Track growth month-over-month.

```powershell
# Import this month's report
$currentMonth = Import-Csv ".\M365Report_2026-07.csv"

# Import last month's report
$lastMonth = Import-Csv ".\M365Report_2026-06.csv"

# Compare Teams sites
$currentTeams = ($currentMonth | Where-Object { $_.ObjectType -eq "Teams sites" }).Count
$lastMonthTeams = ($lastMonth | Where-Object { $_.ObjectType -eq "Teams sites" }).Count
$growth = $currentTeams - $lastMonthTeams

Write-Host "Teams Growth: $growth sites ($lastMonthTeams → $currentTeams)"

# Compare all metrics
$comparison = @()
foreach ($metric in $currentMonth) {
    $lastValue = ($lastMonth | Where-Object { $_.ObjectType -eq $metric.ObjectType }).Count
    $comparison += [PSCustomObject]@{
        Metric = $metric.ObjectType
        LastMonth = $lastValue
        ThisMonth = $metric.Count
        Change = [int]$metric.Count - [int]$lastValue
    }
}

$comparison | Format-Table -AutoSize
```

---

## Example 8: Specific Site Deep Dive

**Scenario:** Check specific site for workflows, custom features, etc.

```powershell
# Connect to specific site
Connect-PnPOnline -Url "https://contoso.sharepoint.com/sites/ProjectX" -Interactive -ClientId "<your-app-id>"

# Check for Workflow 2010
$workflows = Get-PnPWorkflowDefinition
Write-Host "Workflows found: $($workflows.Count)"

# Check for checked-out files
$docLibs = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 }
foreach ($lib in $docLibs) {
    $checkedOut = Get-PnPListItem -List $lib | Where-Object { $_["CheckoutUser"] -ne $null }
    if ($checkedOut) {
        Write-Host "$($lib.Title): $($checkedOut.Count) files checked out"
    }
}

# Check for custom features
$customFeatures = Get-PnPFeature -Scope Site | Where-Object { 
    $_.DefinitionId -notlike "00000000-0000-0000-0000-*" 
}
Write-Host "Custom features: $($customFeatures.Count)"

# Check lookup columns in a list
$lists = Get-PnPList
foreach ($list in $lists) {
    $fields = Get-PnPField -List $list
    $lookups = $fields | Where-Object { $_.TypeAsString -eq "Lookup" }
    if ($lookups.Count -gt 12) {
        Write-Host "⚠️ $($list.Title) has $($lookups.Count) lookup columns (threshold exceeded!)"
    }
}
```

---

## Example 9: Export Multiple Formats

**Scenario:** Create CSV, HTML, and JSON reports.

```powershell
# Run extraction
.\Get-M365ObjectCounts-Fast.ps1 `
    -AdminUrl "https://contoso-admin.sharepoint.com" `
    -ClientId "<your-app-id>" `
    -OutputPath ".\M365Report.csv"

# Import data
$data = Import-Csv ".\M365Report.csv"

# Export as JSON
$data | ConvertTo-Json | Out-File ".\M365Report.json"

# Export as HTML
.\Generate-HTML-Report.ps1 -CsvPath ".\M365Report.csv"

# Export as Excel (requires ImportExcel module)
# Install-Module ImportExcel -Scope CurrentUser
$data | Export-Excel ".\M365Report.xlsx" -AutoSize -TableName "M365Metrics"

Write-Host "Reports generated in CSV, JSON, HTML, and Excel formats!"
```

---

## Example 10: Filter Large Sites Only

**Scenario:** Extract only large sites (>50GB) for capacity planning.

```powershell
# Connect
Connect-PnPOnline -Url "https://contoso-admin.sharepoint.com" -Interactive -ClientId "<your-app-id>"

# Get all sites
$allSites = Get-PnPTenantSite -Detailed

# Filter large sites (>50GB = 51200 MB)
$largeSites = $allSites | Where-Object { $_.StorageUsageCurrent -gt 51200 } | 
    Select-Object Title, Url, 
        @{N='SizeGB';E={[math]::Round($_.StorageUsageCurrent/1024, 2)}},
        @{N='PercentUsed';E={[math]::Round(($_.StorageUsageCurrent/$_.StorageMaximumLevel)*100, 1)}} |
    Sort-Object SizeGB -Descending

# Export
$largeSites | Export-Csv ".\LargeSites.csv" -NoTypeInformation
$largeSites | Format-Table -AutoSize

Write-Host "`nFound $($largeSites.Count) sites over 50GB"
```

---

## 🎯 Quick Reference

| Use Case | Script to Use | Execution Time |
|----------|---------------|----------------|
| Regular monthly check | Get-M365ObjectCounts-Fast.ps1 | 5-15 min |
| Visual report for management | Generate-HTML-Report.ps1 | < 1 min |
| Deep analysis | Get-M365ObjectCounts.ps1 | 1-4 hours |
| Learning/specific checks | M365-Quick-Commands.ps1 | As needed |
| Site-specific analysis | Manual commands | 5-10 min |

---

## 💡 Pro Tips

1. **Always test first** on a dev/test tenant if available
2. **Schedule reports** to run monthly during off-hours
3. **Keep historical data** to track growth trends
4. **Use HTML reports** for stakeholder presentations
5. **Use CSV reports** for data analysis in Excel/Power BI
6. **Filter results** based on your specific needs
7. **Document anomalies** when you find them

---

## 🔗 Related Commands

```powershell
# List all Microsoft 365 Groups
Get-MgGroup -All

# List all Teams
Get-MgGroup -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -All

# List all Guest users
Get-MgUser -Filter "userType eq 'Guest'" -All

# Get SharePoint storage report
Get-PnPTenantSite | Select-Object Url, StorageUsageCurrent, StorageQuota

# Get Teams usage report (requires Reports.Read.All permission)
Get-MgReportTeamsUserActivityUserDetail -Period D30
```

---

**Need more examples?** Check the README-M365-Scripts.md for additional details!
