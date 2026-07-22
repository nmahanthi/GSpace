# Microsoft 365 Object Counts and Volumes - PowerShell Scripts

This repository contains PowerShell scripts to extract Microsoft 365 object counts and volumes as shown in the Microsoft 365 assessment template.

## 📋 Scripts Overview

### 1. `Get-M365ObjectCounts-Fast.ps1` (Recommended)
**Optimized script** for quick extraction of most metrics. Some detailed metrics require manual inspection.

**Pros:**
- Fast execution
- Uses Graph API efficiently
- Good for initial assessment
- Provides accurate counts for most metrics

**Cons:**
- Some detailed metrics marked as "manual check required"
- Uses sampling for chats count

### 2. `Get-M365ObjectCounts.ps1` (Detailed)
**Comprehensive script** that attempts to extract all metrics, including detailed SharePoint site-level checks.

**Pros:**
- Attempts to get all metrics
- More thorough analysis

**Cons:**
- Much slower execution (can take hours for large tenants)
- Requires connecting to each site individually
- May timeout on large tenants

## 🚀 Prerequisites

### Required PowerShell Modules

Install these modules before running the scripts:

```powershell
# Microsoft Graph PowerShell SDK
Install-Module Microsoft.Graph -Scope CurrentUser -Force

# PnP PowerShell for SharePoint Online
Install-Module PnP.PowerShell -Scope CurrentUser -Force

# (Optional) Microsoft Teams module for detailed Teams data
Install-Module MicrosoftTeams -Scope CurrentUser -Force
```

### Register a PnP Entra ID App (One-Time, Required Since Sept 9, 2024)

PnP PowerShell retired its shared multi-tenant app, so `Connect-PnPOnline -Interactive` now requires your own Entra ID App Registration. This does **not** apply to Microsoft Graph (`Connect-MgGraph` uses Microsoft's first-party app).

```powershell
Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP.PowerShell" -Tenant "yourtenant.onmicrosoft.com"
```

Note the **Application (client) ID** it outputs — pass it via `-ClientId` to the scripts, or set it once as an environment variable (`ENTRAID_APP_ID`) so you don't have to specify it every time. See `PERMISSIONS-GUIDE.md` for full details.

### Required Permissions

You need the following permissions:
- **Global Reader** or **SharePoint Administrator** role in Microsoft 365
- **Sites.Read.All** (Application or Delegated)
- **Group.Read.All** (Application or Delegated)
- **Team.ReadBasic.All** (Application or Delegated)
- **User.Read.All** (Application or Delegated)
- **Chat.Read.All** (Delegated) - for chats count

## 📖 Usage

### Quick Start (Recommended)

```powershell
# Run the fast version
.\Get-M365ObjectCounts-Fast.ps1 -AdminUrl "https://yourtenant-admin.sharepoint.com" -ClientId "<your-app-id>"
```

### Detailed Scan

```powershell
# Run the detailed version (may take hours)
.\Get-M365ObjectCounts.ps1 -TenantUrl "https://yourtenant.sharepoint.com" -ClientId "<your-app-id>"
```

### Custom Output Path

```powershell
# Specify custom output path
.\Get-M365ObjectCounts-Fast.ps1 `
    -AdminUrl "https://contoso-admin.sharepoint.com" `
    -ClientId "<your-app-id>" `
    -OutputPath "C:\Reports\M365Counts.csv"
```

## 📊 Metrics Extracted

### Teams Metrics
| Metric | Description | Script Coverage |
|--------|-------------|-----------------|
| Teams sites | Count of Teams provisioned | ✅ Full |
| Groups with Dynamic membership | Groups with dynamic membership rules | ✅ Full |
| Chats | Total chat count | ⚠️ Estimated (sampling) |
| Guests | Guest user count | ✅ Full |

### SharePoint Online Metrics
| Metric | Description | Script Coverage |
|--------|-------------|-----------------|
| Not connected to O365 Groups or Teams | Standalone SharePoint sites | ✅ Full |
| Connected to O365 Groups | Sites with O365 Group (no Teams) | ✅ Full |
| Connected to O365 Groups and Microsoft Teams | Team sites | ✅ Full |

### SharePoint Online - Others
| Metric | Description | Script Coverage |
|--------|-------------|-----------------|
| Large site collection by size | Sites >100GB | ✅ Full |
| Large site collection by numbers | Sites >100K items | ⚠️ Sample (top 50) |
| Workflow 2010 | Legacy workflows | ⚠️ Manual check required |
| Check-out file in use | Currently checked out files | ⚠️ Manual check required |
| Lists With Lookup Threshold Exceeded | Lists with >12 lookup columns | ⚠️ Manual check required |
| Sandbox Solutions | Deprecated feature | ✅ N/A (not supported) |
| Custom Features | Custom site/web features | ⚠️ Manual check required |
| Custom List Templates | Custom list templates | ⚠️ Manual check required |
| Custom Site Templates | Custom site templates | ⚠️ Manual check required |
| Customised Pages | Customized SharePoint pages | ⚠️ Manual check required |

**Legend:**
- ✅ Full: Complete automated extraction
- ⚠️ Sample: Partial extraction or sampling approach
- ⚠️ Manual check required: Requires site-by-site inspection

## 📁 Output

The script generates a CSV file with the following columns:
- **Category**: The main category (Teams, SharePoint Online, etc.)
- **ObjectType**: Specific object type being counted
- **Count**: The count or "N/A" if manual check required

Example output:
```csv
Category,ObjectType,Count
Teams,Teams sites,274
Teams,Groups with Dynamic membership,17
Teams,Chats (estimated),1250
Teams,Guests,45
SharePoint Online,Not connected to O365 Groups or Teams,120
SharePoint Online,Connected to O365 Groups,85
SharePoint Online,Connected to O365 Groups and Microsoft Teams,274
...
```

## 🔍 Manual Checks for Detailed Metrics

For metrics marked as "manual check required", use these approaches:

### Workflow 2010 Count
```powershell
# Connect to specific site
Connect-PnPOnline -Url "https://yourtenant.sharepoint.com/sites/yoursite" -Interactive -ClientId "<your-app-id>"

# Check workflows
Get-PnPWorkflowDefinition
```

### Checked-out Files
```powershell
# Get all checked out files in a library
Get-PnPListItem -List "Documents" | Where-Object { $_["CheckoutUser"] -ne $null }
```

### Lists with Lookup Threshold
```powershell
# Check lookup columns in a list
$fields = Get-PnPField -List "YourList"
$lookupFields = $fields | Where-Object { $_.TypeAsString -eq "Lookup" }
$lookupFields.Count
```

## ⚡ Performance Tips

1. **Run during off-peak hours**: Large tenant scans can take time
2. **Use the Fast version first**: Get initial metrics quickly
3. **Target specific sites**: For detailed checks, focus on high-priority sites
4. **Increase timeout**: For large tenants, consider increasing PnP timeout:
   ```powershell
   Set-PnPRequestTimeout -Timeout 300000
   ```

## 🛠️ Troubleshooting

### Authentication Issues
```powershell
# Clear cached credentials
Disconnect-MgGraph
Disconnect-PnPOnline

# Reconnect with fresh authentication
Connect-MgGraph -Scopes "Group.Read.All","Sites.Read.All" -ForceRefresh
```

### Permission Errors
Ensure you have Global Reader or SharePoint Admin role. Contact your tenant administrator.

### Script Timeout
For large tenants, run the fast version or break the scan into smaller batches.

## 📝 Notes

- **Chats count**: Due to API limitations, we use a sampling approach. For exact count, you'd need to query every user individually (very slow).
- **Sandbox Solutions**: Deprecated in SharePoint Online, script returns 0.
- **Custom features/templates**: Require site-by-site inspection, best done with targeted scans.

## 🔗 References

- [Microsoft Graph PowerShell](https://learn.microsoft.com/en-us/powershell/microsoftgraph/)
- [PnP PowerShell](https://pnp.github.io/powershell/)
- [SharePoint Online Limits](https://learn.microsoft.com/en-us/office365/servicedescriptions/sharepoint-online-service-description/sharepoint-online-limits)

## 📄 License

Free to use and modify for your Microsoft 365 assessments.
