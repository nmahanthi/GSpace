# 🚀 Quick Start Guide - Microsoft 365 Object Counts Extraction

## 📦 What You Got

Three PowerShell scripts to extract Microsoft 365 metrics from your tenant:

1. **Get-M365ObjectCounts-Fast.ps1** - ⚡ Recommended for most users
2. **Get-M365ObjectCounts.ps1** - 🔍 Deep dive (slow but thorough)
3. **M365-Quick-Commands.ps1** - 📝 Individual commands to run manually

## ⏱️ 5-Minute Setup

### Step 1: Install Required Modules (One-Time)

Open PowerShell as Administrator and run:

```powershell
# Install Microsoft Graph
Install-Module Microsoft.Graph -Scope CurrentUser -Force

# Install PnP PowerShell for SharePoint
Install-Module PnP.PowerShell -Scope CurrentUser -Force
```

### Step 2: Register a PnP Entra ID App (One-Time, Required Since Sept 9, 2024)

PnP PowerShell retired its shared multi-tenant app, so `Connect-PnPOnline -Interactive` now requires your own Entra ID App Registration. Register one (needs Application Administrator or Global Admin rights):

```powershell
Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP.PowerShell" -Tenant "yourtenant.onmicrosoft.com"
```

Note the **Application (client) ID** it outputs. Either pass it as `-ClientId` every time, or set it once as an environment variable:

```powershell
[Environment]::SetEnvironmentVariable("ENTRAID_APP_ID", "<your-app-id>", "User")
```

*(Microsoft Graph does NOT need this — `Connect-MgGraph` uses Microsoft's own first-party app.)*

### Step 3: Run the Script

```powershell
# Navigate to the script folder
cd "C:\path\to\scripts"

# Run the fast version (replace 'yourtenant' and 'your-app-id' with your actual values)
.\Get-M365ObjectCounts-Fast.ps1 -AdminUrl "https://yourtenant-admin.sharepoint.com" -ClientId "<your-app-id>"
```

**Example:**
```powershell
.\Get-M365ObjectCounts-Fast.ps1 -AdminUrl "https://contoso-admin.sharepoint.com" -ClientId "11111111-2222-3333-4444-555555555555"
```

*(If you set the `ENTRAID_APP_ID` environment variable in Step 2, you can omit `-ClientId`.)*

### Step 4: Authenticate

When prompted:
1. Sign in with your Microsoft 365 admin account
2. Grant consent to the required permissions
3. Wait for the script to complete

### Step 5: Review Results

The script will:
- Display results in the console
- Export a CSV file: `M365ObjectCounts_YYYYMMDD_HHMMSS.csv`

## 📊 What Metrics You'll Get

### ✅ Automatically Extracted

| Metric | Example Count |
|--------|---------------|
| Teams sites | 274 |
| Groups with Dynamic membership | 17 |
| Guest users | 45 |
| SharePoint sites not connected to Groups | 120 |
| SharePoint sites connected to O365 Groups | 85 |
| SharePoint sites connected to Teams | 274 |
| Large sites (>100GB) | 5 |

### ⚠️ Requires Manual Check

Some metrics require site-by-site inspection and are marked as "manual check required":
- Workflow 2010 count
- Checked-out files
- Lists with lookup threshold exceeded
- Custom features, templates, and customized pages

## 🎯 Usage Examples

### Basic Usage
```powershell
.\Get-M365ObjectCounts-Fast.ps1 -AdminUrl "https://contoso-admin.sharepoint.com" -ClientId "<your-app-id>"
```

### Custom Output Location
```powershell
.\Get-M365ObjectCounts-Fast.ps1 `
    -AdminUrl "https://contoso-admin.sharepoint.com" `
    -ClientId "<your-app-id>" `
    -OutputPath "C:\Reports\M365Report.csv"
```

### Run Individual Commands
If you prefer to run commands one by one:

```powershell
# Open and follow the commands in this file
.\M365-Quick-Commands.ps1
```

## 🔧 Troubleshooting

### Issue: "Module not found"
**Solution:** Install the required modules
```powershell
Install-Module Microsoft.Graph -Force
Install-Module PnP.PowerShell -Force
```

### Issue: "Access denied"
**Solution:** Ensure you have one of these roles:
- Global Administrator
- Global Reader
- SharePoint Administrator

### Issue: PnP asks for "-ClientId" / "No Entra ID App registered" error
**Solution:** PnP PowerShell requires your own Entra ID App Registration since Sept 9, 2024. Run once:
```powershell
Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP.PowerShell" -Tenant "yourtenant.onmicrosoft.com"
```
Then pass the resulting app ID via `-ClientId` (or set `ENTRAID_APP_ID` as an environment variable).

### Issue: "Cannot connect to tenant"
**Solution:** Verify your admin URL format:
- ✅ Correct: `https://contoso-admin.sharepoint.com`
- ❌ Wrong: `https://contoso.sharepoint.com`

### Issue: Script runs too slow
**Solution:** 
- Use the Fast version (Get-M365ObjectCounts-Fast.ps1)
- Run during off-peak hours
- For large tenants (>1000 sites), expect 15-30 minutes

## 📖 Understanding Your Results

### Output CSV Structure

```csv
Category,ObjectType,Count
Teams,Teams sites,274
Teams,Groups with Dynamic membership,17
SharePoint Online,Not connected to O365 Groups or Teams,120
...
```

### Key Metrics Explained

**Teams Sites (274)**: Total number of Microsoft Teams created

**Groups with Dynamic Membership (17)**: Groups that auto-add/remove members based on rules

**Guests (45)**: External users invited to your tenant

**Not Connected to O365 Groups (120)**: Classic SharePoint sites

**Connected to O365 Groups (85)**: Modern SharePoint sites with Groups (no Teams)

**Connected to Teams (274)**: SharePoint sites backing Microsoft Teams

**Large Sites (>100GB)**: Sites consuming significant storage

## 🎓 Next Steps

### For Complete Assessment
1. Run the fast script to get baseline metrics ✅
2. Review the CSV output
3. For detailed metrics, run targeted checks on specific sites
4. Use the detailed script for critical sites only

### For Manual Checks
See the README-M365-Scripts.md file for commands to check:
- Workflow 2010
- Checked-out files
- Custom features and templates

## 📞 Need Help?

**Common Questions:**

**Q: How long does it take?**
A: Fast script: 5-15 minutes. Detailed script: 1-4 hours depending on tenant size.

**Q: Will this modify anything?**
A: No, all scripts are read-only. They only gather information.

**Q: Can I run this on a schedule?**
A: Yes! Use Task Scheduler to run the script weekly/monthly.

**Q: What permissions are needed?**
A: At minimum: Sites.Read.All, Group.Read.All, User.Read.All

**Q: Do I need an app registration?**
A: Not for Microsoft Graph. But for SharePoint/PnP, yes — a one-time, self-service Entra ID App Registration is required since Sept 9, 2024 (see Step 2 above). See PERMISSIONS-GUIDE.md for full details.

## 📂 Files Overview

```
├── Get-M365ObjectCounts-Fast.ps1      # ⚡ RECOMMENDED - Quick extraction
├── Get-M365ObjectCounts.ps1           # 🔍 Detailed extraction (slow)
├── M365-Quick-Commands.ps1            # 📝 Individual commands
├── README-M365-Scripts.md             # 📚 Detailed documentation
└── QUICK-START.md                     # 🚀 This file
```

## ✨ Pro Tips

1. **Test first**: Run on a test tenant if available
2. **Off-peak hours**: Run during nights/weekends for large tenants
3. **Save outputs**: Keep historical reports to track growth
4. **Focus**: Use fast version for regular checks, detailed for deep dives
5. **Batch processing**: For large tenants, consider breaking into batches

---

**Ready to start?** After registering your PnP Entra ID App (Step 2), run this command now:

```powershell
.\Get-M365ObjectCounts-Fast.ps1 -AdminUrl "https://YOUR-TENANT-admin.sharepoint.com" -ClientId "<your-app-id>"
```

Replace `YOUR-TENANT` and `<your-app-id>` with your actual values! 🎉
