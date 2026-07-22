# 🔐 Permissions Guide - Microsoft 365 Object Counts Scripts

## ❓ Do I Need App Registration?

**Short answer: Partially — it depends on which part of the script.**

| Component | App Registration Needed? |
|---|---|
| **Microsoft Graph** (`Connect-MgGraph`) | ❌ **NO** — uses Microsoft's first-party "Microsoft Graph Command Line Tools" app, pre-registered in most tenants |
| **SharePoint/PnP** (`Connect-PnPOnline -Interactive`) | ✅ **YES, as of Sept 9, 2024** — PnP retired its shared multi-tenant app. You must register your own (free, one-time, ~2 minutes) Entra ID App and pass its `-ClientId` |

This is **not** because interactive/MSAL login is deprecated — it's still the default, recommended method for both Graph and PnP. It's specifically that **PnP PowerShell no longer ships a shared app ID** you can piggyback on.

### One-time setup (do this once per tenant):

```powershell
Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP.PowerShell" -Tenant "yourtenant.onmicrosoft.com"
```
- You'll sign in with an account that can create App Registrations (Application Administrator or Global Admin).
- Note the **Application (client) ID** it prints at the end.
- Use it with `-ClientId <app-id>` on the scripts, or set it once as an environment variable so you never type it again:
  ```powershell
  [Environment]::SetEnvironmentVariable("ENTRAID_APP_ID", "<your-app-id>", "User")
  ```

No client secret, certificate, or admin-consent workflow is required for this — it's a delegated-permission app used only for interactive sign-in.

---

## 📋 Required Permissions

### 1️⃣ **User Account Role** (at least ONE of these):

You need one of these Azure AD roles assigned to your account:

| Role | Recommended? | Why? |
|------|-------------|------|
| **Global Reader** | ✅ **YES** | Read-only access to everything (safest) |
| **Global Administrator** | ⚠️ Use if you already have it | Full access (more than needed) |
| **SharePoint Administrator** | ✅ **YES** | Good for SharePoint-focused tasks |
| SharePoint Administrator + Reports Reader | ✅ **YES** | Ideal combination |

**Best Practice:** Use **Global Reader** role for read-only access.

### 2️⃣ **API Permissions** (Delegated - Automatic):

When you run the script and sign in, you'll be asked to consent to these permissions:

#### Microsoft Graph API:
```
✅ Group.Read.All          - Read all groups (Teams, O365 Groups)
✅ Team.ReadBasic.All      - Read Teams information
✅ Sites.Read.All          - Read SharePoint sites
✅ User.Read.All           - Read user information (for Guests count)
✅ Chat.Read.All           - Read chats (for chat count - optional)
```

#### SharePoint Online (PnP):
```
✅ AllSites.Read           - Read all SharePoint sites
✅ TermStore.Read.All      - Read term store (automatic)
```

---

## 🚀 How It Works (Interactive Authentication)

### When You Run the Script:

```powershell
.\Get-M365ObjectCounts-Fast.ps1 -AdminUrl "https://contoso-admin.sharepoint.com"
```

### What Happens:

1. **Browser Opens** - A browser window opens
2. **Sign In** - You sign in with your admin account
3. **Consent Prompt** - You see a permission consent screen (first time only)
4. **Grant Consent** - Click "Accept" to grant permissions
5. **Script Runs** - Script extracts data using your permissions

> **Note:** For the SharePoint/PnP connection step, you must first have registered your own PnP Entra ID App (see above) and pass its ID via `-ClientId`, otherwise `Connect-PnPOnline -Interactive` will fail with an error asking for a Client ID.

### First-Time Consent Screen:

You'll see something like this:

```
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Microsoft Graph Command Line Tools

This application would like to:
  ✓ Read all groups
  ✓ Read basic Teams information
  ✓ Read SharePoint sites
  ✓ Read all users
  ✓ Read user chat messages

[Cancel]  [Accept]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

**Click "Accept"** - This is safe! These are read-only permissions.

---

## 🛡️ Security & Safety

### ✅ What the Scripts CAN Do:
- **Read** Teams, Groups, SharePoint sites
- **Read** user information (names, guest status)
- **Read** site sizes, templates, connections
- **Export** data to CSV files

### ❌ What the Scripts CANNOT Do:
- ❌ Create, modify, or delete anything
- ❌ Change permissions or settings
- ❌ Access user emails or files
- ❌ Modify your tenant configuration
- ❌ Install or deploy anything

**All permissions are READ-ONLY!** 🔒

---

## 📝 Step-by-Step: First Run

### Before Running:

1. **Check your role:**
   ```powershell
   # Connect to Azure AD (if you want to verify)
   Connect-AzureAD
   Get-AzureADDirectoryRole | Where-Object {
       $_.DisplayName -match "Global|SharePoint"
   } | Get-AzureADDirectoryRoleMember | Where-Object {
       $_.UserPrincipalName -eq "your.email@contoso.com"
   }
   ```

2. **Verify you have admin access to SharePoint Admin Center:**
   - Visit: `https://yourtenant-admin.sharepoint.com`
   - If you can access it, you have the right permissions

### During First Run:

1. **Run the script (with your registered ClientId):**
   ```powershell
   .\Get-M365ObjectCounts-Fast.ps1 -AdminUrl "https://contoso-admin.sharepoint.com" -ClientId "<your-app-id>"
   ```

2. **Microsoft Graph Authentication:**
   - Browser opens → Sign in with admin account
   - **Consent screen appears** (first time only)
   - **Review permissions** → Click **"Accept"**
   - Browser shows: "Authentication complete. You can close this window."

3. **SharePoint PnP Authentication:**
   - Another browser tab opens
   - Sign in with the same admin account
   - **Consent screen for your PnP app** (first time only)
   - Click **"Accept"**
   - Window closes automatically

4. **Script Runs:**
   - Console shows progress
   - Data is extracted
   - CSV file is generated

### Subsequent Runs:

**No consent needed!** The script will:
- Use cached credentials (if still valid)
- Or prompt for sign-in only (no consent screen)

---

## 🔧 Troubleshooting Permissions

### Issue 1: "Access Denied" Error

**Cause:** Your account doesn't have sufficient role.

**Solution:**
1. Ask your Global Admin to assign you one of these roles:
   - Global Reader (recommended)
   - SharePoint Administrator
   - Global Administrator

2. Wait 5-10 minutes for role assignment to propagate

3. Try again

---

### Issue 2: "Consent Required" Error

**Cause:** You haven't consented to the required permissions.

**Solution:**
```powershell
# Force re-consent for Microsoft Graph
Connect-MgGraph -Scopes "Group.Read.All", "Sites.Read.All", "User.Read.All" -ForceRefresh

# When browser opens, click "Accept" on the consent screen
```

---

### Issue 3: "Insufficient Privileges" for SharePoint

**Cause:** Your account cannot access SharePoint Admin Center.

**Solution:**
1. Verify URL: Must be `https://TENANT-admin.sharepoint.com` (with `-admin`)
2. Check SharePoint Admin role assignment
3. Try accessing the admin center manually in browser first

---

### Issue 4: "Chat.Read.All Permission Denied"

**Cause:** This permission is optional and may require admin consent.

**Solution 1 (Skip chats):**
- The script will still work; chat count will be skipped

**Solution 2 (Enable chats):**
- Ask Global Admin to pre-consent to Chat.Read.All
- Or remove chat counting from the script

---

## 🎯 Minimal Permissions Approach

If you want to run with **minimal permissions**, modify the script:

```powershell
# Edit Get-M365ObjectCounts-Fast.ps1, line 38
# Remove Chat.Read.All if you don't need chat counts:

Connect-MgGraph -Scopes "Group.Read.All", "Team.ReadBasic.All", "Sites.Read.All", "User.Read.All" -NoWelcome

# The script will skip chat counting but everything else works
```

---

## 📊 Permission Summary Table

| What You Need | Type | How to Get It |
|---------------|------|---------------|
| **Global Reader role** | Azure AD Role | Ask Global Admin to assign |
| **SharePoint Admin role** | Azure AD Role | Ask Global Admin to assign |
| **API Permissions** | Delegated | Auto-granted when you sign in |
| **App Registration (Graph)** | Azure AD App | ❌ NOT NEEDED (uses Microsoft's first-party app) |
| **App Registration (PnP)** | Azure AD App | ✅ **NEEDED** — one-time `Register-PnPEntraIDAppForInteractiveLogin` (since Sept 9, 2024) |
| **Client Secret** | Azure AD App | ❌ **NOT NEEDED** |
| **Certificate** | Azure AD App | ❌ **NOT NEEDED** |

---

## ✅ Pre-Flight Checklist

Before running the script:

- [ ] I have **Global Reader** or **SharePoint Administrator** role
- [ ] I can access `https://TENANT-admin.sharepoint.com` in a browser
- [ ] I have installed PowerShell modules (Microsoft.Graph, PnP.PowerShell)
- [ ] I know my tenant name (e.g., "contoso" in contoso.sharepoint.com)
- [ ] I have registered a PnP Entra ID App (`Register-PnPEntraIDAppForInteractiveLogin`) and have its Client ID, or set the `ENTRAID_APP_ID` environment variable
- [ ] I'm ready to sign in and accept permission consent

If all boxes are checked, **you're ready to run!** 🚀

---

## 🔑 Alternative: App Registration (Advanced)

If you need **unattended/scheduled** runs without interactive login, you can use app registration with application permissions. See **ADVANCED-APP-PERMISSIONS.md** for details.

**Note:** For most users, interactive authentication is simpler and sufficient!

---

## 💡 Quick Reference

**Required PowerShell Modules:**
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module PnP.PowerShell -Scope CurrentUser
```

**Required Azure AD Role:**
- Global Reader (recommended) OR
- SharePoint Administrator OR
- Global Administrator

**Required API Permissions (automatic via consent):**
- Group.Read.All
- Sites.Read.All
- User.Read.All
- Team.ReadBasic.All
- Chat.Read.All (optional)

**App Registration Needed?**
- Graph: ❌ NO
- PnP/SharePoint: ✅ YES — one-time, self-service, via `Register-PnPEntraIDAppForInteractiveLogin` (required since Sept 9, 2024)

---

**Summary:** Have Global Reader or SharePoint Admin role, register a PnP Entra ID app once (`Register-PnPEntraIDAppForInteractiveLogin`), then run the script with `-ClientId <app-id>`, sign in, and accept the consent prompt. That's it! 🎉
