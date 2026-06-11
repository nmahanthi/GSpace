# SPO Navigation Remediation Guide
### Google Sites → SharePoint Online Migration

**Tenant:** `mngenvmcap908272.sharepoint.com`  
**Client ID:** `3834b2e7-ab80-45fc-b4c8-ed5c960076b7`  
**Last Updated:** May 2026

---

## Overview

After migrating from Google Sites to SharePoint Online (SPO), two navigation issues were identified:

| # | Issue | Script |
|---|---|---|
| 1 | Navigation links broken / still pointing to Google Sites URLs | `Fix-SPONavigation.ps1` |
| 2 | Dropdown menus rendered side-by-side (horizontal) instead of vertical | `Set-SPONavigationCascading.ps1` |

---

## Prerequisites

### 1. PowerShell 7+
```powershell
# Check version — must be 7 or higher
pwsh --version

# Download from: https://aka.ms/powershell
```

### 2. PnP PowerShell Module
```powershell
Install-Module PnP.PowerShell -Scope CurrentUser -Force
```

### 3. Azure AD App Registration
- **App ID (Client ID):** `3834b2e7-ab80-45fc-b4c8-ed5c960076b7`
- **Redirect URI:** `http://localhost` (under Authentication in Azure Portal)
- **Required API Permissions:**

| API | Permission | Type |
|---|---|---|
| SharePoint | `Sites.FullControl.All` | Delegated |
| SharePoint | `AllSites.FullControl` | Delegated |

### 4. SiteMapping.csv
Populate with all sites before running bulk scripts:
```
SiteName,GoogleSitesBaseUrl,SPOSiteUrl,RebuildNavigation
FTC Migration Test Site,https://sites.google.com/.../ftcmigrationtestsite,https://mngenvmcap908272.sharepoint.com/sites/ftcmigrationtestsite,FALSE
HR Portal,https://sites.google.com/.../hrportal,https://mngenvmcap908272.sharepoint.com/sites/hrportal,FALSE
```

---

## Script 1 — Fix-SPONavigation.ps1

**Purpose:** Scans every navigation node on a site and fixes broken URLs.

### What It Detects & Fixes

| Rule | Condition | Action |
|---|---|---|
| Rule 1 | URL still points to `sites.google.com` | Converts to SPO `/SitePages/` URL |
| Rule 2 | URL is empty or `#` placeholder | Generates URL from node title |
| Rule 3 | URL contains old Google Sites base domain | Strips old base, maps to SPO |
| Ghost Node | Node has no title, no URL, no ID | Removes via REST API |

### Modes

**Mode A — Auto-Fix (default):** Scans existing nodes, repairs broken URLs only.  
**Mode B — Full Rebuild:** Wipes all top-nav nodes and rebuilds from the defined structure.

### Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `-SiteUrl` | Yes (single) | — | Full SPO site URL |
| `-GoogleSitesBaseUrl` | No | `""` | Old Google Sites base URL for Rule 3 matching |
| `-ClientId` | No | App default | Azure AD App Client ID |
| `-TenantId` | No | Auto-derived | Tenant name, e.g. `contoso.onmicrosoft.com` |
| `-RebuildNavigation` | No | Off | Switch: wipe and rebuild nav from scratch |
| `-WhatIf` | No | Off | Switch: simulate only, no changes written |

### Usage — Single Site

```powershell
# Step 1: Dry-run (safe — nothing is written)
pwsh -ExecutionPolicy Bypass -File .\Fix-SPONavigation.ps1 `
  -SiteUrl "https://mngenvmcap908272.sharepoint.com/sites/ftcmigrationtestsite" `
  -ClientId "3834b2e7-ab80-45fc-b4c8-ed5c960076b7" `
  -WhatIf

# Step 2: Apply fix
pwsh -ExecutionPolicy Bypass -File .\Fix-SPONavigation.ps1 `
  -SiteUrl "https://mngenvmcap908272.sharepoint.com/sites/ftcmigrationtestsite" `
  -ClientId "3834b2e7-ab80-45fc-b4c8-ed5c960076b7"

# Step 3: Full rebuild (use when structure is completely wrong)
pwsh -ExecutionPolicy Bypass -File .\Fix-SPONavigation.ps1 `
  -SiteUrl "https://mngenvmcap908272.sharepoint.com/sites/ftcmigrationtestsite" `
  -ClientId "3834b2e7-ab80-45fc-b4c8-ed5c960076b7" `
  -RebuildNavigation
```

---

## Script 2 — Invoke-BulkNavFix.ps1

**Purpose:** Runs `Fix-SPONavigation.ps1` across all 100 sites in parallel batches.  
**Requires:** `Fix-SPONavigation.ps1` must be in the same folder.

### Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `-CsvPath` | Yes | — | Path to `SiteMapping.csv` |
| `-ClientId` | No | App default | Azure AD App Client ID |
| `-TenantId` | No | Auto-derived | Tenant override |
| `-ThrottleLimit` | No | `5` | Number of sites processed simultaneously |
| `-RebuildNavigation` | No | Off | Apply rebuild mode to every site |
| `-WhatIf` | No | Off | Simulate only |
| `-LogFolder` | No | `.\NavFixLogs` | Output folder for logs |

### Usage — Bulk (100 Sites)

```powershell
# Dry-run across all sites
pwsh -ExecutionPolicy Bypass -File .\Invoke-BulkNavFix.ps1 `
  -CsvPath .\SiteMapping.csv `
  -ClientId "3834b2e7-ab80-45fc-b4c8-ed5c960076b7" `
  -WhatIf

# Apply fixes, 5 sites at a time
pwsh -ExecutionPolicy Bypass -File .\Invoke-BulkNavFix.ps1 `
  -CsvPath .\SiteMapping.csv `
  -ClientId "3834b2e7-ab80-45fc-b4c8-ed5c960076b7" `
  -ThrottleLimit 5

# Full rebuild across all sites, 3 at a time
pwsh -ExecutionPolicy Bypass -File .\Invoke-BulkNavFix.ps1 `
  -CsvPath .\SiteMapping.csv `
  -ClientId "3834b2e7-ab80-45fc-b4c8-ed5c960076b7" `
  -RebuildNavigation -ThrottleLimit 3
```

### Output Files

| File | Description |
|---|---|
| `NavFixLogs\<site>_<timestamp>.log` | Per-site detailed log |
| `NavFixLogs\BulkNavFix_Summary_<timestamp>.csv` | Summary of all sites with status |

### Status Values

| Status | Meaning |
|---|---|
| `FIXED(n)` | Fixed `n` broken nodes |
| `REBUILT` | Nav fully wiped and rebuilt |
| `FAILED` | Could not connect |
| `ERROR` | Unexpected error during processing |

---

## Script 3 — Set-SPONavigationCascading.ps1

**Purpose:** Switches SPO navigation dropdown from **Mega Menu (horizontal/side-by-side)**  
to **Cascading (vertical)** — matching the Google Sites dropdown behaviour.

### Background

| Navigation Style | Behaviour | Setting |
|---|---|---|
| Mega Menu (SPO default) | Child items appear side-by-side in columns | `MegaMenuEnabled = true` |
| Cascading (Google Sites style) | Child items stack vertically in a dropdown | `MegaMenuEnabled = false` |

This script sets `MegaMenuEnabled = false` on each site via the SharePoint REST API.

### Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `-SiteUrl` | Yes (single mode) | — | Full SPO site URL |
| `-CsvPath` | Yes (bulk mode) | — | Path to `SiteMapping.csv` |
| `-ClientId` | No | App default | Azure AD App Client ID |
| `-TenantId` | No | Auto-derived | Tenant override |
| `-ThrottleLimit` | No | `5` | Parallel sites (bulk mode) |
| `-WhatIf` | No | Off | Simulate only |

### Usage — Single Site

```powershell
# Dry-run
pwsh -ExecutionPolicy Bypass -File .\Set-SPONavigationCascading.ps1 `
  -SiteUrl "https://mngenvmcap908272.sharepoint.com/sites/ftcmigrationtestsite" `
  -ClientId "3834b2e7-ab80-45fc-b4c8-ed5c960076b7" `
  -WhatIf

# Apply
pwsh -ExecutionPolicy Bypass -File .\Set-SPONavigationCascading.ps1 `
  -SiteUrl "https://mngenvmcap908272.sharepoint.com/sites/ftcmigrationtestsite" `
  -ClientId "3834b2e7-ab80-45fc-b4c8-ed5c960076b7"
```

### Usage — Bulk (100 Sites)

```powershell
# Apply to all sites, 5 at a time
pwsh -ExecutionPolicy Bypass -File .\Set-SPONavigationCascading.ps1 `
  -CsvPath .\SiteMapping.csv `
  -ClientId "3834b2e7-ab80-45fc-b4c8-ed5c960076b7" `
  -ThrottleLimit 5
```

### Status Values

| Status | Meaning |
|---|---|
| `FIXED` | Mega menu disabled, cascading applied |
| `ALREADY_CASCADING` | Site was already cascading, no change made |
| `DRYRUN` | Simulation only, no changes written |
| `WARN_NOT_APPLIED` | Patch sent but setting did not change (check permissions) |
| `FAILED` | Could not connect |
| `ERROR` | Unexpected error |

---

## Recommended Run Order

Run these steps in sequence for each migration batch:

```
1. Fix-SPONavigation.ps1   (WhatIf)   -> Review broken links
2. Fix-SPONavigation.ps1               -> Apply URL fixes
3. Set-SPONavigationCascading.ps1      -> Switch to vertical dropdowns
4. Verify in browser
```

Or for all 100 sites at once:

```
1. Invoke-BulkNavFix.ps1      -WhatIf  -> Dry-run across all sites
2. Invoke-BulkNavFix.ps1               -> Apply URL fixes to all sites
3. Set-SPONavigationCascading.ps1      -> Fix dropdown layout for all sites
```

---

## Authentication Flow

The scripts use **interactive browser authentication** via the registered Azure AD App:

1. `Connect-PnPOnline` is called with `-Interactive` and `-ClientId`
2. The tenant is **auto-derived** from the SPO URL:  
   `mngenvmcap908272.sharepoint.com` → `mngenvmcap908272.onmicrosoft.com`
3. A browser pop-up appears — sign in with your Microsoft 365 account
4. The OAuth token is reused for all PnP and REST API calls within that session
5. `Disconnect-PnPOnline` is called after each site to clean up the session

> **Note:** Google Sites authentication is NOT required. The Google Sites URL is used  
> only as a string pattern to detect and replace broken navigation links.

---

## Troubleshooting

| Error | Cause | Fix |
|---|---|---|
| `AADSTS700016: Application not found in directory` | Wrong tenant used | Ensure `-TenantId` matches your SPO tenant or omit it for auto-derive |
| `Set-PnPNavigationNode is not recognized` | Does not exist in PnP.PowerShell 1.x | Script now uses REST PATCH — no action needed |
| Unicode parse errors when running with `powershell.exe` | PowerShell 5.1 used | Always run with `pwsh` (PowerShell 7+) |
| `Cannot bind argument to parameter 'Identity' because it is null` | Ghost nav node has no ID | Script handles this via REST API fallback |
| `WARN_NOT_APPLIED` on cascading fix | Insufficient permissions | Ensure App has `Sites.FullControl.All` permission |

---

## File Reference

```
SPORemidation\
  Fix-SPONavigation.ps1          # Core URL-fix engine (single & bulk)
  Invoke-BulkNavFix.ps1          # Bulk runner for Fix-SPONavigation.ps1
  Set-SPONavigationCascading.ps1 # Switches mega menu to cascading (vertical)
  SiteMapping.csv                # Input: all 100 site mappings
  NavFixLogs\                    # Auto-created: per-site logs + summary CSVs
```
