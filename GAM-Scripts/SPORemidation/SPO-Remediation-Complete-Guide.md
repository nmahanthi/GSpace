# SPO Remediation — Complete Process Guide

## Table of Contents
1. [Overview](#1-overview)
2. [Background — Why Remediation is Needed](#2-background)
3. [Prerequisites](#3-prerequisites)
4. [Folder Structure](#4-folder-structure)
5. [Input File — SiteMapping.csv](#5-input-file--sitemappingcsv)
6. [Authentication — How Single Sign-In Works](#6-authentication)
7. [Script Reference](#7-script-reference)
   - [Invoke-FullRemediation.ps1 — Master Orchestrator](#71-invoke-fullremediationps1--master-orchestrator)
   - [Fix-SPONavigation.ps1 — Navigation URL Fixer](#72-fix-sponavigationps1--navigation-url-fixer)
   - [Set-SPONavigationCascading.ps1 — Layout Switcher](#73-set-sponavigationcascadingps1--layout-switcher)
   - [Fix-SPOPageFormatting.ps1 — Page Text Fixer](#74-fix-spopageformattingps1--page-text-fixer)
8. [Step-by-Step Run Guide](#8-step-by-step-run-guide)
9. [Understanding the Output](#9-understanding-the-output)
10. [Reports and Logs](#10-reports-and-logs)
11. [Status Code Reference](#11-status-code-reference)
12. [Troubleshooting](#12-troubleshooting)

---

## 1. Overview

This remediation suite fixes three categories of issues found in **SharePoint Online (SPO) sites that were migrated from Google Sites**:

| # | Problem | Script That Fixes It |
|---|---------|----------------------|
| 1 | Navigation links still point to Google Sites URLs, empty `#` anchors, or ghost nodes | `Fix-SPONavigation.ps1` |
| 2 | Navigation dropdown renders side-by-side (Mega Menu) instead of vertically (like Google Sites) | `Set-SPONavigationCascading.ps1` |
| 3 | Numbered sub-lists show letters (a, b, c) instead of numbers, and bullet markers are squares instead of circles | `Fix-SPOPageFormatting.ps1` |

All three fixes are orchestrated by a **single master script** — `Invoke-FullRemediation.ps1` — which processes all sites listed in `SiteMapping.csv` in one automated run.

---

## 2. Background

### What Was Migrated
Approximately 100 internal sites were moved from **Google Sites** to **SharePoint Online** using an automated migration tool. The migration copied content and structure but introduced several defects.

### Defect 1 — Broken Navigation URLs
During migration, navigation links were not always updated correctly. Three patterns of broken URLs occur:

- **Google Sites URL still present** — e.g. `https://sites.google.com/.../pagename`
- **Empty or placeholder URL** — node URL is blank or just `#`
- **Old base URL embedded** — the Google Sites base URL is still part of the path

Additionally, some migration runs left behind **ghost nodes** — navigation entries with no Title, no URL, and no accessible ID. These cannot be removed through the standard PnP navigation cmdlets and require a direct SharePoint REST API call.

### Defect 2 — Mega Menu Layout
SharePoint Online's default navigation style is **Mega Menu**, which renders child links horizontally side-by-side across columns. Google Sites renders dropdowns **vertically** (one item below another). After migration, the SPO sites look structurally different from the originals, confusing users.

The fix sets `MegaMenuEnabled = false` on each site's web properties, which switches SPO to **Cascading** navigation — matching the original Google Sites vertical dropdown behaviour.

### Defect 3 — HTML List Formatting
The migration tool embedded inline CSS styles in page text web parts. Two specific CSS values were mapped incorrectly:

- `list-style-type: lower-alpha` — causes nested ordered list items to display as `a.`, `b.`, `c.` instead of `1.`, `2.`, `3.`
- `list-style-type: square` — causes bullet points to appear as solid squares instead of round circles

The fix applies targeted regex replacements directly to the HTML stored inside each Text Web Part on every page.

---

## 3. Prerequisites

### Software
| Requirement | Minimum Version | Install Command |
|---|---|---|
| PowerShell | 7.0+ (`pwsh`) | [Download from Microsoft](https://aka.ms/powershell) |
| PnP.PowerShell module | 1.12.0+ | `Install-Module PnP.PowerShell -Scope CurrentUser` |

> **Why PowerShell 7?** The scripts use features not available in Windows PowerShell 5.x, including improved error handling, `Set-StrictMode -Version Latest`, and modern PnP module compatibility.

### Azure AD App Registration
The scripts authenticate using an **Azure AD (Entra ID) App Registration** with delegated permissions. The app must have:

| Permission | Type | Purpose |
|---|---|---|
| `Sites.ReadWrite.All` | Delegated | Read and write all SPO site collections |
| `User.Read` | Delegated | Read signed-in user profile (required for interactive login) |

**Client ID used:** `3834b2e7-ab80-45fc-b4c8-ed5c960076b7`

The Redirect URI must be set to `http://localhost` in the app registration to support interactive browser login.

### Permissions on SharePoint
The account used to sign in must have **Site Collection Administrator** or **Full Control** on all sites being remediated. Without this, the REST API PATCH calls for navigation and web properties will fail with a 403 Forbidden error.

---

## 4. Folder Structure

```
SPORemidation\
│
├── Invoke-FullRemediation.ps1       ← Master orchestrator (run this for all 100 sites)
├── Fix-SPONavigation.ps1            ← Navigation URL fix logic
├── Set-SPONavigationCascading.ps1   ← Mega Menu → Cascading layout fix
├── Fix-SPOPageFormatting.ps1        ← Page text HTML formatting fix
├── SiteMapping.csv                  ← Input: list of all 100 sites
│
└── RemediationLogs\                 ← Created automatically on first run
    ├── ftcmigrationtestsite_YYYYMMDD_HHmmss.log   ← Per-site detailed log
    ├── hrportal_YYYYMMDD_HHmmss.log
    ├── ...
    ├── FullRemediation_YYYYMMDD_HHmmss.csv        ← Machine-readable summary
    └── FullRemediation_YYYYMMDD_HHmmss.html       ← Human-readable HTML report
```

> All four `.ps1` scripts must be in the **same folder**. The orchestrator uses `$PSScriptRoot` to locate the sibling scripts automatically.

---

## 5. Input File — SiteMapping.csv

The CSV file drives the entire bulk run. Each row represents one SharePoint site to remediate.

### Columns

| Column | Required | Description | Example |
|---|---|---|---|
| `SiteName` | Yes | Human-readable display name (used in reports only) | `HR Portal` |
| `GoogleSitesBaseUrl` | Yes | The original Google Sites base URL for this site | `https://sites.google.com/company.com/hrportal` |
| `SPOSiteUrl` | Yes | Full URL of the SharePoint Online site collection | `https://tenant.sharepoint.com/sites/hrportal` |
| `RebuildNavigation` | Yes | `TRUE` = wipe and rebuild navigation from scratch; `FALSE` = scan and fix only | `FALSE` |

### Example File

```csv
SiteName,GoogleSitesBaseUrl,SPOSiteUrl,RebuildNavigation
FTC Migration Test Site,https://sites.google.com/censftmigsme.microsoft-int.com/ftcmigrationtestsite,https://mngenvmcap908272.sharepoint.com/sites/ftcmigrationtestsite,FALSE
HR Portal,https://sites.google.com/censftmigsme.microsoft-int.com/hrportal,https://mngenvmcap908272.sharepoint.com/sites/hrportal,FALSE
IT Help Desk,https://sites.google.com/censftmigsme.microsoft-int.com/ithelpdesk,https://mngenvmcap908272.sharepoint.com/sites/ithelpdesk,FALSE
Finance,https://sites.google.com/censftmigsme.microsoft-int.com/finance,https://mngenvmcap908272.sharepoint.com/sites/finance,FALSE
```

### RebuildNavigation Flag — When to Use TRUE
Set `RebuildNavigation = TRUE` only for sites where navigation is so badly broken that it's easier to wipe all nodes and rebuild from a known-good template. This removes **all** existing navigation nodes and adds the default Google Sites mirror structure. Use `FALSE` for most sites — it is safer as it only changes broken links.

---

## 6. Authentication

### Single Sign-In Design
The orchestrator is designed so that **the browser opens exactly once**, regardless of how many sites are being processed (e.g., 100 sites = 1 browser prompt).

### How It Works — Step by Step

```
Script starts
    │
    ├─► Loads SiteMapping.csv
    │
    ├─► PRE-AUTH: Connects to the first site in the CSV with -Interactive
    │       └─► Browser window opens → you sign in with your M365 account
    │       └─► MSAL stores the access token + refresh token in memory
    │       └─► Connection immediately disconnected (token stays cached)
    │
    └─► FOR EACH SITE in CSV:
            ├─► Connect-PnPOnline called with -Interactive
            │       └─► MSAL finds the cached token → connects SILENTLY (no browser)
            ├─► Run Step 1: Nav fix     (uses the open connection)
            ├─► Run Step 2: Layout fix  (uses the open connection)
            ├─► Run Step 3: Page fix    (uses the open connection)
            └─► Disconnect-PnPOnline
```

### MSAL Token Caching
PnP PowerShell uses **Microsoft Authentication Library (MSAL)** under the hood. After the first interactive login:
- The access token is valid for **1 hour**
- The refresh token is valid for **24 hours** (or until your session policy expires)
- MSAL automatically uses the refresh token to get a new access token when it expires — no re-prompt needed

For a 100-site run that takes 2–3 hours, the token will be silently refreshed mid-run without any user action.

### Authentication Flow Diagram

```
[First Connect]                       [All Subsequent Connects]
      │                                           │
Connect-PnPOnline -Interactive        Connect-PnPOnline -Interactive
      │                                           │
      ▼                                           ▼
  Browser opens                       MSAL checks in-memory cache
      │                                           │
  You sign in                         Token found → connects silently
      │                                           │
  MSAL caches token ──────────────────────────────┘
```

---

## 7. Script Reference

### 7.1 Invoke-FullRemediation.ps1 — Master Orchestrator

This is the **only script you need to run** for a full bulk remediation. It orchestrates all three fix scripts across all sites from the CSV.

#### Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-CsvPath` | String | *(mandatory)* | Path to `SiteMapping.csv` |
| `-ClientId` | String | `3834b2e7-...` | Azure AD App Registration Client ID |
| `-TenantId` | String | *(auto-derived)* | Tenant identifier. If blank, derived from the SPO URL (e.g. `mngenvmcap908272.onmicrosoft.com`) |
| `-DryRun` | Switch | Off | Simulate all steps — reads data, logs what would change, writes nothing |

#### Execution Flow (per site)

```
FOR EACH row in SiteMapping.csv:
  1. Connect-PnPOnline  (silent — MSAL cached token)
  │
  ├─ STEP 1: Invoke-SPONavFix        (Fix-SPONavigation.ps1)
  │     Scans all top-nav nodes
  │     Detects and fixes Google URLs, empty URLs, ghost nodes
  │     Uses -UseCurrentConnection (does NOT reconnect internally)
  │
  ├─ STEP 2: Cascading nav fix       (inline in orchestrator)
  │     Reads MegaMenuEnabled via REST API GET
  │     If true → PATCH to set MegaMenuEnabled = false
  │
  ├─ STEP 3: Repair-SPOPage          (Fix-SPOPageFormatting.ps1)
  │     Lists all .aspx pages in SitePages library
  │     For each page: reads Text Web Parts, applies HTML regex fixes
  │     Saves and publishes page if any changes were made
  │
  └─ Disconnect-PnPOnline
```

#### Output Files (per run)
- **Per-site log**: `RemediationLogs/<siteSlug>_<timestamp>.log` — full verbose log for each site
- **CSV summary**: `RemediationLogs/FullRemediation_<timestamp>.csv` — one row per site, machine-readable
- **HTML report**: `RemediationLogs/FullRemediation_<timestamp>.html` — colour-coded report, opens in any browser

---

### 7.2 Fix-SPONavigation.ps1 — Navigation URL Fixer

Fixes broken or stale navigation node URLs on a single SPO site. Can be run standalone or dot-sourced by the orchestrator.

#### Dual-Mode Design
```
Standalone mode  →  called directly via pwsh -File
                    Manages its own Connect/Disconnect

Library mode     →  dot-sourced by Invoke-FullRemediation.ps1
                    Orchestrator owns the connection (-UseCurrentConnection)
```

#### The Three URL Fix Rules

| Rule | Detects When | Action |
|---|---|---|
| **GSite URL** | `$url -match "sites\.google\.com"` | Strips the Google Sites domain and path prefix, converts page slug to Title-Case, appends `.aspx`, prepends SPO SitePages base URL |
| **Empty/Placeholder** | URL is blank, whitespace, or `#` | Builds a URL from the node's Title: replaces spaces with hyphens, appends `.aspx` |
| **Old Base URL** | URL contains the original Google Sites base URL | Replaces the Google base with the SPO SitePages base path |

#### Ghost Node Handling
A **ghost node** is a navigation entry where both Title and URL are empty or missing. These are left behind by failed migration steps.

Standard PnP cmdlets (`Remove-PnPNavigationNode`) cannot target these nodes because they have no accessible ID. The script uses a **SharePoint REST API fallback**:

```
GET /_api/web/navigation/topnavigationbar?$select=Id,Title,Url
    → Find nodes where Title="" AND Url=""
    → For each: DELETE /_api/web/navigation/topnavigationbar/getById(<Id>)
```

#### Rebuild Mode
When `RebuildNavigation = TRUE` in the CSV, the script:
1. Deletes **all** existing top-nav nodes
2. Adds nodes from a hard-coded default structure that mirrors the standard Google Sites navigation layout

This is a destructive operation — only use it when normal auto-fix cannot repair the navigation.

#### Strict-Mode Safety
All property accesses on navigation nodes are guarded:
```powershell
$url   = if ($Node.PSObject.Properties['Url'])      { $Node.Url }      else { "" }
$title = if ($Node.PSObject.Properties['Title'])    { $Node.Title }    else { "" }
if ($Node.PSObject.Properties['Children'] -and $Node.Children) { ... }
```
This prevents `Set-StrictMode -Version Latest` errors when SPO returns node objects with missing properties.

#### Standalone Usage
```powershell
# Dry-run (simulate, no changes)
pwsh -ExecutionPolicy Bypass -File .\Fix-SPONavigation.ps1 `
  -SiteUrl  "https://tenant.sharepoint.com/sites/mysite" `
  -ClientId "3834b2e7-ab80-45fc-b4c8-ed5c960076b7" `
  -WhatIf

# Apply fixes
pwsh -ExecutionPolicy Bypass -File .\Fix-SPONavigation.ps1 `
  -SiteUrl  "https://tenant.sharepoint.com/sites/mysite" `
  -ClientId "3834b2e7-ab80-45fc-b4c8-ed5c960076b7"

# Full rebuild
pwsh -ExecutionPolicy Bypass -File .\Fix-SPONavigation.ps1 `
  -SiteUrl  "https://tenant.sharepoint.com/sites/mysite" `
  -ClientId "3834b2e7-ab80-45fc-b4c8-ed5c960076b7" `
  -RebuildNavigation
```

---

### 7.3 Set-SPONavigationCascading.ps1 — Layout Switcher

Switches SPO top navigation from **Mega Menu** (horizontal/side-by-side) to **Cascading** (vertical dropdowns) to match Google Sites behaviour.

#### The Problem It Solves

| Navigation Style | How Dropdowns Appear |
|---|---|
| **Mega Menu (SPO default)** | Child links spread horizontally across columns |
| **Cascading (Google Sites style)** | Child links stack vertically, one below another |

#### How It Works
The script reads and patches a single web property via the SharePoint REST API:

```
GET  /_api/web?$select=MegaMenuEnabled,Title
       → MegaMenuEnabled: true  (Mega Menu is ON)

PATCH /_api/web
       → Body: { "MegaMenuEnabled": false }

GET  /_api/web?$select=MegaMenuEnabled
       → MegaMenuEnabled: false  (Cascading is now active)
```

The script verifies the change was applied after the PATCH and reports `WARN_NOT_APPLIED` if the value did not change (usually a permissions issue).

#### Standalone Usage
```powershell
# Single site — dry-run
pwsh -ExecutionPolicy Bypass -File .\Set-SPONavigationCascading.ps1 `
  -SiteUrl  "https://tenant.sharepoint.com/sites/mysite" `
  -ClientId "3834b2e7-ab80-45fc-b4c8-ed5c960076b7" `
  -WhatIf

# Single site — apply
pwsh -ExecutionPolicy Bypass -File .\Set-SPONavigationCascading.ps1 `
  -SiteUrl  "https://tenant.sharepoint.com/sites/mysite" `
  -ClientId "3834b2e7-ab80-45fc-b4c8-ed5c960076b7"

# Bulk — all sites from CSV
pwsh -ExecutionPolicy Bypass -File .\Set-SPONavigationCascading.ps1 `
  -CsvPath  .\SiteMapping.csv `
  -ClientId "3834b2e7-ab80-45fc-b4c8-ed5c960076b7"
```

---

### 7.4 Fix-SPOPageFormatting.ps1 — Page Text Fixer

Scans every page in a site's SitePages library, reads the HTML stored in **Text Web Parts**, applies three regex fixes, then saves and publishes the page.

#### The Three HTML Fixes

**Fix 1 — Nested ordered list: letters → numbers**
```
Before:  list-style-type: lower-alpha   →   a.  b.  c.
After:   list-style-type: decimal       →   1.  2.  3.
```
Regex: `(list-style-type\s*:\s*)lower-alpha` → `${1}decimal`

**Fix 2 — Ordered list type attribute: a → 1**
```
Before:  <ol type="a">
After:   <ol type="1">
```
Regex: `<ol([^>]*)\btype="a"` → `<ol${1}type="1"`

**Fix 3 — Bullet markers: squares → circles**
```
Before:  list-style-type: square   →   ■  ■  ■
After:   list-style-type: disc     →   •  •  •
```
Regex: `(list-style-type\s*:\s*)square` → `${1}disc`

#### Strict-Mode Safety for Web Part Properties
Text Web Parts in PnP PowerShell do not always expose a `.Text` property. Accessing it directly under `Set-StrictMode` would throw an error. The script guards all access:

```powershell
# In the Where-Object filter:
$_.GetType().Name -like "*PageText*" -or $_.PSObject.Properties['Text']

# When reading the value:
$partText = if ($part.PSObject.Properties['Text']) { $part.Text } else { $null }
```

#### Save and Publish
After applying fixes, the script calls both `.Save()` and `.Publish()` on the page object with `| Out-Null` to suppress return-value output that would otherwise corrupt the result pipeline in the orchestrator.

#### Standalone Usage
```powershell
# Single page — dry-run
pwsh -ExecutionPolicy Bypass -File .\Fix-SPOPageFormatting.ps1 `
  -SiteUrl  "https://tenant.sharepoint.com/sites/mysite" `
  -PageName "Text-Formatting.aspx" `
  -ClientId "3834b2e7-ab80-45fc-b4c8-ed5c960076b7" `
  -WhatIf

# All pages in a site — apply
pwsh -ExecutionPolicy Bypass -File .\Fix-SPOPageFormatting.ps1 `
  -SiteUrl  "https://tenant.sharepoint.com/sites/mysite" `
  -ClientId "3834b2e7-ab80-45fc-b4c8-ed5c960076b7" `
  -AllPages
```

---

## 8. Step-by-Step Run Guide

### Step 1 — Prepare SiteMapping.csv
Open `SiteMapping.csv` and verify all 100 site rows are correct:
- `SPOSiteUrl` points to the SharePoint site (not the Google URL)
- `GoogleSitesBaseUrl` is the original Google Sites root (used for Rule 3 URL matching)
- `RebuildNavigation` is `FALSE` for most sites

### Step 2 — Run a Dry-Run First (Strongly Recommended)
```powershell
pwsh -ExecutionPolicy Bypass -File ".\Invoke-FullRemediation.ps1" `
  -CsvPath  .\SiteMapping.csv `
  -ClientId "3834b2e7-ab80-45fc-b4c8-ed5c960076b7" `
  -DryRun
```
- A browser window opens — sign in with your Microsoft 365 admin account
- The script scans all sites and reports what it **would** fix
- No changes are written to SharePoint
- Review the HTML report in `RemediationLogs\`

### Step 3 — Review the Dry-Run Report
Open the generated `.html` file in `RemediationLogs\`. Check:
- `Nav:FIXED(n)` — n navigation nodes would be repaired (expected)
- `Nav:ERROR` — connection or permission problem (investigate before live run)
- `Pages:n/total fixed` — n pages have formatting issues to fix (expected)
- `Casc:ALREADY_CASCADING` — layout already correct, no change needed

### Step 4 — Run the Live Fix
When the dry-run looks correct, remove `-DryRun` to apply all changes:
```powershell
pwsh -ExecutionPolicy Bypass -File ".\Invoke-FullRemediation.ps1" `
  -CsvPath  .\SiteMapping.csv `
  -ClientId "3834b2e7-ab80-45fc-b4c8-ed5c960076b7"
```
- Sign in once when the browser opens
- The script processes all sites sequentially
- All changes are written and published in SharePoint

### Step 5 — Review the Live Run Report
Open the new `.html` report from `RemediationLogs\` and verify:
- All sites show `Nav:FIXED(n)` or `Nav:FIXED(0)` (0 means no broken links were found — already clean)
- `Casc:FIXED` or `Casc:ALREADY_CASCADING` for every site
- `Pages:n/total fixed` with `0 failed`

---

## 9. Understanding the Output

### Console Output During a Run
```
[INFO]    Loaded 100 site(s). DryRun=False
[INFO]    Authenticating — a browser window will open ONCE...
[SUCCESS] Authentication successful. Token cached — no further prompts needed.

[INFO]    [hrportal] Connecting to https://...sharepoint.com/sites/hrportal ...
[SUCCESS] [hrportal] Connected.
[INFO]    [hrportal] Scanning navigation for broken/Google-Sites URLs ...
[WARN]    [hrportal] BROKEN(GSite): [Resources] -> /sites/hrportal/SitePages/Resources.aspx
[SUCCESS] [hrportal] [FIXED] [Resources] -> /sites/hrportal/SitePages/Resources.aspx
[SUCCESS] [hrportal] Done. 3 node(s) fixed.
[INFO]    [hrportal] Processing page: Home.aspx
[WARN]    [hrportal]   WebPart #2 — 2 fix(es) applied.
[SUCCESS] [hrportal] Page 'Home.aspx' saved and published.
```

### Log Colour Coding
| Colour | Level | Meaning |
|---|---|---|
| Cyan | INFO | Normal progress information |
| Yellow | WARN | Issue detected — was or would be fixed |
| Green | SUCCESS | Step completed successfully |
| Red | ERROR | Step failed — requires attention |

---

## 10. Reports and Logs

### Per-Site Log File
Located at: `RemediationLogs/<siteSlug>_<timestamp>.log`

Contains every action taken for that site, including timestamps. Use this for detailed investigation of any site that shows errors in the summary report.

### CSV Summary
Located at: `RemediationLogs/FullRemediation_<timestamp>.csv`

One row per site with columns: `SiteName`, `SiteSlug`, `NavFix`, `CascadingFix`, `PagesFixed`, `PagesScanned`, `PagesFailed`, `Error`, `LogFile`. Import into Excel for filtering and analysis.

### HTML Report
Located at: `RemediationLogs/FullRemediation_<timestamp>.html`

Colour-coded table showing all sites, their fix status for each step, and links to per-site log files. Open in any browser. Includes a summary badge showing total sites fixed vs total with issues.

---

## 11. Status Code Reference

### Navigation Fix Status (NavFix column)
| Status | Meaning |
|---|---|
| `FIXED(n)` | n nodes were fixed (URLs updated or ghost nodes removed) |
| `FIXED(0)` | Scan completed, no broken nodes found |
| `REBUILT` | Full rebuild mode — all nodes replaced from template |
| `FAILED` | Could not connect to the site |
| `ERROR` | Connection succeeded but fix threw an exception |
| `CONNECT_FAILED` | Authentication or network error |

### Cascading Nav Status (CascadingFix column)
| Status | Meaning |
|---|---|
| `FIXED` | MegaMenuEnabled was true — successfully set to false (cascading) |
| `ALREADY_CASCADING` | MegaMenuEnabled was already false — no change needed |
| `DRYRUN` | Would have set cascading — DryRun mode, no change written |
| `WARN_NOT_APPLIED` | PATCH was sent but MegaMenuEnabled is still true — check permissions |
| `ERROR` | REST API call failed |

### Page Formatting Status (per page)
| Status | Meaning |
|---|---|
| `FIXED(n)` | n HTML fixes applied, page saved and published |
| `DRYRUN(n)` | n fixes detected — DryRun mode, no change written |
| `OK` | Page scanned, no formatting issues found |
| `NO_TEXT_PARTS` | Page has no text web parts — nothing to fix |
| `PAGE_NOT_FOUND` | Page listed in SitePages but Get-PnPPage failed |

---

## 12. Troubleshooting

### Error: `CONNECT_FAILED` for all sites
**Cause:** The Azure AD App Registration is not configured correctly, or the account signing in doesn't have access.
**Fix:** Verify the app has `Sites.ReadWrite.All` delegated permission, admin consent is granted, and the redirect URI includes `http://localhost`.

### Error: `Nav:ERROR` — `The property 'Children' cannot be found`
**Cause:** SPO returned a navigation node object without a `Children` property (e.g., a separator or external link type).
**Fix:** Already handled in the current scripts — all node properties are guarded with `$Node.PSObject.Properties[...]` checks.

### Error: `Cascading fix failed: 404 FILE NOT FOUND`
**Cause:** The site URL in `SiteMapping.csv` does not exist in the tenant.
**Fix:** Verify the `SPOSiteUrl` column contains the correct, existing SharePoint site URL. Check by opening the URL in a browser.

### Error: `Page formatting failed: The property 'Text' cannot be found`
**Cause:** A web part on the page does not expose a `.Text` property (e.g., it is an image or video web part, not a text web part).
**Fix:** Already handled — the script uses `$_.PSObject.Properties['Text']` to safely skip non-text web parts.

### Warning: `WARN_NOT_APPLIED` for cascading fix
**Cause:** The PATCH to `MegaMenuEnabled` was accepted by the API but the value did not change. This typically means the account lacks Site Collection Administrator rights.
**Fix:** Ensure the signed-in account is added as a Site Collection Admin on all sites, or grant the App Registration `FullControl` application permissions.

### Warning: `No text web parts found on 'home1.aspx'`
**Cause:** The page exists but contains no text web parts — only image, news, quick links, or other non-text components.
**Status:** This is informational only (`NO_TEXT_PARTS`). No action required.

### Browser Opens Multiple Times
**Cause:** The MSAL token expired mid-run (after ~1 hour), or `Disconnect-PnPOnline` cleared the cache between sites.
**Fix:** For very long runs (100+ sites, 3+ hours), the token will be silently refreshed via the MSAL refresh token. If you see repeated browser prompts, check if your tenant's Conditional Access policies have short session lifetimes.

### Script Stops Midway With No Error
**Cause:** Usually a throttling response (HTTP 429) from SharePoint Online.
**Fix:** SPO applies throttling when too many requests arrive too quickly. The sequential processing design already prevents most throttling. If it still occurs, add `Start-Sleep -Seconds 5` between site iterations or contact your SharePoint admin to review tenant throttling limits.
