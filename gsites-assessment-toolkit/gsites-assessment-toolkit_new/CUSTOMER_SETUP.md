# GSites Assessment Toolkit — Customer Setup Guide

This document describes every change a customer must make before running the toolkit in their environment.

---

## Prerequisites

| Tool | Version | Purpose |
|---|---|---|
| [GAM](https://github.com/GAM-team/GAM) | 7.x | Google Workspace Admin API export |
| [Node.js](https://nodejs.org/) | 18 or later | Site crawling and API extraction scripts |
| [PowerShell](https://github.com/PowerShell/PowerShell) | 7.x (`pwsh`) | Orchestrator and scoring scripts |
| [gcloud CLI](https://cloud.google.com/sdk/docs/install) | Any | OAuth token for the Sites API (published URLs) |

---

## Step 1 — Configure GAM path

Create (or edit) the file `gam.cfg` in the toolkit root folder.  
This file tells the toolkit where `gam.exe` is installed on this machine.

**`gam.cfg`**
```
GAM_PATH=C:\tools\gam\gam.exe
```

Replace the path with the actual location of `gam.exe` on the customer machine.

> **Alternative:** Set the `GAM_PATH` environment variable instead of using `gam.cfg`.
> ```powershell
> $env:GAM_PATH = "C:\tools\gam\gam.exe"
> ```

> **Alternative:** Add the folder containing `gam.exe` to the system `PATH` — the toolkit will find it automatically.

---

## Step 2 — Install Node.js dependencies

Run once from the toolkit root folder:

```powershell
npm install
npx playwright install chromium
```

---

## Step 3 — Authenticate with Google (browser session)

This step saves a browser login session used by the Playwright crawler.  
Run once per machine, or whenever the session expires.

```powershell
node 02_save_playwright_auth.js
```

A browser window opens. Sign in with the **Google Workspace admin account** that has read access to all Google Sites in the domain. After sign-in completes and a site loads successfully, press **Enter** in the terminal.

This saves credentials to `.auth\state.json`.  
> ⚠️ **Do not commit `.auth\state.json` to source control** — it contains personal session cookies.

---

## Step 4 — Authenticate with gcloud (API token)

Required for Step 4A (published URLs).

```powershell
gcloud auth login
```

The toolkit retrieves the token automatically from gcloud on each run.  
To pass the token explicitly:

```powershell
$token = (gcloud auth print-access-token).Trim()
.\Run-FullAssessment.ps1 -PrimaryDomain "yourcompany.com" -AccessToken $token
```

> ⚠️ **403 Forbidden on Step 4A?** The Sites API v1 requires the calling account to
> have explicit Drive-level Viewer access to each site — being a Workspace super
> admin is not enough. Pass `-SitesAdminEmail` (the email you used for
> `gcloud auth login`) and the toolkit will use GAM's elevated access to grant
> that account Reader access to every site before Step 4A runs:
> ```powershell
> .\Run-FullAssessment.ps1 -PrimaryDomain "yourcompany.com" -SitesAdminEmail "admin@yourcompany.com"
> ```
> Skip this pre-grant step with `-SkipGrantAccess`.

---

## Step 5 — Run the assessment

Always pass `-PrimaryDomain` set to the customer's Google Workspace domain.  
This is used to distinguish internal users from external users in permission analysis.

**Full run (first time):**
```powershell
.\Run-FullAssessment.ps1 -PrimaryDomain "yourcompany.com"
```

**Subsequent runs (GAM export already done):**
```powershell
.\Run-FullAssessment.ps1 -PrimaryDomain "yourcompany.com" -SkipGAMExport -SkipBrowserAuth
```

**Batched run — process 10 sites at a time:**
```powershell
# Batch 1: sites 1–10
.\Run-FullAssessment.ps1 -PrimaryDomain "yourcompany.com" -SkipGAMExport -SkipBrowserAuth -MaxSites 10 -SiteOffset 0

# Batch 2: sites 11–20
.\Run-FullAssessment.ps1 -PrimaryDomain "yourcompany.com" -SkipGAMExport -SkipBrowserAuth -MaxSites 10 -SiteOffset 10
```

**Fast API-based extraction (no browser required):**
```powershell
.\Run-FullAssessment.ps1 -PrimaryDomain "yourcompany.com" -SkipGAMExport -SkipBrowserAuth -UseApiExtract
```

**Run for a selected subset of sites (useful for large tenants):**

The CSV can contain either **site names** or **Google Sites URLs**. If you provide URLs, the site name is automatically extracted from the URL path.

Example CSV by name:
```csv
SiteName
My First Site
Another Site
```

Example CSV by URL:
```csv
SiteUrl
https://sites.google.com/yourcompany.com/my-first-site
https://sites.google.com/yourcompany.com/another-site
```

```powershell
# FRESH RUN (first time on this machine / folder):
# The script will run GAM export to build the inventory, then filter it.
# Do NOT use -SkipGAMExport on a fresh run.
.\Run-FullAssessment.ps1 -PrimaryDomain "yourcompany.com" -SelectedSitesCsv "selected_sites.csv"
```

```powershell
# REPEAT RUN (inventory already exists in the output/ folder):
# Skip GAM export and browser auth since they were already done.
.\Run-FullAssessment.ps1 -PrimaryDomain "yourcompany.com" -SkipGAMExport -SkipBrowserAuth -SelectedSitesCsv "selected_sites.csv"
```

```powershell
# REPEAT RUN with inventory stored elsewhere:
# Point to the existing inventory file with -InventoryCsv.
.\Run-FullAssessment.ps1 -PrimaryDomain "yourcompany.com" -SkipGAMExport -SkipBrowserAuth -SelectedSitesCsv "selected_sites.csv" -InventoryCsv "C:\path\to\GSites_Inventory_Detailed.csv"
```

**Run for a specific list of target users (find all sites owned by these users):**

If you don't know the exact site names but want to find all Google Sites owned by a specific set of users, provide a CSV with their email addresses. This completely skips the full tenant scan.

Example CSV (`target_users.csv`):
```csv
Email
user1@yourcompany.com
user2@yourcompany.com
```

```powershell
# Run the assessment only for these specific users
.\Run-FullAssessment.ps1 -PrimaryDomain "yourcompany.com" -TargetUsersCsv "target_users.csv"
```

---

## Customer Changes Summary

| File | What to Change | Required? |
|---|---|---|
| `gam.cfg` | Set `GAM_PATH` to the local path of `gam.exe` | ✅ Yes |
| `.auth\state.json` | Regenerate by running `node 02_save_playwright_auth.js` | ✅ Yes (per user) |
| `Run-FullAssessment.ps1` | Pass `-PrimaryDomain "yourcompany.com"` at runtime | ✅ Yes (parameter) |
| `05_score_sites.ps1` | Pass `-PrimaryDomain "yourcompany.com"` if run directly | ✅ Yes (parameter) |

**No other files need to be edited.** All paths, tokens, output directories, and batch sizes are resolved dynamically.

---

## Output Files

All output is written to the `output\` folder:

| File | Contents |
|---|---|
| `GSites_Inventory_Detailed.csv` | Full site inventory from GAM |
| `GSites_Permissions.csv` | Site-level permission rows |
| `Pages.csv` | Pages crawled per site |
| `Embeds.csv` | Embedded content found on each page |
| `ExternalDomains.csv` | External domains referenced |
| `Complexity_Report.csv` | Final complexity score per site |
