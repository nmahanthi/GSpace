# GSites Assessment Toolkit — Technical Documentation

This document explains what each script does, how the pipeline fits together,
and the design decisions behind the current behavior. For day-to-day usage
instructions, see `CUSTOMER_SETUP.md`.

---

## 1. Purpose

The toolkit inventories, crawls, and scores every **Google Sites** (classic
"new" Google Sites, `mimeType='application/vnd.google-apps.site'`) file
owned in a Google Workspace domain's **My Drive** storage, so a migration
team can estimate how hard each site will be to rebuild/migrate off of
Google Sites. It produces a per-site "complexity score" based on page
count, embedded content (Sheets, Forms, Apps Script, YouTube, etc.),
external domain references, and sharing/permission risk.

**Scope note:** Shared Drive-hosted sites are intentionally **excluded**
from this main pipeline (see §5.2). Only sites stored in an individual
user's My Drive are assessed here. Shared Drive-hosted sites are assessed
by a **separate, standalone pipeline** — see §3.7–§3.10.

---

## 2. Pipeline Overview

There are **two independent pipelines** in this toolkit, sharing the same
scoring/manifest logic but with different inventory sources:

### 2.1 Main pipeline — My Drive sites (`Run-FullAssessment.ps1`)

Single entry point. Runs 5 steps in order:

| Step | Script | Purpose | Can skip? |
|---|---|---|---|
| 1 | `01_run_gam_exports.cmd` | Export site inventory + permissions via GAM | `-SkipGAMExport` |
| 2 | (inline in orchestrator) | Install Node.js deps (Playwright, csv-parse, csv-stringify) | `-SkipDependencyCheck` |
| 3 | `02_save_playwright_auth.js` | Save a logged-in browser session for crawling | `-SkipBrowserAuth` |
| 4 | `03_crawl_sites.js` **or** `03b_api_extract_embeds.js` | Crawl each site, extract pages/embeds/external links | `-SkipCrawl` |
| 5 | `05_score_sites.ps1` | Compute complexity score per site, write final report | always runs |
| 6 | `06_generate_manifest.ps1` | Consolidate all reports into one per-site manifest CSV | not run by the orchestrator; run manually |

### 2.2 Standalone pipeline — Shared Drive sites (`Run-SharedDriveAssessment.ps1`)

Separate entry point, opt-in, run only when a customer needs Shared
Drive-hosted sites assessed too. Runs 3 steps:

| Step | Script | Purpose | Can skip? |
|---|---|---|---|
| 1 | `01d_run_gam_exports_shareddrive.cmd` | Export site inventory + permissions for a specific list of Shared Drives | `-SkipExport` |
| 2 | `01e_list_shareddrive_metadata.cmd` | Export drive-level sharing settings + organizers | `-SkipMetadata` |
| 3 | `05b_score_shareddrive_sites.ps1` | Dedup + compute security-only complexity score | always runs |
| — | `06_generate_manifest.ps1` | Consolidate reports into one manifest (point `-*File` params at `GSites_SharedDrive_*.csv`) | not run automatically; run manually |

This pipeline has no crawling step — it does not visit site content, only
Drive/permission metadata — so its complexity score reflects sharing risk
only (see §3.10).

Each script in both pipelines writes into a shared `output/` folder so
later steps can pick up earlier CSVs by fixed filename.

---

## 3. Step-by-Step: What Each Script Does

### 3.1 `01_run_gam_exports.cmd` — GAM Exports (Step 1)

Uses [GAM](https://github.com/GAM-team/GAM) (Google Workspace Admin CLI)
to export three CSVs to `output/`:

1. **`GSites_Inventory_Min.csv`** — minimal sanity-check export
   (`id, name, mimetype`) — quick way to confirm GAM can find sites at all.
2. **`GSites_Inventory_Detailed.csv`** — full metadata per site: id, name,
   webViewLink (edit URL), created/modified time, owners, size, sharing
   flags, capabilities, etc.
3. **`GSites_Permissions.csv`** — one row per (site, grantee) permission —
   used later for the security/sharing-risk portion of the score.

It resolves the GAM executable path via `GAM_PATH` env var → `gam.cfg` →
system `PATH` (see `GAM_PATH_FIX.md` for why — avoids hardcoding a
personal Windows path in source control).

**Scan target** is controlled by env vars set by the orchestrator:
- `GAM_TARGET_FILE` set → scans only the specific user emails in that CSV
  (`-TargetUsersCsv` / users discovered from `-SelectedSitesCsv`).
- Otherwise → `all users` (every user in the Workspace domain).

**Sites query filter**: always restricted to
`mimeType='application/vnd.google-apps.site' and trashed=false`, optionally
AND'd with a `name=` filter list built from `-SelectedSitesCsv`
(`GAM_SITES_FILTER`) to shrink scope for large tenants.

`GAM_NUM_THREADS` (from `-GamThreads`, default 10) controls how many
parallel GAM worker processes run the export — higher is faster but uses
more memory/risks `BrokenPipeError` on Windows with large user counts.

At the end, a PowerShell one-liner strips numeric array-index suffixes
(e.g. `owners.0.emailAddress` → `owners.emailAddress`) from CSV headers so
downstream field lookups are consistent regardless of how many array
entries GAM emitted.

### 3.2 Step 2 — Node.js Dependency Check (inline in `Run-FullAssessment.ps1`)

Verifies `node`/`npm` are installed, initializes `package.json` if missing,
installs `playwright`, `csv-parse`, `csv-stringify`, and installs the
Playwright Chromium browser binary. Skippable once already done via
`-SkipDependencyCheck`.

### 3.3 `02_save_playwright_auth.js` — Browser Authentication (Step 3)

Opens a real (non-headless) Chromium window to `https://sites.google.com/`
so a human can sign in interactively (handles MFA/SSO). Once the user
presses Enter in the terminal, it saves the authenticated cookies/local
storage to `.auth/state.json` via Playwright's `context.storageState()`.
This saved session is reused by `03_crawl_sites.js` so the crawler doesn't
need to log in per-site. Skippable via `-SkipBrowserAuth` if
`.auth/state.json` already exists.

### 3.4 Step 4 — Site Crawling (two interchangeable modes)

Both modes read `output/GSites_Inventory_Detailed.csv` and use each site's
**edit URL** (`webViewLink`, GAM field `webviewlink`) as the crawl target
— the Sites API v1 "published URL" concept has been fully removed from the
toolkit (no Sites API scope/DWD/service-account setup is required to reach
this URL). Both modes write the same three output files:
`Pages.csv`, `Embeds.csv`, `ExternalDomains.csv`.

**`03_crawl_sites.js`** (default — Playwright browser crawler):
- Launches headless Chromium using the saved `.auth/state.json` session.
- For each site, does a breadth-first crawl starting at the edit URL, up to
  `MAX_PAGES_PER_SITE` (default 200) pages, following only same-site
  internal links (`sameSiteRoot`/`sameHost`).
- On each page, scans the DOM for `<a>`, `<iframe>`, `<img>`,
  `<embed>/<object>/<source>` elements and classifies each target URL via
  `classifyUrl()` into `Sheet`, `Form`, `AppsScriptWebApp`, `YouTube`,
  `Maps`, `DriveFile`, `GoogleDoc`, `GoogleSlides`, or `Other`.
- Records one row per page in `Pages.csv` (with crawl status/error),
  one row per discovered embed/link-of-interest in `Embeds.csv`, and one
  row per external domain reference in `ExternalDomains.csv`.
- Supports batching via `MAX_SITES` / `SITE_OFFSET` env vars (set from
  `-MaxSites` / `-SiteOffset`) so large tenants can be processed in chunks
  across multiple runs.

**`03b_api_extract_embeds.js`** (optional fast path, `-UseApiExtract`):
- Uses the Google **Sites API v1** (`sites.googleapis.com/v1/sites/{id}/pages`)
  directly instead of a browser — no Playwright/auth session needed, just
  an OAuth token (`-AccessToken` / `GCP_ACCESS_TOKEN`, scope
  `sites.readonly`).
- Recursively walks each page's `pageElements` tree (`walkElement`) to find
  embedded Drive items, images, and hyperlinks — equivalent structural
  output to the DOM-scraping approach above, just sourced from the API's
  JSON page model instead of rendered HTML.
- Processes sites concurrently (`CONCURRENCY`, default 10) and appends CSV
  rows incrementally to disk per site (no large in-memory buffers), so it
  scales to very large tenants without memory blowup.
- Much faster than the browser crawler (minutes vs. hours) since there's no
  page rendering, navigation, or link-following involved — page count and
  structure come directly from the API response.

### 3.5 `05_score_sites.ps1` — Complexity Scoring (Step 5)

Joins `GSites_Inventory_Detailed.csv` with `GSites_Permissions.csv`,
`Pages.csv`, `Embeds.csv`, and `ExternalDomains.csv` on site `id`, then for
each site computes:

- **Structure points** (0–40): page count (capped 20) + crawl depth
  (capped 10) + error-page penalty (capped 10).
- **Embed points** (0–40): embed count (capped 20) + pages-with-embeds
  (capped 10) + distinct external domains (capped 10).
- **Security points** (0–20): public ("anyone") sharing (capped 10) +
  external/domain-shared principals (capped 10), using `-PrimaryDomain` to
  tell internal vs. external grantees apart.

`TotalScore` (0–100) maps to a `Rating` (`Low` ≤25, `Medium` ≤50, `High`
≤75, `Very High` >75) and a migration `Recommendation`. Result is written
to `output/Complexity_Report.csv`, one row per site.

### 3.6 `06_generate_manifest.ps1` — Consolidated Manifest (Step 6, manual)

Not part of `Run-FullAssessment.ps1` — run it yourself after Step 5.
Joins `GSites_Inventory_Detailed.csv`, `GSites_Permissions.csv`,
`Pages.csv`, `Embeds.csv`, `ExternalDomains.csv`, and
`Complexity_Report.csv` on site `id` into a single row per site, so a
reviewer has one file instead of six to cross-reference. Columns include
owner/created/modified/size, page and crawl-error counts, embed
count/types, external domain count/list, permission row count, a
semicolon-joined `Grantees` list (`type:identity:role:internal|external`,
using `-PrimaryDomain` to classify), external grantee count, and the
score/rating/recommendation from the complexity report.

Writes `output/GSites_Manifest.csv`. All input/output file names are
parameters, so the same script also works against the standalone Shared
Drive export set (see `05b_score_shareddrive_sites.ps1`) by pointing the
`-*File` parameters at the `GSites_SharedDrive_*.csv` files:

```powershell
.\06_generate_manifest.ps1 -PrimaryDomain "rocheua.com" `
    -InventoryFile 'GSites_SharedDrive_Inventory.csv' `
    -PermissionsFile 'GSites_SharedDrive_Permissions.csv' `
    -PagesFile 'GSites_SharedDrive_Pages.csv' `
    -EmbedsFile 'GSites_SharedDrive_Embeds.csv' `
    -ExternalDomainsFile 'GSites_SharedDrive_ExternalDomains.csv' `
    -ComplexityReportFile 'GSites_SharedDrive_Complexity_Report.csv' `
    -ManifestFile 'GSites_SharedDrive_Manifest.csv'
```

### 3.7 `01d_run_gam_exports_shareddrive.cmd` — Shared Drive GSites Export (Standalone Step 1)

```
01d_run_gam_exports_shareddrive.cmd <SharedDriveIDs.csv> [DefaultScanningUserEmail]
```

Exports the same two categories of data as the main pipeline's Step 1, but
scoped to a specific, known list of Shared Drives instead of `all users`.

**Input CSV** (`driveId,account`):
```csv
driveId,account
0AbCDeFGhIJKLmnUK9PVA,owner1@rocheua.com
0XyzTeamDriveIdHere123,owner2@rocheua.com
```
`driveId` is the Shared Drive's ID; `account` is the Workspace user GAM
impersonates to scan **that specific drive** (its owner, an organizer, a
member, or a Super Admin). This lets different drives be scanned by
whoever actually has access to them — no single admin needs universal
access. Rows with a blank `account` fall back to `[DefaultScanningUserEmail]`
(or the `GAM_ADMIN_USER` env var); if neither is set, the script fails fast
listing the offending `driveId`s.

Outputs:
- **`GSites_SharedDrive_Inventory.csv`** — uses `select teamdriveid "~driveId"`
  to scope the listing to exactly one Shared Drive per row (not
  `corpora alldrives`, which expands every drive a user can see — the
  behavior that caused per-member duplication in an earlier design, see
  `GAM_DOCUMENTATION.md` §6.1). Includes `driveid`/`drivename` and
  `hasaugmentedpermissions`.
- **`GSites_SharedDrive_Permissions.csv`** — adds `permissiondetails` and
  `showshareddrivepermissions`, which resolve whether each grantee's access
  is **direct** (granted on the file) or **inherited** (granted at the
  drive/folder level) — a distinction only meaningful for Shared Drive
  content.

`GAM_NUM_THREADS` defaults to **5** here (vs. 10 in the main pipeline),
since each worker expands a whole Shared Drive's file tree.

### 3.8 `01e_list_shareddrive_metadata.cmd` — Shared Drive Metadata (Standalone Step 2)

Same input CSV/account-resolution logic as `01d`. Reports on the **drives
themselves** rather than the sites on them:

- **`GSites_SharedDrive_Settings.csv`** — one row per drive: id, name,
  createdTime, and flattened `restrictions` fields
  (`adminManagedRestrictions`, `domainUsersOnly`, `driveMembersOnly`,
  `copyRequiresWriterPermission`, `sharingFoldersRequiresOrganizerPermission`).
  Sourced from `gam ... info shareddrive "~driveId" ... formatjson`, forced
  to `num_threads 1` (concurrent workers writing raw JSON to a shared
  stdout stream would risk interleaving/corrupting output), then
  reassembled and parsed by an inline PowerShell JSON-repair step (GAM's
  `formatjson` output has no separator or enclosing array between records).
- **`GSites_SharedDrive_Organizers.csv`** — one row per (drive, organizer),
  from `print shareddriveorganizers`. Lists Manager-role users/groups per
  drive. Note: full restriction visibility in the Settings export requires
  the scanning `account` to itself be a Manager/Organizer of that drive.

Metadata export failures are non-fatal to the overall run (Settings/
Organizers are not required for scoring) — `Run-SharedDriveAssessment.ps1`
logs a warning and continues.

### 3.9 `Run-SharedDriveAssessment.ps1` — Standalone Orchestrator

```powershell
.\Run-SharedDriveAssessment.ps1 -SharedDriveIdsCsv "SharedDriveIDs.csv" -PrimaryDomain "rocheua.com"
```

Runs `01d` → `01e` → `05b` in order, with the same logged-process/log-tail
pattern as the main orchestrator. Parameters: `-SharedDriveIdsCsv`
(required), `-PrimaryDomain` (required), `-DefaultScanningUserEmail`
(optional fallback account), `-GamThreads` (default 5), `-SkipExport`,
`-SkipMetadata`. Prints a final summary of all 5 possible output files
with row counts (or `[MISSING]` if a file wasn't produced).

### 3.10 `05b_score_shareddrive_sites.ps1` — Shared Drive Scoring (Standalone Step 3)

1. **Dedup** — collapses `GSites_SharedDrive_Inventory.csv` by `id` and
   `GSites_SharedDrive_Permissions.csv` by `id`+`permission.id`. Not
   usually needed for the `select teamdriveid` export pattern (each site is
   looked up directly, not iterated per-member like the old `all users`
   approach), but kept as a safety net in case the same drive ID appears
   twice in the input CSV, or a site is reachable via more than one row.
2. **Score** — delegates to the shared `05_score_sites.ps1` engine (§3.5),
   pointing its file parameters at the `GSites_SharedDrive_*.csv` files and
   writing `GSites_SharedDrive_Complexity_Report.csv`.

**Important:** this pipeline never crawls page content, so `Pages.csv`/
`Embeds.csv`/`ExternalDomains.csv` don't exist for it — `StructurePoints`
and `EmbedPoints` are always 0. Only `SecurityPoints` (derived from
permissions) contributes to `TotalScore`. This is expected; the Shared
Drive pipeline answers "how exposed is this site's sharing?" rather than
"how complex is this site to rebuild?".

---

## 4. Orchestrator Support Functions (`Run-FullAssessment.ps1`)

| Function | Purpose |
|---|---|
| `Deduplicate-GamExports` | Collapses duplicate rows in the inventory/permissions CSVs by `id` (inventory) or `id`+`permission.id` (permissions). Runs automatically after Step 1, and again if `-SkipGAMExport` reuses an existing export. See §5.1. |
| `Filter-InventoryBySelectedSites` | When `-SelectedSitesCsv` is given, narrows the inventory (and permissions) CSVs down to only the named/URL-matched sites, backing up the full inventory to `*.full` first. |
| `Extract-SiteNameFromValue` | Parses a Google Sites URL (`https://sites.google.com/<domain>/<site-name>[...]`) down to the bare site name, or passes through a plain name unchanged. |
| `Build-GamNameFilter` | Turns a `-SelectedSitesCsv` site list into a GAM Drive query `name='...' or name='...'` fragment (`GAM_SITES_FILTER`), so GAM only scans for those sites instead of the whole tenant. Falls back to a full scan if the filter would exceed the ~800-char safe query length. |
| `Build-GamTargetUsersFile` | Extracts owner/user email addresses from `-SelectedSitesCsv` or `-TargetUsersCsv` into a CSV GAM can consume via `csv ... gam user "~Email"` (`GAM_TARGET_FILE`), restricting the GAM scan to only those users' Drives instead of `all users`. |
| `Normalize-CsvHeaders` | Strips numeric array-index infixes from GAM's CSV header row (e.g. `owners.0.emailAddress` → `owners.emailAddress`). |
| `Invoke-LoggedProcess` | Runs a child process (used for the GAM `.cmd`) with stdout/stderr redirected to timestamped log files under `logs/`, returning the exit code for error handling. |

---

## 5. Key Design Decisions / Problems Solved

### 5.1 Duplicate rows from GAM's per-user scan

GAM's `all users ... print filelist` iterates every domain user's Drive
independently. Before the Shared Drive exclusion (§5.2), a site hosted on
a Shared Drive was emitted once **per member with access** to that drive —
e.g. a site shared with 50 people produced 50 identical inventory/
permission rows, massively inflating row counts and causing the crawler/
scorer to redundantly reprocess the same site many times.
`Deduplicate-GamExports` (see §4) collapses these to one row per unique
`id`, preferring the row logged under the site's actual owner when
determinable. This remains as a safety net even after §5.2, for any
residual duplication (e.g. externally supplied `-InventoryCsv` files).

### 5.2 Shared Drive exclusion (memory + scope)

Per customer requirement, **all Shared Drive-related data is excluded**
from the assessment. In `01_run_gam_exports.cmd` this was done by removing
`corpora alldrives` from all three GAM `print filelist` commands (GAM
defaults to My-Drive-only scanning without it), removing
`showshareddrivepermissions` from the permissions export, and dropping the
now-meaningless `driveid`/`drivename` fields. Effects:

- Sites hosted only on a Shared Drive **no longer appear at all** in any
  output file.
- Eliminates the Shared-Drive-driven memory blowup previously observed
  (~97% memory usage) — each GAM worker no longer has to expand a Shared
  Drive's full membership/file tree per user, since only My Drive is
  scanned.
- Removes the per-user duplication described in §5.1 at the source for
  future runs (My Drive files are typically owned/scanned once).

### 5.3 Published URL removal

The original design also crawled sites via their **published URL** (the
public/embedded Sites API v1 `siteUrl`, fetched via
`03a_get_published_urls.js` using a service-account/domain-wide-delegation
token). This required enabling the Sites API and granting DWD scopes,
which repeatedly failed with `HTTP 403 SERVICE_DISABLED` for customers who
lacked Owner/Editor/Service Usage Admin rights to enable the API. This
entire feature was removed: `03a_get_published_urls.js`,
`get_service_account_token.js`, and `01b_grant_site_access.cmd` were
deleted, and both crawl scripts now always use the Drive **edit URL**
(`webViewLink`), which is already returned by the Step 1 GAM export and
requires no additional Google Cloud project configuration.

---

## 6. Output Files (`output/` folder)

**Main (My Drive) pipeline:**

| File | Produced by | Contents |
|---|---|---|
| `GSites_Inventory_Min.csv` | Step 1 | Quick id/name/mimetype sanity check |
| `GSites_Inventory_Detailed.csv` | Step 1 (deduped) | Full per-site metadata, one row per unique site |
| `GSites_Permissions.csv` | Step 1 (deduped) | One row per (site, grantee) permission |
| `Pages.csv` | Step 4 | One row per crawled page per site |
| `Embeds.csv` | Step 4 | One row per embedded/linked artifact found |
| `ExternalDomains.csv` | Step 4 | One row per distinct external domain referenced per page |
| `Complexity_Report.csv` | Step 5 | Final per-site score, rating, and migration recommendation |
| `GSites_Manifest.csv` | Step 6 (manual) | Consolidated per-site manifest joining all reports above |
| `gam_target_users.csv` | Step 1 (orchestrator) | Generated user-email list passed to GAM when scoping to specific users |
| `GSites_Inventory_Detailed.csv.full` | Orchestrator | Backup of the full (unfiltered) inventory when `-SelectedSitesCsv` is used |

**Standalone Shared Drive pipeline:**

| File | Produced by | Contents |
|---|---|---|
| `GSites_SharedDrive_Inventory.csv` | `01d` (deduped) | Per-site metadata for sites on the specified Shared Drives |
| `GSites_SharedDrive_Permissions.csv` | `01d` (deduped) | One row per (site, grantee), including direct-vs-inherited detail |
| `GSites_SharedDrive_Settings.csv` | `01e` | One row per drive: sharing restrictions |
| `GSites_SharedDrive_Organizers.csv` | `01e` | One row per (drive, organizer) |
| `GSites_SharedDrive_Complexity_Report.csv` | `05b` | Per-site score (security-only, see §3.10) |
| `GSites_SharedDrive_Manifest.csv` | `06_generate_manifest.ps1` (manual) | Consolidated per-site manifest for Shared Drive sites |

---

## 7. Prerequisites & Setup

Full step-by-step instructions are in `CUSTOMER_SETUP.md`. Summary of
what's required before a first run:

| Tool | Version | Used by |
|---|---|---|
| [GAM](https://github.com/GAM-team/GAM) | 7.x | Steps 1 (both pipelines) — all Google Workspace/Drive data export |
| [Node.js](https://nodejs.org/) | 18+ | Steps 2–4 — crawling and browser auth |
| [PowerShell](https://github.com/PowerShell/PowerShell) | 7.x (`pwsh`) | Orchestrators, scoring, manifest generation |
| [gcloud CLI](https://cloud.google.com/sdk/docs/install) | any | Only for `-UseApiExtract` (OAuth token minting) |

GAM path is resolved via `GAM_PATH` env var → `gam.cfg` file → system
`PATH` (see `GAM_PATH_FIX.md`). Node dependencies (`playwright`,
`csv-parse`, `csv-stringify`) and the Playwright Chromium browser are
installed automatically by Step 2 of the main orchestrator, or manually via
`npm install` + `npx playwright install chromium`. A one-time interactive
browser login (`02_save_playwright_auth.js`) is required before any
Playwright-based crawl; not needed for `-UseApiExtract` or the Shared
Drive pipeline (neither crawls page content).

---

## 8. Directory Reference

| File | Type | Role |
|---|---|---|
| `Run-FullAssessment.ps1` | PowerShell | Main pipeline orchestrator |
| `Run-SharedDriveAssessment.ps1` | PowerShell | Shared Drive pipeline orchestrator |
| `01_run_gam_exports.cmd` | Batch | My Drive GAM export |
| `01d_run_gam_exports_shareddrive.cmd` | Batch | Shared Drive GAM export |
| `01e_list_shareddrive_metadata.cmd` | Batch | Shared Drive metadata export |
| `02_save_playwright_auth.js` | Node.js | Interactive browser login capture |
| `03_crawl_sites.js` | Node.js | Playwright browser crawler |
| `03b_api_extract_embeds.js` | Node.js | Sites API v1 fast extractor |
| `05_score_sites.ps1` | PowerShell | Shared scoring engine (both pipelines) |
| `05b_score_shareddrive_sites.ps1` | PowerShell | Shared Drive dedup + scoring wrapper |
| `06_generate_manifest.ps1` | PowerShell | Consolidated per-site manifest (both pipelines) |
| `gam.cfg` | Config | Local `GAM_PATH` override (not committed with real paths) |
| `package.json` | Config | Node dependencies (`playwright`, `csv-parse`, `csv-stringify`) |
| `CUSTOMER_SETUP.md` | Docs | Step-by-step setup and run instructions |
| `GAM_PATH_FIX.md` | Docs | Why/how `GAM_PATH` resolution was fixed |
| `GAM_DOCUMENTATION.md` | Docs | Deep dive into every GAM command/flag used, and why |
| `TOOLKIT_DOCUMENTATION.md` | Docs | This file — full pipeline design and behavior |
| `output/` | Folder | All generated CSV reports (created on first run) |
| `logs/` | Folder | Per-step stdout/stderr logs from the orchestrators |
| `.auth/state.json` | Generated | Saved Playwright browser session (not committed) |
