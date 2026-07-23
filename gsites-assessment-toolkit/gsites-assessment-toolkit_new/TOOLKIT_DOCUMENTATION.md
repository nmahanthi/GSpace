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
(see §5.2). Only sites stored in an individual user's My Drive are
assessed.

---

## 2. Pipeline Overview

`Run-FullAssessment.ps1` is the single entry point. It runs 5 steps in order:

| Step | Script | Purpose | Can skip? |
|---|---|---|---|
| 1 | `01_run_gam_exports.cmd` | Export site inventory + permissions via GAM | `-SkipGAMExport` |
| 2 | (inline in orchestrator) | Install Node.js deps (Playwright, csv-parse, csv-stringify) | `-SkipDependencyCheck` |
| 3 | `02_save_playwright_auth.js` | Save a logged-in browser session for crawling | `-SkipBrowserAuth` |
| 4 | `03_crawl_sites.js` **or** `03b_api_extract_embeds.js` | Crawl each site, extract pages/embeds/external links | `-SkipCrawl` |
| 5 | `05_score_sites.ps1` | Compute complexity score per site, write final report | always runs |

Each script writes into a shared `output/` folder so later steps can pick
up earlier CSVs by fixed filename.

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

| File | Produced by | Contents |
|---|---|---|
| `GSites_Inventory_Min.csv` | Step 1 | Quick id/name/mimetype sanity check |
| `GSites_Inventory_Detailed.csv` | Step 1 (deduped) | Full per-site metadata, one row per unique site |
| `GSites_Permissions.csv` | Step 1 (deduped) | One row per (site, grantee) permission |
| `Pages.csv` | Step 4 | One row per crawled page per site |
| `Embeds.csv` | Step 4 | One row per embedded/linked artifact found |
| `ExternalDomains.csv` | Step 4 | One row per distinct external domain referenced per page |
| `Complexity_Report.csv` | Step 5 | Final per-site score, rating, and migration recommendation |
| `gam_target_users.csv` | Step 1 (orchestrator) | Generated user-email list passed to GAM when scoping to specific users |
| `GSites_Inventory_Detailed.csv.full` | Orchestrator | Backup of the full (unfiltered) inventory when `-SelectedSitesCsv` is used |
