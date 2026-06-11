# GAM-Scripts — Folder-by-Folder Reference

This directory groups all PowerShell / Node.js automation used in the
**Google Workspace → Microsoft 365 (Google Sites → SharePoint Online)** migration
project. Every sub-folder is independent and solves a different phase of the
migration lifecycle: **discovery → assessment → migration cut-over → post-migration remediation → ancillary inventory**.

| # | Folder | Phase | Primary Output |
|---|--------|-------|----------------|
| 1 | `gsites-assessment-toolkit/` | Pre-migration discovery & complexity scoring of Google Sites | `10_SiteAssessmentSummary.csv` |
| 2 | `SPORemidation/` | Post-migration fixes to SPO navigation, layout & page formatting | `FullRemediation_<ts>.html/.csv` + per-site logs |
| 3 | `SPOembeded link/` | Post-migration recovery of YouTube / Maps / Drive embeds lost by Cloudiway | `EmbedMapping.csv` → injected web parts |
| 4 | `Task Details/` | One-off audit of Google Tasks (creator/assignee derivation) prior to migration | `AllUsers_Tasks_v9.csv` |
| 5 | `connected-Apps/` | Google Workspace OAuth / 3rd-party app / chat-bot inventory | `ConnectedAppsReport_<ts>/*.csv + .html` |

---

## 1. `gsites-assessment-toolkit/`

**Purpose.** End-to-end *non-prod* assessment pack that inventories every Google
Site in a tenant, crawls each site headlessly to count pages and embedded
artifacts, enriches the linked Sheets/Forms/Apps Script files via Google REST
APIs, and finally produces a complexity score used to triage which sites are
easy/medium/hard to migrate.

**Key files**

| File | Role |
|------|------|
| `Run-FullAssessment.ps1` | Orchestrator — runs all 5 numbered steps end-to-end. |
| `01_run_gam_exports.cmd` | GAM7 batch — produces 6 CSVs: site inventory (min + detailed), permissions, candidate Sheets/Forms/Apps-Script lists. |
| `02_save_playwright_auth.js` | Opens a visible Chromium so a tester can sign in to Google once; saves cookies/storage to `.auth/state.json` for headless reuse. |
| `03_crawl_sites.js` | Playwright headless crawler — walks every site, captures `iframe`/`embed`/`data-url` artifacts, external domains, page tree depth. Writes `07_Pages.csv`, `08_Embeds.csv`, `09_ExternalDomains.csv`. |
| `03a_get_published_urls.js` | Calls Google Sites REST API v1 to resolve the **published** URL of each site (edit URLs return 403 to crawlers). |
| `04_enrich_artifacts.ps1` | For every Sheet/Form/Apps-Script discovered, calls Sheets v4 / Forms v1 / Apps-Script API to compute `ComplexityPoints` (cells, named ranges, function count, etc.). |
| `05_score_sites.ps1` | Aggregates everything into `10_SiteAssessmentSummary.csv` with a single complexity score per site. |
| `gam.cfg` / `GAM_PATH_FIX.md` | Externalised GAM path. Three-strategy lookup: `GAM_PATH` env var → `gam.cfg` → system `PATH`. See `GAM_PATH_FIX.md`. |
| `package.json` | Pulls Playwright + `csv-parse` + `csv-stringify`. `node_modules/` is committed for offline runs. |

**Tooling used.** GAM7 (Google Workspace CLI), Playwright (Chromium headless),
Google Sites REST v1, Sheets API v4, Forms API v1, Apps Script API. Auth is OAuth
2.0 bearer (`GCP_ACCESS_TOKEN` env var or `gcloud auth print-access-token`).

**How it works.**
1. `01_run_gam_exports.cmd` issues parallel `gam print filelist` queries filtered by `mimeType=application/vnd.google-apps.site` (plus spreadsheets, forms, scripts).
2. `02_save_playwright_auth.js` produces a reusable `storageState` so all later headless steps skip Google's SSO challenge.
3. `03_crawl_sites.js` reads `02_GSites_Inventory_Detailed.csv`, opens each `webViewLink` in a Playwright context bound to that state, scrolls to trigger lazy-loaded iframes, records every artifact URL and the per-page tree depth.
4. `04_enrich_artifacts.ps1` joins each artifact URL back to its file ID, calls the corresponding Google API, and emits `ComplexityPoints`.
5. `05_score_sites.ps1` combines page count, embed count, external-domain count, permission breadth (`anyone` / `domain` / `external email`), and artifact points to compute a final risk/complexity score per site.

---

## 2. `SPORemidation/`

**Purpose.** After Cloudiway (or equivalent) migrates a Google Site to
SharePoint Online, three classes of defects remain. This folder fixes all three
in one orchestrated, parallel pass across N sites listed in `SiteMapping.csv`.

| Defect | Script |
|--------|--------|
| Navigation links still point at `sites.google.com`, empty `#`, or are ghost nodes with no Id/Title | `Fix-SPONavigation.ps1` |
| Mega-menu (horizontal) nav doesn't match the Google Sites vertical dropdown | `Set-SPONavigationCascading.ps1` |
| Nested `<ol>` lists render as `a, b, c` (Google style) and bullets as squares | `Fix-SPOPageFormatting.ps1` |

**Key files**

| File | Role |
|------|------|
| `Invoke-FullRemediation.ps1` | Master orchestrator. Reads `SiteMapping.csv`, prompts MSAL **once**, then for each site runs nav-URL fix → mega-menu → page-formatting fix, emits per-site `.log`, a CSV summary, and a styled HTML report under `RemediationLogs/`. |
| `Fix-SPONavigation.ps1` | Dual-mode: full rebuild (wipe + recreate hierarchy) **or** auto-repair (rewrites `sites.google.com/*` → `/sites/<slug>/SitePages/*.aspx`, deletes ghost nodes via SP REST `topnavigationbar/getById(...)`). |
| `Set-SPONavigationCascading.ps1` | `PATCH /_api/web {MegaMenuEnabled:false}` per site, verifies, supports `-WhatIf` and parallel bulk via `ForEach-Object -Parallel`. |
| `Fix-SPOPageFormatting.ps1` | Loads each page via `Get-PnPPage`, walks every text web part, applies 3 regex fixes (`list-style-type:lower-alpha`→`decimal`, `<ol type="a">`→`type="1"`, `list-style-type:square`→`disc`), saves and publishes. |
| `Invoke-BulkNavFix.ps1` | Legacy parallel runner for **just** the nav-URL fix (kept for piecemeal runs). |
| `SiteMapping.csv` | Input — columns: `SiteName, GoogleSitesBaseUrl, SPOSiteUrl, RebuildNavigation`. |
| `SPO-Remediation-Complete-Guide.md` | Full 576-line operator manual. |
| `SPO-Navigation-Remediation-Guide.md` | Older nav-only guide kept for reference. |

**Tooling used.** PowerShell 7+ (`ForEach-Object -Parallel`), PnP.PowerShell
(`Connect-PnPOnline -Interactive` with a hardcoded Azure AD App
`ClientId 3834b2e7-…`), SharePoint REST (`/_api/web/navigation/topnavigationbar`,
`/_api/web?$select=MegaMenuEnabled`). MSAL token caching means the browser
prompt fires **once** in the orchestrator and is reused silently for every
subsequent site.

**How it works.**
1. Orchestrator dot-sources `Fix-SPONavigation.ps1` and `Fix-SPOPageFormatting.ps1` to gain their internal functions (`Invoke-SPONavFix`, `Repair-SPOPage`, `Repair-PageHtml`).
2. A throwaway `Connect-PnPOnline` against the **first** site forces the MSAL interactive prompt; the cached token is reused for every site thereafter.
3. For each row in `SiteMapping.csv` a single connect is made, then steps 1/2/3 reuse that connection (`-UseCurrentConnection` switch on `Invoke-SPONavFix`).
4. All results aggregate into a `PSCustomObject` per site (`NavFix`, `CascadingFix`, `PagesFixed`, `PagesScanned`, `PagesFailed`, `Error`, `LogFile`) and feed both the CSV summary and the colour-coded HTML report.

`-DryRun` switch (not `-WhatIf` — PowerShell consumes `-WhatIf` as a common
parameter before the script sees it) simulates without writing changes.

---

## 3. `SPOembeded link/`

**Purpose.** Cloudiway-style Google Sites to SPO migrators routinely drop
embedded content (YouTube videos, Google Maps frames, Drive previews,
arbitrary iframes). This folder discovers what was lost on the Google side,
audits what is missing on the SPO side, builds a mapping CSV, and bulk-injects
the correct PnP web part (`ContentEmbed` or native `YouTube`) into the right
section/column of each migrated SPO page.

**Key files**

| File | Role |
|------|------|
| `Invoke-FullEmbedRemediation.ps1` | End-to-end orchestrator: Google auth -> enhanced crawl -> page-name fuzzy match -> produces `EmbedMapping.csv` -> calls `Add-SPOYouTubeWebParts.ps1`. |
| `Scan-GSitesEmbeds.ps1` | Static scanner over a Google Sites HTML export - regex-extracts `<iframe src>`, `<embed src>`, `data-url`. Classifies as `YouTube`/`GoogleMaps`/`GoogleDrive`/`GoogleDocs`/`GenericIframe`. |
| `Get-GSitesEmbeds-Api.ps1` | Alternative live extraction via `sites.googleapis.com/v1/sites/{siteId}` (`sites.readonly` OAuth scope). |
| `Run-EnhancedCrawl.ps1` + `03_crawl_sites_enhanced.js` | Playwright crawler that scrolls 6x per page to trigger lazy-loaded iframes, then DOM-scans (incl. shadow-DOM) plus a regex pass over raw HTML for YouTube/Maps/Drive URLs. Emits `07_Pages_Enhanced.csv`, `08_Embeds_Enhanced.csv`, `09_ExternalDomains_Enhanced.csv`. |
| `Run-ExtractSiteEmbeds.ps1` + `Extract-SiteEmbeds-Playwright.js` | Lightweight visible-browser single-site variant for ad-hoc checks. |
| `Find-SPOEmptyEmbeds.ps1` | Audits modern SPO pages - flags Embed web parts whose `embedCode` is empty or matches placeholder regexes (`[embed]`, `<iframe`, `not supported`, `placeholder`, ...). |
| `Diagnose-SPOPages.ps1` | Dumps every control on every modern page for layout debugging. |
| `Add-SPOYouTubeWebParts.ps1` | The injector. Reads `EmbedMapping.csv`, auto-detects YouTube vs generic, builds an `&amp;`-escaped iframe HTML, wraps it in `@{embedCode=...}` (serialised with `ConvertTo-Json -Compress`), and calls `Add-PnPPageWebPart -DefaultWebPartType ContentEmbed`. Auto-handles the `OneColumnFullWidth` section conflict by appending a new `OneColumn` section. |
| `EmbedMapping.csv` / `GsitesEmbeds.csv` / `SPOEmptyEmbeds.csv` / `ResourcesEmbedMapping.csv` / `ExtractedEmbeds_FromCrawl.csv` / `SPO_Audit_Results.csv` | Working inventories produced/consumed by the scripts above. |
| `README.md` + `TechnicalDocumentation.md` | Existing operator + technical docs (228 lines of deep notes on every script). |

**Tooling used.** PnP.PowerShell (same client-id pattern as `SPORemidation/`),
Playwright (Chromium), Google Sites REST v1, Node.js + `csv-parse` / `csv-stringify`.

**Critical implementation notes** (from `TechnicalDocumentation.md`):
* All `&` in iframe `src` are escaped to `&amp;` - otherwise SharePoint's embed
  web part throws `Cannot read properties of undefined (reading 'match')`.
* If the target section is `OneColumnFullWidth`, the script appends a new
  `OneColumn` section and retries - `OneColumnFullWidth` rejects standard web parts.
* `DefaultWebPartType = "ContentEmbed"` (the literal value `Embed` is **not**
  valid in current PnP.PowerShell).

---

## 4. `Task Details/`

**Purpose.** The Google Tasks API does **not** expose `creator` or `assignee`
fields on a task. Pre-migration we still need to know **who owns each task**
so the equivalent Microsoft Planner / To-Do task can be re-created against the
right user. This folder derives those fields by correlating Tasks with Drive
files (Docs comments) and Chat-space messages.

**Key files**

| File | Role |
|------|------|
| `Get-GoogleTasksWithCreator.ps1` | Main script (929 lines). Iterates users via `gam print tasks formatjson`. For each task: **Assignee** = user whose tasklist holds it; **Creator** = (Docs surface) owner of the linked Drive file, looked up with `gam user <u> show fileinfo <id>`; (Chat surface) sender of the nearest `Created/Assigned a task ...` chat message correlated by timestamp delta (default tolerance ±10 s). Self-created tasks -> creator = assignee. Also performs orphaned-thread scans (`Origin = SPACESCAN`) and Docs-comment scans (`Origin = DOCSCAN`). |
| `_Get-AssignedTasks.ps1` | PowerShell 7 helper. Mints a user-impersonated JWT from the GAM service-account key (`~/.gam/oauth2service.json`, RS256 via `RSA.ImportFromPem`), exchanges it for an access token, and calls Tasks v1 with `showAssigned=true` - this is the **only** way to retrieve Docs/Chat-assigned tasks (GAM CLI cannot). Supports bulk NDJSON mode parallelised with `ForEach-Object -Parallel`. |
| `Test-GoogleTasksScript.ps1` | Self-contained test harness. Builds a fake `gam.cmd` shim that returns canned CSVs/JSON, runs the main script against it, and asserts the produced CSV row-by-row (5 expected rows: 3 GAM + 1 orphaned SPACESCAN + 1 DOCSCAN). |
| `_v5check.ps1` / `_v6check.ps1` / `_docscan_check.ps1` | Ad-hoc inspectors that group the output CSV by `Origin`, `SurfaceType`, `Shared`. Used during dev iterations (v1 -> v9). |
| `AllUsers_Tasks_v9.csv` | Latest production output. Columns: `AssigneeEmail, TaskListId, TaskListTitle, TaskId, Title, Status, Due, Completed, Updated, WebViewLink, SurfaceType, SourceRef, SourceLink, CreatorEmail, CreatorName, Shared, Origin, Notes`. |
| `AllUsers_Tasks_v1_Summary.csv` | Per-creator rollup: `CreatorEmail, SharedCount, AssigneeEmails, Titles, SurfaceTypes`. |

**Tooling used.** GAM7 CLI (`print tasks`, `print tasklists`, `show fileinfo`,
`print chatmessages`, `info chatspace`, `print filelist`, `print filecomments`,
`print chatspaces`, `print chatmembers`), Google Tasks REST v1 (only via the
helper, scope `https://www.googleapis.com/auth/tasks.readonly`),
PowerShell 7+ for parallelism.

**How it works.**
1. `Resolve-UserList` accepts `-Users a,b,c`, the literal `all` (enumerates via `gam print users`), a `.txt` of emails, or a CSV with a `primaryEmail` column.
2. A "tenant bulk" optimisation calls `gam <select-many-users> print tasks` once instead of N times, keyed back to each user via a hashtable.
3. Per task, the script inspects `assignmentInfo.surfaceType`:
   * `DOCUMENT` -> `Get-DriveFileOwner` returns the Drive owner = creator.
   * `SPACE` -> `Get-SpaceTaskCreator` parses `print chatmessages` for the matching `Created/Assigned a task` row by timestamp delta.
   * (missing) -> `Origin = SELF`, creator = assignee.
4. The REST helper (`_Get-AssignedTasks.ps1`) is invoked once per user (or once for all users in NDJSON bulk mode) to capture tasks visible only through `showAssigned=true`.
5. Optional `SPACESCAN` and `DOCSCAN` passes find orphaned tasks (chat thread mentions or doc comments) that have **no** API task row.

---

## 5. `connected-Apps/`

**Purpose.** Pre-migration inventory of every third-party / OAuth application,
service account, admin role and chat bot active inside the Google Workspace
tenant. Outputs an HTML dashboard + per-section CSVs that security/IT use to
decide what needs re-provisioning (or revoking) on the Microsoft side.

**Key files**

| File | Role |
|------|------|
| `Get-ConnectedAppsReport.ps1` | Main report (1137 lines). Produces 9 sections: domain users + last login, per-user OAuth tokens & scopes, distinct-app aggregate, suspended users still holding tokens (security flag), Marketplace apps (`gam print policies`), service accounts with Domain-Wide Delegation (`gam print svcaccts` with four fallback strategies and a per-strategy timeout), admin roles (`gam print admins`), optional last-activity (Reports API), optional Chat Spaces (requires GAMADV-XTD3 / GAM7). Emits CSVs + a single timestamped HTML dashboard under `ConnectedAppsReport_<ts>/`. |
| `Get-ChatBots (1).ps1` | Standalone Chat-bot audit (450 lines). Pulls `gam report chat` (same data source as Admin Console Reports) over a window (default last 30 days), then enriches with `gam print chatspaces`. Outputs raw activity, per-bot and per-space rollups, plus a last-modified leaderboard. Self-unblocks itself via `Unblock-File` on first run to suppress the Windows Zone-Identifier warning. |

**Tooling used.** GAM / GAMADV-XTD3 / GAM7 only - no Microsoft side at all in
this folder. Auto-detects the GAM binary if `-GamPath` is not supplied; uses
`Start-Job` with per-strategy timeouts so a slow `print svcaccts` call cannot
hang the whole report.

---

## End-to-End Data Flow

```
                  +----------------------------------------------+
                  |  PRE-MIGRATION INVENTORY (read-only)         |
                  |  - gsites-assessment-toolkit/  (Sites)       |
                  |  - Task Details/               (Tasks)       |
                  |  - connected-Apps/             (OAuth/Bots)  |
                  +----------------------------------------------+
                                       |
                                       v  (Cloudiway / equivalent runs the actual migration)
                                       |
                  +----------------------------------------------+
                  |  POST-MIGRATION REMEDIATION (writes to SPO)  |
                  |  - SPORemidation/    (nav, layout, lists)    |
                  |  - SPOembeded link/  (YouTube/Maps/iframes)  |
                  +----------------------------------------------+
```

## Common Conventions Across Folders

* **Azure AD App** - every PnP-PowerShell script defaults to client-id
  `3834b2e7-ab80-45fc-b4c8-ed5c960076b7`. Override with `-ClientId` per call.
* **Tenant auto-derivation** - `https://<prefix>.sharepoint.com/...` ->
  `<prefix>.onmicrosoft.com`. Override with `-TenantId`.
* **`-DryRun` vs `-WhatIf`** - orchestrators use `-DryRun` because PowerShell
  intercepts `-WhatIf` as a common parameter before the param block can read it.
* **CSV-driven bulk runs** - `SiteMapping.csv`, `EmbedMapping.csv`,
  `02_GSites_Inventory_Detailed.csv` all flow into `ForEach-Object -Parallel`
  for throttled (`-ThrottleLimit 5`) execution.
* **Single sign-in** - orchestrators run one MSAL prompt up-front; cached token
  is reused silently for every subsequent site.

## Where to Look First (Deep Docs Already in the Repo)

| Folder | Pre-existing deep doc |
|--------|----------------------|
| `SPORemidation/` | `SPO-Remediation-Complete-Guide.md` (576 lines) + `SPO-Navigation-Remediation-Guide.md` |
| `SPOembeded link/` | `TechnicalDocumentation.md` (228 lines) + `README.md` |
| `gsites-assessment-toolkit/` | `GAM_PATH_FIX.md` |
| `Task Details/` | Inline `.SYNOPSIS` / `.DESCRIPTION` blocks in each `.ps1` |
| `connected-Apps/` | Inline `.SYNOPSIS` / `.DESCRIPTION` blocks in each `.ps1` |
