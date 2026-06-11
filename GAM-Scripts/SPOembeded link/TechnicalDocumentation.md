# Google Sites → SPO Embed Migration Toolkit — Technical Documentation

## 1. `Add-SPOYouTubeWebParts.ps1`

**Purpose:** Reads a CSV mapping file and injects `ContentEmbed` (or historically `YouTube`) web parts into SharePoint Online modern pages.

**Parameters:**
- `SiteUrl` (Mandatory) — Target SPO site URL.
- `MappingCsv` (Mandatory) — Path to CSV with columns: `PageName`, `EmbedUrl`, `SectionIndex`, `ColumnIndex`, `Order`.
- `ClientId` (Optional, default `3834b2e7-ab80-45fc-b4c8-ed5c960076b7`) — Azure AD App Registration ClientId for PnP PowerShell.
- `TenantId` (Optional, derived from `SiteUrl` if omitted) — Tenant name (e.g., `mngenvmcap908272.onmicrosoft.com`).
- `Publish` (Switch) — Calls `$page.Publish()` after injection.
- `DryRun` (Switch) — Logs intended actions without modifying SPO.

**Cmdlet / SDK Functions Used:**
- `Get-Module -ListAvailable -Name "PnP.PowerShell"` — Validates module presence.
- `Connect-PnPOnline` — Authenticates via `ClientId` + `TenantId` + `-Interactive` (MSAL browser flow).
- `Import-Csv` — Parses the mapping CSV.
- `Get-PnPPage -Identity <pageName>` — Loads the modern page object (`IPage`).
- `Add-PnPPageSection -Page $page -SectionTemplate OneColumn` — Appends a new section when the target is `OneColumnFullWidth` or beyond the current section count.
- `Add-PnPPageWebPart -Page $page -DefaultWebPartType $webPartType -WebPartProperties $jsonProps -Section $n -Column $n -Order $n` — Injects the web part.
- `$page.Publish("comment")` — Publishes the page.
- `Disconnect-PnPOnline` — Cleans up the session.

**Internal Logic:**
1. Validates `PnP.PowerShell` is installed.
2. Derives `TenantId` from `SiteUrl` if blank (`Uri.Host` split).
3. Connects interactively (cached MSAL token may skip browser prompt).
4. Iterates CSV rows. For each row:
   - **Web part type detection:** If the URL matches `youtube\.com|youtu\.be`, extracts `videoId` via regex (`v=`, `youtu\.be/`, `embed/`).
   - **HTML iframe construction:** Builds a full `<iframe ...></iframe>` string. Ampersands (`&`) in the `src` attribute are escaped to `&amp;` to prevent the SharePoint embed web part's `_extractEmbedSrc` parser from throwing `Cannot read properties of undefined (reading 'match')`.
   - **Section safety check:** If the target section is `OneColumnFullWidth` (e.g., Section 1 on many migrated pages), the script appends a new `OneColumn` section at the end and redirects the web part there.
   - **JSON serialization:** The iframe string is placed into `@{ embedCode = $embedCode }` and serialized with `ConvertTo-Json -Compress`.
   - **PnP injection:** Uses `Add-PnPPageWebPart` with `DefaultWebPartType = "ContentEmbed"`.
   - **Retry on full-width error:** If the first `Add-PnPPageWebPart` throws a full-width section conflict, catches the exception, appends a new section, and retries.

---

## 2. `Scan-GSitesEmbeds.ps1`

**Purpose:** Static scanner for Google Sites HTML exports (e.g., from Google Takeout or GAM downloads). Extracts all `<iframe src>`, `<embed src>`, and Google Sites `data-url` attributes.

**Parameters:**
- `ExportPath` (Mandatory) — Folder containing `.html` files.
- `OutputCsv` (Optional, default `GsitesEmbeds.csv`)

**Functions / Methods Used:**
- `Get-ChildItem -Path $ExportPath -Recurse -Filter "*.html"` — Enumerates HTML files.
- `Get-Content -Path $file.FullName -Raw` — Reads entire file into a single string.
- `[regex]::Matches($content, '<iframe[^>]*?\s+src=["\'']([^"\''>]+)["\''][^>]*?>')` — Pattern 1: standard iframe.
- `[regex]::Matches($content, '<embed[^>]*?\s+src=["\'']([^"\''>]+)["\''][^>]*?>')` — Pattern 2: `<embed>` tags.
- `[regex]::Matches($content, 'data-url=["\'']([^"\''>]+)["\'']')` — Pattern 3: Google Sites `data-url` attributes.
- `[System.Collections.Generic.List[PSCustomObject]]` — Accumulates results.
- `Export-Csv` — Writes the final inventory.

**Output Columns:**
- `SourceFile`, `FileName`, `PatternType` (`iframe`/`embed`/`gsites-data-url`), `EmbedUrl`, `EmbedType` (`YouTube`/`GoogleMaps`/`GoogleDrive`/`GoogleDocs`/`GenericIframe`), `ContextHtml`

---

## 3. `Find-SPOEmptyEmbeds.ps1`

**Purpose:** Audits every modern page in the `SitePages` library for empty or placeholder embed web parts left by migration tools.

**Parameters:**
- `SiteUrl` (Mandatory)
- `OutputCsv` (Optional, default `SPOEmptyEmbeds.csv`)
- `PlaceholderPatterns` (Optional array of regex strings)
- `ClientId`, `TenantId` (same auth pattern as `Add-SPOYouTubeWebParts.ps1`)

**Cmdlet / SDK Functions Used:**
- `Connect-PnPOnline`
- `Get-PnPListItem -List "SitePages" -PageSize 500` — Retrieves all page list items.
- `Get-PnPPage -Identity $fileName` — Loads each page object to inspect controls.
- `$page.Controls` — Iterates every web part / text control on the page.
- `$wp.PropertiesJson | ConvertFrom-Json` — Deserializes web part properties.
- `Export-Csv` — Writes audit results.

**Detection Logic:**
- For `PageWebPart` controls, checks `embedCode` property.
- `EmptyEmbedUrl`: `embedCode` is null, empty, or whitespace.
- `PlaceholderEmbed`: `embedCode` matches any of the default regex patterns (`\[embed\]`, `<iframe`, `not supported`, `cannot display`, `placeholder`, `google\.com\/embed`, `youtube\.com\/embed`).

---

## 4. `Diagnose-SPOPages.ps1`

**Purpose:** Lists every control on every modern page for troubleshooting page structure and section layout conflicts.

**Key Cmdlets:**
- `Get-PnPListItem -List "SitePages"`
- `Get-PnPPage -Identity $fn`
- `$p.Controls | Format-Table` — Displays `Type.Name`, `Title`, and a text snippet.

---

## 5. `Get-GSitesEmbeds-Api.ps1`

**Purpose:** Alternative extraction using the Google Sites REST API v1 (`sites.googleapis.com/v1/sites/{siteId}`). Requires an OAuth access token with `sites.readonly` scope.

**Key Components:**
- `Invoke-RestMethod` against `https://sites.googleapis.com/v1/sites/{siteId}?path=/&fields=*`
- Parses `embeddedContent` and `textContent` nodes in the JSON response.

---

## 6. `Invoke-FullEmbedRemediation.ps1`

**Purpose:** Master orchestrator that runs the end-to-end workflow in one command.

**Parameters:**
- `SPOUrl`, `Gam7Path`, `ClientId`, `TenantId`
- `SkipGoogleAuth` — Skips re-running `node 02_save_playwright_auth.js`.
- `SkipCrawl` — Skips the enhanced crawl; uses existing CSVs in `gam7/output/`.
- `DryRun` — Passes through to `Add-SPOYouTubeWebParts.ps1`.

**Workflow Steps:**
1. **Google Auth** — Runs `node 02_save_playwright_auth.js` in the `gam7` folder to refresh `storageState`.
2. **Enhanced Crawl** — Calls `Run-EnhancedCrawl.ps1`.
3. **Page Mapping** — Reads `07_Pages_Enhanced.csv` and `08_Embeds_Enhanced.csv`, then fuzzy-matches Google Sites page names to SPO page names by stripping spaces and special characters.
4. **CSV Generation** — Produces `EmbedMapping.csv`.
5. **SPO Injection** — Invokes `Add-SPOYouTubeWebParts.ps1` with the generated mapping.

---

## 7. `Run-EnhancedCrawl.ps1`

**Purpose:** PowerShell wrapper that copies `03_crawl_sites_enhanced.js` into the `gam7` folder (so `node_modules` resolves) and executes it.

**Parameters:**
- `SiteUrl` — Direct URL to a single Google Sites site (overrides CSV inventory).
- `SitesCsv` — Path to `02_GSites_Inventory_Detailed.csv` (Roche domain inventory).
- `Gam7Path` — Default `C:\Users\v-nmahanthi\OneDrive - Microsoft\Documents\gam7`.
- `DryRun` — Validates auth age and prerequisites without crawling.

**Functions / Methods Used:**
- `Get-Item` on `.auth/state.json` — Checks `LastWriteTime` and warns if >24 hours.
- `Copy-Item` — Copies the `.js` into `gam7` to resolve `MODULE_NOT_FOUND`.
- `Push-Location` / `Pop-Location` — Switches working directory to `gam7`.
- `node 03_crawl_sites_enhanced.js` — Executes the Node.js crawl.

---

## 8. `03_crawl_sites_enhanced.js`

**Purpose:** Headless Playwright crawler that navigates Google Sites, triggers lazy loading via scrolling, and extracts embed URLs from the DOM and raw HTML.

**Playwright APIs Used:**
- `chromium.launch({ headless: true })`
- `browser.newContext({ storageState: authFile })` — Reuses cached Google cookies.
- `page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 })`
- `page.evaluate(() => window.scrollBy(...))` — 6 scrolls to trigger lazy-loaded iframes.
- `page.waitForTimeout(ms)` — Waits between scrolls and after load.

**Extraction Logic (inside `page.evaluate`):**
- **DOM Scanner (`scan(root)`):**
  - Queries all `iframe` elements for `src`, `data-src`, `srcdoc`.
  - Queries `embed`, `object`, `video`, `audio`, `source` for `src`/`data`.
  - Queries `[data-url]`, `[data-src]`, `[data-embed-url]` custom attributes.
  - **Recursive shadow DOM:** scans `el.shadowRoot` for every custom element.
- **Regex HTML Scan:**
  - `/(?:https?:\/\/)?(?:www\.)?(?:youtube\.com\/watch\?v=|youtube\.com\/embed\/|youtu\.be\/)([a-zA-Z0-9_-]{11})/g`
  - `/(?:https?:\/\/)?(?:www\.)?google\.com\/maps\/embed[^"'\s]*/g`
  - `/(?:https?:\/\/)?drive\.google\.com\/file\/d\/[^"'\s]+/g`
- **Output:** Writes `07_Pages_Enhanced.csv`, `08_Embeds_Enhanced.csv`, `09_ExternalDomains_Enhanced.csv`, and HTML snapshots to `output/html/`.

---

## 9. `Run-ExtractSiteEmbeds.ps1`

**Purpose:** Lightweight wrapper for crawling a **single** Google Sites URL using `Extract-SiteEmbeds-Playwright.js`.

**Parameters:**
- `SiteUrl` — The target Google Sites URL.
- `Gam7Path` — Where `node_modules` and auth state live.
- `OutputCsv` — Results file.

**Logic:**
- Copies the JS file into `gam7`.
- Opens a **visible** Chromium (`headless: false`) so the user can sign in if the auth state is stale.
- Runs `node Extract-SiteEmbeds-Playwright.js <url> <outputCsv>`.

---

## 10. `Extract-SiteEmbeds-Playwright.js`

**Purpose:** Standalone Node.js single-page extractor. Uses a visible browser window.

**Key Differences from `03_crawl_sites_enhanced.js`:**
- Accepts a single URL via `process.argv[2]`.
- Defaults to `headless: false` (visible window for manual sign-in).
- Does **not** use `storageState` by default; the modified version added `storageState` support for headless reuse.
- Outputs a simpler CSV: `Kind,EmbedUrl,Context`.

---

## Common File Formats

### `EmbedMapping.csv`
```csv
PageName,EmbedUrl,SectionIndex,ColumnIndex,Order
Resources.aspx,"https://docs.google.com/...",1,1,0
```
- URLs containing commas **must** be double-quoted to prevent `Import-Csv` column shift.

### `08_Embeds_Enhanced.csv`
```csv
SiteId,SiteName,SiteUrl,PageUrl,PageTitle,ItemKind,ArtifactType,ArtifactUrl
```

---

## Key PnP PowerShell WebPart Types Used

| Enum Value | Purpose |
|------------|---------|
| `ContentEmbed` | Generic iframe embed (replaced the invalid `Embed` value) |
| `YouTube` | Native YouTube web part (required `embedCode`, not `videoId`, in this environment) |

---

## Troubleshooting Notes

- **Auth state expiry:** `gam7/.auth/state.json` older than 24 hours may cause Google sign-in pages to load. Re-run `node 02_save_playwright_auth.js`.
- **Section conflicts:** `OneColumnFullWidth` sections reject standard web parts. The script auto-detects this and appends a new `OneColumn` section.
- **HTML `&` escaping:** Unescaped `&` inside `embedCode` iframe `src` attributes causes `_extractEmbedSrc` JS crash in the SharePoint embed web part.
- **Duplicate web parts:** Failed script runs can leave broken web parts. Use `Remove-PnPPageComponent -Page <name> -InstanceId <guid> -Force` to clean them.
