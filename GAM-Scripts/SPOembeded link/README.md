# Google Sites to SharePoint Online Embed Migration Toolkit

This toolkit helps recover embedded content (YouTube, Google Maps, etc.) that was lost during a Google Sites to SharePoint Online migration (e.g., via Cloudiway).

## Prerequisites

- **PowerShell 5.1 or PowerShell 7+**
- **PnP.PowerShell** module (for SPO scripts):
  ```powershell
  Install-Module PnP.PowerShell -Force
  ```
- **Google Sites HTML export** (for `Scan-GSitesEmbeds.ps1`)
- SharePoint Online permissions: **Site Owner / Designer** to modify modern pages

## Toolkit Overview

| Script | Purpose |
|--------|---------|
| `Scan-GSitesEmbeds.ps1` | Extract all embed/iframe URLs from your Google Sites HTML export |
| `Find-SPOEmptyEmbeds.ps1` | Audit migrated SPO pages for empty or placeholder embed web parts |
| `Add-SPOYouTubeWebParts.ps1` | Bulk-add Embed/YouTube web parts to SPO pages using a CSV mapping |

## Workflow

### Step 1: Inventory What Was Lost

Run against your **Google Sites HTML export folder**:

```powershell
.\Scan-GSitesEmbeds.ps1 -ExportPath "C:\GsitesExport" -OutputCsv "GsitesEmbeds.csv"
```

Output columns:
- `SourceFile` — relative path to the HTML file
- `FileName` — page file name
- `PatternType` — `iframe`, `embed`, or `gsites-data-url`
- `EmbedUrl` — the extracted URL
- `EmbedType` — auto-detected: `YouTube`, `GoogleMaps`, `GoogleDrive`, `GenericIframe`
- `ContextHtml` — snippet for reference

### Step 2: Audit Migrated SPO Pages

Connect to your **migrated SharePoint site** to find broken embeds:

```powershell
.\Find-SPOEmptyEmbeds.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/migrated" -OutputCsv "SPO_Issues.csv"
```

This detects:
- Empty **Embed** web parts
- Placeholder text like `[embed]`, `<iframe`, or `not supported`
- Broken iframe HTML dumped into **Text** web parts by migration tools

### Step 3: Build a Mapping CSV

Create a CSV named `EmbedMapping.csv` with these columns:

```csv
PageName,EmbedUrl,SectionIndex,ColumnIndex,Order
Home.aspx,https://www.youtube.com/embed/dQw4w9WgXcQ,1,1,0
About.aspx,https://www.google.com/maps/embed?pb=...,1,1,0
```

Notes:
- `SectionIndex` and `ColumnIndex` default to `1` if omitted
- `Order` controls position within the column (default `0` = top)
- The script auto-detects YouTube URLs and uses the **YouTube** web part; all others use **Embed**

### Step 4: Apply the Embeds

Preview without making changes:

```powershell
.\Add-SPOYouTubeWebParts.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/migrated" -MappingCsv "EmbedMapping.csv" -DryRun
```

Apply and publish pages:

```powershell
.\Add-SPOYouTubeWebParts.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/migrated" -MappingCsv "EmbedMapping.csv" -Publish
```

Apply without publishing (manual review first):

```powershell
.\Add-SPOYouTubeWebParts.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/migrated" -MappingCsv "EmbedMapping.csv"
```

## Tips

1. **Run in a test site first** — always validate on a copy of your production site.
2. **Match page names exactly** — `PageName` in the CSV must match the file name in the SPO Site Pages library (e.g., `Home.aspx`).
3. **YouTube ID extraction** — the script parses standard YouTube URLs. If your Google Sites used shortened or custom URLs, verify the `videoId` detection logic.
4. **Authentication** — scripts use `-Interactive` (web login). If you need app-only auth, replace `Connect-PnPOnline` with your preferred PnP connection method.
5. **Column/section positioning** — modern SPO pages can have multiple sections and columns. Use `Get-PnPPage` manually to inspect existing layouts before setting `SectionIndex`/`ColumnIndex`.

## Troubleshooting

| Issue | Resolution |
|-------|------------|
| `PnP.PowerShell not found` | Run `Install-Module PnP.PowerShell -Force` |
| `Failed to load page` | Ensure the page exists in the Site Pages library and is a modern page |
| `Access denied` | Confirm you are a Site Owner or have Edit permissions on the site |
| `YouTube web part not showing` | Check that the video ID was extracted correctly; fall back to Embed web part if needed |
