/**
 * 03b_api_extract_embeds.js
 *
 * Extracts embedded content from Google Sites using the Sites API v1.
 * Produces the same output files as 03_crawl_sites.js but requires no
 * browser, no Playwright auth session, and runs in minutes instead of hours.
 *
 * Requires: GCP_ACCESS_TOKEN env var (scope: sites.readonly)
 * Input:    output/GSites_Inventory_Detailed.csv
 * Output:   output/Pages.csv
 *           output/Embeds.csv
 *           output/ExternalDomains.csv
 */

const fs = require('fs');
const path = require('path');
const https = require('https');
const { parse } = require('csv-parse/sync');
const { stringify } = require('csv-stringify/sync');

const inputCsv = path.resolve(__dirname, 'output', 'GSites_Inventory_Detailed.csv');
const outputDir = path.resolve(__dirname, 'output');
const accessToken = process.env.GCP_ACCESS_TOKEN || process.argv[2];
const maxSites = Number(process.env.MAX_SITES || 0);   // 0 = all
const siteOffset = Number(process.env.SITE_OFFSET || 0);
const CONCURRENCY = Number(process.env.CONCURRENCY || 10);  // parallel sites

if (!accessToken) {
  console.error('ERROR: No access token. Set GCP_ACCESS_TOKEN or pass as argument.');
  console.error('  GCP_ACCESS_TOKEN=<token> node 03b_api_extract_embeds.js');
  process.exit(1);
}
if (!fs.existsSync(inputCsv)) {
  console.error(`ERROR: Input not found: ${inputCsv}`);
  process.exit(1);
}

// ─── Helpers ─────────────────────────────────────────────────────────────────
const sleep = ms => new Promise(r => setTimeout(r, ms));

function readCsv(file) {
  return parse(fs.readFileSync(file, 'utf8'), { columns: true, skip_empty_lines: true });
}

// ─── Incremental CSV writers (no in-memory buffering) ────────────────────────
// Write column headers once at startup, then append rows per site as they finish.
// Node.js is single-threaded so concurrent appendFileSync calls never interleave.
const OUTPUT_FILES = {
  pages: path.join(outputDir, 'Pages.csv'),
  embeds: path.join(outputDir, 'Embeds.csv'),
  domains: path.join(outputDir, 'ExternalDomains.csv'),
};
const CSV_COLS = {
  pages: ['SiteId', 'SiteName', 'SiteUrl', 'PageUrl', 'PageTitle', 'Depth', 'InternalLinksDiscovered', 'EmbedCount', 'CrawlStatus'],
  embeds: ['SiteId', 'SiteName', 'SiteUrl', 'PageUrl', 'PageTitle', 'Depth', 'ItemKind', 'ArtifactType', 'ArtifactUrl', 'AnchorText'],
  domains: ['SiteId', 'SiteName', 'PageUrl', 'ExternalDomain'],
};
function initOutputFiles() {
  for (const [key, file] of Object.entries(OUTPUT_FILES)) {
    fs.writeFileSync(file, CSV_COLS[key].join(',') + '\n', 'utf8');
  }
}
function appendRows(fileKey, rows) {
  if (!rows.length) return;
  fs.appendFileSync(OUTPUT_FILES[fileKey], stringify(rows, { header: false, columns: CSV_COLS[fileKey] }), 'utf8');
}

// ─── HTTP with exponential-backoff retry (replaces fixed 150ms sleep) ────────
// Only pauses when the API actually rate-limits (HTTP 429/503). Zero wait otherwise.
async function httpsGet(url, attempt = 0) {
  const { status, body } = await new Promise((resolve, reject) => {
    const opts = { headers: { Authorization: `Bearer ${accessToken}`, Accept: 'application/json' } };
    https.get(url, opts, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        let body;
        try { body = JSON.parse(data); } catch { body = data; }
        resolve({ status: res.statusCode, body });
      });
    }).on('error', reject);
  });

  if ((status === 429 || status === 503) && attempt < 5) {
    const delay = Math.min(1000 * Math.pow(2, attempt), 32000); // 1s→2s→4s→8s→16s→32s
    console.warn(`  ⏳ Rate limited (HTTP ${status}), retry ${attempt + 1}/5 in ${delay}ms`);
    await sleep(delay);
    return httpsGet(url, attempt + 1);
  }
  return { status, body };
}

// ─── Concurrency pool — runs fn over items with at most N in-flight at once ──
async function runConcurrent(items, concurrency, fn) {
  let idx = 0;
  const workers = Array.from({ length: Math.min(concurrency, items.length) }, async () => {
    while (idx < items.length) {
      await fn(items[idx++]);
    }
  });
  await Promise.all(workers);
}

// ─── URL classification (mirrors 03_crawl_sites.js) ──────────────────────────
function classifyUrl(url) {
  const u = (url || '').toLowerCase();
  if (u.includes('docs.google.com/spreadsheets/')) return 'Sheet';
  if (u.includes('docs.google.com/forms/') || u.includes('forms.gle/')) return 'Form';
  if (u.includes('script.google.com/macros/s/')) return 'AppsScriptWebApp';
  if (u.includes('youtube.com/embed/') || u.includes('youtu.be/')) return 'YouTube';
  if (u.includes('google.com/maps') || u.includes('maps.google.')) return 'Maps';
  if (u.includes('drive.google.com/')) return 'DriveFile';
  if (u.includes('docs.google.com/document/')) return 'GoogleDoc';
  if (u.includes('docs.google.com/presentation/')) return 'GoogleSlides';
  return 'Other';
}

// ─── Drive MIME type → ArtifactType + canonical URL ─────────────────────────
function driveInfo(fileId, mimeType) {
  const m = (mimeType || '').toLowerCase();
  if (m.includes('spreadsheet')) return { type: 'Sheet', url: `https://docs.google.com/spreadsheets/d/${fileId}` };
  if (m.includes('form')) return { type: 'Form', url: `https://docs.google.com/forms/d/${fileId}` };
  if (m.includes('document')) return { type: 'GoogleDoc', url: `https://docs.google.com/document/d/${fileId}` };
  if (m.includes('presentation')) return { type: 'GoogleSlides', url: `https://docs.google.com/presentation/d/${fileId}` };
  if (m.includes('script')) return { type: 'AppsScriptWebApp', url: `https://script.google.com/d/${fileId}` };
  return { type: 'DriveFile', url: `https://drive.google.com/file/d/${fileId}` };
}

// ─── Sites API — list all pages (handles pagination) ─────────────────────────
async function listPages(siteId) {
  const pages = [];
  let url = `https://sites.googleapis.com/v1/sites/${siteId}/pages`;
  while (url) {
    const { status, body } = await httpsGet(url); // retry/backoff handled inside httpsGet
    if (status !== 200) {
      console.warn(`  ⚠ listPages HTTP ${status} for ${siteId}: ${JSON.stringify(body).slice(0, 120)}`);
      break;
    }
    if (Array.isArray(body.pages)) pages.push(...body.pages);
    url = body.nextPageToken
      ? `https://sites.googleapis.com/v1/sites/${siteId}/pages?pageToken=${encodeURIComponent(body.nextPageToken)}`
      : null;
  }
  return pages;
}

// ─── Recursive element walker ─────────────────────────────────────────────────
// Walks the full pageElements tree and appends any embeds found to `results`.
function walkElement(node, results) {
  if (!node || typeof node !== 'object') return;
  if (Array.isArray(node)) { node.forEach(n => walkElement(n, results)); return; }

  // Embedded Drive file or web address (iframe-style)
  if (node.embeddedItem) {
    const ei = node.embeddedItem;
    if (ei.embedType === 'GOOGLE_DRIVE' && ei.driveItem) {
      const rawId = ei.driveItem.driveFile || ei.driveItem.name || '';
      const fileId = rawId.replace(/^drivefiles\//, '');
      if (fileId) {
        const { type, url } = driveInfo(fileId, ei.driveItem.mimeType || '');
        results.push({ itemKind: 'iframe', artifactType: type, artifactUrl: url, anchorText: ei.driveItem.title || '' });
      }
    } else if (ei.chosenUrl) {
      results.push({ itemKind: 'iframe', artifactType: classifyUrl(ei.chosenUrl), artifactUrl: ei.chosenUrl, anchorText: '' });
    }
  }

  // Image widget (Drive-hosted or external URL)
  if (node.image) {
    const img = node.image;
    if (img.driveItem) {
      const rawId = img.driveItem.driveFile || img.driveItem.name || '';
      const fileId = rawId.replace(/^drivefiles\//, '');
      if (fileId) {
        results.push({
          itemKind: 'image', artifactType: 'DriveFile',
          artifactUrl: `https://drive.google.com/file/d/${fileId}`, anchorText: img.driveItem.title || ''
        });
      }
    } else if (img.sourceUrl) {
      results.push({ itemKind: 'image', artifactType: classifyUrl(img.sourceUrl), artifactUrl: img.sourceUrl, anchorText: '' });
    }
  }

  // Text widget — extract hyperlinks from textRun elements
  if (node.text && Array.isArray(node.text.paragraphs)) {
    for (const para of node.text.paragraphs) {
      for (const el of (para.elements || [])) {
        const tr = el.textRun;
        if (tr && tr.link && tr.link.url) {
          results.push({
            itemKind: 'link', artifactType: classifyUrl(tr.link.url),
            artifactUrl: tr.link.url, anchorText: (tr.content || '').trim()
          });
        }
      }
    }
  }

  // Recurse into all child object properties to handle any nesting depth
  for (const val of Object.values(node)) {
    if (val && typeof val === 'object') walkElement(val, results);
  }
}

// ─── Per-site processor (called concurrently) ────────────────────────────────
// Writes its results straight to disk — no global arrays, no memory pressure.
let sitesProcessed = 0;

async function processSite(site, totalInRun) {
  const siteUrl = site.EditUrl;

  let pages = [];
  try {
    pages = await listPages(site.SiteId);
  } catch (err) {
    appendRows('pages', [{
      SiteId: site.SiteId, SiteName: site.SiteName, SiteUrl: siteUrl,
      PageUrl: siteUrl, PageTitle: site.SiteName, Depth: 0,
      InternalLinksDiscovered: 0, EmbedCount: 0,
      CrawlStatus: `API-Error: ${err.message}`
    }]);
    sitesProcessed++;
    return;
  }

  if (pages.length === 0) {
    appendRows('pages', [{
      SiteId: site.SiteId, SiteName: site.SiteName, SiteUrl: siteUrl,
      PageUrl: siteUrl, PageTitle: site.SiteName, Depth: 0,
      InternalLinksDiscovered: 0, EmbedCount: 0, CrawlStatus: 'API-NoPagesFound'
    }]);
    sitesProcessed++;
    return;
  }

  const pageRows = [];
  const embedRows = [];
  const domainRows = [];

  for (let pi = 0; pi < pages.length; pi++) {
    const pg = pages[pi];
    const pageId = (pg.name || '').split('/').pop();
    const pageTitle = pg.title || '(untitled)';
    const pageUrl = siteUrl.replace(/\/$/, '') + '/p/' + pageId;

    const rawEmbeds = [];
    walkElement(pg.pageElements || pg, rawEmbeds);

    // Deduplicate by itemKind + artifactUrl
    const seen = new Set();
    const deduped = rawEmbeds.filter(e => {
      const key = `${e.itemKind}|${e.artifactUrl}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });

    let internalLinks = 0;
    const externalDomains = new Set();

    for (const embed of deduped) {
      if (embed.itemKind === 'link' && embed.artifactUrl.includes('sites.google.com')) {
        internalLinks++;
      } else {
        try { externalDomains.add(new URL(embed.artifactUrl).host.toLowerCase()); } catch { }
      }

      if (embed.artifactType !== 'Other' || embed.itemKind === 'iframe' || embed.itemKind === 'image') {
        embedRows.push({
          SiteId: site.SiteId, SiteName: site.SiteName, SiteUrl: siteUrl,
          PageUrl: pageUrl, PageTitle: pageTitle, Depth: pi === 0 ? 0 : 1,
          ItemKind: embed.itemKind, ArtifactType: embed.artifactType,
          ArtifactUrl: embed.artifactUrl, AnchorText: embed.anchorText
        });
      }
    }

    for (const domain of externalDomains) {
      domainRows.push({ SiteId: site.SiteId, SiteName: site.SiteName, PageUrl: pageUrl, ExternalDomain: domain });
    }

    const embedCount = deduped.filter(e => e.itemKind === 'iframe' || e.artifactType !== 'Other').length;
    pageRows.push({
      SiteId: site.SiteId, SiteName: site.SiteName, SiteUrl: siteUrl,
      PageUrl: pageUrl, PageTitle: pageTitle, Depth: pi === 0 ? 0 : 1,
      InternalLinksDiscovered: internalLinks, EmbedCount: embedCount,
      CrawlStatus: 'API'
    });
  }

  // Append to disk immediately — safe because Node.js is single-threaded
  appendRows('pages', pageRows);
  appendRows('embeds', embedRows);
  appendRows('domains', domainRows);

  sitesProcessed++;
  // Progress log every 100 sites (avoids flooding the console for 100k sites)
  if (sitesProcessed % 100 === 0 || sitesProcessed === totalInRun) {
    const pct = ((sitesProcessed / totalInRun) * 100).toFixed(1);
    console.log(`Progress: ${sitesProcessed}/${totalInRun} sites (${pct}%) — ${pages.length} page(s) for "${site.SiteName}"`);
  }
}

// ─── Main ─────────────────────────────────────────────────────────────────────
(async () => {
  fs.mkdirSync(outputDir, { recursive: true });

  // Read site inventory
  let sitesData = readCsv(inputCsv).map(r => ({
    SiteId: r.id || r.SiteId || '',
    SiteName: r.name || r.SiteName || '',
    EditUrl: r.webViewLink || r.webviewlink || ''
  })).filter(r => r.SiteId);

  // Apply offset / batch limit (same contract as 03_crawl_sites.js)
  const totalSites = sitesData.length;
  if (siteOffset > 0 || maxSites > 0) {
    const start = Math.min(siteOffset, totalSites);
    const end = maxSites > 0 ? Math.min(start + maxSites, totalSites) : totalSites;
    sitesData = sitesData.slice(start, end);
    console.log(`Batch: processing sites ${start + 1}–${start + sitesData.length} of ${totalSites}`);
  } else {
    console.log(`Processing all ${totalSites} sites`);
  }

  console.log(`Concurrency : ${CONCURRENCY} parallel sites`);
  console.log(`Starting at : ${new Date().toISOString()}`);

  // Create output files with headers before any concurrent writes
  initOutputFiles();

  const startTime = Date.now();
  await runConcurrent(sitesData, CONCURRENCY, site => processSite(site, sitesData.length));
  const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);

  console.log('\n========================================');
  console.log('  API Embed Extraction Complete');
  console.log('========================================');
  console.log(`Sites processed : ${sitesData.length}`);
  console.log(`Elapsed time    : ${elapsed}s`);
  console.log(`Output folder   : ${outputDir}`);
  console.log('Output files    : Pages.csv, Embeds.csv, ExternalDomains.csv');
})();
