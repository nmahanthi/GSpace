const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const { parse } = require('csv-parse/sync');
const { stringify } = require('csv-stringify/sync');

const inputCsv = process.argv[2] || path.resolve(__dirname, 'output', '02_GSites_Inventory_Detailed.csv');
const publishedUrlsCsv = path.resolve(__dirname, 'output', '02a_Sites_Published_URLs.csv');
const outputDir = path.resolve(__dirname, 'output');
const authFile = path.resolve(__dirname, '.auth', 'state.json');
const maxPagesPerSite = Number(process.argv[3] || process.env.MAX_PAGES_PER_SITE || 200);

if (!fs.existsSync(inputCsv)) {
  throw new Error(`Input CSV not found: ${inputCsv}`);
}
if (!fs.existsSync(authFile)) {
  throw new Error(`Auth file not found: ${authFile}. Run 02_save_playwright_auth.js first.`);
}

function readCsv(file) {
  const raw = fs.readFileSync(file, 'utf8');
  return parse(raw, { columns: true, skip_empty_lines: true });
}

function writeCsv(file, rows) {
  const csv = stringify(rows, { header: true });
  fs.writeFileSync(file, csv, 'utf8');
}

function firstValue(row, keys) {
  for (const key of keys) {
    const exact = row[key];
    if (exact !== undefined && exact !== null && String(exact).trim() !== '') return String(exact).trim();
    const altKey = Object.keys(row).find(k => k.toLowerCase() === key.toLowerCase());
    if (altKey && row[altKey] !== undefined && row[altKey] !== null && String(row[altKey]).trim() !== '') {
      return String(row[altKey]).trim();
    }
  }
  return '';
}

function normalizeUrl(url) {
  try {
    const u = new URL(url);
    u.hash = '';
    return u.toString();
  } catch {
    return null;
  }
}

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

function sameSiteRoot(candidate, rootUrl) {
  try {
    const c = new URL(candidate);
    const r = new URL(rootUrl);
    return c.host === r.host && c.pathname.startsWith(r.pathname.replace(/\/$/, ''));
  } catch {
    return false;
  }
}

function sameHost(candidate, rootUrl) {
  try {
    return new URL(candidate).host === new URL(rootUrl).host;
  } catch {
    return false;
  }
}

(async () => {
  fs.mkdirSync(outputDir, { recursive: true });
  fs.mkdirSync(path.join(outputDir, 'html'), { recursive: true });

  // Read sites from main inventory
  const sitesData = readCsv(inputCsv).map(r => ({
    SiteId: firstValue(r, ['id', 'SiteId']),
    SiteName: firstValue(r, ['name', 'SiteName']),
    EditUrl: firstValue(r, ['webviewlink', 'webViewLink', 'alternateLink'])
  })).filter(r => r.SiteId && r.EditUrl);

  // Try to load published URLs if available
  let publishedUrlsMap = new Map();
  if (fs.existsSync(publishedUrlsCsv)) {
    console.log('Loading published URLs from:', publishedUrlsCsv);
    const publishedData = readCsv(publishedUrlsCsv);
    for (const row of publishedData) {
      const siteId = firstValue(row, ['SiteId', 'id']);
      const publishedUrl = firstValue(row, ['PublishedUrl', 'publishedUrl']);
      if (siteId && publishedUrl) {
        publishedUrlsMap.set(siteId, publishedUrl);
      }
    }
    console.log(`Loaded ${publishedUrlsMap.size} published URLs`);
  } else {
    console.log('⚠️  Published URLs file not found:', publishedUrlsCsv);
    console.log('⚠️  Will attempt to crawl using edit URLs (may result in 403 errors)');
    console.log('⚠️  Run: node 03a_get_published_urls.js <access_token> to get published URLs');
  }

  // Merge published URLs with site data
  const sites = sitesData.map(site => {
    const publishedUrl = publishedUrlsMap.get(site.SiteId);
    return {
      SiteId: site.SiteId,
      SiteName: site.SiteName,
      EditUrl: site.EditUrl,
      PublishedUrl: publishedUrl || null,
      // Use published URL if available, otherwise fall back to edit URL
      SiteUrl: publishedUrl || site.EditUrl,
      UrlType: publishedUrl ? 'Published' : 'Edit'
    };
  });

  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ storageState: authFile });

  const pagesOut = [];
  const embedsOut = [];
  const externalDomainsOut = [];
  const requestsOut = [];

  context.on('requestfinished', req => {
    try {
      requestsOut.push({
        Timestamp: new Date().toISOString(),
        Method: req.method(),
        Url: req.url(),
        ResourceType: req.resourceType()
      });
    } catch { }
  });

  for (const site of sites) {
    console.log(`Crawling site: ${site.SiteName} | ${site.SiteUrl} [${site.UrlType} URL]`);

    const visited = new Set();
    const queue = [{ url: site.SiteUrl, depth: 0 }];
    let pageCounter = 0;

    while (queue.length > 0 && pageCounter < maxPagesPerSite) {
      const current = queue.shift();
      const currentUrl = normalizeUrl(current.url);
      if (!currentUrl || visited.has(currentUrl)) continue;
      visited.add(currentUrl);

      const page = await context.newPage();
      try {
        await page.goto(currentUrl, { waitUntil: 'domcontentloaded', timeout: 60000 });
        await page.waitForLoadState('load', { timeout: 60000 });

        const title = await page.title();
        const html = await page.content();
        pageCounter += 1;

        const htmlFile = `${site.SiteId}_${pageCounter}.html`.replace(/[^a-zA-Z0-9._-]/g, '_');
        fs.writeFileSync(path.join(outputDir, 'html', htmlFile), html, 'utf8');

        const discovered = await page.evaluate(() => {
          const rows = [];
          const add = (kind, url, text) => {
            if (!url) return;
            rows.push({ kind, url, text: (text || '').trim() });
          };

          for (const a of document.querySelectorAll('a[href]')) add('link', a.href, a.textContent || '');
          for (const f of document.querySelectorAll('iframe[src]')) add('iframe', f.src, '');
          for (const i of document.querySelectorAll('img[src]')) add('image', i.src, i.alt || '');
          for (const e of document.querySelectorAll('embed[src], object[data], source[src]')) {
            add('embed', e.getAttribute('src') || e.getAttribute('data') || '', '');
          }
          return rows;
        });

        const internalLinks = [];
        const externalDomains = new Set();
        let embedCount = 0;

        for (const item of discovered) {
          const normalized = normalizeUrl(item.url);
          if (!normalized) continue;

          const type = classifyUrl(normalized);
          const isInternal = sameSiteRoot(normalized, site.SiteUrl) || sameHost(normalized, site.SiteUrl);

          if (item.kind === 'link' && isInternal && !visited.has(normalized)) {
            internalLinks.push(normalized);
          }

          if (type !== 'Other' || item.kind === 'iframe' || item.kind === 'embed') {
            embedCount += 1;
            embedsOut.push({
              SiteId: site.SiteId,
              SiteName: site.SiteName,
              SiteUrl: site.SiteUrl,
              PageUrl: currentUrl,
              PageTitle: title,
              Depth: current.depth,
              ItemKind: item.kind,
              ArtifactType: type,
              ArtifactUrl: normalized,
              AnchorText: item.text || ''
            });
          }

          if (!isInternal) {
            try {
              externalDomains.add(new URL(normalized).host.toLowerCase());
            } catch { }
          }
        }

        for (const nextUrl of [...new Set(internalLinks)]) {
          queue.push({ url: nextUrl, depth: current.depth + 1 });
        }

        for (const domain of externalDomains) {
          externalDomainsOut.push({
            SiteId: site.SiteId,
            SiteName: site.SiteName,
            PageUrl: currentUrl,
            ExternalDomain: domain
          });
        }

        pagesOut.push({
          SiteId: site.SiteId,
          SiteName: site.SiteName,
          SiteUrl: site.SiteUrl,
          PageUrl: currentUrl,
          PageTitle: title,
          Depth: current.depth,
          InternalLinksDiscovered: [...new Set(internalLinks)].length,
          EmbedCount: embedCount,
          HtmlSnapshot: htmlFile,
          CrawlStatus: 'Success'
        });
      } catch (err) {
        pagesOut.push({
          SiteId: site.SiteId,
          SiteName: site.SiteName,
          SiteUrl: site.SiteUrl,
          PageUrl: currentUrl,
          PageTitle: '',
          Depth: current.depth,
          InternalLinksDiscovered: 0,
          EmbedCount: 0,
          HtmlSnapshot: '',
          CrawlStatus: `Error: ${String(err.message || err)}`
        });
      } finally {
        await page.close();
      }
    }
  }

  await browser.close();

  writeCsv(path.join(outputDir, '07_Pages.csv'), pagesOut);
  writeCsv(path.join(outputDir, '08_Embeds.csv'), embedsOut);
  writeCsv(path.join(outputDir, '09_ExternalDomains.csv'), externalDomainsOut);
  writeCsv(path.join(outputDir, '10_NetworkRequests.csv'), requestsOut);

  console.log('Crawl complete. Outputs written to output folder.');
})();
