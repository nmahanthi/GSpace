const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const { parse } = require('csv-parse/sync');
const { stringify } = require('csv-stringify/sync');

const arg = process.argv[2] || '';
const outputDir = path.resolve(__dirname, 'output');
const authFile = path.resolve(__dirname, '.auth', 'state.json');
const maxPagesPerSite = Number(process.argv[3] || process.env.MAX_PAGES_PER_SITE || 200);

if (!fs.existsSync(authFile)) throw new Error(`Auth file not found: ${authFile}. Run "node Save-GoogleAuth.js" first to capture a Google session.`);

function readCsv(file) {
  return parse(fs.readFileSync(file, 'utf8'), { columns: true, skip_empty_lines: true });
}
function extractSiteIdFromUrl(url) {
  const m = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : '';
}
function writeCsv(file, rows) { fs.writeFileSync(file, stringify(rows, { header: true }), 'utf8'); }

function classifyUrl(url) {
  const u = (url || '').toLowerCase();
  if (u.includes('youtube.com/embed/') || u.includes('youtube.com/watch') || u.includes('youtu.be/')) return 'YouTube';
  if (u.includes('google.com/maps') || u.includes('maps.google.')) return 'Maps';
  if (u.includes('drive.google.com/')) return 'DriveFile';
  if (u.includes('docs.google.com/document/')) return 'GoogleDoc';
  if (u.includes('docs.google.com/presentation/')) return 'GoogleSlides';
  if (u.includes('docs.google.com/spreadsheets/')) return 'Sheet';
  if (u.includes('docs.google.com/forms/') || u.includes('forms.gle/')) return 'Form';
  if (u.includes('calendar.google.com') || u.includes('google.com/calendar')) return 'Calendar';
  if (u.includes('datastudio.google.com') || u.includes('lookerstudio.google.com')) return 'DataStudio';
  if (u.includes('script.google.com/macros/s/')) return 'AppsScriptWebApp';
  return 'Other';
}

function normalizeUrl(url) {
  try { const u = new URL(url); u.hash = ''; return u.toString(); } catch { return null; }
}

// Build a scope descriptor so we can tell whether a link belongs to the same site.
// Handles: /d/<siteId>/... (editor/preview), /view/<siteId>/... and
// published sites (sites.google.com/<domain>/<name>/...).
function buildSiteScope(siteUrl, siteId) {
  if (siteId) return { type: 'id', siteId };
  try {
    const u = new URL(siteUrl);
    const parts = u.pathname.split('/').filter(Boolean);
    return { type: 'path', origin: u.origin, prefix: '/' + parts.slice(0, 2).join('/') + '/' };
  } catch { return { type: 'none' }; }
}

// Returns true when `url` is a subpage of the same Google Site.
function isInternalPage(scope, url) {
  try {
    const u = new URL(url);
    if (u.host !== 'sites.google.com') return false;
    if (scope.type === 'id')
      return u.pathname.includes(`/d/${scope.siteId}/`) || u.pathname.includes(`/view/${scope.siteId}/`);
    if (scope.type === 'path')
      return u.origin === scope.origin && u.pathname.startsWith(scope.prefix);
    return false;
  } catch { return false; }
}

// Collect every <a href> link visible in the live page (incl. shadow DOM),
// resolve relative hrefs against baseUrl, strip fragments, and drop
// non-page targets (assets, embeds, mailto, etc.).
async function extractLinks(page, baseUrl) {
  const raw = await page.evaluate(() => {
    const hrefs = new Set();
    function scan(root) {
      root.querySelectorAll('a[href]').forEach(a => {
        const href = a.getAttribute('href');
        if (href) hrefs.add(href);
      });
      root.querySelectorAll('*').forEach(el => { if (el.shadowRoot) scan(el.shadowRoot); });
    }
    scan(document);
    return [...hrefs];
  });

  const links = new Set();
  for (const href of raw) {
    if (!href || href.startsWith('#') || /^(mailto|tel|javascript):/i.test(href)) continue;
    try {
      const resolved = new URL(href, baseUrl);
      resolved.hash = '';
      const p = resolved.pathname;
      if (/\.(png|jpe?g|gif|svg|webp|ico|css|js|woff2?|ttf|pdf|zip)$/i.test(p)) continue;
      if (/\/(embed|_\/|viewer|export|thumbnail|uc)\b/i.test(p)) continue;
      links.add(resolved.toString());
    } catch { /* ignore malformed hrefs */ }
  }
  return links;
}

async function extractEmbeds(page) {
  return await page.evaluate(() => {
    const results = [];
    const add = (kind, url, ctx = '') => { if (url) results.push({ kind, url: url.trim(), context: ctx.substring(0, 300) }); };

    function scan(root) {
      const iframes = root.querySelectorAll('iframe');
      for (const f of iframes) {
        const src = f.src || f.getAttribute('data-src') || f.getAttribute('srcdoc') || '';
        if (src) add('iframe', src, f.outerHTML);
      }
      const embeds = root.querySelectorAll('embed[src], object[data], video[src], audio[src], source[src]');
      for (const e of embeds) {
        const url = e.src || e.getAttribute('data') || e.getAttribute('src') || '';
        if (url) add(e.tagName.toLowerCase(), url, e.outerHTML);
      }
      const dataEls = root.querySelectorAll('[data-url], [data-src], [data-embed-url], [data-href]');
      for (const e of dataEls) {
        const url = e.getAttribute('data-url') || e.getAttribute('data-src') || e.getAttribute('data-embed-url') || e.getAttribute('data-href') || '';
        if (url && url.startsWith('http')) add('data-embed', url, e.outerHTML);
      }
      // Search for YouTube patterns in text/HTML
      const html = root.innerHTML || '';
      const ytMatches = html.match(/(?:https?:\/\/)?(?:www\.)?(?:youtube\.com\/watch\?v=|youtube\.com\/embed\/|youtu\.be\/)([a-zA-Z0-9_-]{11})/g);
      if (ytMatches) ytMatches.forEach(m => add('youtube-pattern', m, ''));
      // Shadow DOM
      const all = root.querySelectorAll('*');
      for (const el of all) {
        if (el.shadowRoot) scan(el.shadowRoot);
      }
    }
    scan(document);
    return results;
  });
}

(async () => {
  fs.mkdirSync(outputDir, { recursive: true });
  fs.mkdirSync(path.join(outputDir, 'html'), { recursive: true });

  let sites = [];
  if (arg.startsWith('http')) {
    const siteId = extractSiteIdFromUrl(arg);
    if (!siteId) throw new Error(`Could not extract site ID from URL: ${arg}`);
    sites = [{ SiteId: siteId, SiteName: siteId, SiteUrl: arg }];
    console.log(`Using direct site URL: ${arg} (siteId=${siteId})`);
  } else {
    const inputCsv = arg || path.resolve(__dirname, 'output', '02_GSites_Inventory_Detailed.csv');
    if (!fs.existsSync(inputCsv)) throw new Error(`Input CSV not found: ${inputCsv}`);
    sites = readCsv(inputCsv).map(r => {
      // Support both the simple SelectedSites.csv format (SiteUrl, SiteName)
      // and the full gam7 inventory format (id/SiteId, name/SiteName, webViewLink).
      const siteUrl = (r.SiteUrl || r.webViewLink || r.webviewlink || '').trim();
      const siteId  = (r.SiteId  || r.id  || extractSiteIdFromUrl(siteUrl) || '').trim();
      const siteName = (r.SiteName || r.name || siteUrl).trim();
      return { SiteId: siteId, SiteName: siteName, SiteUrl: siteUrl };
    }).filter(r => {
      if (!r.SiteUrl) { console.warn(`  Skipping row with no SiteUrl: ${JSON.stringify(r)}`); return false; }
      return true;
    });
    console.log(`Loaded ${sites.length} site(s) from: ${inputCsv}`);
  }

  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ storageState: authFile });

  const pagesOut = [], embedsOut = [], externalDomainsOut = [];

  try {
    for (const site of sites) {
      console.log(`Crawling site: ${site.SiteName} | ${site.SiteUrl}`);
      const siteScope = buildSiteScope(site.SiteUrl, site.SiteId);
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
          // Scroll to trigger lazy loading
          for (let i = 0; i < 5; i++) {
            await page.evaluate(() => window.scrollBy(0, window.innerHeight));
            await page.waitForTimeout(1500);
          }
          // Wait a bit more for embeds
          await page.waitForTimeout(3000);

          const title = await page.title();
          const html = await page.content();
          pageCounter += 1;

          const htmlFile = `${site.SiteId}_${pageCounter}.html`.replace(/[^a-zA-Z0-9._-]/g, '_');
          fs.writeFileSync(path.join(outputDir, 'html', htmlFile), html, 'utf8');

          const discovered = await extractEmbeds(page);

          // Collect <a href> links from the live page and filter to same-site subpages.
          const pageLinks = await extractLinks(page, currentUrl);
          const internalLinks = new Set([...pageLinks].filter(u => isInternalPage(siteScope, u)));
          const externalDomains = new Set();
          let embedCount = 0;

          for (const item of discovered) {
            const normalized = normalizeUrl(item.url);
            if (!normalized) continue;
            const type = classifyUrl(normalized);
            let itemHost;
            try { itemHost = new URL(normalized).host.toLowerCase(); } catch { continue; }

            if (type !== 'Other' || item.kind === 'iframe' || item.kind === 'embed' || item.kind === 'object' || item.kind === 'youtube-pattern') {
              embedCount += 1;
              embedsOut.push({
                SiteId: site.SiteId, SiteName: site.SiteName, SiteUrl: site.SiteUrl,
                PageUrl: currentUrl, PageTitle: title, Depth: current.depth,
                ItemKind: item.kind, ArtifactType: type, ArtifactUrl: normalized, ContextHtml: item.context
              });
            }

            if (itemHost !== 'sites.google.com') {
              externalDomains.add(itemHost);
            }
          }

          // Enqueue unvisited subpages discovered on this page.
          let newLinks = 0;
          for (const nextUrl of internalLinks) {
            if (!visited.has(nextUrl)) { queue.push({ url: nextUrl, depth: current.depth + 1 }); newLinks++; }
          }

          for (const domain of externalDomains) externalDomainsOut.push({ SiteId: site.SiteId, SiteName: site.SiteName, PageUrl: currentUrl, ExternalDomain: domain });

          pagesOut.push({
            SiteId: site.SiteId, SiteName: site.SiteName, SiteUrl: site.SiteUrl,
            PageUrl: currentUrl, PageTitle: title, Depth: current.depth,
            InternalLinksDiscovered: internalLinks.size, EmbedCount: embedCount,
            HtmlSnapshot: htmlFile, CrawlStatus: 'Success'
          });
          console.log(`  [depth=${current.depth}] Page ${pageCounter}: "${title}" | embeds=${embedCount} | links=${internalLinks.size} (${newLinks} new) | queue=${queue.length}`);
        } catch (err) {
          console.error(`  ERROR on ${currentUrl}: ${err.message || err}`);
          pagesOut.push({
            SiteId: site.SiteId, SiteName: site.SiteName, SiteUrl: site.SiteUrl,
            PageUrl: currentUrl, PageTitle: '', Depth: current.depth,
            InternalLinksDiscovered: 0, EmbedCount: 0, HtmlSnapshot: '',
            CrawlStatus: `Error: ${String(err.message || err)}`
          });
        } finally {
          await page.close();
        }
      }
    }
  } finally {
    // Always save whatever was collected — even if an error stopped the crawl early.
    await browser.close().catch(() => {});
    writeCsv(path.join(outputDir, '07_Pages_Enhanced.csv'), pagesOut);
    writeCsv(path.join(outputDir, '08_Embeds_Enhanced.csv'), embedsOut);
    writeCsv(path.join(outputDir, '09_ExternalDomains_Enhanced.csv'), externalDomainsOut);
    console.log('\nOutput saved to: ' + outputDir);
    console.log(`Pages: ${pagesOut.length} | Embeds: ${embedsOut.length} | Domains: ${externalDomainsOut.length}`);
  }
})();
