const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const url = process.argv[2];
if (!url) {
  console.error('Usage: node Extract-SiteEmbeds-Playwright.js <google-sites-url> [output.csv] [auth-state-path]');
  console.error('Example: node Extract-SiteEmbeds-Playwright.js "https://sites.google.com/censftmigsme.microsoft-int.com/ftcmigrationtestsite/resources"');
  process.exit(1);
}

const outputCsv = process.argv[3] || 'ExtractedEmbeds.csv';
const authFile = process.argv[4] || path.resolve(__dirname, '.auth', 'state.json');

(async () => {
  const hasAuth = fs.existsSync(authFile);
  console.log(`Launching browser (headless:true, auth=${hasAuth})...`);
  const browser = await chromium.launch({ headless: true });
  const context = hasAuth ? await browser.newContext({ storageState: authFile }) : await browser.newContext();
  const page = await context.newPage();

  console.log(`Navigating to: ${url}`);
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });

  console.log('Page loaded. If you see a sign-in prompt, sign in now.');
  console.log('Waiting 10 seconds for any sign-in / redirects...');
  await page.waitForTimeout(10000);

  console.log('Scrolling to trigger lazy-loaded embeds...');
  for (let i = 0; i < 6; i++) {
    await page.evaluate(() => window.scrollBy(0, window.innerHeight));
    await page.waitForTimeout(2000);
  }
  await page.waitForTimeout(3000);

  const embeds = await page.evaluate(() => {
    const results = [];
    const add = (kind, url, ctx = '') => {
      if (!url) return;
      const trimmed = url.trim();
      if (!trimmed.startsWith('http')) return;
      // Skip Google auth / tracking frames
      if (trimmed.includes('accounts.google.com') || trimmed.includes('bscframe') || trimmed.includes('recaptcha')) return;
      results.push({ kind, url: trimmed, context: ctx.substring(0, 200) });
    };

    function scan(root) {
      root.querySelectorAll('iframe').forEach(f => {
        add('iframe', f.src || f.getAttribute('data-src') || f.getAttribute('srcdoc'), f.outerHTML);
      });
      root.querySelectorAll('embed[src], object[data], video[src], audio[src], source[src]').forEach(e => {
        add(e.tagName.toLowerCase(), e.src || e.getAttribute('data') || e.getAttribute('src'), e.outerHTML);
      });
      root.querySelectorAll('[data-url], [data-src], [data-embed-url]').forEach(e => {
        add('data-embed',
          e.getAttribute('data-url') || e.getAttribute('data-src') || e.getAttribute('data-embed-url'),
          e.outerHTML);
      });
      // Recurse shadow DOM
      root.querySelectorAll('*').forEach(el => { if (el.shadowRoot) scan(el.shadowRoot); });
    }
    scan(document);

    // Regex search for YouTube / Maps / Drive patterns in full HTML
    const html = document.documentElement.innerHTML;
    const yt = html.match(/(?:https?:\/\/)?(?:www\.)?(?:youtube\.com\/watch\?v=|youtube\.com\/embed\/|youtu\.be\/)([a-zA-Z0-9_-]{11})/g);
    if (yt) yt.forEach(m => add('youtube-pattern', m));
    const maps = html.match(/(?:https?:\/\/)?(?:www\.)?google\.com\/maps\/embed[^"'\s]*/g);
    if (maps) maps.forEach(m => add('maps-pattern', m));
    const drive = html.match(/(?:https?:\/\/)?drive\.google\.com\/file\/d\/[^"'\s]+/g);
    if (drive) drive.forEach(m => add('drive-pattern', m));

    return results;
  });

  // Deduplicate
  const unique = [];
  const seen = new Set();
  for (const e of embeds) {
    const key = e.kind + '|' + e.url;
    if (!seen.has(key)) { seen.add(key); unique.push(e); }
  }

  console.log(`\nFound ${unique.length} unique embed(s):`);
  console.table(unique.map(e => ({ Kind: e.kind, URL: e.url.substring(0, 80) + (e.url.length > 80 ? '...' : '') })));

  // Write CSV
  const header = 'Kind,EmbedUrl,Context';
  const rows = unique.map(e => `"${e.kind}","${e.url.replace(/"/g, '""')}","${(e.context || '').replace(/"/g, '""').substring(0, 100)}"`);
  fs.writeFileSync(outputCsv, [header, ...rows].join('\n'), 'utf8');
  console.log(`\nSaved to: ${path.resolve(outputCsv)}`);

  await browser.close();
})();
