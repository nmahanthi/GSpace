/**
 * Save-GoogleAuth.js
 *
 * Opens a visible Chromium browser so you can sign in to Google interactively,
 * then saves the authenticated session to .auth/state.json.
 *
 * Run once before any crawl:
 *   node Save-GoogleAuth.js
 *
 * The saved session is reused by 03_crawl_sites_enhanced.js and
 * Extract-SiteEmbeds-Playwright.js so you do not need to sign in again
 * unless the session expires (~24 hours).
 */
const { chromium } = require('playwright');
const fs           = require('fs');
const path         = require('path');
const readline     = require('readline');

const authDir  = path.resolve(__dirname, '.auth');
const authFile = path.join(authDir, 'state.json');

(async () => {
  fs.mkdirSync(authDir, { recursive: true });

  console.log('Launching Chromium for Google sign-in...');
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext();
  const page    = await context.newPage();

  await page.goto('https://accounts.google.com/', { waitUntil: 'domcontentloaded', timeout: 60000 });

  console.log('\nSteps:');
  console.log('  1. Sign in to Google in the browser window that just opened.');
  console.log('  2. Navigate to one of the Google Sites you want to crawl.');
  console.log('  3. Once the site loads, return here and press Enter.\n');

  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  await new Promise(resolve => rl.question('Press Enter when signed in > ', () => { rl.close(); resolve(); }));

  await context.storageState({ path: authFile });
  console.log(`\nAuth state saved to: ${authFile}`);
  console.log('You can now run the crawl scripts. Re-run this script if auth expires.');
  await browser.close();
})();
