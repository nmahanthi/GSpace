const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const readline = require('readline');

const startUrl = process.argv[2] || 'https://sites.google.com/';
const authDir = path.resolve(__dirname, '.auth');
const authFile = path.join(authDir, 'state.json');

function waitForEnter(promptText) {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  return new Promise(resolve => rl.question(promptText, () => {
    rl.close();
    resolve();
  }));
}

(async () => {
  fs.mkdirSync(authDir, { recursive: true });

  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext();
  const page = await context.newPage();

  console.log(`Opening ${startUrl}`);
  await page.goto(startUrl, { waitUntil: 'domcontentloaded', timeout: 60000 });
  console.log('Sign in with the non-prod test account that has access to the target Google Sites.');
  console.log('After sign-in completes and a target site opens successfully, press Enter here.');

  await waitForEnter('Press Enter to save authenticated browser state... ');

  await context.storageState({ path: authFile });
  await browser.close();

  console.log(`Saved auth state to: ${authFile}`);
})();
