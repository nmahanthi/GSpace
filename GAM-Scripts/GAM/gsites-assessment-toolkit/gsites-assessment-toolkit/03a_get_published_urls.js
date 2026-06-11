/**
 * Get Published URLs for Google Sites using Sites API
 * 
 * This script fetches the published URLs for Google Sites which are needed
 * for crawling (edit URLs return 403 errors).
 * 
 * Input: 02_GSites_Inventory_Detailed.csv
 * Output: 02a_Sites_Published_URLs.csv
 */

const fs = require('fs');
const path = require('path');
const https = require('https');
const { parse } = require('csv-parse/sync');
const { stringify } = require('csv-stringify/sync');

const inputCsv = path.resolve(__dirname, 'output', '02_GSites_Inventory_Detailed.csv');
const outputCsv = path.resolve(__dirname, 'output', '02a_Sites_Published_URLs.csv');
const accessToken = process.env.GCP_ACCESS_TOKEN || process.argv[2];

if (!accessToken) {
  console.error('ERROR: No access token provided');
  console.error('Usage: node 03a_get_published_urls.js <access_token>');
  console.error('   OR: GCP_ACCESS_TOKEN=<token> node 03a_get_published_urls.js');
  console.error('');
  console.error('To get a token:');
  console.error('  gcloud auth print-access-token');
  process.exit(1);
}

if (!fs.existsSync(inputCsv)) {
  console.error(`ERROR: Input file not found: ${inputCsv}`);
  process.exit(1);
}

function readCsv(file) {
  const raw = fs.readFileSync(file, 'utf8');
  return parse(raw, { columns: true, skip_empty_lines: true });
}

function writeCsv(file, rows) {
  const csv = stringify(rows, { header: true });
  fs.writeFileSync(file, csv, 'utf8');
}

function httpsGet(url, headers) {
  return new Promise((resolve, reject) => {
    const options = {
      headers: headers || {}
    };
    
    https.get(url, options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          try {
            resolve(JSON.parse(data));
          } catch (e) {
            reject(new Error(`Failed to parse JSON: ${e.message}`));
          }
        } else {
          reject(new Error(`HTTP ${res.statusCode}: ${data}`));
        }
      });
    }).on('error', reject);
  });
}

async function getPublishedUrl(siteId, token) {
  const apiUrl = `https://sites.googleapis.com/v1/sites/${siteId}`;
  const headers = {
    'Authorization': `Bearer ${token}`,
    'Accept': 'application/json'
  };
  
  try {
    const response = await httpsGet(apiUrl, headers);
    return {
      success: true,
      publishedUrl: response.siteUrl || null,
      title: response.title || null,
      error: null
    };
  } catch (error) {
    return {
      success: false,
      publishedUrl: null,
      title: null,
      error: error.message
    };
  }
}

async function main() {
  console.log('Reading sites from CSV...');
  const sites = readCsv(inputCsv);
  console.log(`Found ${sites.length} sites`);
  
  const results = [];
  let successCount = 0;
  let failCount = 0;
  
  for (let i = 0; i < sites.length; i++) {
    const site = sites[i];
    const siteId = site.id || site.SiteId;
    const siteName = site.name || site.SiteName;
    const editUrl = site.webViewLink || site.webviewlink;
    
    if (!siteId) {
      console.log(`[${i + 1}/${sites.length}] SKIP: No site ID`);
      continue;
    }
    
    console.log(`[${i + 1}/${sites.length}] Fetching: ${siteName} (${siteId})`);
    
    const result = await getPublishedUrl(siteId, accessToken);
    
    if (result.success) {
      console.log(`  ✓ Published URL: ${result.publishedUrl || 'NOT PUBLISHED'}`);
      successCount++;
    } else {
      console.log(`  ✗ Error: ${result.error}`);
      failCount++;
    }
    
    results.push({
      SiteId: siteId,
      SiteName: siteName,
      EditUrl: editUrl,
      PublishedUrl: result.publishedUrl || '',
      SiteTitle: result.title || siteName,
      ApiSuccess: result.success ? 'Yes' : 'No',
      ApiError: result.error || ''
    });
    
    // Rate limiting: 10 requests per second max
    await new Promise(resolve => setTimeout(resolve, 150));
  }
  
  console.log('');
  console.log('Writing results to CSV...');
  writeCsv(outputCsv, results);
  
  console.log('');
  console.log('========================================');
  console.log('  Published URLs Retrieval Complete');
  console.log('========================================');
  console.log(`Total sites:     ${sites.length}`);
  console.log(`Success:         ${successCount}`);
  console.log(`Failed:          ${failCount}`);
  console.log(`Output file:     ${outputCsv}`);
  console.log('');
  
  if (failCount > 0) {
    console.log('⚠️  Some sites failed - check ApiError column in output CSV');
  }
  
  const publishedCount = results.filter(r => r.PublishedUrl).length;
  console.log(`Sites with published URLs: ${publishedCount}`);
  console.log(`Sites without published URLs: ${successCount - publishedCount}`);
}

main().catch(err => {
  console.error('FATAL ERROR:', err);
  process.exit(1);
});

