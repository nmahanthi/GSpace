/**
 * get_service_account_token.js
 *
 * Mints a short-lived OAuth2 access token via domain-wide delegation (DWD),
 * using the SAME service account key file GAM itself uses (oauth2service.json)
 * to impersonate a Workspace user/admin (the "sub" JWT claim).
 *
 * Why this exists: `gcloud auth print-access-token` uses gcloud's own OAuth
 * client, whose scope list is hardcoded by Google and can NEVER include
 * Sites API scopes (sites.readonly), no matter how you re-authenticate. A
 * service-account JWT can request ANY scope that the service account's
 * Client ID has been explicitly authorized for in Admin Console > Security >
 * API controls > Domain-wide delegation - which is the only way to get a
 * token the Sites API v1 will accept.
 *
 * Usage:
 *   node get_service_account_token.js <oauth2service.json> <impersonate-email> [scopes]
 *   OAUTH2SERVICE_JSON=<path> IMPERSONATE_EMAIL=<email> node get_service_account_token.js
 *
 * Prints ONLY the access token to stdout on success (nothing else) so it can
 * be captured cleanly by the calling shell. All diagnostics go to stderr.
 */

const fs = require('fs');
const crypto = require('crypto');
const https = require('https');

const DEFAULT_SCOPES = [
  'https://www.googleapis.com/auth/sites.readonly',
  'https://www.googleapis.com/auth/drive.readonly'
].join(' ');

const keyPath = process.env.OAUTH2SERVICE_JSON || process.argv[2];
const impersonateEmail = process.env.IMPERSONATE_EMAIL || process.argv[3];
const scopes = process.env.SA_SCOPES || process.argv[4] || DEFAULT_SCOPES;

if (!keyPath || !impersonateEmail) {
  console.error('ERROR: Missing required arguments.');
  console.error('Usage: node get_service_account_token.js <oauth2service.json> <impersonate-email> [scopes]');
  process.exit(1);
}

if (!fs.existsSync(keyPath)) {
  console.error(`ERROR: Service account key file not found: ${keyPath}`);
  process.exit(1);
}

let key;
try {
  key = JSON.parse(fs.readFileSync(keyPath, 'utf8'));
} catch (e) {
  console.error(`ERROR: Failed to parse ${keyPath}: ${e.message}`);
  process.exit(1);
}

if (!key.client_email || !key.private_key) {
  console.error(`ERROR: ${keyPath} is missing client_email or private_key.`);
  console.error('  This script requires a classic service account key with an RSA');
  console.error('  private_key - not GAM\'s VM-metadata-based attached service account mode.');
  process.exit(1);
}

function base64url(buf) {
  return Buffer.from(buf).toString('base64')
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

function buildJwt() {
  const now = Math.floor(Date.now() / 1000);
  const header = { alg: 'RS256', typ: 'JWT' };
  const claims = {
    iss: key.client_email,
    scope: scopes,
    aud: 'https://oauth2.googleapis.com/token',
    sub: impersonateEmail, // impersonated user - required for domain-wide delegation
    iat: now,
    exp: now + 3600
  };

  const unsigned = `${base64url(JSON.stringify(header))}.${base64url(JSON.stringify(claims))}`;
  const signer = crypto.createSign('RSA-SHA256');
  signer.update(unsigned);
  signer.end();
  const signature = signer.sign(key.private_key)
    .toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

  return `${unsigned}.${signature}`;
}

function exchangeJwtForToken(jwt) {
  return new Promise((resolve, reject) => {
    const body = `grant_type=${encodeURIComponent('urn:ietf:params:oauth:grant-type:jwt-bearer')}&assertion=${encodeURIComponent(jwt)}`;

    const req = https.request('https://oauth2.googleapis.com/token', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(body)
      }
    }, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        let parsed;
        try {
          parsed = JSON.parse(data);
        } catch (e) {
          return reject(new Error(`Failed to parse token response: ${data}`));
        }
        if (res.statusCode !== 200) {
          return reject(new Error(
            `Token exchange failed (HTTP ${res.statusCode}): ${data}\n` +
            `  Common cause: service account Client ID (${key.client_id || 'unknown'}) is not\n` +
            `  authorized in Admin Console > Security > API controls > Domain-wide delegation\n` +
            `  for scope(s): ${scopes}`
          ));
        }
        resolve(parsed.access_token);
      });
    });

    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

(async () => {
  try {
    const token = await exchangeJwtForToken(buildJwt());
    if (!token) {
      console.error('ERROR: No access_token in response.');
      process.exit(1);
    }
    process.stdout.write(token);
  } catch (err) {
    console.error(`ERROR: ${err.message}`);
    process.exit(1);
  }
})();
