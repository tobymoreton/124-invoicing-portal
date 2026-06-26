/**
 * P124 — Invoicing Portal
 * Azure Function: /api/aureli
 *
 * GET ?action=balance  — reads AureliBalance from PortalConfig SP list
 *                        returns { balance, formatted, asOf }
 *
 * Auth: @tmclegal.co.uk domain check (all methods)
 * Admin-only gate on write operations (future Phase 2+)
 *
 * PortalConfig list GUID: 6661b8ba-3f10-436b-b4de-88b370e8160b
 * SITE_PATH: tmcostings.sharepoint.com:/sites/TMCLegalLimited: (hardcoded)
 *
 * PortalConfig fields:
 *   Title  — config key (e.g. "AureliBalance")
 *   Value  — config value (Text — raw number string, e.g. "323799.36")
 *   Modified — SP auto-updates on every PATCH; used as asOf timestamp
 */

const https   = require('https');
const { URL } = require('url');

const PORTAL_CONFIG_GUID = '6661b8ba-3f10-436b-b4de-88b370e8160b';
const SITE_PATH          = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';
const ALLOWED_DOMAIN     = '@tmclegal.co.uk';
const FINANCE_ADMIN_EMAILS = ['toby@tmclegal.co.uk', 'danielle@tmclegal.co.uk'];

// Module-level cache — valid for the lifetime of this function instance
let _cache = null; // { balance, formatted, asOf, cachedAt }
const CACHE_TTL_MS = 5 * 60 * 1000; // 5 minutes

function getCallerEmail(req) {
  try {
    const header = req.headers && req.headers['x-ms-client-principal'];
    if (!header) return null;
    const decoded  = Buffer.from(header, 'base64').toString('utf8');
    const principal = JSON.parse(decoded);
    if (principal.userDetails) return principal.userDetails.toLowerCase();
    const claim = (principal.claims || []).find(
      c => c.typ === 'preferred_username' || c.typ === 'email' || c.typ === 'upn'
        || c.typ === 'http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress'
    );
    return claim ? claim.val.toLowerCase() : null;
  } catch { return null; }
}

module.exports = async function (context, req) {
  context.log('P124 /api/aureli called — method:', req.method, 'action:', req.query.action);

  // ── Auth ────────────────────────────────────────────────────────────────
  const callerEmail = getCallerEmail(req);
  if (!callerEmail || !FINANCE_ADMIN_EMAILS.includes(callerEmail)) {
    context.res = { status: 403, body: 'Forbidden — restricted to authorised users only.' };
    return;
  }

  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    context.res = { status: 500, body: 'Missing required app settings.' };
    return;
  }

  const action = (req.query.action || 'balance').toLowerCase();

  // ── GET balance ──────────────────────────────────────────────────────────
  if (req.method === 'GET' && action === 'balance') {
    try {
      // Serve from cache if fresh
      if (_cache && (Date.now() - _cache.cachedAt) < CACHE_TTL_MS) {
        context.res = {
          status: 200,
          headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-cache' },
          body: JSON.stringify({ ..._cache, stale: false }),
        };
        return;
      }

      const token = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);

      // Fetch PortalConfig items, filter to AureliBalance
      const url = `https://graph.microsoft.com/v1.0/sites/${SITE_PATH}/lists/${PORTAL_CONFIG_GUID}/items`
                + `?$expand=fields($select=Title,Value)`
                + `&$filter=fields/Title eq 'AureliBalance'`
                + `&$top=5`;

      const result = await graphGet(url, token);
      const items  = result.value || [];

      if (!items.length) {
        context.res = { status: 404, body: 'AureliBalance config item not found in PortalConfig list.' };
        return;
      }

      const item    = items[0];
      const raw     = item.fields?.Value || '0';
      const balance = parseFloat(raw) || 0;
      const asOf    = item.lastModifiedDateTime || new Date().toISOString();

      const formatted = '£' + balance.toLocaleString('en-GB', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      });

      _cache = { balance, formatted, asOf, cachedAt: Date.now() };

      context.res = {
        status: 200,
        headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-cache' },
        body: JSON.stringify({ balance, formatted, asOf, stale: false }),
      };
    } catch (err) {
      context.log.error('Error in /api/aureli GET balance:', err.message);
      // Return stale cache if available rather than a hard error
      if (_cache) {
        context.res = {
          status: 200,
          headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-cache' },
          body: JSON.stringify({ ..._cache, stale: true }),
        };
        return;
      }
      context.res = { status: 500, body: 'Error fetching balance: ' + err.message };
    }
    return;
  }

  context.res = { status: 400, body: 'Unknown action. Supported: action=balance' };
};

// ── Helpers ──────────────────────────────────────────────────────────────────

function getToken(tenantId, clientId, clientSecret) {
  return new Promise((resolve, reject) => {
    const body = new URLSearchParams({
      grant_type:    'client_credentials',
      client_id:     clientId,
      client_secret: clientSecret,
      scope:         'https://graph.microsoft.com/.default',
    }).toString();
    const options = {
      hostname: 'login.microsoftonline.com',
      path:     `/${tenantId}/oauth2/v2.0/token`,
      method:   'POST',
      headers: {
        'Content-Type':   'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(body),
      },
    };
    const req = https.request(options, res => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          const json = JSON.parse(data);
          if (json.access_token) resolve(json.access_token);
          else reject(new Error(`Token error: ${json.error_description || data}`));
        } catch (e) { reject(e); }
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

function graphGet(url, token) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    const req = https.request({
      hostname: u.hostname,
      path:     u.pathname + u.search,
      method:   'GET',
      headers: {
        Authorization: `Bearer ${token}`,
        Accept:        'application/json',
      },
    }, res => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        if (res.statusCode >= 400) {
          reject(new Error(`Graph GET ${res.statusCode}: ${data.slice(0, 300)}`));
          return;
        }
        try { resolve(JSON.parse(data)); } catch (e) { reject(new Error(`JSON parse: ${e.message}`)); }
      });
    });
    req.on('error', reject);
    req.end();
  });
}
