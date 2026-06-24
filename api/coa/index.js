/**
 * P124 — Invoicing Portal
 * Azure Function: /api/coa
 */

const https   = require('https');
const { URL } = require('url');

const TT2_GUID   = '67db204c-30a5-4f4d-b276-60852d9967e1';
const CASES_GUID = 'ae420bda-e550-499c-b337-90e4f33617c1';
const SITE_PATH  = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';

function getCallerEmail(req) {
  try {
    const header = req.headers && req.headers['x-ms-client-principal'];
    if (!header) return null;
    const decoded = Buffer.from(header, 'base64').toString('utf8');
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
  context.log('P124 /api/coa called');

  const callerEmail = getCallerEmail(req);
  if (!callerEmail) {
    context.res = { status: 403, body: 'Forbidden — you must be signed in.' };
    return;
  }

  const ref  = (req.query.ref  || '').trim();
  const mode = (req.query.mode || 'unbilled').toLowerCase();

  if (!ref) {
    context.res = { status: 400, body: 'Missing required param: ref' };
    return;
  }
  if (mode !== 'unbilled' && mode !== 'all') {
    context.res = { status: 400, body: 'Invalid mode — must be "unbilled" or "all"' };
    return;
  }

  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    context.res = { status: 500, body: 'Missing required app settings.' };
    return;
  }

  try {
    const token = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);

    const [entries, limit] = await Promise.all([
      fetchTT2Entries(token, ref, mode),
      fetchCaseLimit(token, ref),
    ]);

    const sum = entries.reduce((acc, entry) => {
      const f   = entry.fields || {};
      const amt = parseFloat(f['Num_BillableAmount_x00a3_']);
      const val = !isNaN(amt) && amt > 0
        ? amt
        : (parseFloat(f['TimeSpentMirror']) || 0) * (parseFloat(f['field_6']) || 0);
      return acc + val;
    }, 0);

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-cache' },
      body: JSON.stringify({
        sum:   Math.round(sum * 100) / 100,
        count: entries.length,
        limit: limit,
        mode:  mode,
      }),
    };
  } catch (err) {
    context.log.error('Error in /api/coa:', err.message);
    context.res = { status: 500, body: 'Error: ' + err.message };
  }
};

async function fetchTT2Entries(token, ref, mode) {
  const selectFields = [
    'field_16',
    'Billable_x003f_',
    'Billed_x003f_',
    'Num_BillableAmount_x00a3_',
    'TimeSpentMirror',
    'field_6',
  ].join(',');

  const base = `https://graph.microsoft.com/v1.0/sites/${SITE_PATH}/lists/${TT2_GUID}/items`
             + `?$expand=fields($select=${encodeURIComponent(selectFields)})`
             + `&$top=999`;

  let url = base;
  let all = [];

  while (url) {
    const page = await graphGet(url, token, false);
    all = all.concat(page.value || []);
    url = page['@odata.nextLink'] || null;
  }

  return all.filter(item => {
    const f = item.fields || {};
    if ((f['field_16'] || '').toString().trim() !== ref) return false;
    if (!f['Billable_x003f_']) return false;
    if (mode === 'unbilled' && f['Billed_x003f_']) return false;
    return true;
  });
}

async function fetchCaseLimit(token, ref) {
  const selectFields = 'Ourreference_x0028_text_x0029_,Limit_x0020_Costs_x0020_of_x0020';
  const base = `https://graph.microsoft.com/v1.0/sites/${SITE_PATH}/lists/${CASES_GUID}/items`
             + `?$expand=fields($select=${encodeURIComponent(selectFields)})`
             + `&$top=500`;

  let url = base;
  while (url) {
    const page = await graphGet(url, token, true);
    const match = (page.value || []).find(item => {
      const f = item.fields || {};
      return (f['Ourreference_x0028_text_x0029_'] || '').toString().trim() === ref;
    });
    if (match) {
      const val = parseFloat(match.fields && match.fields['Limit_x0020_Costs_x0020_of_x0020']);
      return isNaN(val) ? null : val;
    }
    url = page['@odata.nextLink'] || null;
  }
  return null;
}

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

function graphGet(url, token, allowThrottleable = false) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    const headers = {
      Authorization: `Bearer ${token}`,
      Accept:        'application/json',
    };
    if (allowThrottleable) {
      headers['Prefer'] = 'allowthrottleablequeries';
    }

    const options = {
      hostname: u.hostname,
      path:     u.pathname + u.search,
      method:   'GET',
      headers,
    };

    const req = https.request(options, res => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        if (res.statusCode >= 400) {
          reject(new Error(`Graph ${res.statusCode}: ${data.slice(0, 300)}`));
          return;
        }
        try { resolve(JSON.parse(data)); }
        catch (e) { reject(new Error(`JSON parse error: ${e.message}`)); }
      });
    });
    req.on('error', reject);
    req.end();
  });
}
