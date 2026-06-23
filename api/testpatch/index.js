/**
 * P124 — Invoicing Portal
 * Azure Function: /api/testtpatch  *** TEMPORARY — DELETE AFTER DIAGNOSIS ***
 *
 * Patches Billed_x003f_ = true on a single TT2 item and returns the full
 * Graph response (including any error) so we can diagnose why TT2 patching
 * is failing silently in issueinvoice.
 *
 * POST body (JSON):
 *   wipId   string   — TT2 SP list item ID to patch
 *
 * Returns: { status, body } from Graph
 */

const https   = require('https');
const { URL } = require('url');

const SITE_PATH = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';
const TT2       = '67db204c-30a5-4f4d-b276-60852d9967e1';

const ADMIN_EMAILS = ['toby@tmclegal.co.uk', 'danielle@tmclegal.co.uk'];

function getCallerEmail(req) {
  try {
    const header = req.headers && req.headers['x-ms-client-principal'];
    if (!header) return null;
    const decoded   = Buffer.from(header, 'base64').toString('utf8');
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
  context.log('P124 /api/testtpatch called');

  const callerEmail = getCallerEmail(req);
  if (!callerEmail || !ADMIN_EMAILS.includes(callerEmail)) {
    context.res = { status: 403, body: 'Forbidden — admin only.' };
    return;
  }

  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    context.res = { status: 500, body: 'Missing app settings.' };
    return;
  }

  let body;
  try {
    body = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
  } catch (e) {
    context.res = { status: 400, body: 'Invalid JSON.' };
    return;
  }

  const { wipId } = body || {};
  if (!wipId) {
    context.res = { status: 400, body: 'Missing wipId.' };
    return;
  }

  try {
    const token = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
    context.log('Token obtained. Patching TT2 item', wipId);

    const url = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
      + '/lists/' + TT2 + '/items/' + wipId + '/fields';

    context.log('PATCH URL:', url);

    const { status, responseBody } = await graphPatchRaw(url, token, { 'Billed_x003f_': true });

    context.log('Graph response status:', status);
    context.log('Graph response body:', responseBody);

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        wipId,
        graphStatus: status,
        graphResponse: responseBody,
        patchUrl: url,
      }),
    };

  } catch (err) {
    context.log.error('testtpatch error:', err.message);
    context.res = {
      status: 500,
      body: JSON.stringify({ error: err.message }),
    };
  }
};

function getToken(tenantId, clientId, clientSecret) {
  return new Promise(function (resolve, reject) {
    const body = new URLSearchParams({
      grant_type:    'client_credentials',
      client_id:     clientId,
      client_secret: clientSecret,
      scope:         'https://graph.microsoft.com/.default',
    }).toString();

    const options = {
      hostname: 'login.microsoftonline.com',
      path:     '/' + tenantId + '/oauth2/v2.0/token',
      method:   'POST',
      headers: {
        'Content-Type':   'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(body),
      },
    };

    const req = https.request(options, function (res) {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        try {
          const json = JSON.parse(data);
          if (json.access_token) resolve(json.access_token);
          else reject(new Error('Token error: ' + (json.error_description || data)));
        } catch (e) { reject(e); }
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

// Returns raw status + body rather than throwing, so we can see exactly what Graph says
function graphPatchRaw(url, token, payload) {
  return new Promise(function (resolve, reject) {
    const bodyStr = JSON.stringify(payload);
    const u = new URL(url);
    const options = {
      hostname: u.hostname,
      path:     u.pathname + u.search,
      method:   'PATCH',
      headers: {
        Authorization:    'Bearer ' + token,
        'Content-Type':   'application/json',
        'Content-Length': Buffer.byteLength(bodyStr),
      },
    };
    const req = https.request(options, function (res) {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => resolve({ status: res.statusCode, responseBody: data }));
    });
    req.on('error', reject);
    req.write(bodyStr);
    req.end();
  });
}
