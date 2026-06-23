/**
 * P124 — Invoicing Portal
 * Azure Function: /api/caseupdate
 *
 * PATCHes fields on a Cases list item via Graph (app-only client-credentials auth).
 * All authenticated users can call this (draftsmen, finance, admin).
 *
 * POST body (JSON):
 *   itemId   string   — SP list item ID of the Cases list item
 *   fields   object   — key/value pairs to PATCH (InternalName: value)
 *
 * Returns: { updated: true }
 *
 * List: Cases
 * GUID: ae420bda-e550-499c-b337-90e4f33617c1
 */

const https   = require('https');
const { URL } = require('url');

const SITE_PATH  = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';
const CASES_LIST = 'ae420bda-e550-499c-b337-90e4f33617c1';

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
  context.log('P124 /api/caseupdate called');

  const callerEmail = getCallerEmail(req);
  if (!callerEmail) {
    context.res = { status: 403, body: 'Forbidden — you must be signed in.' };
    return;
  }

  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    context.res = { status: 500, body: 'Missing required app settings.' };
    return;
  }

  let body;
  try {
    body = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
  } catch {
    context.res = { status: 400, body: 'Invalid JSON body.' };
    return;
  }

  const { itemId, fields } = body || {};
  if (!itemId || !fields || typeof fields !== 'object' || Object.keys(fields).length === 0) {
    context.res = { status: 400, body: 'itemId and fields are required.' };
    return;
  }

  try {
    const token = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
    await patchCaseFields(token, itemId, fields);

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-cache' },
      body: JSON.stringify({ updated: true, itemId, callerEmail }),
    };
  } catch (err) {
    context.log.error('Error updating case:', err.message);
    context.res = { status: 500, body: 'Error updating case: ' + err.message };
  }
};

function patchCaseFields(token, itemId, fields) {
  return new Promise(function (resolve, reject) {
    const body = JSON.stringify({ fields });
    const path = '/v1.0/sites/' + SITE_PATH + '/lists/' + CASES_LIST + '/items/' + itemId;

    const options = {
      hostname: 'graph.microsoft.com',
      path,
      method: 'PATCH',
      headers: {
        Authorization:    'Bearer ' + token,
        'Content-Type':   'application/json',
        'Content-Length': Buffer.byteLength(body),
      },
    };

    const req = https.request(options, function (res) {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', function () {
        if (res.statusCode >= 400) {
          reject(new Error('Graph PATCH ' + res.statusCode + ': ' + data.slice(0, 400)));
          return;
        }
        try { resolve(JSON.parse(data)); }
        catch { resolve({}); }
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

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
      res.on('end', function () {
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
