/**
 * P124 — Invoicing Portal
 * Azure Function: /api/courts
 *
 * Returns all court names from the SP Courts list, sorted alphabetically.
 * Used to populate the Court Name dropdown in the Opponent & Court tab.
 *
 * No query params required.
 *
 * Returns:
 *   { courts: ["County Court at Central London", "High Court of Justice", ...] }
 *
 * The Courts list is discovered by display name ("Courts") rather than GUID,
 * so no hardcoded GUID is needed. Results are sorted alphabetically client-side.
 *
 * Cached for 1 hour via Cache-Control header — courts list changes rarely.
 */

const https   = require('https');
const { URL } = require('url');

const SITE_PATH = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';

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
  context.log('P124 /api/courts called');

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

  try {
    const token = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);

    // 1. Find the Courts list by display name
    const listsUrl = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
                   + '/lists?$filter=displayName eq \'Courts\'&$select=id,displayName';
    const listsResp = await graphGet(listsUrl, token);
    const courtsList = (listsResp.value || [])[0];
    if (!courtsList) {
      context.res = { status: 404, body: 'Courts list not found in SharePoint.' };
      return;
    }

    // 2. Fetch all items — courts list is small, no pagination needed in practice
    const itemsUrl = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
                   + '/lists/' + courtsList.id
                   + '/items?$expand=fields($select=Title)&$top=500';
    const itemsResp = await graphGet(itemsUrl, token);
    const courts = (itemsResp.value || [])
      .map(item => (item.fields && item.fields.Title) || '')
      .filter(Boolean)
      .sort((a, b) => a.localeCompare(b));

    context.res = {
      status: 200,
      headers: {
        'Content-Type':  'application/json',
        'Cache-Control': 'public, max-age=3600', // cache 1 hour — courts rarely change
      },
      body: JSON.stringify({ courts }),
    };
  } catch (err) {
    context.log.error('Error fetching courts:', err.message);
    context.res = { status: 500, body: 'Error: ' + err.message };
  }
};

function getToken(tenantId, clientId, clientSecret) {
  return new Promise(function(resolve, reject) {
    var body = new URLSearchParams({
      grant_type:    'client_credentials',
      client_id:     clientId,
      client_secret: clientSecret,
      scope:         'https://graph.microsoft.com/.default',
    }).toString();

    var options = {
      hostname: 'login.microsoftonline.com',
      path:     '/' + tenantId + '/oauth2/v2.0/token',
      method:   'POST',
      headers: {
        'Content-Type':   'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(body),
      },
    };

    var req = https.request(options, function(res) {
      var data = '';
      res.on('data', function(chunk) { data += chunk; });
      res.on('end', function() {
        try {
          var json = JSON.parse(data);
          if (json.access_token) resolve(json.access_token);
          else reject(new Error('Token error: ' + (json.error_description || data)));
        } catch(e) { reject(e); }
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

function graphGet(url, token) {
  return new Promise(function(resolve, reject) {
    var u = new URL(url);
    var options = {
      hostname: u.hostname,
      path:     u.pathname + u.search,
      method:   'GET',
      headers: {
        Authorization: 'Bearer ' + token,
        Accept:        'application/json',
      },
    };

    var req = https.request(options, function(res) {
      var data = '';
      res.on('data', function(chunk) { data += chunk; });
      res.on('end', function() {
        if (res.statusCode >= 400) {
          reject(new Error('Graph ' + res.statusCode + ': ' + data.slice(0, 300)));
          return;
        }
        try { resolve(JSON.parse(data)); }
        catch(e) { reject(new Error('JSON parse error: ' + e.message)); }
      });
    });
    req.on('error', reject);
    req.end();
  });
}
