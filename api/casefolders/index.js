/**
 * P124 — Invoicing Portal
 * Azure Function: /api/casefolders
 *
 * Finds SharePoint FOLDERS whose name contains a case reference, across both
 * document sites, and returns links to them.
 *
 * ⚠️  READ-ONLY. This function performs NO writes of any kind and never reads
 *     folder CONTENTS — it returns the folder's own name, path and webUrl only.
 *     TMC-File is a live working library; nothing here opens, lists or alters it.
 *
 * ⚠️  S81 STATUS: PROBE. Whether Graph app-only auth permits driveItem
 *     `search(q=)` on this tenant is NOT PROVEN. It is UNCERTAIN, and the only way
 *     to establish it is to deploy this and call it. `?debug=1` returns the raw
 *     per-site outcome (status + first 300 chars of any error) precisely so a
 *     failure is diagnosable in one call instead of presenting as "no results".
 *     Do NOT build UI against this until a live call has been seen to work.
 *
 * Sites searched (Toby, 2026-07-18): BOTH.
 *   - tmcostings.sharepoint.com:/sites/TMCLegalLimited:   (the portal's own site)
 *   - tmcostings.sharepoint.com:/sites/TMC-File:          (the case file library)
 *
 * GET /api/casefolders?ref=<Our Reference>[&debug=1]
 * Returns: { ref, folders: [ { name, site, path, webUrl } ], sites: [...debug] }
 */

const https   = require('https');
const { URL } = require('url');

// Both document sites. `key` is what the client shows as the source label.
const SITES = [
  { key: 'TMCLegalLimited', path: 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:' },
  { key: 'TMC-File',        path: 'tmcostings.sharepoint.com:/sites/TMC-File:' },
];

const ALLOWED_DOMAIN = '@tmclegal.co.uk';

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
  context.log('P124 /api/casefolders called');

  const callerEmail = getCallerEmail(req);
  if (!callerEmail || callerEmail.indexOf(ALLOWED_DOMAIN) === -1) {
    context.res = { status: 403, body: 'Forbidden — TMC staff only.' };
    return;
  }

  const ref = ((req.query && req.query.ref) || '').trim();
  if (!ref) {
    context.res = { status: 400, body: 'Missing required query parameter: ref.' };
    return;
  }
  // Guard against a wildcard-ish search returning half the tenant.
  if (ref.length < 4) {
    context.res = { status: 400, body: 'ref must be at least 4 characters.' };
    return;
  }

  const debug = String((req.query && req.query.debug) || '') === '1';

  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    context.res = { status: 500, body: 'Missing required app settings.' };
    return;
  }

  try {
    const token   = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
    const folders = [];
    const sites   = [];

    for (const site of SITES) {
      const outcome = { site: site.key, ok: false, driveId: null, hits: 0, error: null };
      try {
        // Resolve the site's default document library drive id first — the same
        // two-step /api/attachments needed (a drive id, not a list route).
        const drive = await graphGet(
          'https://graph.microsoft.com/v1.0/sites/' + site.path + '/drive?$select=id', token);
        outcome.driveId = drive && drive.id ? drive.id : null;
        if (!outcome.driveId) throw new Error('no drive id returned for site');

        const q   = encodeURIComponent("'" + ref.replace(/'/g, "''") + "'");
        const url = 'https://graph.microsoft.com/v1.0/drives/' + outcome.driveId
                  + '/root/search(q=' + q + ')'
                  + '?$select=id,name,webUrl,folder,parentReference&$top=50';

        const res   = await graphGet(url, token);
        const items = (res && res.value) || [];

        // FOLDERS ONLY, and only where the NAME carries the ref (Graph search also
        // matches file CONTENT, which would return documents that merely mention it).
        const refLower = ref.toLowerCase();
        items.filter(it => it.folder && (it.name || '').toLowerCase().indexOf(refLower) !== -1)
             .forEach(it => {
               folders.push({
                 name:   it.name || null,
                 site:   site.key,
                 path:   (it.parentReference && it.parentReference.path) || null,
                 webUrl: it.webUrl || null,
               });
               outcome.hits++;
             });
        outcome.ok = true;
      } catch (e) {
        outcome.error = (e && e.message ? e.message : String(e)).slice(0, 300);
      }
      sites.push(outcome);
    }

    const payload = { ref: ref, folders: folders };
    if (debug) payload.sites = sites;

    context.res = {
      status: 200,
      headers: {
        'Content-Type': 'application/json',
        'Cache-Control': 'no-cache',
      },
      body: JSON.stringify(payload),
    };
  } catch (err) {
    context.log.error('Error searching case folders:', err.message);
    context.res = { status: 500, body: 'Error: ' + err.message };
  }
};

// ─── TOKEN (client-credentials) ──────────────────────────
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

// ─── GRAPH GET ───────────────────────────────────────────
function graphGet(url, token) {
  return new Promise(function (resolve, reject) {
    const u = new URL(url);
    const options = {
      hostname: u.hostname,
      path:     u.pathname + u.search,
      method:   'GET',
      headers: {
        Authorization:    'Bearer ' + token,
        Accept:           'application/json',
        ConsistencyLevel: 'eventual',
      },
    };

    const req = https.request(options, function (res) {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        if (res.statusCode >= 400) {
          reject(new Error('Graph ' + res.statusCode + ': ' + data.slice(0, 300)));
          return;
        }
        try { resolve(JSON.parse(data)); }
        catch (e) { reject(new Error('JSON parse error: ' + e.message)); }
      });
    });
    req.on('error', reject);
    req.end();
  });
}
