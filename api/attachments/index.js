/**
 * P124 — Invoicing Portal
 * Azure Function: /api/attachments
 *
 * Lists files in the shared `caseAttachments` SharePoint document library for a given case,
 * via Graph (app-only client-credentials auth). Read-only.
 *
 * A document library is a Graph drive; its files are driveItems, each carrying a listItem that
 * holds the custom columns (ourRef, docType). We fetch the library root children and expand the
 * listItem fields to get file metadata + columns in ONE call, then filter by ourRef server-side
 * (matches the existing full-scan + client-side-filter pattern used by /api/cases; volume is low).
 *
 * v1 lists the library ROOT only — subfolders are not traversed. If files end up organised into
 * folders later, add recursion (or switch to /lists/{guid}/items?$expand=driveItem).
 *
 * Query params:
 *   ref — case Our Reference (required); matched against the listItem `ourRef` column
 *
 * Library: caseAttachments
 * GUID: 710dea64-11ae-4ae7-8fde-d4508206e1c1
 * Columns: ourRef (Single line of text), docType (Choice)
 */

const https   = require('https');
const { URL } = require('url');

const LIST_GUID      = '710dea64-11ae-4ae7-8fde-d4508206e1c1';
const SITE_PATH      = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';
const ALLOWED_DOMAIN = '@tmclegal.co.uk';

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
  context.log('P124 /api/attachments called');

  const callerEmail = getCallerEmail(req);
  if (!callerEmail || callerEmail.indexOf(ALLOWED_DOMAIN) === -1) {
    context.res = { status: 403, body: 'Forbidden — you must be signed in with a TMC account.' };
    return;
  }

  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    context.res = { status: 500, body: 'Missing required app settings.' };
    return;
  }

  const ref = (req.query.ref || '').trim();
  if (!ref) {
    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-cache' },
      body: JSON.stringify({ value: [] }),
    };
    return;
  }

  try {
    const token  = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
    const result = await listAttachments(token, ref);
    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-cache' },
      body: JSON.stringify(result),
    };
  } catch (err) {
    context.log.error('Error listing attachments:', err.message);
    context.res = { status: 500, body: 'Error: ' + err.message };
  }
};

async function listAttachments(token, ref) {
  // Query the document library's LIST ITEMS, expanding fields (ourRef, docType) + driveItem
  // (file metadata) in one call. This is the supported route for a doc library; the
  // /drive/root/children + nested listItem-expand route returns Graph 400 BadRequest.
  // The list-items endpoint spans the whole library (all folders) — no recursion needed.
  var base = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH + '/lists/' + LIST_GUID + '/items'
           + '?$expand=fields($select=OurRef,docType),driveItem&$top=200';

  var url = base;
  var all = [];
  while (url) {
    var page = await graphGet(url, token);
    all = all.concat(page.value || []);
    url = page['@odata.nextLink'] || null;
  }

  var refLc = ref.toLowerCase();

  var files = all.filter(function(item) {
    var d = item.driveItem;
    if (!d || !d.file) return false; // skip folders / non-file list items
    var f = item.fields || {};
    // NB Graph returns the column key as `OurRef` (capital O) even though SharePoint's Edit Column
    // Field= param shows `ourRef`. Read the Graph key, not the SP internal-name casing.
    return String(f.OurRef || '').trim().toLowerCase() === refLc;
  });

  return {
    value: files.map(function(item) {
      var f = item.fields    || {};
      var d = item.driveItem || {};
      return {
        id:          d.id || item.id,
        name:        d.name || '',
        size:        d.size || 0,
        webUrl:      d.webUrl || '',
        downloadUrl: d['@microsoft.graph.downloadUrl'] || '',
        modified:    d.lastModifiedDateTime || '',
        ourRef:      f.OurRef || '',
        docType:     f.docType || '',
      };
    }),
  };
}

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
