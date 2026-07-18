/**
 * P124 — Invoicing Portal
 * Azure Function: /api/casefolders
 *
 * Finds SharePoint FILES AND FOLDERS whose NAME contains a case reference, across
 * every document library on both sites, newest-modified first.
 *
 * ⚠️  READ-ONLY. No writes of any kind. Returns each item's own name, library,
 *     path, size, modified date and webUrl — it never opens or downloads content.
 *
 * ── S81 findings, all EXTRACTED from live probes 2026-07-18 ──────────────────
 *  1. App-only Graph `search(q=)` on a drive IS permitted on this tenant.
 *     (First probe: both drives resolved, ok:true, no permission error.)
 *  2. Searching `/sites/{site}:/drive` searches ONLY the site's DEFAULT document
 *     library. The first probe did exactly that and returned 0 for
 *     TMCLegalLimited — not because nothing matched, but because Toby's files sit
 *     in a DIFFERENT library on that site. FIX: enumerate /drives and search every
 *     one. This is why v1 reported "nothing found" for the site holding most of
 *     the documents.
 *  3. Ref-named FILES definitely exist (ref 1736735 returned 9, e.g.
 *     "1736735-Forster-FormalServiceLetter.pdf"); ref-named FOLDERS were not found
 *     in the default TMC-File library. Toby: a handful of TMC-File folders and
 *     files carry the ref today, and the naming convention will be tightened once
 *     this surfaces in the portal. So: match on NAME for both files and folders,
 *     and do not assume a per-case folder exists.
 *  4. Graph search also matches file CONTENT and PATH, which returns items that
 *     merely mention the ref. Everything is therefore filtered to NAME contains ref.
 *
 * GET /api/casefolders?ref=<Our Reference>[&debug=1]
 * Returns: { ref, count, items: [ { name, isFolder, library, site, path, webUrl,
 *                                   modified, size } ], drives?: [...debug] }
 */

const https   = require('https');
const { URL } = require('url');

const SITES = [
  { key: 'TMC Legal Limited', path: 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:' },
  { key: 'TMC-File',          path: 'tmcostings.sharepoint.com:/sites/TMC-File:' },
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
    const token    = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
    const refLower = ref.toLowerCase();
    const items    = [];
    const dbg      = [];
    const seen     = {};   // dedupe by driveItem id

    for (const site of SITES) {
      let drives = [];
      try {
        // EVERY document library on the site, not just the default one.
        const dl = await graphGet(
          'https://graph.microsoft.com/v1.0/sites/' + site.path + '/drives?$select=id,name', token);
        drives = (dl && dl.value) || [];
      } catch (e) {
        dbg.push({ site: site.key, library: '(drive list)', error: errText(e) });
        continue;
      }

      for (const drv of drives) {
        const row = { site: site.key, library: drv.name || drv.id, raw: 0, kept: 0, error: null };
        try {
          const q   = encodeURIComponent("'" + ref.replace(/'/g, "''") + "'");
          // No $select — naming fields in a $select on search() results has proved
          // unreliable on this tenant (cf. the boolean-dropping $select in /api/wip).
          const url = 'https://graph.microsoft.com/v1.0/drives/' + drv.id
                    + '/root/search(q=' + q + ')?$top=200';
          const res = await graphGet(url, token);
          const hits = (res && res.value) || [];
          row.raw = hits.length;

          hits.forEach(it => {
            const name = it.name || '';
            if (name.toLowerCase().indexOf(refLower) === -1) return;  // NAME must carry the ref
            if (seen[it.id]) return;
            seen[it.id] = true;
            items.push({
              name:     name,
              isFolder: !!it.folder,
              site:     site.key,
              library:  drv.name || null,
              path:     (it.parentReference && it.parentReference.path) || null,
              webUrl:   it.webUrl || null,
              modified: it.lastModifiedDateTime || null,
              size:     typeof it.size === 'number' ? it.size : null,
            });
            row.kept++;
          });
        } catch (e) {
          row.error = errText(e);
        }
        dbg.push(row);
      }
    }

    // Most recently modified first; anything without a date goes last.
    items.sort((a, b) => {
      const da = a.modified ? new Date(a.modified).getTime() : -Infinity;
      const db = b.modified ? new Date(b.modified).getTime() : -Infinity;
      return db - da;
    });

    const payload = { ref: ref, count: items.length, items: items };
    if (debug) payload.drives = dbg;

    context.res = {
      status: 200,
      headers: {
        'Content-Type': 'application/json',
        'Cache-Control': 'no-cache',
      },
      body: JSON.stringify(payload),
    };
  } catch (err) {
    context.log.error('Error searching case files:', err.message);
    context.res = { status: 500, body: 'Error: ' + err.message };
  }
};

function errText(e) {
  return (e && e.message ? e.message : String(e)).slice(0, 300);
}

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
