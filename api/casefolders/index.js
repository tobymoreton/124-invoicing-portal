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

// ── S82: THE TWO-LIBRARY SHORTLIST ──────────────────────────────────────────
// Across every probe on 2026-07-18, only these two libraries EVER returned a
// hit; the other 33 returned raw=0 every single time. Searching all 35 is what
// made this endpoint unusable: v3 sequential 15.9-17.2 s; v4 Promise.all = Graph
// 429 'activityLimitReached' on 31 of 35 and ZERO results; v5 pool-of-4 + retry
// = correct but 43.2 s. Pooling did not avoid the throttle, it converted
// throttling into waiting. Two calls cannot be throttled.
//
// Matched by NAME (case-insensitive, trimmed). If NEITHER name is found — e.g.
// a library is renamed — we fall back to the full sweep, so a rename degrades to
// SLOW, never to BROKEN. `?deep=1` forces the full sweep on demand, so nothing
// is ever permanently unreachable.
// S84 CORRECTION: the 2026-07-18 probe was run against Toby's own cases, whose files
// live in Working Drafts. Draftsmen's case folders live in the TMC-File site's DEFAULT
// library, named 'Documents' (e.g. 'McCalla - 1741294'). That library was never in the
// shortlist, so for those users the endpoint returned count=0 with failed=0 - a silent,
// convincing "nothing found". Tracy reported it 2026-07-23; confirmed live the next day
// (ref 1741294: shortlist raw=0; deep raw=29 kept=1, site TMC-File, library Documents).
// Lesson: a shortlist derived from one user's data is a shortlist fitted to one user.
const SHORTLIST = ['working drafts (current)', 'email attachments', 'documents'];

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
  // ?deep=1 forces the full sweep of every library (slow — see SHORTLIST above).
  const deep  = String((req.query && req.query.deep)  || '') === '1';

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

    // 1. Enumerate every document library on both sites (two cheap calls).
    const allDrives = [];
    for (const site of SITES) {
      try {
        const dl = await graphGet(
          'https://graph.microsoft.com/v1.0/sites/' + site.path + '/drives?$select=id,name', token);
        ((dl && dl.value) || []).forEach(d => allDrives.push({ site: site.key, drive: d }));
      } catch (e) {
        dbg.push({ site: site.key, library: '(drive list)', error: errText(e) });
      }
    }

    // 2. Pick the targets. Shortlist unless ?deep=1, or unless neither name exists.
    const shortlisted = allDrives.filter(
      x => SHORTLIST.indexOf(String(x.drive.name || '').trim().toLowerCase()) !== -1);

    let mode, targets;
    if (deep)                    { mode = 'deep';          targets = allDrives; }
    else if (shortlisted.length) { mode = 'shortlist';     targets = shortlisted; }
    else                         { mode = 'fallback-full'; targets = allDrives; }

    {
      // Bounded concurrency + one 429 retry. In shortlist mode the pool is the
      // whole (tiny) target list, so both searches go out together.
      const CONCURRENCY = mode === 'shortlist' ? Math.max(1, targets.length) : 4;
      const queue = targets.slice();

      const worker = async () => {
        while (queue.length) {
          const t = queue.shift();
          if (!t) break;
          const drv = t.drive;
          const row = { site: t.site, library: drv.name || drv.id, raw: 0, kept: 0, retried: false, error: null };
          try {
            const q   = encodeURIComponent("'" + ref.replace(/'/g, "''") + "'");
            // No $select — naming fields in a $select on search() results has proved
            // unreliable on this tenant (cf. the boolean-dropping $select in /api/wip).
            const url = 'https://graph.microsoft.com/v1.0/drives/' + drv.id
                      + '/root/search(q=' + q + ')?$top=200';

            // S83: two retries with escalating backoff. A search that is
            // throttled to death must NOT look like a genuine empty result —
            // see `failed` in the payload below.
            let res;
            let attempt = 0;
            for (;;) {
              try {
                res = await graphGet(url, token);
                break;
              } catch (e1) {
                if (String(e1 && e1.message).indexOf('Graph 429') === -1) throw e1;
                attempt++;
                if (attempt > 2) throw e1;
                row.retried = true;
                await sleep(attempt === 1 ? 1500 : 4000);
              }
            }

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
                site:     t.site,
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
      };

      const workers = [];
      for (let i = 0; i < Math.min(CONCURRENCY, targets.length); i++) workers.push(worker());
      await Promise.all(workers);
    }

    // Folders first (they are the case's home, and Toby's naming convention is
    // tightening around them), then most recently modified; no date sorts last.
    items.sort((a, b) => {
      if (a.isFolder !== b.isFolder) return a.isFolder ? -1 : 1;
      const da = a.modified ? new Date(a.modified).getTime() : -Infinity;
      const db = b.modified ? new Date(b.modified).getTime() : -Infinity;
      return db - da;
    });

    const payload = {
      ref:       ref,
      count:     items.length,
      mode:      mode,                 // shortlist | deep | fallback-full
      searched:  targets.length,       // libraries actually searched
      // S83: libraries whose search FAILED (Graph 429 throttling, mostly). If
      // this is > 0 the result set is INCOMPLETE and count:0 does not mean
      // "nothing exists" — case.html must say "search failed", not "nothing found".
      failed:    dbg.filter(function (r) { return !!r.error; }).length,
      available: allDrives.length,     // libraries that exist
      items:     items,
    };
    if (debug) payload.drives = dbg;

    context.res = {
      status: 200,
      headers: {
        'Content-Type': 'application/json',
        'Cache-Control': 'no-cache',
        // S81: build marker. Three separate measurements this session were taken
        // against a build that had not deployed yet, and one produced a wrong
        // conclusion. Any response can now be attributed to a specific build in
        // one call. BUMP THIS ON EVERY CHANGE TO THIS FILE.
        'X-Api-Build': 'S84-v8-documents-in-shortlist',
        'X-Casefolders-Mode': mode,
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

function sleep(ms) {
  return new Promise(function (r) { setTimeout(r, ms); });
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
