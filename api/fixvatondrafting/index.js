/**
 * P124 — Invoicing Portal
 * Azure Function: /api/fixvatondrafting
 *
 * S122. One-off repair: sets VatOnDraftingFee_x003f_ = true on every case where it is
 * explicitly false.
 *
 * Why: until 05/09/2026 case.html displayed a BLANK "VAT on Drafting Fee" as "No"
 * (`f['VatOnDraftingFee_x003f_'] ? 'Yes' : 'No'`), pre-selected No in the edit dropdown
 * (`!!(f[...])`), and wrote the dropdown's value back on EVERY Overview save. So on any case
 * where the field had never been set, the first unrelated Overview edit — an assignee, a
 * status, a date — silently persisted `false`. invoice-create.html then read that hard false
 * and produced an invoice with no VAT on the drafting fee. Nobody chose No. S116j fixed the
 * display default to Yes; this fixes the values already written.
 *
 * 27 cases of 2,085 were in this state when the fault was found (2026-09-08), across nine
 * firms and including the AAA test case — the spread is what proves it was not a decision.
 *
 * GET  /api/fixvatondrafting            → DRY RUN. Lists what would change. Writes nothing.
 * POST /api/fixvatondrafting            → applies the change.
 *
 * Admin only (Toby, Danielle).
 *
 * Deliberately narrow: it only ever writes `true`, only to items currently holding an explicit
 * `false`, and touches no other field. A case genuinely exempt from VAT on its drafting fee can
 * be set back to No by hand afterwards — the dry run gives the list to check that against first.
 *
 * GUIDs:
 *   Cases list:  ae420bda-e550-499c-b337-90e4f33617c1
 */

const https   = require('https');
const { URL } = require('url');

const SITE_PATH = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';
const LIST_GUID = 'ae420bda-e550-499c-b337-90e4f33617c1';

const ADMIN_EMAILS = ['toby@tmclegal.co.uk', 'danielle@tmclegal.co.uk'];

// ── Entry point ───────────────────────────────────────────────────────────────

module.exports = async function (context, req) {
  const dryRun = String(req.method || '').toUpperCase() !== 'POST';
  context.log('P124 /api/fixvatondrafting called — ' + (dryRun ? 'DRY RUN' : 'APPLY'));

  const callerEmail = getCallerEmail(req);
  if (!callerEmail || !ADMIN_EMAILS.includes(callerEmail)) {
    context.res = { status: 403, body: 'Forbidden — admin only.' };
    return;
  }

  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    context.res = { status: 500, body: 'Missing required app settings.' };
    return;
  }

  try {
    const token = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);

    // ── 1. Read every case ────────────────────────────────────────────────────
    // $expand=fields with NO $select: SharePoint booleans drop silently out of a
    // projection, which would make every case read blank and this repair find nothing.
    // No $filter either — the column is not indexed, so filtering server-side fails.
    let url = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
      + '/lists/' + LIST_GUID + '/items?$expand=fields&$top=500';
    const cases = [];
    let pages = 0;
    while (url && pages < 40) {
      const page = await graphGet(url, token);
      (page.value || []).forEach(it => cases.push(it));
      url = page['@odata.nextLink'] || null;
      pages++;
    }
    context.log('Read', cases.length, 'cases over', pages, 'page(s).');

    // ── 2. Pick out the explicit falses ───────────────────────────────────────
    // Strictly `=== false`. A blank or missing value is left alone: it already reads as
    // Yes everywhere, and writing to it would touch cases this fault never affected.
    const targets = cases.filter(it => (it.fields || {})['VatOnDraftingFee_x003f_'] === false);

    const listed = targets.map(it => {
      const f = it.fields || {};
      return {
        itemId: String(it.id),
        ref:    f['Ourreference_x0028_text_x0029_'] || null,
        name:   f.Title || null,
        firm:   f['Firm_x0028_text_x0029_'] || null,
        status: f.StatusMirror || null,
        laip:   f.InterPartesorLegalAid || null,
      };
    });

    if (dryRun) {
      context.res = {
        status:  200,
        headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-cache' },
        body: JSON.stringify({
          dryRun: true, casesRead: cases.length, wouldUpdate: targets.length, cases: listed,
        }),
      };
      return;
    }

    // ── 3. Apply ──────────────────────────────────────────────────────────────
    // One at a time, recording each outcome. A failure on one case must not abort the rest
    // or leave the caller guessing which ones landed.
    const updated = [];
    const failed  = [];
    for (const t of listed) {
      try {
        const patchUrl = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
          + '/lists/' + LIST_GUID + '/items/' + t.itemId + '/fields';
        await graphPatch(patchUrl, token, { 'VatOnDraftingFee_x003f_': true });
        updated.push(t);
      } catch (e) {
        context.log.error('Failed on item', t.itemId, e.message);
        failed.push({ ...t, error: e.message });
      }
    }
    context.log('Updated', updated.length, '— failed', failed.length);

    context.res = {
      status:  200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-cache' },
      body: JSON.stringify({
        dryRun: false, casesRead: cases.length,
        updatedCount: updated.length, failedCount: failed.length,
        updated, failed, updatedBy: callerEmail,
      }),
    };
  } catch (err) {
    context.log.error('fixvatondrafting error:', err.message, err.stack);
    context.res = { status: 500, body: 'Error: ' + err.message };
  }
};

// ── Helpers ───────────────────────────────────────────────────────────────────

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

function graphGet(url, token) {
  return new Promise(function (resolve, reject) {
    const u = new URL(url);
    const options = {
      hostname: u.hostname,
      path:     u.pathname + u.search,
      method:   'GET',
      headers:  { Authorization: 'Bearer ' + token },
    };
    const req = https.request(options, function (res) {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        if (res.statusCode >= 400) {
          reject(new Error('Graph GET ' + res.statusCode + ': ' + data.slice(0, 400)));
          return;
        }
        try { resolve(JSON.parse(data)); } catch (e) { reject(e); }
      });
    });
    req.on('error', reject);
    req.end();
  });
}

function graphPatch(url, token, bodyObj) {
  return new Promise(function (resolve, reject) {
    const payload = JSON.stringify(bodyObj);
    const u = new URL(url);
    const options = {
      hostname: u.hostname,
      path:     u.pathname + u.search,
      method:   'PATCH',
      headers: {
        Authorization:    'Bearer ' + token,
        'Content-Type':   'application/json',
        'Content-Length': Buffer.byteLength(payload),
      },
    };
    const req = https.request(options, function (res) {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        if (res.statusCode >= 400) {
          reject(new Error('Graph PATCH ' + res.statusCode + ': ' + data.slice(0, 400)));
          return;
        }
        try { resolve(data ? JSON.parse(data) : {}); } catch { resolve({}); }
      });
    });
    req.on('error', reject);
    req.write(payload);
    req.end();
  });
}
