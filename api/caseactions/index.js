/**
 * P124 — Invoicing Portal
 * Azure Function: /api/caseactions
 *
 * GET  ?ref={ourRef}   — fetch all TT2 entries for a case ref
 * POST (body)          — add a new TT2 entry
 * PATCH (body)         — edit Work Done / Time Spent / Rate on an existing TT2 entry
 *
 * Auth: @tmclegal.co.uk domain check on all methods.
 *
 * TT2 GUID: 67db204c-30a5-4f4d-b276-60852d9967e1
 * SITE_PATH: tmcostings.sharepoint.com:/sites/TMCLegalLimited: (hardcoded)
 *
 * TT2 field map:
 *   field_1   — Date Entered (DateTime)
 *   field_2   — Work Done description (Note)
 *   field_3   — Time Spent hrs (Number)
 *   field_6   — Rate £/hr (Number)
 *   field_9   — Brief (Choice)
 *   field_12  — Date Work Done (DateTime)
 *   field_16  — Our Reference (Text) — filter key
 *   field_18  — Email (Text)
 *   Completedby_x0028_text_x0029_ — Completed by name (Text)
 *   Billable_x003f_               — Billable? (Boolean)
 *   Billed_x003f_                 — Billed? (Boolean) — filter client-side only
 *   Num_BillableAmount_x00a3_     — Billable Amount £ (Currency)
 *   TimeSpentMirror               — Time Spent mirror (Number)
 *   CaseName                       — Case Name (Lookup → Cases list; written as integer SP item ID)
 *   Casename_x0028_text_x0029_    — Case name text mirror (separate column)
 */

const https   = require('https');
const { URL } = require('url');

const TT2_GUID      = '67db204c-30a5-4f4d-b276-60852d9967e1';
const SITE_PATH     = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';
const ALLOWED_DOMAIN = '@tmclegal.co.uk';

// No $select on fields expand — boolean fields (Billable_x003f_, Billed_x003f_) are
// not reliably returned when named in $select via Graph. Fetching all fields ensures
// booleans come through. Per-case TT2 item count is small so no perf concern.

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
  context.log('P124 /api/caseactions called — method:', req.method);

  const callerEmail = getCallerEmail(req);
  if (!callerEmail || !callerEmail.endsWith(ALLOWED_DOMAIN)) {
    context.res = { status: 403, body: 'Forbidden — TMC Legal staff only.' };
    return;
  }

  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    context.res = { status: 500, body: 'Missing required app settings.' };
    return;
  }

  try {
    const token = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);

    // ── GET ──────────────────────────────────────────────────────────────────
    if (req.method === 'GET') {
      // ── Diagnostic: ?schema=1 returns TT2 column internal names — use to find correct lookup field name
      if (req.query.schema === '1') {
        const schemaUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_PATH}/lists/${TT2_GUID}/columns`;
        const schema = await graphGet(schemaUrl, token);
        const cols = (schema.value || []).map(c => ({
          internalName: c.name,
          displayName:  c.displayName,
          required:     c.required,
          isLookup:     !!c.lookup,
          lookupList:   c.lookup?.listId || null,
        })).sort((a, b) => (a.displayName||'').localeCompare(b.displayName||''));
        context.res = { status: 200, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-cache' }, body: JSON.stringify(cols, null, 2) };
        return;
      }

      const ref     = (req.query.ref || '').trim();
      const allMode = (req.query.all || '') === '1';
      if (!ref && !allMode) { context.res = { status: 400, body: 'Missing required param: ref (or all=1)' }; return; }

      const base = `https://graph.microsoft.com/v1.0/sites/${SITE_PATH}/lists/${TT2_GUID}/items`
                 + `?$expand=fields`
                 + `&$top=999`;

      let filtered = [];

      if (!allMode && ref) {
        // Per-case fast path: server-side $filter on field_16 (requires column indexed in SP).
        // Falls back automatically to full list scan if Graph returns 400 (not yet indexed).
        try {
          const escapedRef = ref.replace(/'/g, "''");
          let url = base + `&$filter=fields/field_16 eq '${escapedRef}'`;
          while (url) {
            const page = await graphGet(url, token);
            filtered = filtered.concat(page.value || []);
            url = page['@odata.nextLink'] || null;
          }
        } catch (filterErr) {
          context.log.warn('field_16 filter failed (not indexed?), falling back to full scan:', filterErr.message);
          filtered = [];
          let url = base;
          let all = [];
          while (url) { const page = await graphGet(url, token); all = all.concat(page.value || []); url = page['@odata.nextLink'] || null; }
          filtered = all.filter(item => (item.fields?.['field_16'] || '').toString().trim() === ref);
        }
      } else {
        // all=1 mode: walk entire list, return every UNBILLED entry across all cases.
        // The dashboard uses one call here rather than per-case calls that throttle Graph.
        let url = base;
        let all = [];
        while (url) { const page = await graphGet(url, token); all = all.concat(page.value || []); url = page['@odata.nextLink'] || null; }
        filtered = all.filter(item => item.fields?.['Billed_x003f_'] !== true);
      }

      filtered.sort((a, b) => {
        const da = new Date(a.fields?.['field_12'] || a.fields?.['field_1'] || 0);
        const db = new Date(b.fields?.['field_12'] || b.fields?.['field_1'] || 0);
        return db - da;
      });

      context.res = {
        status: 200,
        headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-cache' },
        body: JSON.stringify({ value: filtered }),
      };
      return;
    }

    // ── POST — add new TT2 entry ─────────────────────────────────────────────
    if (req.method === 'POST') {
      const b = req.body || {};
      const ref      = (b.ref      || '').trim();
      const workDone = (b.workDone || '').trim();
      const timeSpent = parseFloat(b.timeSpent);
      const rate      = parseFloat(b.rate);

      if (!ref)     { context.res = { status: 400, body: 'Missing ref' };     return; }
      if (!workDone){ context.res = { status: 400, body: 'Missing workDone' };return; }
      if (isNaN(timeSpent) || timeSpent <= 0) {
        context.res = { status: 400, body: 'timeSpent must be a positive number' };
        return;
      }

      const now          = new Date().toISOString();
      const dateEntered  = b.dateEntered  ? new Date(b.dateEntered).toISOString()  : now;
      const dateWorkDone = b.dateWorkDone ? new Date(b.dateWorkDone).toISOString() : now;

      context.log('POST caseactions: ref=', ref, 'caseItemId=', b.caseItemId);

      const fields = {
        Title:    b.caseName || ref,
        // Case_x0020_Name is a SP Lookup column (required). SP internal name confirmed from FldEditEx URL: Field=Case_x0020_Name
        // Graph requires LookupId suffix: Case_x0020_NameLookupId
        ...(b.caseItemId && !isNaN(parseInt(b.caseItemId)) ? { 'Case_x0020_NameLookupId': parseInt(b.caseItemId) } : {}),
        field_1:  dateEntered,
        field_2:  workDone,
        field_3:  timeSpent,
        field_6:  isNaN(rate) ? 0 : rate,
        field_9:  b.brief || 'Preparation/Perusal',
        field_12: dateWorkDone,
        field_16: ref,
        field_18: b.email || callerEmail,
        Completedby_x0028_text_x0029_: b.completedBy || '',
        Casename_x0028_text_x0029_:    b.caseName    || '',
        TimeSpentMirror: timeSpent,  // PA mirror field — also write here so COA/WIP don't need to wait for PA run
        // NOTE: Billable_x003f_ and Billed_x003f_ deliberately omitted — boolean writes via Graph are unreliable
        // on this tenant and may trigger SP-side deletion. SP column defaults (Billable=Yes, Billed=No) apply.
      };

      const createUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_PATH}/lists/${TT2_GUID}/items`;
      const created   = await graphPost(createUrl, token, { fields });

      context.res = {
        status: 201,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: created.id, fields: created.fields, _debug: { caseItemId: b.caseItemId, fieldsSent: Object.keys(fields) } }),
      };
      return;
    }

    // ── PATCH — edit Work Done / Time Spent / Rate ───────────────────────────
    if (req.method === 'PATCH') {
      const b      = req.body || {};
      const itemId = (b.itemId || '').toString().trim();
      if (!itemId) { context.res = { status: 400, body: 'Missing itemId' }; return; }

      const fields = {};
      if (b.workDone  !== undefined) fields['field_2'] = b.workDone;
      if (b.timeSpent !== undefined) { fields['field_3'] = parseFloat(b.timeSpent) || 0; fields['TimeSpentMirror'] = parseFloat(b.timeSpent) || 0; }
      if (b.rate      !== undefined) fields['field_6'] = parseFloat(b.rate)      || 0;
      if (b.brief     !== undefined) fields['field_9'] = b.brief;
      if (b.billable  !== undefined) fields['Billable_x003f_'] = !!b.billable;

      if (Object.keys(fields).length === 0) {
        context.res = { status: 400, body: 'No editable fields provided' };
        return;
      }

      const patchUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_PATH}/lists/${TT2_GUID}/items/${itemId}/fields`;
      await graphPatch(patchUrl, token, fields);

      context.res = {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ok: true }),
      };
      return;
    }

    context.res = { status: 405, body: 'Method not allowed' };

  } catch (err) {
    context.log.error('Error in /api/caseactions:', err.message);
    context.res = { status: 500, body: 'Error: ' + err.message };
  }
};

function getToken(tenantId, clientId, clientSecret) {
  return new Promise((resolve, reject) => {
    const body = new URLSearchParams({
      grant_type: 'client_credentials', client_id: clientId,
      client_secret: clientSecret, scope: 'https://graph.microsoft.com/.default',
    }).toString();
    const options = {
      hostname: 'login.microsoftonline.com',
      path: `/${tenantId}/oauth2/v2.0/token`,
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(body) },
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
      hostname: u.hostname, path: u.pathname + u.search, method: 'GET',
      headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
    }, res => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        if (res.statusCode >= 400) { reject(new Error(`Graph GET ${res.statusCode}: ${data.slice(0,300)}`)); return; }
        try { resolve(JSON.parse(data)); } catch (e) { reject(new Error(`JSON parse: ${e.message}`)); }
      });
    });
    req.setTimeout(20000, () => { req.destroy(new Error('Graph GET timeout (20s) — SP may be throttling')); });
    req.on('error', reject);
    req.end();
  });
}

function graphPost(url, token, body) {
  return new Promise((resolve, reject) => {
    const bodyStr = JSON.stringify(body);
    const u = new URL(url);
    const req = https.request({
      hostname: u.hostname, path: u.pathname + u.search, method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json', Accept: 'application/json', 'Content-Length': Buffer.byteLength(bodyStr) },
    }, res => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        if (res.statusCode >= 400) { reject(new Error(`Graph POST ${res.statusCode}: ${data.slice(0,300)}`)); return; }
        try { resolve(JSON.parse(data)); } catch (e) { reject(new Error(`JSON parse: ${e.message}`)); }
      });
    });
    req.on('error', reject);
    req.write(bodyStr);
    req.end();
  });
}

function graphPatch(url, token, body) {
  return new Promise((resolve, reject) => {
    const bodyStr = JSON.stringify(body);
    const u = new URL(url);
    const req = https.request({
      hostname: u.hostname, path: u.pathname + u.search, method: 'PATCH',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json', Accept: 'application/json', 'Content-Length': Buffer.byteLength(bodyStr) },
    }, res => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        if (res.statusCode >= 400) { reject(new Error(`Graph PATCH ${res.statusCode}: ${data.slice(0,300)}`)); return; }
        resolve(data ? JSON.parse(data) : {});
      });
    });
    req.on('error', reject);
    req.write(bodyStr);
    req.end();
  });
}
