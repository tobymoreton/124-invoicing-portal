/**
 * P124 — Invoicing Portal
 * Azure Function: /api/caseactions
 *
 * GET  ?ref={ourRef}   — fetch all TT2 entries for a case ref
 * POST (body)          — add a new TT2 entry
 * PATCH (body)         — edit Work Done / Time Spent / Rate on an existing TT2 entry
 * DELETE (body)        — delete a TT2 entry by itemId
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
 *   Case_x0020_Name               — Case Name (Lookup → Cases list; written as Case_x0020_NameLookupId: integer SP item ID)
 *   Casename_x0028_text_x0029_    — Case name text mirror (separate column)
 */

const https   = require('https');
const { URL } = require('url');

const TT2_GUID      = '67db204c-30a5-4f4d-b276-60852d9967e1';
const SITE_PATH     = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';
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

    // -- GET --------------------------------------------------------------------
    if (req.method === 'GET') {
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
      let billedFiltered = null;

      if (!allMode && ref) {
        try {
          const escapedRef = ref.replace(/'/g, "''");
          let url = base + `&$filter=fields/field_16 eq '${escapedRef}'`;
          while (url) {
            const page = await graphGet(url, token);
            filtered = filtered.concat(page.value || []);
            url = page['@odata.nextLink'] || null;
          }
        } catch (filterErr) {
          context.log.warn('field_16 filter failed, falling back to full scan:', filterErr.message);
          filtered = [];
          let url = base;
          let all = [];
          while (url) { const page = await graphGet(url, token); all = all.concat(page.value || []); url = page['@odata.nextLink'] || null; }
          filtered = all.filter(item => (item.fields?.['field_16'] || '').toString().trim() === ref);
        }
      } else {
        // all=1 — the dashboard's bulk read. The unselected $expand=fields payload (every
        // column on 3,600+ rows across 4 sequential Graph pages) outgrew the SWA gateway
        // timeout. S76 measured one call DIE after 143,157 ms with NO http status at all,
        // while an identical call minutes later returned 200 / 3,500 rows in 12,975 ms —
        // i.e. it is now sitting on the limit and failing intermittently. loadActions()
        // then rendered the failure as "No actions recorded." Narrow the projection to the
        // columns the dashboard actually reads.
        // WATCH: boolean columns have previously been seen to drop OUT of a
        // fields($select=...) projection on this tenant. Billed_x003f_ is the filter key
        // here, so the number of rows it removes is returned as _billedFiltered. It should
        // be ~177. If it comes back 0, the projection has eaten the booleans and this change
        // MUST be reverted — otherwise billed work is reported as unbilled WIP.
        const ALL_FIELDS = [
          'field_1','field_2','field_3','field_6','field_9','field_12','field_16','field_18',
          'Completedby_x0028_text_x0029_','Casename_x0028_text_x0029_',
          'Billable_x003f_','Billed_x003f_','Num_BillableAmount_x00a3_',
          'TimeSpentMirror','CreatedUTC','Title',
        ].join(',');
        let url = `https://graph.microsoft.com/v1.0/sites/${SITE_PATH}/lists/${TT2_GUID}/items`
                + `?$select=id,createdDateTime`
                + `&$expand=fields($select=${ALL_FIELDS})`
                + `&$top=999`;
        let all = [];
        while (url) { const page = await graphGet(url, token); all = all.concat(page.value || []); url = page['@odata.nextLink'] || null; }
        filtered = all.filter(item => item.fields?.['Billed_x003f_'] !== true);
        billedFiltered = all.length - filtered.length;
      }

      // ── Optional 'since' narrowing (additive, non-breaking) ─────────────────
      // Omit to preserve current behaviour (fetch everything). Pass an ISO date
      // (e.g. ?since=2026-01-01) to narrow to actions on/after that date — for
      // future use by any view that doesn't need full history (e.g. a paged or
      // date-scoped dashboard), without changing what existing callers get.
      const sinceParam = (req.query.since || '').trim();
      if (sinceParam) {
        const sinceDate = new Date(sinceParam);
        if (!isNaN(sinceDate)) {
          filtered = filtered.filter(item => {
            const d = new Date(item.fields?.['field_12'] || item.fields?.['field_1'] || 0);
            return d >= sinceDate;
          });
        }
      }

      filtered.sort((a, b) => {
        const da = new Date(a.fields?.['field_12'] || a.fields?.['field_1'] || 0);
        const db = new Date(b.fields?.['field_12'] || b.fields?.['field_1'] || 0);
        if (db - da !== 0) return db - da;
        // Tie-break: field_12 (Date Work Done) is stored as midnight-only (T00:00:00Z), so
        // every action logged on the same calendar day has an identical timestamp here.
        // Prefer the real creation timestamp (CreatedUTC, added 2026-07-05 — stamped
        // server-side on POST, never influenced by client input) when BOTH rows have it.
        // Rows created before this field existed won't have it, so fall back to SharePoint
        // item ID descending — ID always increases with creation order, so it remains a
        // correct proxy for those older rows.
        const caU = a.fields?.['CreatedUTC'];
        const cbU = b.fields?.['CreatedUTC'];
        if (caU && cbU) {
          const diff = new Date(cbU) - new Date(caU);
          if (diff !== 0) return diff;
        }
        return (parseInt(b.id, 10) || 0) - (parseInt(a.id, 10) || 0);
      });

      context.res = {
        status: 200,
        headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-cache' },
        body: JSON.stringify({ value: filtered, ...(billedFiltered !== null ? { _billedFiltered: billedFiltered } : {}) }),
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
        // Case_x0020_Name is a required SP Lookup column. SP internal name: Case_x0020_Name (confirmed from FldEditEx URL).
        // Graph requires LookupId suffix. caseItemId = SP integer item ID of the Cases list record.
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
        TimeSpentMirror: timeSpent,
        // CreatedUTC — true server-side creation timestamp. Unlike field_1 (Date Entered),
        // which the caller/UI can backdate, this is always `now` and never client-influenced.
        // Confirmed internal name matches the display name (Toby, 2026-07-05).
        CreatedUTC: now,
        // Billable_x003f_ and Billed_x003f_ deliberately omitted — boolean writes unreliable on this tenant
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

    // ── PATCH — edit existing TT2 entry ─────────────────────────────────────
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
      if (b.completedBy !== undefined) fields['Completedby_x0028_text_x0029_'] = b.completedBy;
      if (b.dateWorkDone !== undefined) fields['field_12'] = b.dateWorkDone;

      if (Object.keys(fields).length === 0) {
        context.res = { status: 400, body: 'No editable fields provided' };
        return;
      }

      // Audit trail — stamp on every successful edit. Confirmed internal names match
      // display names (Toby, 2026-07-05).
      fields['LastEditedByEmail'] = callerEmail;
      fields['LastEditedUTC']     = new Date().toISOString();

      const patchUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_PATH}/lists/${TT2_GUID}/items/${itemId}/fields`;
      await graphPatch(patchUrl, token, fields);

      context.res = {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ok: true }),
      };
      return;
    }

    // ── DELETE — remove TT2 item ─────────────────────────────────────────────
    if (req.method === 'DELETE') {
      const itemId = ((req.body || {}).itemId || '').toString().trim();
      if (!itemId) { context.res = { status: 400, body: 'Missing itemId' }; return; }
      // Audit — the item is gone after delete so there's no field to stamp; log to
      // Azure Function logs (Application Insights) instead.
      context.log('AUDIT caseactions DELETE — itemId=' + itemId + ' deletedBy=' + callerEmail + ' at=' + new Date().toISOString());
      const deleteUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_PATH}/lists/${TT2_GUID}/items/${itemId}`;
      await graphDelete(deleteUrl, token);
      context.res = { status: 200, headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ deleted: true, itemId }) };
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
    req.setTimeout(20000, () => { req.destroy(new Error('Token request timeout (20s)')); });
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
    req.setTimeout(20000, () => { req.destroy(new Error('Graph GET timeout (20s)')); });
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
    req.setTimeout(20000, () => { req.destroy(new Error('Graph POST timeout (20s)')); });
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
    req.setTimeout(20000, () => { req.destroy(new Error('Graph PATCH timeout (20s)')); });
    req.on('error', reject);
    req.write(bodyStr);
    req.end();
  });
}

function graphDelete(url, token) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    const req = https.request({
      hostname: u.hostname, path: u.pathname + u.search, method: 'DELETE',
      headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
    }, res => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        if (res.statusCode >= 400) { reject(new Error(`Graph DELETE ${res.statusCode}: ${data.slice(0,300)}`)); return; }
        resolve({});
      });
    });
    req.setTimeout(20000, () => { req.destroy(new Error('Graph DELETE timeout (20s)')); });
    req.on('error', reject);
    req.end();
  });
}
