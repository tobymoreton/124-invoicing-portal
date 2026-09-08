/**
 * P124 — Invoicing Portal
 * Azure Function: /api/applycaserate
 *
 * Cascades a case's current hourly rate (Cases.RateForCase) onto that case's
 * UNBILLED Time Tracking2 actions — rewriting field_6 (rate) and the
 * Num_BillableAmount_x00a3_ snapshot — so a rate change reaches the actions
 * immediately instead of waiting for the weekly PA555 job. BILLED actions are
 * never touched (their stored amount is what was invoiced).
 *
 * GET  ?ref={ourRef}   — preview only (dry run): what WOULD change
 * POST ?ref={ourRef}   — apply the change
 *
 * TT2 GUID:   67db204c-30a5-4f4d-b276-60852d9967e1
 * Cases GUID: ae420bda-e550-499c-b337-90e4f33617c1
 */

const https   = require('https');
const { URL } = require('url');

const TT2_GUID       = '67db204c-30a5-4f4d-b276-60852d9967e1';
const CASES_GUID     = 'ae420bda-e550-499c-b337-90e4f33617c1';
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
  context.log('P124 /api/applycaserate called — method:', req.method);

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

  const ref = (req.query.ref || '').trim();
  if (!ref) { context.res = { status: 400, body: 'Missing required param: ref' }; return; }
  // GET is always a preview. POST applies (unless explicitly ?dryRun=1).
  const dryRun = req.method === 'GET' || req.query.dryRun === '1';

  try {
    const token = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);

    // 1) The case's current rate
    const rate = await fetchCaseRate(token, ref);
    if (!(rate > 0)) {
      context.res = { status: 400, body: `No usable case rate (RateForCase) set for ${ref}. Set the case rate first.` };
      return;
    }

    // 2) That case's TT2 actions
    const items = await fetchTT2ForRef(token, ref);

    // 3) Classify. Only unbilled, billable rows are eligible; billed rows keep the
    //    invoiced snapshot and are never restated. Mismatch guard avoids needless writes.
    const toUpdate = [];
    let skippedSameRate = 0, skippedNoHours = 0, billedSkipped = 0, nonBillableSkipped = 0;

    for (const it of items) {
      const f = it.fields || {};
      if (f['Billed_x003f_'] === true)    { billedSkipped++;      continue; }
      if (f['Billable_x003f_'] === false) { nonBillableSkipped++; continue; }
      const hrs = parseFloat(f['TimeSpentMirror'] || f['field_3']) || 0;
      if (hrs <= 0)                       { skippedNoHours++;     continue; }
      const curRate = parseFloat(f['field_6']) || 0;
      if (Math.abs(curRate - rate) < 0.005) { skippedSameRate++; continue; }
      const pro    = parseFloat(f['ProRataApportionment']);
      const factor = (isNaN(pro) || pro <= 0) ? 1 : pro / 100;
      toUpdate.push({
        id: it.id,
        hrs,
        fromRate: curRate,
        toRate: rate,
        oldAmount: parseFloat(f['Num_BillableAmount_x00a3_']) || 0,
        newAmount: Math.round(hrs * rate * factor * 100) / 100,
      });
    }

    // 4) Apply (unless dry run). Sequential — a case has tens of actions, not thousands.
    let updated = 0;
    const errors = [];
    if (!dryRun) {
      const stampBy  = callerEmail;
      const stampUTC = new Date().toISOString();
      for (const u of toUpdate) {
        try {
          await graphPatch(
            `https://graph.microsoft.com/v1.0/sites/${SITE_PATH}/lists/${TT2_GUID}/items/${u.id}/fields`,
            token,
            {
              field_6: u.toRate,
              Num_BillableAmount_x00a3_: u.newAmount,
              LastEditedByEmail: stampBy,
              LastEditedUTC: stampUTC,
            },
          );
          updated++;
        } catch (e) {
          errors.push({ id: u.id, error: e.message });
        }
      }
    }

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-cache' },
      body: JSON.stringify({
        ref, rate, dryRun,
        eligible: toUpdate.length,
        updated: dryRun ? 0 : updated,
        skippedSameRate, skippedNoHours, billedSkipped, nonBillableSkipped,
        errors,
        items: toUpdate,
      }),
    };
  } catch (err) {
    context.log.error('Error in /api/applycaserate:', err.message);
    context.res = { status: 500, body: 'Error: ' + err.message };
  }
};

// ─── The case's current rate from the Cases list ─────────────────────────────
async function fetchCaseRate(token, ref) {
  const sel = 'Ourreference_x0028_text_x0029_,RateForCase';
  let url = `https://graph.microsoft.com/v1.0/sites/${SITE_PATH}/lists/${CASES_GUID}/items`
          + `?$expand=fields($select=${encodeURIComponent(sel)})&$top=500`;
  while (url) {
    const page = await graphGet(url, token, true);
    const match = (page.value || []).find(it =>
      (it.fields?.['Ourreference_x0028_text_x0029_'] || '').toString().trim() === ref);
    if (match) { const v = parseFloat(match.fields?.['RateForCase']); return isNaN(v) ? 0 : v; }
    url = page['@odata.nextLink'] || null;
  }
  return 0;
}

// ─── That case's TT2 actions (server filter on field_16, full-scan fallback) ──
async function fetchTT2ForRef(token, ref) {
  const base = `https://graph.microsoft.com/v1.0/sites/${SITE_PATH}/lists/${TT2_GUID}/items`
             + `?$expand=fields&$top=999`;
  try {
    const escaped = ref.replace(/'/g, "''");
    let url = base + `&$filter=fields/field_16 eq '${escaped}'`;
    let out = [];
    while (url) { const p = await graphGet(url, token); out = out.concat(p.value || []); url = p['@odata.nextLink'] || null; }
    return out;
  } catch (e) {
    let url = base, all = [];
    while (url) { const p = await graphGet(url, token); all = all.concat(p.value || []); url = p['@odata.nextLink'] || null; }
    return all.filter(it => (it.fields?.['field_16'] || '').toString().trim() === ref);
  }
}

// ─── Token / Graph helpers (identical to /api/caseactions) ───────────────────
function getToken(tenantId, clientId, clientSecret) {
  return new Promise((resolve, reject) => {
    const body = new URLSearchParams({
      grant_type: 'client_credentials', client_id: clientId,
      client_secret: clientSecret, scope: 'https://graph.microsoft.com/.default',
    }).toString();
    const req = https.request({
      hostname: 'login.microsoftonline.com',
      path: `/${tenantId}/oauth2/v2.0/token`, method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(body) },
    }, res => {
      let data = ''; res.on('data', c => data += c);
      res.on('end', () => { try { const j = JSON.parse(data); j.access_token ? resolve(j.access_token) : reject(new Error(`Token error: ${j.error_description || data}`)); } catch (e) { reject(e); } });
    });
    req.setTimeout(20000, () => req.destroy(new Error('Token request timeout (20s)')));
    req.on('error', reject); req.write(body); req.end();
  });
}
function graphGet(url, token, allowThrottleable = false) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    const headers = { Authorization: `Bearer ${token}`, Accept: 'application/json' };
    if (allowThrottleable) headers['Prefer'] = 'allowthrottleablequeries';
    const req = https.request({ hostname: u.hostname, path: u.pathname + u.search, method: 'GET', headers }, res => {
      let data = ''; res.on('data', c => data += c);
      res.on('end', () => { if (res.statusCode >= 400) return reject(new Error(`Graph GET ${res.statusCode}: ${data.slice(0,300)}`)); try { resolve(JSON.parse(data)); } catch (e) { reject(new Error(`JSON parse: ${e.message}`)); } });
    });
    req.setTimeout(20000, () => req.destroy(new Error('Graph GET timeout (20s)')));
    req.on('error', reject); req.end();
  });
}
function graphPatch(url, token, body) {
  return new Promise((resolve, reject) => {
    const s = JSON.stringify(body); const u = new URL(url);
    const req = https.request({ hostname: u.hostname, path: u.pathname + u.search, method: 'PATCH',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json', Accept: 'application/json', 'Content-Length': Buffer.byteLength(s) } }, res => {
      let data = ''; res.on('data', c => data += c);
      res.on('end', () => { if (res.statusCode >= 400) return reject(new Error(`Graph PATCH ${res.statusCode}: ${data.slice(0,300)}`)); resolve(data ? JSON.parse(data) : {}); });
    });
    req.setTimeout(20000, () => req.destroy(new Error('Graph PATCH timeout (20s)')));
    req.on('error', reject); req.write(s); req.end();
  });
}
