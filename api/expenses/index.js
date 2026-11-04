/**
 * P124 — Invoicing Portal
 * Azure Function: /api/expenses   (S159, 04/11/2026)
 *
 * The case Expenses tab. One row per expense (court fee, courier, copying …), each
 * recorded as paid by the CLIENT or by TMC.
 *
 *   - Paid by TMC  → offered on the draft-invoice form as its own invoice line and
 *                    recharged. Stamped with the invoice when drafted/issued; released on cancel.
 *   - Paid by client → never invoiced.
 *
 * The Settlement tab no longer takes a court fee typed by hand. After EVERY write here the
 * case's two existing columns are rewritten from the expense totals (gross = net + VAT):
 *     CourtFeePaidByTMC  = Σ TMC rows       ClientPaidExpenses = Σ client rows
 * so Settlement, Scenarios, the settlement PDF and invoice-create read the same columns
 * they always have — no change to any settlement arithmetic.
 *
 * Routes:
 *   GET  /api/expenses?ref=1741297                 → { rows, totals }
 *   GET  /api/expenses?ref=1741297&invoiceable=1   → TMC rows not on any draft/invoice only
 *   POST { action:'add',    ref, caseItemId, date, type, description, net, vat, paidBy }
 *   POST { action:'update', id, caseItemId, date?, type?, description?, net?, vat?, paidBy? }
 *   POST { action:'delete', id, caseItemId }
 *
 * A row on a draft or issued invoice (InvoiceItemId set) is LOCKED: no update, no delete.
 * Cancel the invoice (api/cancelinvoice releases it) to change it.
 *
 * List: CaseExpenses (created in the SP UI, no spaces → internal name == display name).
 * Columns — plain types only, never Yes/No / Lookup / Person / Calculated:
 *   Title (default)        description
 *   Ourref                 single line text  — case reference
 *   ExpenseDate            date only
 *   ExpenseType            single line text  — Court fee | Courier / postage | Copying | Other
 *   AmountNet              number, 2dp
 *   VATAmount              number, 2dp
 *   PaidBy                 single line text  — Client | TMC
 *   InvoiceItemId          single line text  — Invoice Library item id of the draft/invoice it is on
 *   InvoiceNumber          single line text  — stamped by api/issueinvoice
 *   AddedByEmail           single line text
 */

const https   = require('https');
const { URL } = require('url');

const API_BUILD   = 'S159-expenses';
const SITE_PATH   = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';
// CaseExpenses list GUID (created 04/11/2026) — same value in createdraft,
// issueinvoice and cancelinvoice (EXPENSES_LIST).
const EXPENSES_LIST = '14d9fb67-1c55-484d-85ba-730ce73569a0';  // CaseExpenses
const CASES_LIST    = 'ae420bda-e550-499c-b337-90e4f33617c1';
const REF_FIELD     = 'Ourreference_x0028_text_x0029_';
const ALLOWED_DOMAIN = '@tmclegal.co.uk';

const TYPES   = ['Court fee', 'Courier / postage', 'Copying', 'Other'];
const PAYERS  = ['Client', 'TMC'];
const SELECT  = 'id,Title,Ourref,ExpenseDate,ExpenseType,AmountNet,VATAmount,PaidBy,InvoiceItemId,InvoiceNumber,AddedByEmail,Created';

module.exports = async function (context, req) {
  const headers = { 'Content-Type': 'application/json', 'X-Api-Build': API_BUILD };
  const send = (status, obj) => { context.res = { status, headers, body: JSON.stringify(obj) }; };
  const fail = (status, message) => send(status, { error: message });

  const callerEmail = getCallerEmail(req);
  if (!callerEmail || !callerEmail.endsWith(ALLOWED_DOMAIN)) return fail(403, 'Forbidden.');

  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) return fail(500, 'Missing required app settings.');

  const base = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH + '/lists/';
  const listUrl = base + EXPENSES_LIST;

  try {
    const token = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);

    // ── GET ────────────────────────────────────────────────────────────────
    if (req.method === 'GET') {
      const ref = String((req.query && req.query.ref) || '').trim();
      if (!ref) return fail(400, 'ref is required.');
      let rows = await rowsForRef(token, listUrl, ref);
      if (req.query.invoiceable === '1') rows = rows.filter(r => r.paidBy === 'TMC' && r.status === 'open');
      return send(200, { ref, rows, totals: totals(rows) });
    }

    // ── POST ───────────────────────────────────────────────────────────────
    let body;
    try { body = typeof req.body === 'string' ? JSON.parse(req.body) : (req.body || {}); }
    catch (e) { return fail(400, 'Invalid JSON body.'); }
    const action = String(body.action || '');

    let ref;
    if (action === 'add') {
      ref = String(body.ref || '').trim();
      if (!ref) return fail(400, 'ref is required.');
      const v = validate(body, true);
      if (v.error) return fail(400, v.error);
      await graphPost(listUrl + '/items', token, { fields: Object.assign(v.fields, {
        Ourref: ref, AddedByEmail: callerEmail,
      }) });
    } else if (action === 'update' || action === 'delete') {
      const id = String(body.id || '').trim();
      if (!/^\d+$/.test(id)) return fail(400, 'id is required.');
      const cur = shape(await graphGet(listUrl + '/items/' + id + '?$expand=fields($select=' + SELECT + ')', token));
      ref = cur.ref;
      if (cur.status !== 'open') {
        return fail(409, 'This expense is on ' + (cur.invoiceNumber ? 'invoice ' + cur.invoiceNumber : 'a draft invoice')
          + ' and cannot be changed. Cancel that ' + (cur.invoiceNumber ? 'invoice' : 'draft') + ' first.');
      }
      if (action === 'delete') {
        await graphDelete(listUrl + '/items/' + id, token);
      } else {
        const v = validate(Object.assign({
          date: cur.date, type: cur.type, description: cur.description, net: cur.net, vat: cur.vat, paidBy: cur.paidBy,
        }, stripUndefined({ date: body.date, type: body.type, description: body.description, net: body.net, vat: body.vat, paidBy: body.paidBy })), false);
        if (v.error) return fail(400, v.error);
        await graphPatch(listUrl + '/items/' + id + '/fields', token, v.fields);
      }
    } else {
      return fail(400, 'Unknown action.');
    }

    // ── Re-read and mirror the totals onto the case ─────────────────────────
    const rows = await rowsForRef(token, listUrl, ref);
    const t = totals(rows);
    let mirror = null, mirrorError = null;
    const caseItemId = String(body.caseItemId || '').trim();
    if (!/^\d+$/.test(caseItemId)) {
      mirrorError = 'No case id supplied — Settlement not updated.';
    } else {
      try {
        const cf = await graphGet(base + CASES_LIST + '/items/' + caseItemId + '/fields?$select=' + REF_FIELD, token);
        if (String(cf[REF_FIELD] || '').trim().toLowerCase() !== ref.toLowerCase()) {
          mirrorError = 'Case id does not match reference ' + ref + ' — Settlement not updated.';
        } else {
          mirror = { CourtFeePaidByTMC: t.tmc, ClientPaidExpenses: t.client };
          await graphPatch(base + CASES_LIST + '/items/' + caseItemId + '/fields', token, mirror);
        }
      } catch (e) {
        context.log.error('Expense mirror to Cases failed:', e.message);
        mirrorError = e.message; mirror = null;
      }
    }
    return send(200, { ref, rows, totals: t, mirror, mirrorError });

  } catch (err) {
    context.log.error('expenses error:', err.message);
    return fail(500, err.message);
  }
};

// ── Helpers ─────────────────────────────────────────────────────────────────

async function rowsForRef(token, listUrl, ref) {
  const url = listUrl + '/items?$expand=fields($select=' + SELECT + ')'
    + '&$filter=fields/Ourref eq \'' + ref.replace(/'/g, "''") + '\'&$top=500';
  const items = await graphGetAll(url, token);
  return items.map(shape)
    .filter(r => r.ref.toLowerCase() === ref.toLowerCase())
    .sort((a, b) => (a.date || '').localeCompare(b.date || '') || a.id - b.id);
}

function shape(item) {
  const f = item.fields || {};
  const net = round2(num(f.AmountNet)), vat = round2(num(f.VATAmount));
  const invoiceItemId = String(f.InvoiceItemId || '').trim();
  const invoiceNumber = String(f.InvoiceNumber || '').trim();
  return {
    id:            parseInt(item.id || f.id, 10),
    ref:           String(f.Ourref || '').trim(),
    date:          f.ExpenseDate ? String(f.ExpenseDate).slice(0, 10) : '',
    type:          String(f.ExpenseType || ''),
    description:   String(f.Title || ''),
    net, vat, gross: round2(net + vat),
    paidBy:        String(f.PaidBy || '') === 'TMC' ? 'TMC' : 'Client',
    invoiceItemId, invoiceNumber,
    status:        invoiceNumber ? 'invoiced' : (invoiceItemId ? 'draft' : 'open'),
    addedBy:       String(f.AddedByEmail || ''),
  };
}

function totals(rows) {
  const t = { client: 0, tmc: 0, tmcOpen: 0 };
  for (const r of rows) {
    if (r.paidBy === 'TMC') { t.tmc += r.gross; if (r.status === 'open') t.tmcOpen += r.gross; }
    else t.client += r.gross;
  }
  t.client = round2(t.client); t.tmc = round2(t.tmc); t.tmcOpen = round2(t.tmcOpen);
  return t;
}

function validate(b, isNew) {
  const description = String(b.description || '').trim();
  const type   = String(b.type || '').trim();
  const paidBy = String(b.paidBy || '').trim();
  const net    = round2(num(b.net)), vat = round2(num(b.vat));
  const date   = String(b.date || '').slice(0, 10);
  if (!description)                     return { error: 'Enter a description.' };
  if (!TYPES.includes(type))            return { error: 'Choose a type: ' + TYPES.join(', ') + '.' };
  if (!PAYERS.includes(paidBy))         return { error: 'Say who paid it: the client or TMC.' };
  if (!(net > 0))                       return { error: 'Enter the amount (ex VAT), above £0.' };
  if (vat < 0)                          return { error: 'VAT cannot be negative.' };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return { error: 'Enter the date paid.' };
  return { fields: {
    Title: description.slice(0, 255), ExpenseType: type, PaidBy: paidBy,
    AmountNet: net, VATAmount: vat, ExpenseDate: date + 'T12:00:00Z',
  } };
}

function stripUndefined(o) { const r = {}; for (const k in o) if (o[k] !== undefined) r[k] = o[k]; return r; }
function round2(n) { return Math.round((Number(n) || 0) * 100) / 100; }
function num(v) { const n = parseFloat(v); return isNaN(n) ? 0 : n; }

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

// ── Token ───────────────────────────────────────────────────────────────────────
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
    const r = https.request(options, function (res) {
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
    r.on('error', reject);
    r.write(body);
    r.end();
  });
}

// ── HTTP helpers ────────────────────────────────────────────────────────────────
function graphGet(url, token) {
  return new Promise(function (resolve, reject) {
    const u = new URL(url);
    const options = {
      hostname: u.hostname,
      path:     u.pathname + u.search,
      method:   'GET',
      headers: { Authorization: 'Bearer ' + token, Accept: 'application/json', Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' },
    };
    const r = https.request(options, function (res) {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        if (res.statusCode >= 400) {
          reject(new Error('Graph GET ' + res.statusCode + ' — ' + url.split('?')[0] + ': ' + data.slice(0, 400)));
          return;
        }
        try { resolve(JSON.parse(data)); }
        catch (e) { reject(new Error('JSON parse: ' + e.message)); }
      });
    });
    r.on('error', reject);
    r.end();
  });
}

// Follows @odata.nextLink so a >page list is never silently truncated.
async function graphGetAll(url, token) {
  let out = [];
  let next = url;
  let guard = 0;
  while (next && guard < 20) {
    const page = await graphGet(next, token);
    out = out.concat(page.value || []);
    next = page['@odata.nextLink'] || null;
    guard++;
  }
  return out;
}

function graphPost(url, token, body) {
  return new Promise(function (resolve, reject) {
    const payload = JSON.stringify(body);
    const u = new URL(url);
    const options = {
      hostname: u.hostname,
      path:     u.pathname + u.search,
      method:   'POST',
      headers: {
        Authorization:    'Bearer ' + token,
        'Content-Type':   'application/json',
        'Content-Length': Buffer.byteLength(payload),
      },
    };
    const r = https.request(options, function (res) {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        if (res.statusCode >= 400) {
          reject(new Error('Graph POST ' + res.statusCode + ' — ' + url.split('?')[0] + ': ' + data.slice(0, 400)));
          return;
        }
        try { resolve(data ? JSON.parse(data) : {}); }
        catch (e) { resolve({}); }
      });
    });
    r.on('error', reject);
    r.write(payload);
    r.end();
  });
}

function graphPatch(url, token, body, retries) {
  retries = retries === undefined ? 1 : retries;
  return new Promise(function (resolve, reject) {
    const payload = JSON.stringify(body);
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
    const r = https.request(options, function (res) {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        if (res.statusCode === 409 && retries > 0) {
          setTimeout(() => {
            graphPatch(url, token, body, retries - 1).then(resolve).catch(reject);
          }, 500);
          return;
        }
        if (res.statusCode >= 400) {
          reject(new Error('Graph PATCH ' + res.statusCode + ' — ' + url.split('?')[0] + ': ' + data.slice(0, 400)));
          return;
        }
        try { resolve(data ? JSON.parse(data) : {}); }
        catch (e) { resolve({}); }
      });
    });
    r.on('error', reject);
    r.write(payload);
    r.end();
  });
}

function graphDelete(url, token) {
  return new Promise(function (resolve, reject) {
    const u = new URL(url);
    const r = https.request({ hostname: u.hostname, path: u.pathname + u.search, method: 'DELETE',
      headers: { Authorization: 'Bearer ' + token } }, function (res) {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        if (res.statusCode >= 400) reject(new Error('Graph DELETE ' + res.statusCode + ': ' + data.slice(0, 400)));
        else resolve();
      });
    });
    r.on('error', reject);
    r.end();
  });
}
