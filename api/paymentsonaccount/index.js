/**
 * P124 — Invoicing Portal
 * Azure Function: /api/paymentsonaccount
 *
 * The "unallocated payments" / payments-on-account (suspense) register.
 * A client pays a sum with NO remittance advice, so it is unknown which
 * invoice(s) it settles. This records that money as a first-class row that
 * stays visible on the accounts page until it is allocated — so it can be
 * chased instead of quietly lost.
 *
 * Management only (Toby, Danielle) — same gate as /api/payment, deliberately
 * NOT the ISSUE_EMAILS list (Lesley issues invoices but does not apply money).
 *
 * ⚠️  This function ONLY manages rows in the PaymentsOnAccount list.
 *     ALLOCATING money onto an invoice is done SEPARATELY by the caller via
 *     the existing /api/payment (which fills the invoice's PaymentAmountN slot).
 *     The 'allocate' action here just records, on the on-account row, how much
 *     of it has now been placed and against which invoice(s). Two writes, one
 *     per source of truth — never double-book.
 *
 * ⚠️  Status is a plain CHOICE column here (Unallocated / Part allocated /
 *     Allocated) — the portal owns it. Remaining (= AmountReceived −
 *     AmountAllocated) is computed here, never stored (no Calculated column,
 *     which Graph app-only cannot read).
 *
 * Routes:
 *   GET  /api/paymentsonaccount            → all rows (+ computed remaining)
 *   GET  /api/paymentsonaccount?open=1     → rows not fully allocated only
 *   POST /api/paymentsonaccount            → body.action:
 *        'add'      : { clientFirmName, amountReceived, dateReceived, note? }
 *        'allocate' : { id, amount, invoiceRef, confirmOverAllocate? }
 *
 * List: PaymentsOnAccount — e7d7f036-6242-4e01-856b-b43b59b83b4d
 *
 * ⚠️  FIELD NAMES BELOW ARE INFERRED (created via the SP UI with no spaces, so
 *     internal name == display name). If any Graph write 400s on a field, that
 *     is the first thing to verify against _api/web/lists/getbytitle(...)/fields.
 */

const https   = require('https');
const { URL } = require('url');

const SITE_PATH = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';
const POA_LIST  = 'e7d7f036-6242-4e01-856b-b43b59b83b4d';

// Management tier only — mirrors /api/payment (PAYMENT_EMAILS, S73).
const POA_EMAILS = ['toby@tmclegal.co.uk', 'danielle@tmclegal.co.uk'];

// Field internal names — single point of change. INFERRED == display names.
const F = {
  title:     'Title',
  firm:      'ClientFirmName',
  received:  'AmountReceived',
  allocated: 'AmountAllocated',
  date:      'DateReceived',
  note:      'Note',
  status:    'Status',
  allocTo:   'AllocatedTo',
  addedBy:   'AddedByEmail',
};

const STATUS = { UNALLOC: 'Unallocated', PART: 'Part allocated', DONE: 'Allocated' };

module.exports = async function (context, req) {
  context.log('P124 /api/paymentsonaccount called (' + req.method + ')');

  const callerEmail = getCallerEmail(req);
  if (!callerEmail || !POA_EMAILS.includes(callerEmail)) {
    context.res = { status: 403, body: 'Forbidden — not authorised for payments on account.' };
    return;
  }

  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    context.res = { status: 500, body: 'Missing required app settings.' };
    return;
  }

  const listUrl = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
    + '/lists/' + POA_LIST;

  try {
    const token = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);

    // ── GET — list rows ───────────────────────────────────────────────────────
    if (req.method === 'GET') {
      const openOnly = String((req.query && req.query.open) || '') === '1';
      const url = listUrl + '/items?$expand=fields&$top=500&$orderby=fields/'
        + F.date + ' desc';
      const rows = await graphGetAll(url, token);

      const out = rows.map(r => shape(r)).filter(r => !openOnly || r.remaining > 0.005);
      const totalUnallocated = out.reduce((s, r) => s + (r.remaining > 0 ? r.remaining : 0), 0);

      context.res = {
        status: 200,
        headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
        body: JSON.stringify({
          ok: true,
          count: out.length,
          totalUnallocated: Math.round(totalUnallocated * 100) / 100,
          rows: out,
        }),
      };
      return;
    }

    // ── POST — add / allocate ─────────────────────────────────────────────────
    if (req.method === 'POST') {
      let body;
      try { body = typeof req.body === 'string' ? JSON.parse(req.body) : req.body; }
      catch (e) { bad(context, 'Invalid JSON body.', 'BAD_INPUT'); return; }
      body = body || {};
      const action = (body.action || 'add').toLowerCase();

      // ---- ADD -----------------------------------------------------------------
      if (action === 'add') {
        const firm = String(body.clientFirmName || '').trim();
        const amount = Math.round(parseFloat(body.amountReceived) * 100) / 100;
        if (!firm)                       { bad(context, 'clientFirmName is required.', 'BAD_INPUT'); return; }
        if (isNaN(amount) || amount <= 0){ bad(context, 'amountReceived must be a positive number.', 'BAD_INPUT'); return; }

        const d = new Date(String(body.dateReceived || '') + 'T00:00:00Z');
        if (!body.dateReceived || isNaN(d.getTime())) {
          bad(context, 'A valid dateReceived (YYYY-MM-DD) is required.', 'BAD_INPUT'); return;
        }
        const yyyymmdd = body.dateReceived.replace(/-/g, '');
        const slug = firm.replace(/[^A-Za-z0-9]+/g, '').slice(0, 20) || 'FIRM';

        const fields = {};
        fields[F.title]     = 'POA-' + yyyymmdd + '-' + slug;
        fields[F.firm]      = firm;
        fields[F.received]  = amount;
        fields[F.allocated] = 0;
        fields[F.date]      = d.toISOString();
        fields[F.note]      = String(body.note || '').trim();
        fields[F.status]    = STATUS.UNALLOC;
        fields[F.allocTo]   = '';
        fields[F.addedBy]   = callerEmail;

        const created = await graphPost(listUrl + '/items', token, { fields: fields });
        context.log('POA added by ' + callerEmail + ': £' + amount + ' from ' + firm);
        context.res = {
          status: 200,
          headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
          body: JSON.stringify({ ok: true, action: 'add', row: shape(created) }),
        };
        return;
      }

      // ---- ALLOCATE ------------------------------------------------------------
      // Records, on the on-account row, that `amount` has now been placed against
      // `invoiceRef`. The money itself must ALREADY have been written to the
      // invoice slot via /api/payment by the caller — this only tracks it.
      if (action === 'allocate') {
        const id = body.id;
        const amount = Math.round(parseFloat(body.amount) * 100) / 100;
        const invoiceRef = String(body.invoiceRef || '').trim();
        if (!id)                          { bad(context, 'id is required.', 'BAD_INPUT'); return; }
        if (isNaN(amount) || amount <= 0) { bad(context, 'amount must be a positive number.', 'BAD_INPUT'); return; }
        if (!invoiceRef)                  { bad(context, 'invoiceRef is required.', 'BAD_INPUT'); return; }

        const itemUrl = listUrl + '/items/' + id + '/fields';
        const cur = await graphGet(itemUrl, token);

        const received  = num(cur[F.received]);
        const already   = num(cur[F.allocated]);
        const remaining = Math.round((received - already) * 100) / 100;

        if (amount - remaining > 0.005 && body.confirmOverAllocate !== true) {
          context.res = {
            status: 400,
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              error: 'This would allocate £' + amount.toFixed(2)
                + ' but only £' + remaining.toFixed(2) + ' is unallocated on this payment.',
              code: 'OVER_ALLOCATE',
              remaining: remaining,
            }),
          };
          return;
        }

        const newAllocated = Math.round((already + amount) * 100) / 100;
        const newRemaining = Math.round((received - newAllocated) * 100) / 100;
        const newStatus = newRemaining <= 0.005 ? STATUS.DONE : STATUS.PART;

        const stamp = new Date().toISOString().slice(0, 10);
        const priorLog = String(cur[F.allocTo] || '').trim();
        const newLog = (priorLog ? priorLog + '\n' : '')
          + stamp + ': £' + amount.toFixed(2) + ' → ' + invoiceRef + ' (' + callerEmail + ')';

        const patch = {};
        patch[F.allocated] = newAllocated;
        patch[F.status]    = newStatus;
        patch[F.allocTo]   = newLog;

        await graphPatch(itemUrl, token, patch);
        const after = await graphGet(itemUrl, token);
        context.log('POA ' + id + ' allocated £' + amount + ' → ' + invoiceRef + ' by ' + callerEmail);
        context.res = {
          status: 200,
          headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
          body: JSON.stringify({ ok: true, action: 'allocate', row: shape({ id: id, fields: after }) }),
        };
        return;
      }

      bad(context, "Unknown action '" + action + "'. Use 'add' or 'allocate'.", 'BAD_INPUT');
      return;
    }

    context.res = { status: 405, body: 'Method not allowed.' };

  } catch (err) {
    context.log.error('paymentsonaccount error:', err.message, err.stack);
    context.res = { status: 500, body: 'Error: ' + err.message };
  }
};

// ── Shape a Graph list item into the portal row the UI renders ─────────────────
function shape(item) {
  const f = (item && item.fields) || {};
  const received  = num(f[F.received]);
  const allocated = num(f[F.allocated]);
  return {
    id:        item.id,
    firm:      f[F.firm]   || '',
    received:  received,
    allocated: allocated,
    remaining: Math.round((received - allocated) * 100) / 100,
    date:      f[F.date]   || null,
    note:      f[F.note]   || '',
    status:    f[F.status] || STATUS.UNALLOC,
    allocatedTo: f[F.allocTo] || '',
    addedBy:   f[F.addedBy] || '',
    title:     f[F.title]  || '',
  };
}

function num(v) { const n = parseFloat(v); return isNaN(n) ? 0 : n; }

function bad(context, message, code) {
  context.res = {
    status: 400,
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ error: message, code: code }),
  };
}

// ── Identity ───────────────────────────────────────────────────────────────────
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
      headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' },
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
