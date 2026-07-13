/**
 * P124 — Invoicing Portal
 * Azure Function: /api/payment
 *
 * Records, edits and clears payments against an Invoice Library item.
 * Management only (Toby, Danielle).
 *
 * The Invoice Library has THREE payment slots: PaymentAmount1/2/3 + PaymentDate1/2/3.
 * There is no fourth. A request that needs one is REJECTED, not silently dropped.
 *
 * ⚠️  Status and AmountOutstanding are CALCULATED SP columns. This function NEVER
 *     writes either of them. Writing the payment is enough — SharePoint recomputes.
 *     Do not "help" by patching Status.
 *
 * POST body (JSON):
 *   listItemId      string        — Invoice Library SP list item ID   (required)
 *   slot            1 | 2 | 3     — target slot. Omit to use the first free one.
 *   amount          number | null — payment amount. null/0/'' = CLEAR the slot.
 *   date            'YYYY-MM-DD'  — payment date. Required when amount > 0.
 *   confirmOverpay  boolean       — must be true to book a payment that takes the
 *                                   invoice below zero outstanding.
 *
 * Returns 200: { ok, slot, cleared, invoice: { ...read-back fields... } }
 * Returns 400: { error, code }  code = NO_FREE_SLOT | OVERPAY | BAD_INPUT | NOT_ISSUED
 *
 * List: Invoice Library — 5c366b19-0da9-4be9-b68f-60e6a0209cdb
 */

const https   = require('https');
const { URL } = require('url');

const SITE_PATH   = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';
const INVOICE_LIB = '5c366b19-0da9-4be9-b68f-60e6a0209cdb';

// Management tier only. Deliberately NOT the ISSUE_EMAILS list — Lesley issues and
// cancels invoices but does not apply payments (Toby's call, S73).
const PAYMENT_EMAILS = ['toby@tmclegal.co.uk', 'danielle@tmclegal.co.uk'];

const READ_BACK = [
  'id', 'OrderDetails', 'AmountDue', 'Cancelled', 'InvoiceDate', 'Status', 'Status_Text',
  'PaymentAmount1', 'PaymentAmount2', 'PaymentAmount3',
  'PaymentDate1', 'PaymentDate2', 'PaymentDate3',
];

module.exports = async function (context, req) {
  context.log('P124 /api/payment called');

  const callerEmail = getCallerEmail(req);
  if (!callerEmail || !PAYMENT_EMAILS.includes(callerEmail)) {
    context.res = { status: 403, body: 'Forbidden — not authorised to apply payments.' };
    return;
  }

  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    context.res = { status: 500, body: 'Missing required app settings.' };
    return;
  }

  let body;
  try {
    body = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
  } catch (e) {
    context.res = { status: 400, body: 'Invalid JSON body.' };
    return;
  }
  body = body || {};

  const listItemId = body.listItemId;
  if (!listItemId) { bad(context, 'Missing required field: listItemId.', 'BAD_INPUT'); return; }

  // amount: null / '' / 0  => CLEAR the slot
  const rawAmount = body.amount;
  const clearing  = (rawAmount === null || rawAmount === undefined || rawAmount === ''
                     || parseFloat(rawAmount) === 0);
  const amount    = clearing ? null : Math.round(parseFloat(rawAmount) * 100) / 100;

  if (!clearing && (isNaN(amount) || amount < 0)) {
    bad(context, 'Amount must be a positive number, or blank to clear the payment.', 'BAD_INPUT');
    return;
  }

  let slot = (body.slot === undefined || body.slot === null) ? null : parseInt(body.slot, 10);
  if (slot !== null && [1, 2, 3].indexOf(slot) === -1) {
    bad(context, 'slot must be 1, 2 or 3.', 'BAD_INPUT');
    return;
  }
  if (clearing && slot === null) {
    bad(context, 'Clearing a payment requires an explicit slot.', 'BAD_INPUT');
    return;
  }

  // Date — required when booking an amount. Stored as SP DateTime, date-only, midnight UTC.
  let dateIso = null;
  if (!clearing) {
    const d = new Date(String(body.date || '') + 'T00:00:00Z');
    if (!body.date || isNaN(d.getTime())) {
      bad(context, 'A valid payment date (YYYY-MM-DD) is required.', 'BAD_INPUT');
      return;
    }
    dateIso = d.toISOString();
  }

  try {
    const token = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);

    // ── 1. Read the invoice as it stands ──────────────────────────────────────
    const fieldsUrl = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
      + '/lists/' + INVOICE_LIB + '/items/' + listItemId + '/fields';
    const f = await graphGet(fieldsUrl, token);

    if (f.Cancelled === true) {
      bad(context, 'This invoice is cancelled — a payment cannot be recorded against it.', 'NOT_ISSUED');
      return;
    }
    if (!f.OrderDetails || !f.InvoiceDate) {
      bad(context, 'This invoice is still a draft — issue it before recording a payment.', 'NOT_ISSUED');
      return;
    }

    const paid = function (n) {
      const v = parseFloat(f['PaymentAmount' + n]);
      return isNaN(v) ? 0 : v;
    };

    // ── 2. Resolve the slot ───────────────────────────────────────────────────
    if (slot === null) {
      slot = [1, 2, 3].find(n => paid(n) === 0 && !f['PaymentDate' + n]) || null;
      if (slot === null) {
        bad(context,
          'All three payment slots on this invoice are in use. Edit an existing payment instead.',
          'NO_FREE_SLOT');
        return;
      }
    }

    // ── 3. Overpayment guard — warn, never silently allow ─────────────────────
    if (!clearing) {
      const due   = parseFloat(f.AmountDue) || 0;
      const other = [1, 2, 3].filter(n => n !== slot).reduce((s, n) => s + paid(n), 0);
      const newOutstanding = Math.round((due - other - amount) * 100) / 100;
      if (newOutstanding < 0 && body.confirmOverpay !== true) {
        context.res = {
          status: 400,
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            error: 'This payment would overpay the invoice by £'
              + Math.abs(newOutstanding).toFixed(2) + '.',
            code: 'OVERPAY',
            amountDue: due,
            otherPayments: other,
            newOutstanding: newOutstanding,
          }),
        };
        return;
      }
    }

    // ── 4. Write. Payment fields ONLY — never Status, never AmountOutstanding. ─
    const patch = {};
    patch['PaymentAmount' + slot] = clearing ? null : amount;
    patch['PaymentDate'   + slot] = clearing ? null : dateIso;

    await graphPatch(fieldsUrl, token, patch);
    context.log((clearing ? 'Cleared' : 'Recorded') + ' payment slot ' + slot
      + ' on invoice ' + (f.OrderDetails || listItemId) + ' by ' + callerEmail);

    // ── 5. Read back — the UI renders from SharePoint, never from optimism ─────
    const after = await graphGet(fieldsUrl + '?$select=' + READ_BACK.join(','), token);

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: JSON.stringify({ ok: true, slot: slot, cleared: clearing, invoice: after }),
    };

  } catch (err) {
    context.log.error('payment error:', err.message, err.stack);
    context.res = { status: 500, body: 'Error: ' + err.message };
  }
};

function bad(context, message, code) {
  context.res = {
    status: 400,
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ error: message, code: code }),
  };
}

// ── Identity ──────────────────────────────────────────────────────────────────

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

// ── Token ─────────────────────────────────────────────────────────────────────

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

// ── HTTP helpers ──────────────────────────────────────────────────────────────

function graphGet(url, token) {
  return new Promise(function (resolve, reject) {
    const u = new URL(url);
    const options = {
      hostname: u.hostname,
      path:     u.pathname + u.search,
      method:   'GET',
      headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' },
    };
    const req = https.request(options, function (res) {
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
    req.on('error', reject);
    req.end();
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
    const req = https.request(options, function (res) {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        // 409 resourceModified — PA099.14 patches this library on every change and
        // bumps the eTag. Retry once.
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
    req.on('error', reject);
    req.write(payload);
    req.end();
  });
}
