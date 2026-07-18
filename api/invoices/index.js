/**
 * P124 — Invoicing Portal
 * Azure Function: /api/invoices
 *
 * Fetches all items from the SP Invoice Library using client-credentials
 * (app-only Graph auth). Returns a normalised JSON array.
 *
 * ⚠️  READ-ONLY — no write operations in this function.
 * ⚠️  AmountOutstanding is a ReadOnly/calculated SP field — computed client-side instead.
 * ⚠️  Do NOT include Update, Cancel, or EditMetadata in any $select —
 *     these are Infowise action triggers on this list.
 * ⚠️  DraftedBy (User field) returns null via Graph app-only — use DraftedByEmail instead.
 *
 * List: Invoice Library
 * GUID: 5c366b19-0da9-4be9-b68f-60e6a0209cdb
 */

const https  = require('https');
const { URL } = require('url');

const LIST_GUID  = '5c366b19-0da9-4be9-b68f-60e6a0209cdb';
const SITE_PATH  = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';

const SELECT_FIELDS = [
  'id',
  'OrderDetails',
  'VendorName',
  'Casename',
  'Ourref',
  'InvoiceDate',
  'DueDate',
  'AmountDue',
  'AmountOutstanding',    // ReadOnly: explicit $select required
  'Status',
  'Status_Text',          // Mirror of calculated Status field — use this
  'Invoicetype',
  'LAorIP',
  'PaymentAmount1',
  'PaymentAmount2',
  'PaymentAmount3',
  'PaymentDate1',
  'PaymentDate2',
  'PaymentDate3',
  'Case_x0020_ID',
  'DraftedByEmail',       // Mirror of DraftedBy/Email — use this, not DraftedBy
  'DraftingFeeElement',   // Drafting fee portion of invoice
  'VAT',                  // VAT amount
  'Net',                  // Net amount (ex VAT)
  'Cancelled',            // Boolean — filter these out
  'Theirref',             // Client's own reference
].join(',');

const ADMIN_EMAILS = [
  'toby@tmclegal.co.uk',
  'danielle@tmclegal.co.uk',
];

// Finance tier: sees all draftsman billing but not the main invoice ledger/table
const FINANCE_EMAILS = [
  'lesley@tmclegal.co.uk',
];

// ---- Default response window (S81) ----------------------------------------
// The full ledger is 3,014 invoices / ~2.08 MB and the client paginates 18 rows
// at a time. By default this function returns only the rows a consumer can
// actually need; ?all=1 returns the complete history, unchanged.
//
// A row is IN the window if ANY limb is true:
//   1. Not settled: |AmountOutstanding| > 0.005. Deliberately "not equal to zero",
//      NOT "greater than zero", so OVERPAID invoices (negative outstanding)
//      survive -- the anomaly panel's 'overpaid' rule depends on them.
//   2. Raised within WINDOW_MONTHS.
//   3. Any payment banked within WINDOW_MONTHS. Without this limb an old, fully
//      settled invoice paid this month drops out and index.html's "Paid this
//      month" stat silently under-reports. Measured 2026-07-18: this limb alone
//      carries 99 of the 509 windowed rows.
//
// Measured live 2026-07-18 as toby@: 3,014 rows / 2,080.8 KB -> 509 / 357.1 KB.
const WINDOW_MONTHS = 12;

function inDefaultWindow(inv, cutoff) {
  const o = parseFloat(inv.AmountOutstanding);
  if (!isNaN(o) && Math.abs(o) > 0.005) return true;
  if (inv.InvoiceDate && new Date(inv.InvoiceDate) >= cutoff) return true;
  for (const k of [1, 2, 3]) {
    const d = inv['PaymentDate' + k];
    if (d && new Date(d) >= cutoff) return true;
  }
  return false;
}

// Decode the x-ms-client-principal header injected by Azure SWA
function getCallerEmail(req) {
  try {
    const header = req.headers && req.headers['x-ms-client-principal'];
    if (!header) return null;
    const decoded = Buffer.from(header, 'base64').toString('utf8');
    const principal = JSON.parse(decoded);
    // userDetails is the most reliable field for AAD (always contains UPN/email)
    if (principal.userDetails) return principal.userDetails.toLowerCase();
    // Fallback: hunt through claims
    const claim = (principal.claims || []).find(
      c => c.typ === 'preferred_username' || c.typ === 'email' || c.typ === 'upn'
        || c.typ === 'http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress'
    );
    return claim ? claim.val.toLowerCase() : null;
  } catch { return null; }
}

module.exports = async function (context, req) {
  context.log('P124 /api/invoices called');

  // Identity check
  const callerEmail = getCallerEmail(req);
  if (!callerEmail) {
    context.res = { status: 403, body: 'Forbidden — could not determine caller identity.' };
    return;
  }
  const isAdmin   = ADMIN_EMAILS.includes(callerEmail);
  const isFinance  = FINANCE_EMAILS.includes(callerEmail);

  // Case-scoped read (S68). ?ref=<Our Reference> returns every invoice raised on that ONE
  // case to any authenticated TMC user, ignoring the drafted-by scope below. Deliberate:
  // the case Invoicing tab must show the case's full billing picture, not just the caller's
  // own bills. Scope is a single case — this is not a back door to the whole ledger.
  const refParam = ((req.query && req.query.ref) || '').trim();

  // ?all=1 -- full history, no window. Used by index.html's admin-only background
  // pull that feeds the anomaly panel (its 'zeroValue' and 'payNoDate' rules can
  // hit fully-settled invoices of any age).
  const wantAll = String((req.query && req.query.all) || '') === '1';

  // ?billedrefs=1 -- distinct Our Refs carrying a drafting fee on a live invoice.
  // Serves cases.html's billedRefs set (the 4%-of-outstanding-PC card excludes
  // already-billed cases and needs EVERY ref ever billed, at any age). 1,538 refs
  // / ~15 KB against 2,080.8 KB for the same job. Measured 2026-07-18.
  const wantBilledRefs = String((req.query && req.query.billedrefs) || '') === '1';

  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    context.res = { status: 500, body: 'Missing required app settings.' };
    return;
  }

  try {
    const token    = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
    // Admins + finance get all invoices; non-admins get only their own (by DraftedByEmail)
    const invoices = await fetchAllInvoices(token, (isAdmin || isFinance || refParam) ? null : callerEmail);
    if (wantBilledRefs) {
      const refs = [...new Set(
        invoices
          .filter(inv => !inv.Cancelled && (parseFloat(inv.DraftingFeeElement) || 0) > 0)
          .map(inv => inv.Ourref)
          .filter(Boolean)
      )];
      context.res = {
        status: 200,
        headers: {
          'Content-Type': 'application/json',
          'Cache-Control': 'no-cache',
          'X-Invoice-Window': 'billedrefs',
        },
        body: JSON.stringify(refs),
      };
      return;
    }

    let payload = refParam
      ? invoices.filter(inv => (inv.Ourref || '').toString().trim().toLowerCase() === refParam.toLowerCase())
      : invoices;

    // Case-scoped reads are NEVER windowed -- the case Invoicing tab must show the
    // whole billing history for that one case.
    let windowed = false;
    if (!refParam && !wantAll) {
      const cutoff = new Date();
      cutoff.setMonth(cutoff.getMonth() - WINDOW_MONTHS);
      payload = payload.filter(inv => inDefaultWindow(inv, cutoff));
      windowed = true;
    }

    context.res = {
      status: 200,
      headers: {
        'Content-Type': 'application/json',
        'Cache-Control': 'no-cache',
        'X-Invoice-Window': windowed ? (WINDOW_MONTHS + 'm') : 'all',
      },
      body: JSON.stringify(payload),
    };
  } catch (err) {
    context.log.error('Error fetching invoices:', err.message);
    context.res = { status: 500, body: `Error: ${err.message}` };
  }
};

// ─── TOKEN (client-credentials) ──────────────────────────
function getToken(tenantId, clientId, clientSecret) {
  return new Promise((resolve, reject) => {
    const body = new URLSearchParams({
      grant_type:    'client_credentials',
      client_id:     clientId,
      client_secret: clientSecret,
      scope:         'https://graph.microsoft.com/.default',
    }).toString();

    const options = {
      hostname: 'login.microsoftonline.com',
      path:     `/${tenantId}/oauth2/v2.0/token`,
      method:   'POST',
      headers: {
        'Content-Type':   'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(body),
      },
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

// ─── FETCH ALL (handles pagination) ──────────────────────
// callerEmailFilter: null = return all (admin); string = return only that draftsman's invoices
async function fetchAllInvoices(token, callerEmailFilter) {
  const base = `https://graph.microsoft.com/v1.0/sites/${SITE_PATH}/lists/${LIST_GUID}/items` +
               `?$expand=fields($select=${SELECT_FIELDS})&$top=500`;

  let url = base;
  let all = [];

  while (url) {
    const page  = await graphGet(url, token);
    const items = (page.value || []).map(normalise);
    all = all.concat(items);
    url = page['@odata.nextLink'] || null;
  }

  // Non-admin/non-finance: filter to caller's own invoices only (server-side security)
  if (callerEmailFilter) {
    all = all.filter(inv => (inv.DraftedByEmail || '').toLowerCase() === callerEmailFilter);
  }

  // Sort by DueDate asc — null dates go last
  all.sort((a, b) => {
    const da = a.DueDate ? new Date(a.DueDate).getTime() : Infinity;
    const db = b.DueDate ? new Date(b.DueDate).getTime() : Infinity;
    return da - db;
  });

  return all;
}

// ─── GRAPH GET ───────────────────────────────────────────
function graphGet(url, token) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    const options = {
      hostname: u.hostname,
      path:     u.pathname + u.search,
      method:   'GET',
      headers: {
        Authorization:    `Bearer ${token}`,
        Accept:           'application/json',
        ConsistencyLevel: 'eventual',
      },
    };

    const req = https.request(options, res => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        if (res.statusCode >= 400) {
          reject(new Error(`Graph ${res.statusCode}: ${data.slice(0, 200)}`));
          return;
        }
        try { resolve(JSON.parse(data)); }
        catch (e) { reject(new Error(`JSON parse error: ${e.message}`)); }
      });
    });
    req.on('error', reject);
    req.end();
  });
}

// ─── NORMALISE ───────────────────────────────────────────
function normalise(item) {
  const f = item.fields || {};
  return {
    _id:                String(item.id),
    OrderDetails:       f.OrderDetails        || null,
    VendorName:         f.VendorName          || null,
    Casename:           f.Casename            || null,
    Ourref:             f.Ourref              || null,
    InvoiceDate:        f.InvoiceDate         || null,
    DueDate:            f.DueDate             || null,
    AmountDue:          toNum(f.AmountDue),
    AmountOutstanding:  computeOutstanding(f),
    Status:             f.Status              || null,
    Status_Text:        f.Status_Text         || null,
    Invoicetype:        f.Invoicetype         || null,
    LAorIP:             f.LAorIP              || null,
    PaymentAmount1:     toNum(f.PaymentAmount1),
    PaymentAmount2:     toNum(f.PaymentAmount2),
    PaymentAmount3:     toNum(f.PaymentAmount3),
    PaymentDate1:       f.PaymentDate1        || null,
    PaymentDate2:       f.PaymentDate2        || null,
    PaymentDate3:       f.PaymentDate3        || null,
    Case_x0020_ID:      f.Case_x0020_ID       || null,
    DraftedByEmail:     f.DraftedByEmail      || null,
    DraftingFeeElement: toNum(f.DraftingFeeElement),
    VAT:                toNum(f.VAT),
    Net:                toNum(f.Net),
    Cancelled:          f.Cancelled           || false,
    Theirref:           f.Theirref            || null,
    FileUrl:            item.webUrl           || null,
  };
}

function toNum(v) {
  const n = parseFloat(v);
  return isNaN(n) ? null : n;
}

// AmountOutstanding is a calculated SP field — Graph returns null for it.
// Compute client-side: AmountDue minus all recorded payments.
function computeOutstanding(f) {
  const due = parseFloat(f.AmountDue);
  if (isNaN(due)) return null;
  const p1 = parseFloat(f.PaymentAmount1) || 0;
  const p2 = parseFloat(f.PaymentAmount2) || 0;
  const p3 = parseFloat(f.PaymentAmount3) || 0;
  return Math.round((due - p1 - p2 - p3) * 100) / 100;
}
