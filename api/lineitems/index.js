/**
 * P124 — Invoicing Portal
 * Azure Function: /api/lineitems
 *
 * Fetches all billable items from the SP Invoice | Line Items list using
 * client-credentials (app-only Graph auth). Returns a normalised JSON array.
 *
 * ⚠️  READ-ONLY — no write operations in this function.
 * ⚠️  CompletedBy (User field) returns null via Graph app-only auth.
 *     Use CompletedByEmail (Text mirror) instead — populated by PA099.12.
 * ⚠️  Value (Calculated field) returns null via Graph.
 *     Use ValueMirror (Currency mirror) instead — populated by PA099.12.
 *
 * Required App Settings (same as /api/invoices):
 *   TENANT_ID      — Entra tenant ID (GUID)
 *   CLIENT_ID      — App registration client ID
 *   CLIENT_SECRET  — App registration client secret
 *   SITE_ID        — SharePoint site ID
 *
 * List: Invoice | Line Items
 * GUID: 496468a5-e2ed-48db-8826-58cb08844eee
 */

const https  = require('https');
const { URL } = require('url');

const LIST_GUID  = '496468a5-e2ed-48db-8826-58cb08844eee';
// Invoice Library — read only, and only for the OrderDetails → InvoiceDate map (S107b).
const INVOICE_LIB = '5c366b19-0da9-4be9-b68f-60e6a0209cdb';
const SITE_PATH  = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';

// Fields to retrieve
// CompletedByEmail and ValueMirror are mirror fields populated by PA099.12
// CompletedBy and Value return null via Graph app-only — do not use
const SELECT_FIELDS = [
  'id',
  'field_1',              // Work done
  'field_2',              // Time spent (units)
  'field_3',              // Rate (£/hr)
  'ProRataApportionment', // Percentage (e.g. 57 = 57%)
  'CompletedByEmail',     // Mirror of CompletedBy/Email — use this, not CompletedBy
  'ValueMirror',          // Mirror of Value calculated field — use this, not Value
  'InvoiceIDRef',         // Links to Invoice Library OrderDetails (invoice number)
  'CaseName',             // Case name
  'field_5',              // Our reference
  'Completed_x0020_on',   // Date completed
  'BillableYorN_x0020__x2753_', // Billable boolean
  'InvoiceType',          // Choice: Time Only, Drafting & Time
  'InvoiceDate',          // Invoice date
].join(',');

const ADMIN_EMAILS = [
  'toby@tmclegal.co.uk',
  'danielle@tmclegal.co.uk',
];

// Finance tier: sees all draftsman billing but not the main invoice ledger/table
const FINANCE_EMAILS = [
  'lesley@tmclegal.co.uk',
];

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
  context.log('P124 /api/lineitems called');

  // Resolve caller identity
  const callerEmail = getCallerEmail(req);
  const isAdmin     = callerEmail && ADMIN_EMAILS.includes(callerEmail);
  const isFinance   = callerEmail && FINANCE_EMAILS.includes(callerEmail);

  // Non-admins must be authenticated
  if (!callerEmail) {
    context.res = { status: 403, body: 'Forbidden — authentication required.' };
    return;
  }

  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    context.res = { status: 500, body: 'Missing required app settings.' };
    return;
  }

  try {
    const token     = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
    // S107b: the parent invoice's date, keyed by invoice number. See fetchInvoiceDates.
    const invDates  = await fetchInvoiceDates(token);
    const lineItems = await fetchAllLineItems(token, isAdmin, isFinance, callerEmail, invDates);

    context.res = {
      status: 200,
      headers: {
        'Content-Type': 'application/json',
        'Cache-Control': 'no-cache',
        // BUMP ON EVERY CHANGE TO THIS FILE (standing rule, S81).
        'X-Api-Build': 'S107b-linevalue-and-date-fallback',
        // Rows whose ValueMirror was blank and had to be computed. Non-zero means PA099.12
        // is not keeping up — worth watching, not an error.
        'X-Derived-Lines': String(lineItems.filter(i => i.ValueDerived).length),
        // Rows whose own InvoiceDate was blank and was taken from the parent invoice.
        'X-Derived-Dates': String(lineItems.filter(i => i.InvoiceDateDerived).length),
      },
      body: JSON.stringify(lineItems),
    };
  } catch (err) {
    context.log.error('Error fetching line items:', err.message);
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
// S107b: `issueinvoice` stamps the invoice number onto a Line Item but NOT the invoice date
// (step 9 patches InvoiceIDRef only). 1,888 of 3,864 billed line items therefore have no
// InvoiceDate of their own — and index.html's Draftsman Billing panel opens with
// `if (!li.InvoiceDate) return;`, so those lines are dropped before the month filter and a
// draftsman's timed work silently disappears from the table, leaving only invoice-level VAT.
//
// The date is taken from the parent invoice, which is the authoritative source anyway: the
// line item's copy is only ever a snapshot of it. Costs one extra paged read of the Invoice
// Library per call. Deliberate — doing this in the page instead would fix one consumer and
// leave the next one broken, which is the mistake made earlier in this same session when
// /api/invoices was fixed and this endpoint was not.
async function fetchInvoiceDates(token) {
  const dates = new Map();
  let url = `https://graph.microsoft.com/v1.0/sites/${SITE_PATH}/lists/${INVOICE_LIB}/items` +
            `?$expand=fields($select=OrderDetails,InvoiceDate)&$top=999`;

  while (url) {
    const page = await graphGet(url, token);
    for (const item of (page.value || [])) {
      const f   = item.fields || {};
      const num = String(f.OrderDetails || '').trim();
      if (num && f.InvoiceDate) dates.set(num, f.InvoiceDate);
    }
    url = page['@odata.nextLink'] || null;
  }

  return dates;
}

async function fetchAllLineItems(token, isAdmin, isFinance, callerEmail, invDates) {
  // Admins + finance get all line items; non-admins filtered by CompletedByEmail server-side
  // BillableYorN filter removed from Graph query — not indexed, filter client-side instead
  const emailFilter = (isAdmin || isFinance)
    ? ''
    : `fields/CompletedByEmail eq '${callerEmail}'`;

  const filterParam = emailFilter ? `&$filter=${encodeURIComponent(emailFilter)}` : '';

  const base = `https://graph.microsoft.com/v1.0/sites/${SITE_PATH}/lists/${LIST_GUID}/items` +
               `?$expand=fields($select=${SELECT_FIELDS})${filterParam}&$top=500`;

  let url = base;
  let all = [];

  while (url) {
    const page  = await graphGet(url, token);
    const items = (page.value || []).map(normalise).filter(i => i.Billable);
    // Fill a blank InvoiceDate from the parent invoice. Only ever FILLS — a date the row
    // already holds is left exactly as it is.
    if (invDates && invDates.size) {
      for (const i of items) {
        if (i.InvoiceDate || !i.InvoiceIDRef) continue;
        const d = invDates.get(String(i.InvoiceIDRef).trim());
        if (d) { i.InvoiceDate = d; i.InvoiceDateDerived = true; }
      }
    }
    all = all.concat(items);
    url = page['@odata.nextLink'] || null;
  }

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
// Flattens a SP list item into a plain object.
// Uses mirror fields (CompletedByEmail, ValueMirror) not the originals.
//
// S107: ValueMirror is written by PA099.12. When that flow misses a row the mirror is blank
// and every consumer of this endpoint scores the line at £0 — index.html's Draftsman Billing
// panel reads `li.Value||0`, so a draftsman's whole month of timed work showed as a dash with
// only the VAT column populated. Measured 2026-08-06: 62 of 3,894 rows had no mirror.
// So derive hours × rate × pro-rata when the mirror is missing. A stored figure that IS
// present is never overridden — on a billed line that is what was actually invoiced.
// `ValueDerived` says which happened, so a consumer that cares can tell them apart.
function normalise(item) {
  const f       = item.fields || {};
  const stored  = toNum(f.ValueMirror);
  const derived = stored === null;
  return {
    _id:               String(item.id),
    WorkDone:          f.field_1                          || null,
    TimeSpent:         toNum(f.field_2),                           // units (1 unit = 6 min)
    Rate:              toNum(f.field_3),                           // £/hr
    ProRata:           toNum(f.ProRataApportionment),              // percentage
    CompletedByEmail:  f.CompletedByEmail                 || null, // mirror — not CompletedBy
    Value:             derived ? lineValueFromInputs(f) : stored,  // mirror — not Value
    ValueDerived:      derived,                                    // true = mirror was blank
    InvoiceIDRef:      f.InvoiceIDRef                     || null,
    CaseName:          f.CaseName                         || null,
    OurRef:            f.field_5                          || null,
    CompletedOn:       f.Completed_x0020_on               || null,
    Billable:          f.BillableYorN_x0020__x2753_       || false,
    InvoiceType:       f.InvoiceType                      || null,
    InvoiceDate:       f.InvoiceDate                      || null,
    InvoiceDateDerived: false,                                     // true if filled from parent
  };
}

// Hours × rate × pro-rata, matching the SharePoint calculated column `Billable_x0020_amount`,
// which applies ProRataApportionment as a /100 factor. Blank or zero pro-rata means 100%
// — a missing apportionment must not zero the line. Same helper as /api/invoices.
function lineValueFromInputs(f) {
  const hrs  = parseFloat(f.field_2) || 0;
  const rate = parseFloat(f.field_3) || 0;
  let   pro  = parseFloat(f.ProRataApportionment);
  if (isNaN(pro) || pro <= 0) pro = 100;
  return Math.round(hrs * rate * (pro / 100) * 100) / 100;
}

function toNum(v) {
  const n = parseFloat(v);
  return isNaN(n) ? null : n;
}
