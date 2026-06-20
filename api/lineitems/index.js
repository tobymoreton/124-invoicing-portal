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

const LIST_GUID = '496468a5-e2ed-48db-8826-58cb08844eee';

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
  'lesley@tmclegal.co.uk',
  'danielle@tmclegal.co.uk',
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

  // Non-admins must be authenticated
  if (!callerEmail) {
    context.res = { status: 403, body: 'Forbidden — authentication required.' };
    return;
  }

  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET, SITE_ID } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET || !SITE_ID) {
    context.res = { status: 500, body: 'Missing required app settings.' };
    return;
  }

  try {
    const token     = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
    const lineItems = await fetchAllLineItems(token, SITE_ID, isAdmin, callerEmail);

    context.res = {
      status: 200,
      headers: {
        'Content-Type': 'application/json',
        'Cache-Control': 'no-cache',
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
async function fetchAllLineItems(token, siteId, isAdmin, callerEmail) {
  // Admins get all billable items; draftsmen get only their own
  const emailFilter = isAdmin
    ? 'fields/BillableYorN_x0020__x2753_ eq true'
    : `fields/BillableYorN_x0020__x2753_ eq true and fields/CompletedByEmail eq '${callerEmail}'`;

  const base = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${LIST_GUID}/items` +
               `?$expand=fields($select=${SELECT_FIELDS})&$filter=${encodeURIComponent(emailFilter)}&$top=500`;

  let url = base;
  let all = [];

  while (url) {
    const page  = await graphGet(url, token);
    const items = (page.value || []).map(normalise);
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
function normalise(item) {
  const f = item.fields || {};
  return {
    _id:               String(item.id),
    WorkDone:          f.field_1                          || null,
    TimeSpent:         toNum(f.field_2),                           // units (1 unit = 6 min)
    Rate:              toNum(f.field_3),                           // £/hr
    ProRata:           toNum(f.ProRataApportionment),              // percentage
    CompletedByEmail:  f.CompletedByEmail                 || null, // mirror — not CompletedBy
    Value:             toNum(f.ValueMirror),                       // mirror — not Value
    InvoiceIDRef:      f.InvoiceIDRef                     || null,
    CaseName:          f.CaseName                         || null,
    OurRef:            f.field_5                          || null,
    CompletedOn:       f.Completed_x0020_on               || null,
    Billable:          f.BillableYorN_x0020__x2753_       || false,
    InvoiceType:       f.InvoiceType                      || null,
    InvoiceDate:       f.InvoiceDate                      || null,
  };
}

function toNum(v) {
  const n = parseFloat(v);
  return isNaN(n) ? null : n;
}
