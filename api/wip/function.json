/**
 * P124 — Invoicing Portal
 * Azure Function: /api/wip
 *
 * Fetches unbilled billable actions from the SP Time Tracking2 list using
 * client-credentials (app-only Graph auth). Returns a normalised JSON array.
 *
 * WIP definition: Billable? = true AND Billed? = false
 * Both fields are indexed in SP — filtered server-side via Graph $filter.
 *
 * ⚠️  READ-ONLY — no write operations in this function.
 * ⚠️  Completedby (User field) returns null via Graph app-only auth.
 *     Use Completedby_x0028_text_x0029_ (Text mirror) instead.
 * ⚠️  BillableAmount£ (Num_BillableAmount_x00a3_) is a native Currency field — safe for Graph.
 *
 * Three-tier access (must stay in sync with api/invoices and api/lineitems):
 *   ADMIN   — all WIP items
 *   FINANCE — all WIP items (leaderboard view, no invoice table)
 *   DRAFTSMAN — own WIP only, filtered server-side by Completedby_x0028_text_x0029_
 *
 * List: Time Tracking2
 * GUID: 67db204c-30a5-4f4d-b276-60852d9967e1
 */

const https  = require('https');
const { URL } = require('url');

const LIST_GUID = '67db204c-30a5-4f4d-b276-60852d9967e1';
const SITE_PATH = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';

const SELECT_FIELDS = [
  'id',
  'Completedby_x0028_text_x0029_', // Draftsman name (text mirror of User field)
  'field_18',                       // Email (text) — draftsman email
  'Casename_x0028_text_x0029_',     // Case name (text mirror)
  'field_16',                       // Our reference
  'field_12',                       // Date Completed
  'TimeSpentMirror',                // Hours spent (Number mirror)
  'Num_BillableAmount_x00a3_',      // WIP value £ (native Currency field)
  'Billable_x003f_',                // Billable? boolean (indexed)
  'Billed_x003f_',                  // Billed? boolean (indexed)
].join(',');

const ADMIN_EMAILS = [
  'toby@tmclegal.co.uk',
  'danielle@tmclegal.co.uk',
];

// Finance tier: sees all draftsman WIP but not the main invoice ledger/table
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
  context.log('P124 /api/wip called');

  // Identity check
  const callerEmail = getCallerEmail(req);
  if (!callerEmail) {
    context.res = { status: 403, body: 'Forbidden — could not determine caller identity.' };
    return;
  }

  const isAdmin   = ADMIN_EMAILS.includes(callerEmail);
  const isFinance = FINANCE_EMAILS.includes(callerEmail);

  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    context.res = { status: 500, body: 'Missing required app settings.' };
    return;
  }

  try {
    const token   = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
    const items   = await fetchAllWIP(token, isAdmin, isFinance, callerEmail);

    context.res = {
      status: 200,
      headers: {
        'Content-Type': 'application/json',
        'Cache-Control': 'no-cache',
      },
      body: JSON.stringify(items),
    };
  } catch (err) {
    context.log.error('Error fetching WIP:', err.message);
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

// ─── FETCH ALL WIP (handles pagination) ──────────────────
// Both Billable? and Billed? are indexed — safe to use in Graph $filter.
// Draftsmen are also filtered server-side by email where possible.
// NOTE: Completedby_x0028_text_x0029_ contains the draftsman name (text), not email.
//       field_18 contains the email. Both selected; draftsman filter uses field_18.
async function fetchAllWIP(token, isAdmin, isFinance, callerEmail) {
  // Base WIP filter: billable and not yet billed
  const wipFilter = `fields/Billable_x003f_ eq true and fields/Billed_x003f_ eq false`;

  // Draftsmen additionally filtered by their email (field_18)
  // UNCERTAIN if field_18 is indexed — if Graph rejects it, fall back to client-side filter
  const emailClause = (!isAdmin && !isFinance)
    ? ` and fields/field_18 eq '${callerEmail}'`
    : '';

  const filter = encodeURIComponent(wipFilter + emailClause);

  const base = `https://graph.microsoft.com/v1.0/sites/${SITE_PATH}/lists/${LIST_GUID}/items` +
               `?$expand=fields($select=${SELECT_FIELDS})&$filter=${filter}&$top=500`;

  let url = base;
  let all = [];

  while (url) {
    const page  = await graphGet(url, token);
    const items = (page.value || []).map(normalise);
    all = all.concat(items);
    url = page['@odata.nextLink'] || null;
  }

  // Fallback client-side filter for draftsmen in case field_18 is not indexed
  if (!isAdmin && !isFinance) {
    all = all.filter(item =>
      (item.Email || '').toLowerCase() === callerEmail ||
      (item.DraftsmanName || '').toLowerCase().startsWith(callerEmail.split('@')[0].replace('.', ' '))
    );
  }

  // Sort by DateCompleted desc (most recent first)
  all.sort((a, b) => {
    const da = a.DateCompleted ? new Date(a.DateCompleted).getTime() : 0;
    const db = b.DateCompleted ? new Date(b.DateCompleted).getTime() : 0;
    return db - da;
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
    _id:           String(item.id),
    DraftsmanName: f['Completedby_x0028_text_x0029_'] || null, // text mirror of User field
    Email:         f.field_18                          || null, // draftsman email
    CaseName:      f['Casename_x0028_text_x0029_']    || null, // text mirror of case name
    OurRef:        f.field_16                          || null,
    DateCompleted: f.field_12                          || null,
    HoursSpent:    toNum(f.TimeSpentMirror),
    WIPValue:      toNum(f['Num_BillableAmount_x00a3_']),
    Billable:      f['Billable_x003f_']                || false,
    Billed:        f['Billed_x003f_']                  || false,
  };
}

function toNum(v) {
  const n = parseFloat(v);
  return isNaN(n) ? null : n;
}
