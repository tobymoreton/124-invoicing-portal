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
 * Access: ALL authenticated users (Admin/Finance/Draftsman) see the FULL WIP schedule,
 * unfiltered by who completed the work. (Prior per-draftsman filter on field_18 email
 * REMOVED 2026-07-01 — it silently dropped valid entries, including some the draftsman
 * had completed themselves, and this schedule must show all work to anyone raising an
 * invoice.) isAdmin/isFinance are still computed for potential future use but no longer
 * gate what WIP data is returned.
 *
 * List: Time Tracking2
 * GUID: 67db204c-30a5-4f4d-b276-60852d9967e1
 */

const https  = require('https');
const { URL } = require('url');

const LIST_GUID = '67db204c-30a5-4f4d-b276-60852d9967e1';
const SITE_PATH = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';

// NOTE: Do NOT restrict to a $select list on this tenant.
// Boolean fields (Billable_x003f_, Billed_x003f_) are silently dropped from
// Graph responses when explicitly named in $expand=fields($select=...).
// Restriction removed 2026-07-01 — use $expand=fields (all fields) instead.
// Confirmed pattern: same fix applied to /api/coa and /api/caseactions.
// SELECT_FIELDS constant retained here as documentation of which fields are
// consumed by the normalise() function, but it is NOT passed to Graph.
const SELECT_FIELDS_DOC = [
  'id',
  'Completedby_x0028_text_x0029_', // Draftsman name (text mirror of User field)
  'field_18',                       // Email (text) — draftsman email
  'Casename_x0028_text_x0029_',     // Case name (text mirror)
  'field_16',                       // Our reference
  'field_12',                       // Date Completed
  'TimeSpentMirror',                // Hours spent (Number mirror)
  'field_6',                        // Rate £/hr
  'Num_BillableAmount_x00a3_',      // WIP value £ — only populated post-billing; compute from TimeSpentMirror × field_6 for unbilled
  'field_2',                        // Work done (Note) — free-text description
  'Billable_x003f_',                // Billable? boolean (indexed) — NOT reliable in $select on this tenant
  'Billed_x003f_',                  // Billed? boolean (indexed) — NOT reliable in $select on this tenant
];

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
    const items   = await fetchAllWIP(token);

    context.res = {
      status: 200,
      headers: {
        'Content-Type': 'application/json',
        'Cache-Control': 'no-cache',
        // S81: build marker — lets a live measurement be attributed to a specific
        // build. Added because S81's "$orderby removal changed nothing" conclusion
        // could NOT be trusted: there was no way to tell the new build from the old.
        // BUMP THIS ON EVERY CHANGE TO THIS FILE.
        'X-Api-Build': 'S81-no-orderby',
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
// No caller-based filtering — WIP schedule must show ALL work to ANY authenticated
// user who can create an invoice, regardless of role or who completed the work.
// (Draftsman-only-sees-own-WIP restriction REMOVED 2026-07-01 — was filtering on
//  field_18 (email), which is not a reliable "who did the work" identifier and
//  was silently dropping valid entries, incl. some Tom Winyard himself completed.)
async function fetchAllWIP(token) {
  // Filter: not yet billed (Billable flag unreliable — many valid WIP entries have Billable=false)
  const wipFilter = `fields/Billed_x003f_ eq false`;

  const filter = encodeURIComponent(wipFilter);

  // $expand=fields with no $select — booleans drop silently when named in $select on this tenant.
  // Server-side $filter on Billed_x003f_ still works (it’s an indexed field query, not a response field).
  // S81: $orderby REMOVED — fetchAllWIP() re-sorts by DateCompleted desc in JS below, so the
  // SharePoint sort was pure duplicated work on a 3,562-row query. Output is unchanged.
  const base = `https://graph.microsoft.com/v1.0/sites/${SITE_PATH}/lists/${LIST_GUID}/items` +
               `?$expand=fields&$filter=${filter}&$top=999`;

  let url = base;
  let all = [];

  while (url) {
    const page  = await graphGet(url, token);
    const items = (page.value || []).map(normalise);
    all = all.concat(items);
    url = page['@odata.nextLink'] || null;
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
    Rate:          toNum(f.field_6),
    // WIPValue: use BillableAmount if populated (billed rows), else compute from hours × rate
    WIPValue:      toNum(f['Num_BillableAmount_x00a3_']) ||
                   (toNum(f.TimeSpentMirror) != null && toNum(f.field_6) != null
                     ? Math.round(toNum(f.TimeSpentMirror) * toNum(f.field_6) * 100) / 100
                     : null),
    WorkDone:        f.field_2                           || null,
    Billable:      f['Billable_x003f_']                || false,
    Billed:        f['Billed_x003f_']                  || false,
  };
}

function toNum(v) {
  const n = parseFloat(v);
  return isNaN(n) ? null : n;
}
