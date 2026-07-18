/**
 * P124 — Invoicing Portal
 * Azure Function: /api/calendar        READ PATH ONLY (S82)
 *
 * The firm-wide shared calendar. STANDALONE — deliberately nothing to do with case
 * reminders or key dates. Holidays, court hearings, office closures, training.
 *
 * ── STORE ───────────────────────────────────────────────────────────────────
 * The calendar of the mailbox `automation@tmclegal.co.uk`, read app-only via Graph.
 * NOT a SharePoint list. Chosen so the same calendar appears in everyone's Outlook
 * and on their phones without the portal doing anything, and so there are no SP
 * field names to guess (the S80 `paymentsonaccount` trap).
 *
 * ── AUTHORISATION — READ THIS BEFORE CHANGING ANY PERMISSION ────────────────
 * Calendars.ReadWrite is granted ONLY through Exchange RBAC for Applications,
 * scoped to that one mailbox:
 *   New-ServicePrincipal        -AppId 93e76e96-82d5-469a-8295-3cdcc0780a8b
 *                               -ObjectId bd4076d2-2ac6-4d27-b674-6b434b858d75
 *   New-ManagementScope         -Name "TMC Automation Mailbox"
 *                               -RecipientRestrictionFilter
 *                                 "PrimarySmtpAddress -eq 'automation@tmclegal.co.uk'"
 *   New-ManagementRoleAssignment -App <ObjectId>
 *                               -Role "Application Calendars.ReadWrite"
 *                               -CustomResourceScope "TMC Automation Mailbox"
 * Verified 2026-07-18: Test-ServicePrincipalAuthorization returned InScope: True.
 *
 * 🔴 NEVER CONSENT Calendars.ReadWrite IN ENTRA. Entra grants and Exchange RBAC
 *    assignments are ADDITIVE — a union. An Entra grant is unscoped, so adding one
 *    silently gives this app read/write over EVERY mailbox calendar in the tenant
 *    and the mailbox restriction above stops meaning anything. Microsoft's own docs
 *    are explicit on this. (Application Access Policies are legacy and replaced by
 *    RBAC for Applications — do not use New-ApplicationAccessPolicy.)
 *
 * ⚠️ Permission changes are cached 30 min – 2 hrs. A 403 straight after a change is
 *    not proof of a broken grant; Test-ServicePrincipalAuthorization bypasses the
 *    cache and is the authority.
 *
 * GET /api/calendar?start=YYYY-MM-DD&end=YYYY-MM-DD
 *   start/end optional — defaults to the current month padded by one week each way
 *   so a month grid's leading and trailing days are populated.
 * Returns: { mailbox, start, end, count, events: [...] }
 */

const https   = require('https');
const { URL } = require('url');

const CAL_MAILBOX    = process.env.CALENDAR_MAILBOX || 'automation@tmclegal.co.uk';
const ALLOWED_DOMAIN = '@tmclegal.co.uk';
const TZ             = 'Europe/London';

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

// Everyone at TMC can READ the shared calendar — that is the point of it.
module.exports = async function (context, req) {
  context.log('P124 /api/calendar called');

  const callerEmail = getCallerEmail(req);
  if (!callerEmail || callerEmail.indexOf(ALLOWED_DOMAIN) === -1) {
    context.res = { status: 403, body: 'Forbidden — TMC staff only.' };
    return;
  }

  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    context.res = { status: 500, body: 'Missing required app settings.' };
    return;
  }

  const q     = req.query || {};
  const range = monthRange(q.start, q.end);

  try {
    const token  = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
    const events = [];

    // calendarView expands recurring series into individual occurrences — which is
    // what a calendar grid needs. /events would return the series master instead and
    // every weekly hearing would appear exactly once, on the wrong day.
    let url = 'https://graph.microsoft.com/v1.0/users/' + encodeURIComponent(CAL_MAILBOX)
            + '/calendarView'
            + '?startDateTime=' + encodeURIComponent(range.start)
            + '&endDateTime='   + encodeURIComponent(range.end)
            + '&$top=250&$orderby=start/dateTime';

    let guard = 0;
    while (url && guard++ < 10) {
      const page = await graphGet(url, token);
      ((page && page.value) || []).forEach(e => {
        events.push({
          id:        e.id || null,
          subject:   e.subject || '(no title)',
          start:     (e.start && e.start.dateTime) || null,
          end:       (e.end   && e.end.dateTime)   || null,
          timeZone:  (e.start && e.start.timeZone) || TZ,
          isAllDay:  !!e.isAllDay,
          location:  (e.location && e.location.displayName) || null,
          categories: e.categories || [],
          organizer: (e.organizer && e.organizer.emailAddress && e.organizer.emailAddress.address) || null,
          preview:   e.bodyPreview ? String(e.bodyPreview).slice(0, 300) : null,
          cancelled: !!e.isCancelled,
          webLink:   e.webLink || null,
        });
      });
      url = (page && page['@odata.nextLink']) || null;
    }

    context.res = {
      status: 200,
      headers: {
        'Content-Type':  'application/json',
        'Cache-Control': 'no-cache',
        // BUMP ON EVERY CHANGE TO THIS FILE (standing rule, S81).
        'X-Api-Build':   'S82-cal-v1-read',
        'X-Calendar-Mailbox': CAL_MAILBOX,
      },
      body: JSON.stringify({
        mailbox: CAL_MAILBOX,
        start:   range.start,
        end:     range.end,
        count:   events.length,
        events:  events,
      }),
    };
  } catch (err) {
    const msg = (err && err.message) || String(err);
    context.log.error('Error reading shared calendar:', msg);
    // Fail LOUDLY and specifically. A 403 here means the RBAC assignment is missing,
    // still cached, or the mailbox name is wrong — three different fixes, so say which
    // rather than returning an empty month that looks like "no events". (S75/S80.)
    const status = /Graph 403/.test(msg) ? 403
                 : /Graph 404/.test(msg) ? 404
                 : 500;
    context.res = {
      status: status,
      headers: { 'Content-Type': 'application/json', 'X-Api-Build': 'S82-cal-v1-read' },
      body: JSON.stringify({
        error:   msg.slice(0, 400),
        mailbox: CAL_MAILBOX,
        hint: status === 403
          ? 'App is not authorised for this mailbox. Check the Exchange RBAC role assignment '
            + '(Test-ServicePrincipalAuthorization -Resource ' + CAL_MAILBOX + '); permission '
            + 'changes cache for 30 min to 2 hrs.'
          : status === 404
          ? 'Mailbox not found. Confirm ' + CAL_MAILBOX + ' is a real mailbox, not an alias or '
            + 'distribution group (office@ is an alias — S76).'
          : 'Unexpected error — see message.',
      }),
    };
  }
};

// Default window: the current month, padded a week each side so a month grid's
// leading/trailing days are not silently empty.
function monthRange(startQ, endQ) {
  const iso = d => d.toISOString().slice(0, 19) + 'Z';
  const valid = s => /^\d{4}-\d{2}-\d{2}/.test(String(s || ''));

  if (valid(startQ) && valid(endQ)) {
    return { start: iso(new Date(startQ + 'T00:00:00Z')), end: iso(new Date(endQ + 'T23:59:59Z')) };
  }
  const now = new Date();
  const s   = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), 1));
  const e   = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth() + 1, 0, 23, 59, 59));
  s.setUTCDate(s.getUTCDate() - 7);
  e.setUTCDate(e.getUTCDate() + 7);
  return { start: iso(s), end: iso(e) };
}

// ─── TOKEN (client-credentials) ──────────────────────────
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

// ─── GRAPH GET ───────────────────────────────────────────
function graphGet(url, token) {
  return new Promise(function (resolve, reject) {
    const u = new URL(url);
    const options = {
      hostname: u.hostname,
      path:     u.pathname + u.search,
      method:   'GET',
      headers: {
        Authorization: 'Bearer ' + token,
        Accept:        'application/json',
        // Returns start/end as local wall-clock in this zone rather than UTC, so the
        // grid does not have to re-derive BST and get it wrong twice a year.
        Prefer:        'outlook.timezone="' + TZ + '"',
      },
    };

    const req = https.request(options, function (res) {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        if (res.statusCode >= 400) {
          reject(new Error('Graph ' + res.statusCode + ': ' + data.slice(0, 300)));
          return;
        }
        try { resolve(JSON.parse(data)); }
        catch (e) { reject(new Error('JSON parse error: ' + e.message)); }
      });
    });
    req.on('error', reject);
    req.end();
  });
}
