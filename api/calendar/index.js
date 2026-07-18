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
 *
 * POST /api/calendar        create an event (S82 write path)
 *   { subject, date, endDate?, startTime?, endTime?, allDay?, who?, location?, notes?, category? }
 *   `who` is who the entry is FOR (validated against the staff roster; blank = firm-wide)
 *   and is prefixed into the subject so it is legible in Outlook and on phones.
 *   Everyone at TMC may create, and may create FOR anyone — an admin booking leave on
 *   a draftsman's behalf is a normal case, and the creator is recorded regardless.
 *   The creator's email is written into the body because app-only writes appear in
 *   Outlook as the mailbox itself (the S73 identity-trail lesson).
 *
 * PATCH /api/calendar?id=<eventId>     amend an event (same body shape as POST)
 * DELETE /api/calendar?id=<eventId>    remove an event
 *   Anyone at TMC may amend or remove anything — it is a shared calendar, and an entry
 *   nobody but its author can correct is an entry that stays wrong. Every amendment
 *   appends to the audit trail in the body rather than overwriting it. Deletes go to
 *   the mailbox's Deleted Items and are recoverable from Outlook.
 */

const https   = require('https');
const { URL } = require('url');

const CAL_MAILBOX    = process.env.CALENDAR_MAILBOX || 'automation@tmclegal.co.uk';
const ALLOWED_DOMAIN = '@tmclegal.co.uk';
const TZ             = 'Europe/London';
// BUMP ON EVERY CHANGE TO THIS FILE (standing rule, S81).
const BUILD          = 'S82-cal-v4-crud';

// Who an entry is FOR — distinct from who created it. Mirrors the portal roster
// (case.html PERSON_EMAILS) minus David, a leaver: historic data is not an issue
// here because this calendar starts empty. Validated SERVER-side so a hand-rolled
// POST cannot inject arbitrary text into the subject line everyone sees.
const STAFF = {
  toby:     'Toby',     danielle: 'Danielle', lesley: 'Lesley',
  joanna:   'Joanna',   tracy:    'Tracy',    kelly:  'Kelly',
  tom:      'Tom',      julie:    'Julie',    daniel: 'Daniel',
};
function resolveWho(raw) {
  const s = String(raw || '').trim().toLowerCase();
  if (!s || s === 'firm-wide' || s === 'everyone') return { ok: true, name: null };
  const token = s.indexOf('@') > -1 ? s.split('@')[0] : s;
  // Exact token match only — never substring: 'daniel' is a substring of 'danielle'
  // and that exact bug bled Daniel's cases into Danielle's bucket (S65).
  if (STAFF[token]) return { ok: true, name: STAFF[token] };
  return { ok: false, name: null };
}

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
  context.log('P124 /api/calendar called, method=' + req.method);

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

  const method = (req.method || 'GET').toUpperCase();
  if (method === 'POST') {
    await createEvent(context, req, callerEmail, TENANT_ID, CLIENT_ID, CLIENT_SECRET);
    return;
  }
  if (method === 'PATCH') {
    await updateEvent(context, req, callerEmail, TENANT_ID, CLIENT_ID, CLIENT_SECRET);
    return;
  }
  if (method === 'DELETE') {
    await deleteEvent(context, req, callerEmail, TENANT_ID, CLIENT_ID, CLIENT_SECRET);
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
          preview:   e.bodyPreview ? String(e.bodyPreview).slice(0, 1000) : null,
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
        'X-Api-Build':   BUILD,
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
      headers: { 'Content-Type': 'application/json', 'X-Api-Build': BUILD },
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

// ─── SHARED PAYLOAD BUILDER ──────────────────────────────
// Used by BOTH create and amend. Kept in one place deliberately: two copies of the
// all-day half-open boundary is two chances to get it wrong, and only one of them
// would be noticed.
function buildEventPayload(b, callerEmail, verb) {
  const subject = String(b.subject || '').trim();
  const date    = String(b.date || '').trim();
  if (!subject)                          return { error: 'subject is required' };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return { error: 'date is required as YYYY-MM-DD' };

  const who = resolveWho(b.who);
  if (!who.ok) {
    return { error: 'Unrecognised person: ' + String(b.who).slice(0, 60),
             hint:  'Must be a current member of staff, or blank for firm-wide.' };
  }

  const endDate = /^\d{4}-\d{2}-\d{2}$/.test(String(b.endDate || '')) ? String(b.endDate) : date;
  if (endDate < date) return { error: 'endDate is before date' };

  const allDay = b.allDay !== false && !b.startTime;   // default TRUE unless a time is given
  let start, end;

  if (allDay) {
    // Graph all-day events are half-open: end is the day AFTER the last day, at 00:00.
    // Getting this wrong is the classic off-by-one that makes a one-day holiday vanish.
    const dayAfter = new Date(endDate + 'T00:00:00Z');
    dayAfter.setUTCDate(dayAfter.getUTCDate() + 1);
    start = { dateTime: date + 'T00:00:00',                                timeZone: TZ };
    end   = { dateTime: dayAfter.toISOString().slice(0, 10) + 'T00:00:00', timeZone: TZ };
  } else {
    const st = /^\d{2}:\d{2}$/.test(String(b.startTime || '')) ? String(b.startTime) : null;
    const et = /^\d{2}:\d{2}$/.test(String(b.endTime   || '')) ? String(b.endTime)   : null;
    if (!st) return { error: 'startTime is required as HH:MM when the entry is not all-day' };
    const endT = et || addMinutes(st, 60);
    if (endDate === date && endT <= st) return { error: 'endTime is not after startTime' };
    start = { dateTime: date    + 'T' + st   + ':00', timeZone: TZ };
    end   = { dateTime: endDate + 'T' + endT + ':00', timeZone: TZ };
  }

  // The person's name goes in the SUBJECT, not just the body — it has to be legible
  // in Outlook and on a phone, where nobody opens the entry to find out whose leave
  // it is. 'Kelly — Annual leave'.
  const stamp = new Date().toISOString().slice(0, 16).replace('T', ' ');
  const notes = String(b.notes || '').trim();
  const trail = String(b.trail || '').trim();   // prior audit lines, preserved on amend
  const body  = (notes ? notes + '\n\n' : '')
              + (who.name ? 'For: ' + who.name + '\n' : 'For: firm-wide\n')
              + '\u2014\n'
              + (trail ? trail + '\n' : '')
              + (verb === 'amend' ? 'Amended' : 'Added')
              + ' via the TMC portal by ' + callerEmail + ' on ' + stamp + ' UTC.';

  const fullSubject = (who.name ? who.name + ' \u2014 ' : '') + subject;

  const payload = {
    subject:   fullSubject.slice(0, 255),
    isAllDay:  allDay,
    start:     start,
    end:       end,
    body:      { contentType: 'text', content: body },
    showAs:    allDay ? 'free' : 'busy',
  };
  // Sent even when blank on amend, so clearing a location or category actually clears it.
  payload.location   = { displayName: String(b.location || '').slice(0, 255) };
  payload.categories = b.category ? [String(b.category).slice(0, 64)] : [];

  return { payload: payload, who: who.name, subject: subject };
}

// ─── CREATE (POST) ───────────────────────────────────────
async function createEvent(context, req, callerEmail, tenantId, clientId, clientSecret) {
  const b = req.body || {};
  const fail = (status, error, hint) => {
    context.res = {
      status: status,
      headers: { 'Content-Type': 'application/json', 'X-Api-Build': BUILD },
      body: JSON.stringify({ error: error, hint: hint || null }),
    };
  };

  const built = buildEventPayload(b, callerEmail, 'create');
  if (built.error) return fail(400, built.error, built.hint);
  const payload = built.payload;

  try {
    const token = await getToken(tenantId, clientId, clientSecret);
    const url   = 'https://graph.microsoft.com/v1.0/users/'
                + encodeURIComponent(CAL_MAILBOX) + '/events';
    const made  = await graphPost(url, token, payload);

    context.log('Calendar event created by ' + callerEmail + ': ' + built.subject);
    context.res = {
      status: 201,
      headers: { 'Content-Type': 'application/json', 'X-Api-Build': BUILD },
      body: JSON.stringify({
        ok: true,
        id: made && made.id,
        subject: made && made.subject,
        start: made && made.start,
        end: made && made.end,
        isAllDay: !!(made && made.isAllDay),
        webLink: made && made.webLink,
        who: built.who,
        createdBy: callerEmail,
      }),
    };
  } catch (err) {
    const msg = (err && err.message) || String(err);
    context.log.error('Calendar create failed:', msg);
    const status = /Graph 403/.test(msg) ? 403 : /Graph 404/.test(msg) ? 404 : 500;
    fail(status, msg.slice(0, 400), graphHint(status));
  }
}

// ─── AMEND (PATCH) ───────────────────────────────────────
async function updateEvent(context, req, callerEmail, tenantId, clientId, clientSecret) {
  const b  = req.body || {};
  const id = String((req.query && req.query.id) || b.id || '').trim();
  const fail = (status, error, hint) => {
    context.res = {
      status: status,
      headers: { 'Content-Type': 'application/json', 'X-Api-Build': BUILD },
      body: JSON.stringify({ error: error, hint: hint || null }),
    };
  };
  if (!id) return fail(400, 'id is required');

  const built = buildEventPayload(b, callerEmail, 'amend');
  if (built.error) return fail(400, built.error, built.hint);

  try {
    const token = await getToken(tenantId, clientId, clientSecret);
    const url   = 'https://graph.microsoft.com/v1.0/users/'
                + encodeURIComponent(CAL_MAILBOX) + '/events/' + encodeURIComponent(id);

    // Read the existing body FIRST so the audit trail survives the amendment. Doing
    // this server-side rather than round-tripping it through the browser means the
    // trail cannot be truncated, dropped or edited by whoever is amending.
    let trail = '';
    try {
      const existing = await graphGet(url + '?$select=body', token);
      const prior = (existing && existing.body && existing.body.content) || '';
      const marker = prior.indexOf('\u2014');
      if (marker > -1) {
        trail = prior.slice(marker + 1)
          .replace(/<[^>]*>/g, '')      // body may come back as HTML
          .split(/\r?\n/).map(s => s.trim())
          .filter(s => /via the TMC portal by /.test(s))
          .join('\n');
      }
    } catch (e) { /* no trail recoverable — the amendment line still gets written */ }

    const built2 = trail
      ? buildEventPayload(Object.assign({}, b, { trail: trail }), callerEmail, 'amend')
      : built;
    const done = await graphPatch(url, token, built2.payload || built.payload);

    context.log('Calendar event amended by ' + callerEmail + ': ' + built.subject);
    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'X-Api-Build': BUILD },
      body: JSON.stringify({
        ok: true,
        id: done && done.id,
        subject: done && done.subject,
        start: done && done.start,
        end: done && done.end,
        isAllDay: !!(done && done.isAllDay),
        who: built.who,
        amendedBy: callerEmail,
      }),
    };
  } catch (err) {
    const msg = (err && err.message) || String(err);
    context.log.error('Calendar amend failed:', msg);
    const status = /Graph 40[34]/.test(msg) ? (/Graph 403/.test(msg) ? 403 : 404) : 500;
    fail(status, msg.slice(0, 400),
      status === 404 ? 'That entry no longer exists — it may already have been deleted.'
                     : graphHint(status));
  }
}

// ─── DELETE ──────────────────────────────────────────────
// A real delete, not a cancelled-flag: this calendar has no attendees to notify, and
// Outlook keeps the item in the mailbox's Deleted Items, so it is recoverable.
async function deleteEvent(context, req, callerEmail, tenantId, clientId, clientSecret) {
  const id = String((req.query && req.query.id) || '').trim();
  const fail = (status, error, hint) => {
    context.res = {
      status: status,
      headers: { 'Content-Type': 'application/json', 'X-Api-Build': BUILD },
      body: JSON.stringify({ error: error, hint: hint || null }),
    };
  };
  if (!id) return fail(400, 'id is required');

  try {
    const token = await getToken(tenantId, clientId, clientSecret);
    const url   = 'https://graph.microsoft.com/v1.0/users/'
                + encodeURIComponent(CAL_MAILBOX) + '/events/' + encodeURIComponent(id);
    await graphDelete(url, token);

    context.log('Calendar event deleted by ' + callerEmail + ': ' + id.slice(0, 40));
    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'X-Api-Build': BUILD },
      body: JSON.stringify({ ok: true, deleted: true, deletedBy: callerEmail }),
    };
  } catch (err) {
    const msg = (err && err.message) || String(err);
    context.log.error('Calendar delete failed:', msg);
    const status = /Graph 403/.test(msg) ? 403 : /Graph 404/.test(msg) ? 404 : 500;
    // 404 on delete is not an error worth shouting about — the entry is gone either way.
    if (status === 404) {
      context.res = {
        status: 200,
        headers: { 'Content-Type': 'application/json', 'X-Api-Build': BUILD },
        body: JSON.stringify({ ok: true, deleted: false, note: 'Entry was already gone.' }),
      };
      return;
    }
    fail(status, msg.slice(0, 400), graphHint(status));
  }
}

function graphHint(status) {
  return status === 403
    ? 'RBAC assignment missing or still cached (30 min – 2 hrs). Calendars.ReadWrite must '
      + 'be granted via Exchange RBAC only — never consented in Entra.'
    : status === 404
    ? 'Not found — confirm ' + CAL_MAILBOX + ' is a mailbox, not an alias.'
    : 'Unexpected error — see message.';
}

function addMinutes(hhmm, mins) {
  const [h, m] = hhmm.split(':').map(Number);
  const t = ((h * 60 + m + mins) % 1440 + 1440) % 1440;
  return String(Math.floor(t / 60)).padStart(2, '0') + ':' + String(t % 60).padStart(2, '0');
}

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

// ─── GRAPH POST ────────────────────────────────────────
function graphPost(url, token, payload) {
  return new Promise(function (resolve, reject) {
    const u    = new URL(url);
    const data = JSON.stringify(payload);
    const options = {
      hostname: u.hostname,
      path:     u.pathname + u.search,
      method:   'POST',
      headers: {
        Authorization:   'Bearer ' + token,
        Accept:          'application/json',
        'Content-Type':  'application/json',
        'Content-Length': Buffer.byteLength(data),
        Prefer:          'outlook.timezone="' + TZ + '"',
      },
    };

    const req = https.request(options, function (res) {
      let out = '';
      res.on('data', chunk => { out += chunk; });
      res.on('end', () => {
        if (res.statusCode >= 400) {
          reject(new Error('Graph ' + res.statusCode + ': ' + out.slice(0, 300)));
          return;
        }
        try { resolve(JSON.parse(out)); }
        catch (e) { reject(new Error('JSON parse error: ' + e.message)); }
      });
    });
    req.on('error', reject);
    req.write(data);
    req.end();
  });
}

// ─── GRAPH PATCH / DELETE ────────────────────────────────
function graphSend(method, url, token, payload) {
  return new Promise(function (resolve, reject) {
    const u    = new URL(url);
    const data = payload ? JSON.stringify(payload) : null;
    const headers = {
      Authorization: 'Bearer ' + token,
      Accept:        'application/json',
      Prefer:        'outlook.timezone="' + TZ + '"',
    };
    if (data) {
      headers['Content-Type']   = 'application/json';
      headers['Content-Length'] = Buffer.byteLength(data);
    }

    const req = https.request(
      { hostname: u.hostname, path: u.pathname + u.search, method: method, headers: headers },
      function (res) {
        let out = '';
        res.on('data', chunk => { out += chunk; });
        res.on('end', () => {
          if (res.statusCode >= 400) {
            reject(new Error('Graph ' + res.statusCode + ': ' + out.slice(0, 300)));
            return;
          }
          // DELETE returns 204 with an empty body — not an error, and not JSON.
          if (!out) { resolve({}); return; }
          try { resolve(JSON.parse(out)); }
          catch (e) { reject(new Error('JSON parse error: ' + e.message)); }
        });
      }
    );
    req.on('error', reject);
    if (data) req.write(data);
    req.end();
  });
}
function graphPatch(url, token, payload) { return graphSend('PATCH', url, token, payload); }
function graphDelete(url, token)         { return graphSend('DELETE', url, token, null); }

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
