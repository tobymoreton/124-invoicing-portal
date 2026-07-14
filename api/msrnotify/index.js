/**
 * P124 — Invoicing Portal
 * Azure Function: /api/msrnotify
 *
 * Emails the draftsmen who still have UNVERIFIED matters on a Matter Status Report, with a deep
 * link straight to their own filtered view of it, so they can check, edit and verify their rows
 * before the report goes to the client.
 *
 *   POST body (JSON):
 *     firm        string   — instructing firm, for the subject line
 *     link        string   — base MSR URL (e.g. https://<swa>/msr.html) — used to build each
 *                            recipient's own ?firm=&draftsman= deep link
 *     recipients  array    — [{ name: 'Tom', count: 3 }, ...]  (canonical roster first names)
 *
 * Management only (Toby + Danielle) — this sends mail to other people.
 *
 * ⚠ RECIPIENTS ARE NOT TRUSTED. The client sends NAMES, never addresses. Every name is resolved
 *   against the hardcoded roster below and anything not on it is dropped. The portal therefore
 *   cannot be made to email an arbitrary address, even if the page were tampered with.
 *
 * ⚠ REQUIRES `Mail.Send` (Application) on the app registration, with admin consent.
 *   Without it Graph returns 403 and no code change here will help. Added 2026-07-14.
 *
 * Sends FROM office@tmclegal.co.uk — the same mailbox PA099.14 uses for invoice notifications,
 * so it reads as a system notification rather than a personal email.
 */

const https   = require('https');
const { URL } = require('url');

// The mailbox these emails are SENT FROM.
//
// ⚠ This must be a real MAILBOX, not a distribution group, an alias or a mail-enabled security
//   group. Graph /users/{addr}/sendMail resolves the address to a user object; anything that is not
//   one comes back 404 ErrorInvalidUser ("The requested user '...' is invalid").
//   office@tmclegal.co.uk was tried first, on the strength of PA099.14 sending invoice
//   notifications to it — but PA099.14 DELIVERS TO it, which proves nothing about whether it can
//   be SENT AS. It failed with exactly that 404 (2026-07-14).
//
// automation@tmclegal.co.uk is the Power Automate service account (the same identity the PA
// loop-breakers test for), so it is a real account rather than an alias — and it reads as a system
// notification rather than a personal email, which is the point.
//
// Overridable without a code change via the MSR_SEND_FROM app setting on the Function App.
const SEND_FROM = process.env.MSR_SEND_FROM || 'automation@tmclegal.co.uk';

// Who gets told when a report is fully verified. Management, not the draftsmen.
const COMPLETE_TO = 'toby@tmclegal.co.uk';

// Management only — this sends mail on behalf of the firm.
//
// EXCEPTION: the 'complete' mode is open to ANY signed-in draftsman, because it is fired by the
// browser of whoever happens to verify LAST — and that will usually not be Management. It can only
// ever send to COMPLETE_TO above, so it cannot be used to mail anyone else.
const ALLOWED_EMAILS = [
  'toby@tmclegal.co.uk',
  'danielle@tmclegal.co.uk',
];

// The ONLY addresses this function will ever send to. Canonical roster first name -> mailbox.
//
// Danielle is deliberately absent: assigning a case to her is the hand-back / unallocated bucket,
// not a draftsman with work to verify. David is a leaver. Lesley is Finance, not a draftsman.
const ROSTER_EMAILS = {
  Toby:   'toby@tmclegal.co.uk',
  Tom:    'tom@tmclegal.co.uk',
  Tracy:  'tracy@tmclegal.co.uk',
  Joanna: 'joanna@tmclegal.co.uk',
  Kelly:  'kelly@tmclegal.co.uk',
  Julie:  'julie@tmclegal.co.uk',
  Daniel: 'daniel@tmclegal.co.uk',
};

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

function esc(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

module.exports = async function (context, req) {
  context.log('P124 /api/msrnotify called');

  const callerEmail = getCallerEmail(req);
  if (!callerEmail || !callerEmail.endsWith('@tmclegal.co.uk')) {
    context.res = { status: 403, body: 'Forbidden — you must be signed in.' };
    return;
  }
  // Chase mode mails other people, so it stays Management-only. Completion mode can only ever mail
  // COMPLETE_TO, so any signed-in draftsman may trigger it (see the note above).
  const isComplete = req.body && String(req.body.mode || '') === 'complete';
  if (!isComplete && !ALLOWED_EMAILS.includes(callerEmail)) {
    context.res = { status: 403, body: 'Forbidden — notifying the draftsmen is restricted to Management.' };
    return;
  }

  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    context.res = { status: 500, body: 'Missing required app settings.' };
    return;
  }

  const body = req.body || {};
  const mode = String(body.mode || 'chase').trim();
  const firm = String(body.firm || '').trim();
  const link = String(body.link || '').trim();
  const list = Array.isArray(body.recipients) ? body.recipients : [];

  if (!firm || !link) {
    context.res = { status: 400, body: 'Missing firm or link.' };
    return;
  }
  if (!/^https:\/\//i.test(link)) {
    context.res = { status: 400, body: 'link must be an https URL.' };
    return;
  }

  // ── mode: 'complete' ── every draftsman with matters on this report has now verified them.
  // Fired by the browser of whoever verifies LAST, once their write lands and the reload comes back
  // with nothing outstanding.
  //
  // ⚠ NOT idempotent. If two draftsmen verify their last rows within the same few seconds, both
  //   browsers can see a clean report and both will fire — the email may arrive twice. Deduping
  //   properly needs a "completion notified" stamp in SharePoint; not built. A duplicate email is a
  //   far cheaper failure than a missed one, so it fires optimistically.
  if (mode === 'complete') {
    const total = parseInt(body.total, 10) || 0;
    try {
      const token = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
      const deep  = link + (link.indexOf('?') === -1 ? '?' : '&') + 'firm=' + encodeURIComponent(firm);
      const html =
        '<p>Every draftsman with matters on the <strong>' + esc(firm) + '</strong> Matter Status Report '
        + 'has now checked and verified their entries.</p>'
        + (total ? '<p><strong>' + total + '</strong> ' + (total === 1 ? 'matter is' : 'matters are') + ' on the report.</p>' : '')
        + '<p><a href="' + esc(deep) + '">Open the report</a></p>'
        + '<p style="color:#666;font-size:12px;">Sent from the TMC portal. A verified row loses its tick if it is '
        + 'edited on the report afterwards — glance at the report before you send it.</p>';

      await sendMail(token, COMPLETE_TO,
        'Matter Status Report — ' + firm + ' — all draftsmen have verified', html);
      context.log('AUDIT msrnotify COMPLETE firm=' + firm + ' to=' + COMPLETE_TO + ' triggeredBy=' + callerEmail);

      context.res = {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ complete: true, sentTo: COMPLETE_TO }),
      };
    } catch (err) {
      context.log.error('Error sending completion email:', err.message);
      context.res = { status: 500, body: 'Error: ' + err.message };
    }
    return;
  }

  // ── mode: 'chase' (default) ── nudge the draftsmen who still have unverified matters.
  if (!list.length) {
    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sent: 0, skipped: [], note: 'Nobody has unverified matters on this report.' }),
    };
    return;
  }

  // Resolve names -> addresses against the roster. Anything unknown is DROPPED and reported back,
  // never guessed into an address.
  const targets = [];
  const skipped = [];
  list.forEach(r => {
    const name  = String((r && r.name) || '').trim();
    const count = parseInt((r && r.count), 10);
    const to    = ROSTER_EMAILS[name];
    if (!to)              { skipped.push(name || '(blank)'); return; }
    if (!(count > 0))     { skipped.push(name + ' (no unverified matters)'); return; }
    targets.push({ name, to, count });
  });

  if (!targets.length) {
    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sent: 0, skipped }),
    };
    return;
  }

  try {
    const token = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
    const sent  = [];
    const failed = [];

    for (const t of targets) {
      const deep = link + (link.indexOf('?') === -1 ? '?' : '&')
                 + 'firm=' + encodeURIComponent(firm) + '&draftsman=' + encodeURIComponent(t.name);
      const subject = 'Matter Status Report — ' + firm + ' — ' + t.count
                    + (t.count === 1 ? ' matter needs' : ' matters need') + ' your check';
      const html =
        '<p>Hi ' + esc(t.name) + ',</p>'
        + '<p>The Matter Status Report for <strong>' + esc(firm) + '</strong> is going out to the client. '
        + 'You have <strong>' + t.count + '</strong> ' + (t.count === 1 ? 'matter' : 'matters')
        + ' on it that you have not yet verified.</p>'
        + '<p><a href="' + esc(deep) + '">Open your matters on the report</a></p>'
        + '<p>Check the figures and the Current Position, edit anything that is wrong straight on the page, '
        + 'then press <strong>Verify my entries</strong>. If a matter should not be on the report at all, '
        + 'use the &#10005; on its row \u2014 it is not deleted, it moves to the Excluded list at the bottom.</p>'
        + '<p style="color:#666;font-size:12px;">Sent from the TMC portal.</p>';

      try {
        await sendMail(token, t.to, subject, html);
        sent.push(t.name);
        context.log('AUDIT msrnotify sent to=' + t.to + ' firm=' + firm + ' count=' + t.count + ' by=' + callerEmail);
      } catch (e) {
        // One bad address must not stop the rest. Report exactly who did NOT get it — never
        // report a send as successful when it was not.
        var hint = /ErrorInvalidUser/i.test(e.message)
          ? ' — the SEND-FROM mailbox (' + SEND_FROM + ') is not a valid mailbox, not the recipient'
          : '';
        failed.push(t.name + ': ' + e.message + hint);
        context.log.error('msrnotify FAILED to=' + t.to + ': ' + e.message);
      }
    }

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sent: sent.length, sentTo: sent, failed, skipped }),
    };
  } catch (err) {
    context.log.error('Error notifying draftsmen:', err.message);
    context.res = { status: 500, body: 'Error: ' + err.message };
  }
};

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

function sendMail(token, to, subject, html) {
  return graphPost(
    'https://graph.microsoft.com/v1.0/users/' + encodeURIComponent(SEND_FROM) + '/sendMail',
    token,
    {
      message: {
        subject: subject,
        body: { contentType: 'HTML', content: html },
        toRecipients: [{ emailAddress: { address: to } }],
      },
      saveToSentItems: true,
    }
  );
}

function graphPost(url, token, body) {
  return new Promise(function (resolve, reject) {
    const payload = JSON.stringify(body);
    const u       = new URL(url);
    const options = {
      hostname: u.hostname, path: u.pathname + u.search, method: 'POST',
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
        if (res.statusCode >= 400) { reject(new Error('Graph POST ' + res.statusCode + ': ' + data.slice(0, 300))); return; }
        try { resolve(JSON.parse(data)); } catch (e) { resolve({}); }
      });
    });
    req.on('error', reject);
    req.write(payload);
    req.end();
  });
}
