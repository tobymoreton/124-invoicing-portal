/**
 * P124 — Invoicing Portal
 * Azure Function: /api/cancelinvoice
 *
 * Cancels a draft or issued invoice. Admin + finance (Toby, Danielle, Lesley).
 *
 * Sequence:
 *   1. Auth check (admin + finance)
 *   2. Read DraftWipIds from Invoice Library item
 *   3. Revert each TT2 entry: Billed_x003f_ = false
 *   4. Delete each Line Item where field_7 is in DraftWipIds
 *   5. Patch Invoice Library item: Cancelled = true, AmountDue = 0, Net = 0,
 *      DraftingFeeElement = 0, Expenses = 0
 *   6. Send cancellation email to office@tmclegal.co.uk via Graph Mail.Send
 *
 * POST body (JSON):
 *   listItemId   string   — Invoice Library SP list item ID
 *
 * Returns: { cancelled: true, invoiceTitle }
 *
 * GUIDs:
 *   Invoice Library:  5c366b19-0da9-4be9-b68f-60e6a0209cdb
 *   Line Items:       496468a5-e2ed-48db-8826-58cb08844eee
 *   Time Tracking2:   67db204c-30a5-4f4d-b276-60852d9967e1
 */

const https   = require('https');
const { URL } = require('url');

const SITE_PATH   = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';
const INVOICE_LIB = '5c366b19-0da9-4be9-b68f-60e6a0209cdb';
const LINE_ITEMS  = '496468a5-e2ed-48db-8826-58cb08844eee';
const TT2         = '67db204c-30a5-4f4d-b276-60852d9967e1';

const CANCEL_EMAILS = ['toby@tmclegal.co.uk', 'danielle@tmclegal.co.uk', 'lesley@tmclegal.co.uk'];
const NOTIFY_EMAIL  = 'office@tmclegal.co.uk';
// Graph Mail.Send requires a licensed mailbox to send from.
// Set MAIL_SENDER in app settings to e.g. 'toby@tmclegal.co.uk'
// If not set, the email step is skipped and a warning is logged.

// ── Entry point ───────────────────────────────────────────────────────────────

module.exports = async function (context, req) {
  context.log('P124 /api/cancelinvoice called');

  const callerEmail = getCallerEmail(req);
  if (!callerEmail || !CANCEL_EMAILS.includes(callerEmail)) {
    context.res = { status: 403, body: 'Forbidden — not authorised to cancel invoices.' };
    return;
  }

  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET, MAIL_SENDER } = process.env;
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

  const { listItemId } = body || {};
  if (!listItemId) {
    context.res = { status: 400, body: 'Missing required field: listItemId.' };
    return;
  }

  try {
    const token = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);

    // ── 1. Read Invoice Library item fields ───────────────────────────────────
    context.log('Reading Invoice Library item fields…');
    const fieldsUrl  = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
      + '/lists/' + INVOICE_LIB + '/items/' + listItemId + '/fields';
    const itemFields = await graphGet(fieldsUrl, token);

    const wipIdsCsv    = itemFields.DraftWipIds  || '';
    const invoiceTitle = itemFields.Title         || listItemId;
    const invoiceRef   = itemFields.OrderDetails  || '(draft — no number assigned)';
    const caseName     = itemFields.Casename      || '';
    const wipIds       = wipIdsCsv.split(',').map(s => s.trim()).filter(Boolean);

    context.log('DraftWipIds:', wipIds);
    context.log('Cancelling:', invoiceTitle);

    // ── 2. Revert TT2 entries ─────────────────────────────────────────────────
    if (wipIds.length > 0) {
      context.log('Reverting', wipIds.length, 'TT2 entries…');
      let tt2Done = 0, tt2Failed = 0;
      for (const wipId of wipIds) {
        try {
          await patchListItem(token, TT2, wipId, { 'Billed_x003f_': false });
          tt2Done++;
        } catch (e) {
          context.log.error('TT2 revert failed for id', wipId, ':', e.message);
          tt2Failed++;
        }
      }
      context.log('TT2 reverted: done=' + tt2Done + ', failed=' + tt2Failed);
    }

    // ── 3. Delete Line Items where field_7 in DraftWipIds ────────────────────
    if (wipIds.length > 0) {
      context.log('Deleting Line Items…');
      let liDeleted = 0, liFailed = 0;
      for (const wipId of wipIds) {
        try {
          const liSearchUrl = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
            + '/lists/' + LINE_ITEMS + '/items'
            + '?$expand=fields($select=id,field_7)'
            + '&$filter=fields/field_7 eq \'' + wipId + '\'';
          const liResult = await graphGetUnindexed(liSearchUrl, token);
          const liItems  = liResult.value || [];

          for (const liItem of liItems) {
            await deleteListItem(token, LINE_ITEMS, liItem.id);
            liDeleted++;
          }
        } catch (e) {
          context.log.error('Line Item delete failed for wipId', wipId, ':', e.message);
          liFailed++;
        }
      }
      context.log('Line Items: deleted=' + liDeleted + ', failed=' + liFailed);
    }

    // ── 4. Patch Invoice Library item — mark Cancelled ────────────────────────
    context.log('Patching Invoice Library item as cancelled…');
    await patchListItem(token, INVOICE_LIB, listItemId, {
      Cancelled:          true,
      AmountDue:          0,
      Net:                0,
      DraftingFeeElement: 0,
      Expenses:           0,
    });
    context.log('Invoice Library item patched — Status will flip to Cancelled.');

    // ── 5. Send cancellation email ────────────────────────────────────────────
    if (MAIL_SENDER) {
      try {
        context.log('Sending cancellation email…');
        const subject = 'Invoice Cancelled — ' + invoiceRef + (caseName ? ' | ' + caseName : '');
        const htmlBody = '<p>An invoice has been cancelled in the TMC Legal invoicing portal.</p>'
          + '<p><strong>Invoice:</strong> ' + invoiceRef + '<br>'
          + (caseName ? '<strong>Case:</strong> ' + caseName + '<br>' : '')
          + '<strong>Cancelled by:</strong> ' + callerEmail + '</p>'
          + '<p>The associated WIP entries have been un-billed and line items deleted.</p>';

        await sendMail(token, MAIL_SENDER, NOTIFY_EMAIL, subject, htmlBody);
        context.log('Cancellation email sent.');
      } catch (mailErr) {
        // Email failure is non-fatal — log and continue
        context.log.error('Cancellation email failed:', mailErr.message);
      }
    } else {
      context.log('MAIL_SENDER not set — skipping cancellation email.');
    }

    // ── Done ──────────────────────────────────────────────────────────────────
    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ cancelled: true, invoiceTitle }),
    };

  } catch (err) {
    context.log.error('cancelinvoice error:', err.message, err.stack);
    context.res = { status: 500, body: 'Error: ' + err.message };
  }
};

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

// Like graphGet but adds the Prefer header to allow filtering on non-indexed SP columns.
function graphGetUnindexed(url, token) {
  return new Promise(function (resolve, reject) {
    const u = new URL(url);
    const options = {
      hostname: u.hostname,
      path:     u.pathname + u.search,
      method:   'GET',
      headers: {
        Authorization: 'Bearer ' + token,
        Accept:        'application/json',
        Prefer:        'HonorNonIndexedQueriesWarningMayFailRandomly',
      },
    };
    const req = https.request(options, function (res) {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        if (res.statusCode >= 400) {
          reject(new Error('Graph GET (unindexed) ' + res.statusCode + ' — ' + url.split('?')[0] + ': ' + data.slice(0, 400)));
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
        if (res.statusCode === 409 && retries > 0) {
          // resourceModified — another process updated the item; retry once after brief delay
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

function graphDelete(url, token) {
  return new Promise(function (resolve, reject) {
    const u = new URL(url);
    const options = {
      hostname: u.hostname,
      path:     u.pathname + u.search,
      method:   'DELETE',
      headers: { Authorization: 'Bearer ' + token },
    };
    const req = https.request(options, function (res) {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        if (res.statusCode >= 400) {
          reject(new Error('Graph DELETE ' + res.statusCode + ' — ' + url.split('?')[0] + ': ' + data.slice(0, 400)));
          return;
        }
        resolve({});
      });
    });
    req.on('error', reject);
    req.end();
  });
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
    const req = https.request(options, function (res) {
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
    req.on('error', reject);
    req.write(payload);
    req.end();
  });
}

async function patchListItem(token, listGuid, itemId, fields) {
  const url = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
    + '/lists/' + listGuid + '/items/' + itemId + '/fields';
  await graphPatch(url, token, fields);
}

async function deleteListItem(token, listGuid, itemId) {
  const url = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
    + '/lists/' + listGuid + '/items/' + itemId;
  await graphDelete(url, token);
}

async function sendMail(token, from, to, subject, htmlBody) {
  const url = 'https://graph.microsoft.com/v1.0/users/' + from + '/sendMail';
  await graphPost(url, token, {
    message: {
      subject,
      body: { contentType: 'HTML', content: htmlBody },
      toRecipients: [{ emailAddress: { address: to } }],
    },
    saveToSentItems: true,
  });
}
