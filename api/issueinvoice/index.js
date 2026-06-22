/**
 * P124 — Invoicing Portal
 * Azure Function: /api/issueinvoice
 *
 * Issues a draft invoice. Finance-only (Lesley).
 *
 * Sequence:
 *   1. Finance-only auth check
 *   2. Read next invoice number from counter list (single item, InvoiceNumber field)
 *   3. Increment counter (patch InvoiceNumber + 1)
 *   4. Overwrite draft HTML file in Invoice Library with issued HTML (real number + date)
 *   5. Patch Invoice Library item: OrderDetails = number, InvoiceDate = today, AmountDue = grandTotal
 *   6. Read DraftWipIds CSV from Invoice Library item
 *   7. Mark each TT2 entry Billed_x003f_ = true
 *   8. Update each Line Item where field_7 in DraftWipIds: set InvoiceIDRef = number
 *
 * POST body (JSON):
 *   listItemId   string   — Invoice Library SP list item ID
 *   pdfBase64    string   — issued invoice HTML (base64), real number + date baked in
 *   grandTotal   number   — amount to write to AmountDue
 *
 * Returns: { invoiceNumber, invoiceDate }
 *
 * GUIDs:
 *   Invoice Library:  5c366b19-0da9-4be9-b68f-60e6a0209cdb
 *   Line Items:       496468a5-e2ed-48db-8826-58cb08844eee
 *   Time Tracking2:   67db204c-30a5-4f4d-b276-60852d9967e1
 *   Counter list:     7366b72d-02c4-4db1-a48b-907dfc7a33c7
 */

const https   = require('https');
const { URL } = require('url');

const SITE_PATH   = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';
const INVOICE_LIB = '5c366b19-0da9-4be9-b68f-60e6a0209cdb';
const LINE_ITEMS  = '496468a5-e2ed-48db-8826-58cb08844eee';
const TT2         = '67db204c-30a5-4f4d-b276-60852d9967e1';
const COUNTER     = '7366b72d-02c4-4db1-a48b-907dfc7a33c7';

const FINANCE_EMAILS = ['lesley@tmclegal.co.uk'];

// ── Entry point ───────────────────────────────────────────────────────────────

module.exports = async function (context, req) {
  context.log('P124 /api/issueinvoice called');

  const callerEmail = getCallerEmail(req);
  if (!callerEmail || !FINANCE_EMAILS.includes(callerEmail)) {
    context.res = { status: 403, body: 'Forbidden — finance access required.' };
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

  const { listItemId, pdfBase64, grandTotal } = body || {};
  if (!listItemId || !pdfBase64 || grandTotal === undefined) {
    context.res = { status: 400, body: 'Missing required fields: listItemId, pdfBase64, grandTotal.' };
    return;
  }

  try {
    const token = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);

    // ── 1. Read + increment invoice counter ───────────────────────────────────
    context.log('Reading invoice counter…');
    const counterUrl = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
      + '/lists/' + COUNTER + '/items?$expand=fields($select=InvoiceNumber,id)&$top=1';
    const counterResult = await graphGet(counterUrl, token);
    const counterItem   = counterResult.value && counterResult.value[0];
    if (!counterItem) throw new Error('Counter list is empty — cannot get invoice number.');

    const counterItemId  = counterItem.id;
    const invoiceNumber  = counterItem.fields.InvoiceNumber;
    if (!invoiceNumber) throw new Error('InvoiceNumber field is blank in counter list.');

    context.log('Invoice number:', invoiceNumber);

    // Increment counter immediately to prevent double-use
    await patchListItem(token, COUNTER, counterItemId, { InvoiceNumber: invoiceNumber + 1 });
    context.log('Counter incremented to', invoiceNumber + 1);

    // ── 2. Overwrite draft file with issued HTML ───────────────────────────────
    context.log('Overwriting draft file with issued HTML…');
    const htmlBuffer = Buffer.from(pdfBase64, 'base64');

    // Get the drive item ID from the list item so we can overwrite the file
    const listItemUrl = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
      + '/lists/' + INVOICE_LIB + '/items/' + listItemId + '?$expand=driveItem';
    const listItemData = await graphGet(listItemUrl, token);
    const driveItemId  = listItemData.driveItem && listItemData.driveItem.id;
    if (!driveItemId) throw new Error('Could not get driveItem ID for list item ' + listItemId);

    const driveUrl  = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
      + '/lists/' + INVOICE_LIB + '/drive';
    const driveInfo = await graphGet(driveUrl, token);
    const driveId   = driveInfo.id;
    if (!driveId) throw new Error('Could not get drive ID for Invoice Library.');

    const uploadUrl = 'https://graph.microsoft.com/v1.0/drives/' + driveId
      + '/items/' + driveItemId + '/content';
    await graphPut(uploadUrl, token, htmlBuffer, 'text/html');
    context.log('Issued HTML uploaded.');

    // ── 3. Patch Invoice Library item ─────────────────────────────────────────
    const invoiceDate = new Date().toISOString();
    context.log('Patching Invoice Library item…');
    await patchListItem(token, INVOICE_LIB, listItemId, {
      OrderDetails: String(invoiceNumber),
      InvoiceDate:  invoiceDate,
      AmountDue:    grandTotal,
    });
    context.log('Invoice Library item patched — Status will flip to Issued.');

    // ── 4. Read DraftWipIds from Invoice Library item ─────────────────────────
    context.log('Reading DraftWipIds…');
    const fieldsUrl  = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
      + '/lists/' + INVOICE_LIB + '/items/' + listItemId + '/fields';
    const itemFields = await graphGet(fieldsUrl, token);
    const wipIdsCsv  = itemFields.DraftWipIds || '';
    const wipIds     = wipIdsCsv.split(',').map(s => s.trim()).filter(Boolean);
    context.log('DraftWipIds:', wipIds);

    // ── 5. Mark TT2 entries as Billed ─────────────────────────────────────────
    if (wipIds.length > 0) {
      context.log('Marking', wipIds.length, 'TT2 entries as billed…');
      let tt2Done = 0, tt2Failed = 0;
      for (const wipId of wipIds) {
        try {
          await patchListItem(token, TT2, wipId, { 'Billed_x003f_': true });
          tt2Done++;
        } catch (e) {
          context.log.error('TT2 patch failed for id', wipId, ':', e.message);
          tt2Failed++;
        }
      }
      context.log('TT2: billed=' + tt2Done + ', failed=' + tt2Failed);
    }

    // ── 6. Update Line Items with real invoice number ─────────────────────────
    // Find Line Items where field_7 (TT2 source ID back-link) is in our wipIds list
    if (wipIds.length > 0) {
      context.log('Updating Line Items with invoice number…');

      // Fetch all Line Items for this invoice via field_7 filter
      // Graph doesn't support "in" filter — fetch by each ID individually
      let liDone = 0, liFailed = 0;
      for (const wipId of wipIds) {
        try {
          // Find the Line Item where field_7 = wipId
          const liSearchUrl = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
            + '/lists/' + LINE_ITEMS + '/items'
            + '?$expand=fields($select=id,field_7,InvoiceIDRef)'
            + '&$filter=fields/field_7 eq \'' + wipId + '\'';
          const liResult = await graphGet(liSearchUrl, token);
          const liItems  = liResult.value || [];

          for (const liItem of liItems) {
            await patchListItem(token, LINE_ITEMS, liItem.id, {
              InvoiceIDRef: String(invoiceNumber),
            });
            liDone++;
          }
        } catch (e) {
          context.log.error('Line Item update failed for wipId', wipId, ':', e.message);
          liFailed++;
        }
      }
      context.log('Line Items: updated=' + liDone + ', failed=' + liFailed);
    }

    // ── Done ──────────────────────────────────────────────────────────────────
    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        invoiceNumber: String(invoiceNumber),
        invoiceDate,
      }),
    };

  } catch (err) {
    context.log.error('issueinvoice error:', err.message, err.stack);
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

function graphPatch(url, token, body) {
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

function graphPut(url, token, buffer, contentType) {
  return new Promise(function (resolve, reject) {
    const u = new URL(url);
    const options = {
      hostname: u.hostname,
      path:     u.pathname + u.search,
      method:   'PUT',
      headers: {
        Authorization:    'Bearer ' + token,
        'Content-Type':   contentType,
        'Content-Length': buffer.length,
      },
    };
    const req = https.request(options, function (res) {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        if (res.statusCode >= 400) {
          reject(new Error('Graph PUT ' + res.statusCode + ' — ' + url.split('?')[0] + ': ' + data.slice(0, 400)));
          return;
        }
        try { resolve(data ? JSON.parse(data) : {}); }
        catch (e) { resolve({}); }
      });
    });
    req.on('error', reject);
    req.write(buffer);
    req.end();
  });
}

async function patchListItem(token, listGuid, itemId, fields) {
  const url = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
    + '/lists/' + listGuid + '/items/' + itemId + '/fields';
  await graphPatch(url, token, fields);
}
