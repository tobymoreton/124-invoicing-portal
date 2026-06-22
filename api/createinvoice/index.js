/**
 * P124 — Invoicing Portal
 * Azure Function: /api/createinvoice
 *
 * Creates a draft invoice. Admin-only.
 *
 * POST body (JSON):
 *   pdfBase64       string   — base64-encoded PDF bytes
 *   invoiceDate     string   — ISO date string (from browser, today)
 *   computed        object   — selectedCase.computed from invoice-create.html
 *   caseFields      object   — selectedCase.fields from invoice-create.html
 *   caseSpId        string   — SharePoint item ID of the case
 *   checkedWipIds   string[] — SP item IDs of checked (billable) WIP entries
 *
 * Sequence:
 *   1. Get next invoice number from counter list (read → delete → write N+1)
 *   2. Upload PDF to Invoice Library
 *   3. Patch Invoice Library item metadata
 *   4. Patch Cases list item
 *   5. Patch each checked Time Tracking2 entry (Billed = true)
 *
 * Returns: { invoiceNumber, pdfUrl }
 */

const https   = require('https');
const { URL } = require('url');

const SITE_PATH      = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';
const INVOICE_LIB    = '5c366b19-0da9-4be9-b68f-60e6a0209cdb';
const CASES_LIST     = 'ae420bda-e550-499c-b337-90e4f33617c1';
const TT2_LIST       = '67db204c-30a5-4f4d-b276-60852d9967e1';
const COUNTER_LIST   = '7366b72d-02c4-4db1-a48b-907dfc7a33c7';

const ADMIN_EMAILS = ['toby@tmclegal.co.uk', 'danielle@tmclegal.co.uk'];

// ── Entry point ───────────────────────────────────────────────────────────────

module.exports = async function (context, req) {
  context.log('P124 /api/createinvoice called');

  // Admin only
  const callerEmail = getCallerEmail(req);
  if (!callerEmail || !ADMIN_EMAILS.includes(callerEmail)) {
    context.res = { status: 403, body: 'Forbidden — admin access required.' };
    return;
  }

  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    context.res = { status: 500, body: 'Missing required app settings.' };
    return;
  }

  // Parse body
  let body;
  try {
    body = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
  } catch (e) {
    context.res = { status: 400, body: 'Invalid JSON body.' };
    return;
  }

  const { pdfBase64, invoiceDate, computed, caseFields, caseSpId, checkedWipIds } = body || {};

  if (!pdfBase64 || !computed || !caseFields || !caseSpId) {
    context.res = { status: 400, body: 'Missing required fields: pdfBase64, computed, caseFields, caseSpId.' };
    return;
  }

  try {
    const token = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);

    // ── Step 1: Get next invoice number ──────────────────────────────────────
    context.log('Step 1: Getting invoice number from counter list');
    const invoiceNumber = await consumeInvoiceNumber(token, context);
    context.log('Invoice number assigned:', invoiceNumber);

    // ── Step 2: Upload PDF to Invoice Library ─────────────────────────────────
    context.log('Step 2: Uploading PDF');
    const fileName  = invoiceNumber + '.html';
    const pdfBuffer = Buffer.from(pdfBase64, 'base64');
    const uploadResult = await uploadPdf(token, fileName, pdfBuffer, context);
    const driveItemId = uploadResult.id;
    const pdfUrl      = uploadResult.webUrl;
    context.log('PDF uploaded, driveItemId:', driveItemId);

    // ── Step 3: Patch Invoice Library item metadata ───────────────────────────
    context.log('Step 3: Patching Invoice Library metadata');
    const invDate    = invoiceDate ? new Date(invoiceDate) : new Date();
    const dueDateObj = new Date(invDate); dueDateObj.setDate(dueDateObj.getDate() + 30);
    const invDateISO = invDate.toISOString().split('T')[0];
    const dueDateISO = dueDateObj.toISOString().split('T')[0];

    const invoiceType = resolveInvoiceType(computed);

    // Get the SP list item ID for the uploaded file (need it for PATCH /fields)
    const listItemId = await getListItemIdForDriveItem(token, driveItemId, context);
    context.log('Invoice Library list item ID:', listItemId);

    const metadataFields = {
      Title:                      invoiceNumber,
      InvoiceDate:                invDateISO,
      Casename:                   caseFields.Title || '',
      Invoicetype:                invoiceType,
      LAorIP:                     caseFields.LAorIP || '',
      AmountDue:                  computed.grand || 0,
      DraftedByEmail:             callerEmail,
      Net:                        computed.subTotal || 0,
      OData__x0056_AT20:          computed.vat || 0,
      DraftingFeeElement:         computed.draftingFee || 0,
      Expenses:                   (computed.bespokeExp || 0) + (computed.bespokeExpVat || 0),
      OrderDetails:               computed.bespokeInvoiceLine || '',
      Ourref:                     caseFields.Ourreference_x0028_text_x0029_ || '',
      Theirref:                   caseFields.ClientCaseReference || '',
      Case_x0020_ID:              caseFields.caseID_text || '',
      OData__ExtendedDescription: computed.timedWorkLine || '',
      DueDate:                    dueDateISO,
    };

    await patchListItem(token, INVOICE_LIB, listItemId, metadataFields, context);
    context.log('Step 3 complete');

    // ── Step 4: Patch Cases list item ─────────────────────────────────────────
    context.log('Step 4: Patching Cases list item', caseSpId);
    const caseFieldsPatch = {
      Invoice_x0020_number:              invoiceNumber,
      Invoice_x0020_date:                invDateISO,
      Drafting_x0020_fee_x0020_amount:   computed.draftingFee || 0,
      Timed_x0020_work_x0020_fee_x0020:  computed.timedFee || 0,
      DraftingYetToBeInvoiced:           false,
    };
    await patchListItem(token, CASES_LIST, caseSpId, caseFieldsPatch, context);
    context.log('Step 4 complete');

    // ── Step 5: Mark Time Tracking2 entries as Billed ─────────────────────────
    if (checkedWipIds && checkedWipIds.length > 0) {
      context.log('Step 5: Marking', checkedWipIds.length, 'WIP entries as Billed');
      for (const wipId of checkedWipIds) {
        await patchListItem(token, TT2_LIST, wipId, {
          'Billed_x003f_': true,
          InvoiceNumber:   invoiceNumber,
          InvoiceDate:     invDateISO,
        }, context);
      }
      context.log('Step 5 complete');
    } else {
      context.log('Step 5: No WIP entries to mark');
    }

    // ── Done ──────────────────────────────────────────────────────────────────
    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ invoiceNumber, pdfUrl }),
    };

  } catch (err) {
    context.log.error('createinvoice error:', err.message, err.stack);
    context.res = { status: 500, body: 'Error: ' + err.message };
  }
};

// ── Invoice number counter ────────────────────────────────────────────────────
// Counter list holds one item. We read it, capture the number, delete it,
// write N+1 back. All server-side so it is atomic within this function instance.

async function consumeInvoiceNumber(token, context) {
  const listUrl = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
    + '/lists/' + COUNTER_LIST + '/items?$expand=fields&$top=1';
  const page = await graphGet(listUrl, token);
  const items = page.value || [];
  if (!items.length) throw new Error('Counter list is empty — cannot assign invoice number.');

  const item          = items[0];
  const itemId        = item.id;
  const currentNumber = item.fields.InvoiceNumber || item.fields.Title;
  if (!currentNumber) throw new Error('Counter list item has no InvoiceNumber or Title field.');

  const invoiceNumber = String(currentNumber).trim();

  // Parse next number: handles purely numeric ("1042") or prefixed ("TMC-1042")
  const match = invoiceNumber.match(/^(.*?)(\d+)$/);
  if (!match) throw new Error('Cannot parse invoice number format: ' + invoiceNumber);
  const prefix         = match[1];
  const nextNum        = parseInt(match[2], 10) + 1;
  const nextPadded     = String(nextNum).padStart(match[2].length, '0');
  const nextInvoiceNum = prefix + nextPadded;

  // Delete current item
  await graphDelete(
    'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
    + '/lists/' + COUNTER_LIST + '/items/' + itemId,
    token
  );

  // Write N+1 back
  await graphPost(
    'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
    + '/lists/' + COUNTER_LIST + '/items',
    token,
    { fields: { Title: nextInvoiceNum, InvoiceNumber: nextInvoiceNum } }
  );

  return invoiceNumber;
}

// ── PDF upload ────────────────────────────────────────────────────────────────
// Simple upload (<4 MB) via PUT to Invoice Library drive root.

async function uploadPdf(token, fileName, pdfBuffer, context) {
  const driveUrl  = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
    + '/lists/' + INVOICE_LIB + '/drive';
  const driveInfo = await graphGet(driveUrl, token);
  const driveId   = driveInfo.id;
  if (!driveId) throw new Error('Could not get drive ID for Invoice Library.');

  const uploadUrl = 'https://graph.microsoft.com/v1.0/drives/' + driveId
    + '/root:/' + encodeURIComponent(fileName) + ':/content';

  const result = await graphPut(uploadUrl, token, pdfBuffer, 'text/html');
  if (!result.id) throw new Error('PDF upload did not return a DriveItem id.');
  return result;
}

// ── Get SP list item ID for a DriveItem ───────────────────────────────────────

async function getListItemIdForDriveItem(token, driveItemId, context) {
  const driveUrl  = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
    + '/lists/' + INVOICE_LIB + '/drive';
  const driveInfo = await graphGet(driveUrl, token);
  const driveId   = driveInfo.id;

  const itemUrl   = 'https://graph.microsoft.com/v1.0/drives/' + driveId
    + '/items/' + driveItemId + '?$expand=listItem';
  const driveItem = await graphGet(itemUrl, token);
  const spId      = driveItem.listItem && driveItem.listItem.id;
  if (!spId) throw new Error('Could not get SP list item ID for DriveItem ' + driveItemId);
  return spId;
}

// ── Invoice type ──────────────────────────────────────────────────────────────

function resolveInvoiceType(computed) {
  const d = computed.draftingFee > 0;
  const t = computed.timedFee    > 0;
  if (d && t) return 'Drafting and timed work';
  if (d)      return 'Drafting only';
  if (t)      return 'Timed work only';
  return 'Other';
}

// ── Patch SP list item fields ─────────────────────────────────────────────────

async function patchListItem(token, listGuid, itemId, fields, context) {
  const url = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
    + '/lists/' + listGuid + '/items/' + itemId + '/fields';
  await graphPatch(url, token, fields);
}

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
      headers: {
        Authorization: 'Bearer ' + token,
        Accept:        'application/json',
      },
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

function graphDelete(url, token) {
  return new Promise(function (resolve, reject) {
    const u = new URL(url);
    const options = {
      hostname: u.hostname,
      path:     u.pathname + u.search,
      method:   'DELETE',
      headers: {
        Authorization: 'Bearer ' + token,
      },
    };
    const req = https.request(options, function (res) {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        if (res.statusCode >= 400) {
          reject(new Error('Graph DELETE ' + res.statusCode + ' — ' + url.split('?')[0] + ': ' + data.slice(0, 400)));
          return;
        }
        resolve();
      });
    });
    req.on('error', reject);
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
        try { resolve(JSON.parse(data)); }
        catch (e) { reject(new Error('JSON parse on PUT response: ' + e.message)); }
      });
    });
    req.on('error', reject);
    req.write(buffer);
    req.end();
  });
}
