/**
 * P124 — Invoicing Portal
 * Azure Function: /api/createdraft
 *
 * Creates a DRAFT invoice. Admin-only.
 *
 * Two-stage workflow:
 *   DRAFT (this function)  — admin creates draft. Writes invoice HTML + figures
 *                            to Invoice Library with InvoiceDate BLANK and
 *                            AmountDue = 0, so the calculated Status field reads
 *                            "Draft". Creates one Line Item per checked WIP entry.
 *                            NO invoice number consumed. NO Cases patch. NO WIP marking.
 *   ISSUE (api/issueinvoice, Lesley) — consumes invoice number, sets InvoiceDate
 *                            + AmountDue (Status flips to "Issued"), marks WIP billed,
 *                            updates Line Items with real invoice number.
 *
 * POST body (JSON):
 *   pdfBase64        string    — base64-encoded invoice HTML (print-ready)
 *   computed         object    — selectedCase.computed from invoice-create.html
 *   caseFields       object    — selectedCase.fields from invoice-create.html
 *   checkedWipIds    string[]  — SP item IDs of checked WIP entries
 *   checkedWipEntries object[] — full entry data for Line Items creation:
 *                                [{_id, DateCompleted, WorkDone, HoursSpent, Rate,
 *                                   CaseName, OurRef, Email}]
 *
 * Returns: { fileName, pdfUrl, status: 'Draft' }
 *
 * Invoice Library GUID:  5c366b19-0da9-4be9-b68f-60e6a0209cdb
 * Line Items list GUID:  496468a5-e2ed-48db-8826-58cb08844eee
 */

const https   = require('https');
const { URL } = require('url');

const SITE_PATH    = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';
const INVOICE_LIB  = '5c366b19-0da9-4be9-b68f-60e6a0209cdb';
const LINE_ITEMS   = '496468a5-e2ed-48db-8826-58cb08844eee';

const ADMIN_EMAILS = ['toby@tmclegal.co.uk', 'danielle@tmclegal.co.uk'];

// ── Entry point ───────────────────────────────────────────────────────────────

module.exports = async function (context, req) {
  context.log('P124 /api/createdraft called');

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

  let body;
  try {
    body = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
  } catch (e) {
    context.res = { status: 400, body: 'Invalid JSON body.' };
    return;
  }

  const { pdfBase64, computed, caseFields, checkedWipIds, checkedWipEntries } = body || {};
  if (!pdfBase64 || !computed || !caseFields) {
    context.res = { status: 400, body: 'Missing required fields: pdfBase64, computed, caseFields.' };
    return;
  }

  try {
    const token = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);

    const invoiceType = resolveInvoiceType(computed);

    // ── Draft filename ────────────────────────────────────────────────────────
    // No invoice number at draft stage (assigned by Lesley at issue).
    // Use: DRAFT - {caseName} - {YYYY-MM-DD HHmm}.html
    const now      = new Date();
    const stamp    = now.toISOString().slice(0, 16).replace('T', ' ').replace(':', '');
    const safeName = (caseFields.Title || 'Case').replace(/[\\\/:*?"<>|]/g, '-').slice(0, 80);
    const fileName = 'DRAFT - ' + safeName + ' - ' + stamp + '.html';

    // ── Upload draft HTML to Invoice Library ──────────────────────────────────
    context.log('Uploading draft:', fileName);
    const pdfBuffer    = Buffer.from(pdfBase64, 'base64');
    const uploadResult = await uploadFile(token, fileName, pdfBuffer);
    const driveItemId  = uploadResult.id;
    const pdfUrl       = uploadResult.webUrl;
    context.log('Draft uploaded, driveItemId:', driveItemId);

    // ── Patch metadata ────────────────────────────────────────────────────────
    // Figures populated; InvoiceDate intentionally omitted (blank) + AmountDue
    // intentionally omitted (defaults to 0) → calculated Status field = "Draft"
    const listItemId = await getListItemIdForDriveItem(token, driveItemId);
    context.log('Invoice Library list item ID:', listItemId);

    // Store checked WIP ids as CSV so the issue stage can mark exactly those billed
    const wipIdsCsv = Array.isArray(checkedWipIds) ? checkedWipIds.join(',') : '';

    const metadataFields = {
      Title:                invoiceType + ' draft — ' + (caseFields.Title || ''),
      Casename:             caseFields.Title || '',
      Invoicetype:          invoiceType,
      Net:                  computed.subTotal  || 0,
      VAT:                  computed.vat       || 0,
      DraftingFeeElement:   computed.draftingFee || 0,
      Expenses:             (computed.bespokeExp || 0) + (computed.bespokeExpVat || 0),
      Ourref:               caseFields.Ourreference_x0028_text_x0029_ || '',
      Theirref:             caseFields.ClientCaseReference || '',
      _ExtendedDescription: computed.timedWorkLine || '',
      DraftedByEmail:       callerEmail,
      VendorName:           caseFields.Firm_x0028_text_x0029_ || '',
      DraftWipIds:          wipIdsCsv,
      // DELIBERATELY OMITTED to keep Status = "Draft":
      //   InvoiceDate  — blank
      //   AmountDue    — 0 / omitted
      //   OrderDetails (invoice number) — assigned at issue
    };

    await patchListItem(token, INVOICE_LIB, listItemId, metadataFields);
    context.log('Draft metadata written. Calculated Status = "Draft".');

    // ── Create Line Items ─────────────────────────────────────────────────────
    // One Line Item per checked WIP entry. InvoiceIDRef left blank — the real
    // invoice number is assigned at issue stage and written back then.
    // field_7 stores the TT2 source ID as a back-link for the issue stage.
    const entriesToBill = Array.isArray(checkedWipEntries) ? checkedWipEntries : [];
    if (entriesToBill.length > 0) {
      context.log('Creating', entriesToBill.length, 'Line Item(s)…');
      let liCreated = 0;
      let liFailed  = 0;
      for (const entry of entriesToBill) {
        try {
          const liFields = {
            field_1:  entry.WorkDone   || '',       // Work Done (text)
            field_2:  entry.HoursSpent || 0,        // Time (hours)
            field_3:  entry.Rate       || 0,        // Rate £/hr
            field_5:  entry.OurRef     || caseFields.Ourreference_x0028_text_x0029_ || '', // Our Reference
            field_7:  String(entry._id || ''),      // TT2 source ID (back-link)
            CaseName: entry.CaseName   || caseFields.Title || '',
            'caseName_x0020__x0001f455_': entry.CaseName || caseFields.Title || '',
            'BillableYorN_x0020__x2753_': true,
            InvoiceType: invoiceType,               // plain string — Graph app-only
            // InvoiceIDRef intentionally blank (assigned at issue)
            // CompletedBy (Person) — cannot write via Graph app-only; omitted
            // caseID (Lookup)      — cannot reliably write via Graph app-only; omitted
          };
          // Date: write as ISO string if present
          if (entry.DateCompleted) {
            liFields['Completed_x0020_on'] = new Date(entry.DateCompleted).toISOString();
          }
          await createListItem(token, LINE_ITEMS, liFields);
          liCreated++;
        } catch (liErr) {
          context.log.error('Line Item creation failed for TT2 id', entry._id, ':', liErr.message);
          liFailed++;
          // Don't abort — continue with remaining entries
        }
      }
      context.log('Line Items: created=' + liCreated + ', failed=' + liFailed);
    }

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ fileName, pdfUrl, status: 'Draft' }),
    };

  } catch (err) {
    context.log.error('createdraft error:', err.message, err.stack);
    context.res = { status: 500, body: 'Error: ' + err.message };
  }
};

// ── Invoice type ──────────────────────────────────────────────────────────────

function resolveInvoiceType(computed) {
  const d = computed.draftingFee > 0;
  const t = computed.timedFee    > 0;
  if (d && t) return 'Drafting and timed work';
  if (d)      return 'Drafting only';
  if (t)      return 'Timed work only';
  return 'Other';
}

// ── File upload ───────────────────────────────────────────────────────────────

async function uploadFile(token, fileName, buffer) {
  const driveUrl  = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
    + '/lists/' + INVOICE_LIB + '/drive';
  const driveInfo = await graphGet(driveUrl, token);
  const driveId   = driveInfo.id;
  if (!driveId) throw new Error('Could not get drive ID for Invoice Library.');

  const uploadUrl = 'https://graph.microsoft.com/v1.0/drives/' + driveId
    + '/root:/' + encodeURIComponent(fileName) + ':/content';

  const result = await graphPut(uploadUrl, token, buffer, 'text/html');
  if (!result.id) throw new Error('Draft upload did not return a DriveItem id.');
  return result;
}

async function getListItemIdForDriveItem(token, driveItemId) {
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

async function patchListItem(token, listGuid, itemId, fields) {
  const url = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
    + '/lists/' + listGuid + '/items/' + itemId + '/fields';
  await graphPatch(url, token, fields);
}

async function createListItem(token, listGuid, fields) {
  const url = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
    + '/lists/' + listGuid + '/items';
  return await graphPost(url, token, { fields });
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
