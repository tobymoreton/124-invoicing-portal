/**
 * P124 — Invoicing Portal
 * Azure Function: /api/issueinvoice
 *
 * Issues a draft invoice. Finance + admin (Lesley, Toby, Danielle).
 *
 * Sequence:
 *   1. Auth check (finance + admin)
 *   2. Read next invoice number from counter list (single item, InvoiceNumber field)
 *   3. Increment counter (patch InvoiceNumber + 1)
 *   4. Read existing draft HTML file from Invoice Library
 *   5. Do server-side find/replace to bake in real number + date (lossless — preserves
 *      schedule, line descriptions, etc. exactly as drafted)
 *   6. Overwrite draft file with issued HTML
 *   6a. Rename DriveItem to {invoiceNumber}.html
 *   7. Patch Invoice Library item: OrderDetails = number, InvoiceDate = today, DueDate = today + 30, AmountDue = grandTotal
 *   8. Read DraftWipIds CSV from Invoice Library item
 *   9. Mark each TT2 entry Billed_x003f_ = true
 *  10. Update each Line Item where field_7 in DraftWipIds: set InvoiceIDRef = number
 *
 * POST body (JSON):
 *   listItemId   string   — Invoice Library SP list item ID
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

const ISSUE_EMAILS = ['lesley@tmclegal.co.uk', 'toby@tmclegal.co.uk', 'danielle@tmclegal.co.uk'];

// ── Entry point ───────────────────────────────────────────────────────────────

module.exports = async function (context, req) {
  context.log('P124 /api/issueinvoice called');

  const callerEmail = getCallerEmail(req);
  if (!callerEmail || !ISSUE_EMAILS.includes(callerEmail)) {
    context.res = { status: 403, body: 'Forbidden — not authorised to issue invoices.' };
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

  const { listItemId, grandTotal } = body || {};
  if (!listItemId || grandTotal === undefined) {
    context.res = { status: 400, body: 'Missing required fields: listItemId, grandTotal.' };
    return;
  }

  try {
    const token = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);

    // ── DraftedByEmail safety guard ───────────────────────────────────────────
    // If the invoice has a drafting fee, DraftedByEmail is required for wage
    // calculation. Block issue BEFORE consuming the invoice number.
    context.log('Checking DraftedByEmail safety guard…');
    const guardUrl    = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
      + '/lists/' + INVOICE_LIB + '/items/' + listItemId
      + '/fields?$select=DraftedByEmail,DraftingFeeElement';
    const guardFields = await graphGet(guardUrl, token);
    const hasDraftingFee = parseFloat(guardFields.DraftingFeeElement) > 0;
    const draftedByEmail = (guardFields.DraftedByEmail || '').trim();
    if (hasDraftingFee && !draftedByEmail) {
      context.res = {
        status: 400,
        body: 'Cannot issue — Drafted By is blank. Please set the Drafted By field on this invoice before issuing.',
      };
      return;
    }
    context.log('DraftedByEmail guard passed.');

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

    // ── 2. Get driveItem ID and drive ID ──────────────────────────────────────
    context.log('Fetching driveItem details…');
    const listItemUrl = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
      + '/lists/' + INVOICE_LIB + '/items/' + listItemId + '?$expand=driveItem';
    const listItemData = await graphGet(listItemUrl, token);
    const driveItemId  = listItemData.driveItem && listItemData.driveItem.id;
    const driveItemEtag = (listItemData.driveItem && listItemData.driveItem.eTag) || null;
    if (!driveItemId) throw new Error('Could not get driveItem ID for list item ' + listItemId);

    const driveUrl  = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
      + '/lists/' + INVOICE_LIB + '/drive';
    const driveInfo = await graphGet(driveUrl, token);
    const driveId   = driveInfo.id;
    if (!driveId) throw new Error('Could not get drive ID for Invoice Library.');

    // ── 3. Read existing draft HTML ───────────────────────────────────────────
    context.log('Reading existing draft HTML…');
    const downloadUrl = 'https://graph.microsoft.com/v1.0/drives/' + driveId
      + '/items/' + driveItemId + '/content';
    const draftHtml = await graphGetFile(downloadUrl, token);
    context.log('Draft HTML read, length:', draftHtml.length);

    // ── 4. Replace DRAFT placeholders with real number + date ─────────────────
    const invoiceDate    = new Date();
    const issuedHtml     = doHtmlReplace(draftHtml, String(invoiceNumber), invoiceDate);
    const htmlBuffer     = Buffer.from(issuedHtml, 'utf8');
    context.log('HTML replacements applied.');

    // ── 5. Overwrite draft file with issued HTML ───────────────────────────────
    context.log('Uploading issued HTML…');
    const uploadUrl = 'https://graph.microsoft.com/v1.0/drives/' + driveId
      + '/items/' + driveItemId + '/content';
    // Overwrite unconditionally — no If-Match. We are deliberately replacing the
    // draft with the issued HTML. A concurrent metadata touch (e.g. PA099.14
    // setting DraftNotificationSent) changes the driveItem eTag and would otherwise
    // cause a spurious 412 resourceModified. Content is only ever written by this
    // function, so dropping the concurrency check carries no lost-update risk.
    await graphPut(uploadUrl, token, htmlBuffer, 'text/html', null);
    context.log('Issued HTML uploaded.');

    // ── 5a. Rename DriveItem to {invoiceNumber}.html ────────────────────────────
    // Renames the file from 'DRAFT - ...' to match the convention used by all
    // other invoices in the library (e.g. '13299.html').
    context.log('Renaming DriveItem to ' + invoiceNumber + '.html…');
    const renameUrl = 'https://graph.microsoft.com/v1.0/drives/' + driveId
      + '/items/' + driveItemId;
    await graphPatch(renameUrl, token, { name: String(invoiceNumber) + '.html' });
    context.log('DriveItem renamed.');

    // ── 6. Patch Invoice Library item ─────────────────────────────────────────
    const invoiceDateIso = invoiceDate.toISOString();
    // Due date = invoice date + 30 days — the SAME rule doHtmlReplace() uses to print
    // the due date on the client's invoice document. Written to SP so the invoice can
    // age, be chased, and appear in Overdue/aging views. Before S68 it was printed and
    // thrown away: every invoice from 13290 onwards has no DueDate in SharePoint.
    const dueDate = new Date(invoiceDate);
    dueDate.setDate(dueDate.getDate() + 30);
    const dueDateIso = dueDate.toISOString();
    context.log('Patching Invoice Library item…');
    await patchListItem(token, INVOICE_LIB, listItemId, {
      OrderDetails: String(invoiceNumber),
      InvoiceDate:  invoiceDateIso,
      DueDate:      dueDateIso,
      AmountDue:    grandTotal,
    });
    context.log('Invoice Library item patched.');

    // ── 7. Read DraftWipIds from Invoice Library item ─────────────────────────
    context.log('Reading DraftWipIds…');
    const fieldsUrl  = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
      + '/lists/' + INVOICE_LIB + '/items/' + listItemId + '/fields';
    const itemFields = await graphGet(fieldsUrl, token);
    const wipIdsCsv  = itemFields.DraftWipIds || '';
    const wipIds     = wipIdsCsv.split(',').map(s => s.trim()).filter(Boolean);
    context.log('DraftWipIds:', wipIds);

    // ── 8. Mark TT2 entries as Billed ─────────────────────────────────────────
    const tt2Errors = [];
    if (wipIds.length > 0) {
      context.log('Marking', wipIds.length, 'TT2 entries as billed…');
      let tt2Done = 0, tt2Failed = 0;
      for (const wipId of wipIds) {
        try {
          await patchListItem(token, TT2, wipId, { 'Billed_x003f_': true });
          tt2Done++;
        } catch (e) {
          context.log.error('TT2 patch failed for id', wipId, ':', e.message);
          tt2Errors.push({ wipId, error: e.message });
          tt2Failed++;
        }
      }
      context.log('TT2: billed=' + tt2Done + ', failed=' + tt2Failed);
    }

    // ── 9. Update Line Items with real invoice number ─────────────────────────
    if (wipIds.length > 0) {
      context.log('Updating Line Items with invoice number…');
      let liDone = 0, liFailed = 0;
      for (const wipId of wipIds) {
        try {
          const liSearchUrl = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
            + '/lists/' + LINE_ITEMS + '/items'
            + '?$expand=fields($select=id,field_7,InvoiceIDRef)'
            + '&$filter=fields/field_7 eq \'' + wipId + '\'';
          const liResult = await graphGetUnindexed(liSearchUrl, token);
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
    // Collect TT2 and Line Item errors for diagnostics
    const liErrors  = [];

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        invoiceNumber: String(invoiceNumber),
        invoiceDate:   invoiceDateIso,
        tt2Done:       wipIds.length - tt2Errors.length,
        tt2Failed:     tt2Errors.length,
        tt2Errors,
        wipIds,
      }),
    };

  } catch (err) {
    context.log.error('issueinvoice error:', err.message, err.stack);
    context.res = { status: 500, body: 'Error: ' + err.message };
  }
};

// ── HTML replace ──────────────────────────────────────────────────────────────
//
// Replaces all DRAFT placeholders in the draft HTML with real values.
// Operates on string — no DOM parsing needed.
//
// Replacements:
//   1. <title>Invoice DRAFT</title>           → <title>Invoice {N}</title>
//   2. >DRAFT<  (invoice number cell)          → >{N}<
//   3. DRAFT — date assigned on issue          → DD/MM/YYYY
//   4. Due date cell containing only —         → DD/MM/YYYY  (3rd meta-table row right cell)
//   5. <div class="draft-banner">…</div>       → (removed)

function doHtmlReplace(html, invoiceNumber, invoiceDate) {
  const dateStr    = _fmtDate(invoiceDate);
  const dueDate    = new Date(invoiceDate);
  dueDate.setDate(dueDate.getDate() + 30);
  const dueDateStr = _fmtDate(dueDate);

  // 1. Page title
  html = html.replace(
    '<title>Invoice DRAFT</title>',
    '<title>Invoice ' + invoiceNumber + '</title>'
  );

  // 2. Invoice number cell — the cell contains exactly >DRAFT< as its text node
  //    The meta-table renders: <td>DRAFT</td>  (right-hand cell, first row)
  //    We replace only the first occurrence to avoid touching any other DRAFT text
  html = html.replace('>DRAFT<', '>' + invoiceNumber + '<');

  // 3. Invoice date placeholder
  html = html.replace(
    'DRAFT — date assigned on issue',
    dateStr
  );

  // 4. Due date — the draft renders as plain '—' inside the right-hand <td> of the
  //    Due Date row. To avoid replacing other '—' cells, match the specific em-dash
  //    that appears immediately after the Due Date label cell.
  //    The HTML pattern is: <td>Due Date</td><td>—</td>
  html = html.replace(
    '<td>Due Date</td><td>—</td>',
    '<td>Due Date</td><td>' + dueDateStr + '</td>'
  );

  // 5. Remove draft banner — matches from opening tag to closing </div>
  //    Uses a non-greedy match so we don't eat past the first </div>
  html = html.replace(/<div class="draft-banner">[\s\S]*?<\/div>/, '');

  return html;
}

function _fmtDate(d) {
  const dt = (d instanceof Date) ? d : new Date(d);
  const dd = String(dt.getDate()).padStart(2, '0');
  const mm = String(dt.getMonth() + 1).padStart(2, '0');
  const yyyy = dt.getFullYear();
  return dd + '/' + mm + '/' + yyyy;
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

// Downloads a file as a UTF-8 string (follows the 302 redirect that Graph returns
// for /content requests — Node's https does NOT auto-follow redirects).
function graphGetFile(url, token) {
  return new Promise(function (resolve, reject) {
    function doRequest(requestUrl, authToken) {
      const u = new URL(requestUrl);
      const options = {
        hostname: u.hostname,
        path:     u.pathname + u.search,
        method:   'GET',
        headers:  { Authorization: 'Bearer ' + authToken },
      };
      const req = https.request(options, function (res) {
        // Graph /content returns a 302 redirect to a pre-authenticated download URL
        if (res.statusCode === 302 || res.statusCode === 301) {
          const location = res.headers['location'];
          if (!location) { reject(new Error('Redirect with no Location header')); return; }
          // Follow redirect — no auth header needed on the redirect target
          doRedirect(location);
          return;
        }
        if (res.statusCode >= 400) {
          let data = '';
          res.on('data', chunk => { data += chunk; });
          res.on('end', () => reject(new Error('Graph GET file ' + res.statusCode + ': ' + data.slice(0, 400))));
          return;
        }
        const chunks = [];
        res.on('data', chunk => { chunks.push(chunk); });
        res.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
      });
      req.on('error', reject);
      req.end();
    }

    function doRedirect(redirectUrl) {
      const u = new URL(redirectUrl);
      const options = {
        hostname: u.hostname,
        path:     u.pathname + u.search,
        method:   'GET',
      };
      const req = https.request(options, function (res) {
        if (res.statusCode >= 400) {
          let data = '';
          res.on('data', chunk => { data += chunk; });
          res.on('end', () => reject(new Error('File redirect GET ' + res.statusCode + ': ' + data.slice(0, 200))));
          return;
        }
        const chunks = [];
        res.on('data', chunk => { chunks.push(chunk); });
        res.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
      });
      req.on('error', reject);
      req.end();
    }

    doRequest(url, token);
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

// etag param: pass the driveItem eTag to satisfy Graph's concurrency check.
// If null, omit the If-Match header entirely (Graph accepts new files without it).
function graphPut(url, token, buffer, contentType, etag) {
  return new Promise(function (resolve, reject) {
    const u = new URL(url);
    const headers = {
      Authorization:    'Bearer ' + token,
      'Content-Type':   contentType,
      'Content-Length': buffer.length,
    };
    if (etag) headers['If-Match'] = etag;
    const options = {
      hostname: u.hostname,
      path:     u.pathname + u.search,
      method:   'PUT',
      headers,
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
