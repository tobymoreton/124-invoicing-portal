/**
 * P124 — Invoicing Portal
 * Azure Function: /api/invoicepdf
 *
 * S120. Returns an invoice as a real PDF, converted on demand.
 *
 * Why this exists: every invoice in the Invoice Library is stored as HTML
 * ({invoiceNumber}.html — see issueinvoice-index.js). The link on the invoice drawer
 * opened that HTML in SharePoint, which clients cannot open or read when it is sent to
 * them. Rather than storing a second copy of every invoice, this converts on request:
 * Microsoft Graph will render a driveItem into PDF, and html is one of its supported
 * source formats.
 *
 * Nothing is written. No second file lands in the Invoice Library (which would create a
 * second list item and show up as a phantom row in every invoice list and statement), and
 * the PDF can never drift out of step with the HTML because it is made from it each time.
 *
 * GET /api/invoicepdf?listItemId=1234
 *   → 200 application/pdf, Content-Disposition: attachment; filename="{name}.pdf"
 *
 * Any signed-in portal user may download an invoice — the drawer that links here is
 * already visible to everyone.
 *
 * GUIDs:
 *   Invoice Library:  5c366b19-0da9-4be9-b68f-60e6a0209cdb
 */

const https   = require('https');
const { URL } = require('url');

const SITE_PATH   = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';
const INVOICE_LIB = '5c366b19-0da9-4be9-b68f-60e6a0209cdb';

// ── Entry point ───────────────────────────────────────────────────────────────

module.exports = async function (context, req) {
  context.log('P124 /api/invoicepdf called');

  const callerEmail = getCallerEmail(req);
  if (!callerEmail) {
    context.res = { status: 403, body: 'Forbidden — sign in to download an invoice.' };
    return;
  }

  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    context.res = { status: 500, body: 'Missing required app settings.' };
    return;
  }

  const listItemId = String((req.query && req.query.listItemId) || '').trim();
  if (!listItemId || !/^\d+$/.test(listItemId)) {
    context.res = { status: 400, body: 'listItemId is required and must be numeric.' };
    return;
  }

  try {
    const token = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);

    // ── 1. Resolve the list item to its underlying file ───────────────────────
    const itemUrl = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
      + '/lists/' + INVOICE_LIB + '/items/' + listItemId + '?$expand=driveItem';
    const listItemData = await graphGet(itemUrl, token);
    const driveItem    = listItemData.driveItem;
    const driveItemId  = driveItem && driveItem.id;
    if (!driveItemId) throw new Error('No file attached to Invoice Library item ' + listItemId);

    const driveUrl = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
      + '/lists/' + INVOICE_LIB + '/drive';
    const driveInfo = await graphGet(driveUrl, token);
    const driveId   = driveInfo.id;
    if (!driveId) throw new Error('Could not get drive ID for Invoice Library.');

    // Name the download after the invoice, not after the stored file: '13299.html' → '13299.pdf'.
    const sourceName = String(driveItem.name || ('invoice-' + listItemId));
    const baseName   = sourceName.replace(/\.[^.]+$/, '');
    const pdfName    = sanitiseFileName(baseName) + '.pdf';

    // ── 2. Ask Graph for the PDF rendering ────────────────────────────────────
    // Graph converts to PDF from a list of source formats that includes htm and html.
    // The call answers 302 with a short-lived pre-authenticated URL; graphGetBinary
    // follows it without the Authorization header, as Graph requires.
    context.log('Converting', sourceName, 'to PDF…');
    const convertUrl = 'https://graph.microsoft.com/v1.0/drives/' + driveId
      + '/items/' + driveItemId + '/content?format=pdf';
    const pdfBuffer = await graphGetBinary(convertUrl, token);

    // A conversion that quietly returns something other than a PDF must not be handed to
    // the client as one — it would download as a broken file with no explanation.
    if (!pdfBuffer || pdfBuffer.length < 100 || pdfBuffer.slice(0, 4).toString('latin1') !== '%PDF') {
      throw new Error('Graph did not return a PDF for ' + sourceName
        + ' (' + (pdfBuffer ? pdfBuffer.length : 0) + ' bytes).');
    }
    context.log('PDF returned,', pdfBuffer.length, 'bytes.');

    context.res = {
      status:  200,
      headers: {
        'Content-Type':        'application/pdf',
        'Content-Disposition': 'attachment; filename="' + pdfName + '"',
        'Content-Length':      String(pdfBuffer.length),
        'Cache-Control':       'no-store',
      },
      isRaw: true,
      body:  pdfBuffer,
    };
  } catch (err) {
    context.log.error('invoicepdf error:', err.message, err.stack);
    context.res = { status: 500, body: 'Could not produce a PDF for this invoice: ' + err.message };
  }
};

// ── Helpers ───────────────────────────────────────────────────────────────────

function sanitiseFileName(s) {
  return String(s || 'invoice').replace(/[\\/:*?"<>|]/g, '-').slice(0, 80) || 'invoice';
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

function graphGet(url, token) {
  return new Promise(function (resolve, reject) {
    const u = new URL(url);
    const options = {
      hostname: u.hostname,
      path:     u.pathname + u.search,
      method:   'GET',
      headers:  { Authorization: 'Bearer ' + token },
    };
    const req = https.request(options, function (res) {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        if (res.statusCode >= 400) {
          reject(new Error('Graph GET ' + res.statusCode + ': ' + data.slice(0, 400)));
          return;
        }
        try { resolve(JSON.parse(data)); } catch (e) { reject(e); }
      });
    });
    req.on('error', reject);
    req.end();
  });
}

// Binary sibling of issueinvoice's graphGetFile. That one decodes to utf8, which would
// corrupt a PDF; this returns the Buffer untouched. Redirects are followed without the
// Authorization header — the pre-authenticated URL rejects it — and are capped so a
// redirect loop cannot hang the function.
function graphGetBinary(url, token) {
  return new Promise(function (resolve, reject) {
    function go(requestUrl, authToken, hops) {
      if (hops > 5) { reject(new Error('Too many redirects fetching the converted file.')); return; }
      const u = new URL(requestUrl);
      const headers = authToken ? { Authorization: 'Bearer ' + authToken } : {};
      const options = {
        hostname: u.hostname,
        path:     u.pathname + u.search,
        method:   'GET',
        headers,
      };
      const req = https.request(options, function (res) {
        if (res.statusCode === 301 || res.statusCode === 302 || res.statusCode === 303 || res.statusCode === 307) {
          const location = res.headers['location'];
          if (!location) { reject(new Error('Redirect with no Location header')); return; }
          res.resume();
          go(location, null, hops + 1);
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
        res.on('end', () => resolve(Buffer.concat(chunks)));
      });
      req.on('error', reject);
      req.end();
    }
    go(url, token, 0);
  });
}
