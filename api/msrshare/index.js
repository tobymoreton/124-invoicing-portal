/**
 * P124 — Invoicing Portal
 * Azure Function: /api/msrshare
 *
 * Uploads a generated Matter Status Report .xlsx to the SharePoint **Admin Library**
 * and returns its webUrl, so the file can be circulated to the draftsmen for editing
 * in Excel Online (co-authoring — one shared file, no forks).
 *
 * DELIBERATE DESIGN: this is a **dead snapshot**. Nothing written into the workbook
 * flows back to SharePoint. Current Position and the rest are refreshed by the normal
 * overnight process. Do NOT build a write-back path — that is the divergent-copy bug
 * that S72 removed.
 *
 *   POST — raw .xlsx bytes in the body; filename in header x-file-name (URL-encoded).
 *          Management only (Toby + Danielle).
 *          Returns { uploaded, name, webUrl }.
 *
 * Target: SharePoint document library "Admin Library" on /sites/TMCLegalLimited,
 * subfolder "Matter Status Reports" (created on first use).
 *
 * ⚠ The library is resolved BY NAME at runtime (no GUID recorded for it). If the
 *   library is ever renamed, this errors loudly listing the names it did find —
 *   it never silently writes somewhere else.
 *
 * ⚠ No sharing link is minted. The file relies on the Admin Library's existing
 *   permissions; anyone who can already open that library can open the file.
 *   If draftsmen cannot open it, the fix is a library permission, not this code.
 */

const https   = require('https');
const { URL } = require('url');

const SITE_PATH   = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';
const LIBRARY_NAME = 'Admin Library';
const FOLDER_NAME  = 'Matter Status Reports';

// Management only — this publishes a file others will act on.
const ALLOWED_EMAILS = [
  'toby@tmclegal.co.uk',
  'danielle@tmclegal.co.uk',
];

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

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

module.exports = async function (context, req) {
  context.log('P124 /api/msrshare called');

  const callerEmail = getCallerEmail(req);
  if (!callerEmail || !ALLOWED_EMAILS.includes(callerEmail)) {
    context.res = { status: 403, body: 'Forbidden — publishing a Matter Status Report is restricted to Management.' };
    return;
  }

  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    context.res = { status: 500, body: 'Missing required app settings.' };
    return;
  }

  const fileName = decodeURIComponent((req.headers['x-file-name'] || '').trim());
  if (!fileName) {
    context.res = { status: 400, body: 'Missing required header: x-file-name.' };
    return;
  }
  if (!/\.xlsx$/i.test(fileName)) {
    context.res = { status: 400, body: 'x-file-name must end in .xlsx.' };
    return;
  }

  // Body is the raw workbook bytes.
  let buffer = req.body;
  if (!Buffer.isBuffer(buffer)) {
    if (typeof buffer === 'string') buffer = Buffer.from(buffer, 'binary');
    else if (buffer == null)        buffer = Buffer.alloc(0);
    else                            buffer = Buffer.from(buffer);
  }
  if (!buffer.length) {
    context.res = { status: 400, body: 'Empty workbook body.' };
    return;
  }

  try {
    const token = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);

    // 1) Resolve the Admin Library drive BY NAME. There is no recorded GUID for this
    //    library, and guessing one would be worse than looking it up.
    const drives = await graphGet(
      'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH + '/drives?$select=id,name', token);
    const list   = (drives && drives.value) || [];
    const drive  = list.find(d => (d.name || '').toLowerCase() === LIBRARY_NAME.toLowerCase());
    if (!drive || !drive.id) {
      throw new Error('Could not find a document library named "' + LIBRARY_NAME + '". Libraries found: '
        + list.map(d => d.name).join(', '));
    }
    const driveId = drive.id;

    // 2) Ensure the target folder exists. conflictBehavior=replace is safe for a folder —
    //    it does not touch the folder's contents.
    try {
      await graphPost('https://graph.microsoft.com/v1.0/drives/' + driveId + '/root/children', token, {
        name: FOLDER_NAME,
        folder: {},
        '@microsoft.graph.conflictBehavior': 'replace',
      });
    } catch (e) {
      context.log('Folder ensure returned: ' + e.message + ' (continuing — it may already exist)');
    }

    // 3) PUT the bytes. Same-name upload replaces + versions the previous snapshot,
    //    so re-publishing today's report keeps the SAME link. The filename carries the
    //    date, so yesterday's snapshot is a separate, untouched file.
    //    NOTE: /lists/{guid}/drive/root:/... is NOT a valid Graph address — always
    //    address the drive canonically via /drives/{driveId} (proven S63).
    const putUrl = 'https://graph.microsoft.com/v1.0/drives/' + driveId
                 + '/root:/' + encodeURIComponent(FOLDER_NAME) + '/' + encodeURIComponent(fileName)
                 + ':/content';
    const item = await graphPut(putUrl, token, buffer, XLSX_MIME);
    if (!item || !item.id) throw new Error('Upload returned no driveItem.');

    context.log('AUDIT msrshare POST name=' + fileName + ' by=' + callerEmail + ' url=' + (item.webUrl || ''));
    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        uploaded: true,
        name:   item.name   || fileName,
        webUrl: item.webUrl || '',
      }),
    };
  } catch (err) {
    context.log.error('Error publishing MSR:', err.message);
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

function graphGet(url, token) {
  return new Promise(function (resolve, reject) {
    const u = new URL(url);
    const options = {
      hostname: u.hostname, path: u.pathname + u.search, method: 'GET',
      headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' },
    };
    const req = https.request(options, function (res) {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        if (res.statusCode >= 400) { reject(new Error('Graph GET ' + res.statusCode + ': ' + data.slice(0, 300))); return; }
        try { resolve(JSON.parse(data)); } catch (e) { reject(new Error('JSON parse error: ' + e.message)); }
      });
    });
    req.on('error', reject);
    req.end();
  });
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

function graphPut(url, token, buffer, contentType) {
  return new Promise(function (resolve, reject) {
    const u = new URL(url);
    const options = {
      hostname: u.hostname, path: u.pathname + u.search, method: 'PUT',
      headers: {
        Authorization:    'Bearer ' + token,
        'Content-Type':   contentType || 'application/octet-stream',
        'Content-Length': buffer.length,
      },
    };
    const req = https.request(options, function (res) {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        if (res.statusCode >= 400) { reject(new Error('Graph PUT ' + res.statusCode + ': ' + data.slice(0, 300))); return; }
        try { resolve(JSON.parse(data)); } catch (e) { resolve({}); }
      });
    });
    req.on('error', reject);
    req.write(buffer);
    req.end();
  });
}
