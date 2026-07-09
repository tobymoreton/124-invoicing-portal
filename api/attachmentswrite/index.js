/**
 * P124 — Invoicing Portal
 * Azure Function: /api/attachmentswrite
 *
 * Write operations on the shared `caseAttachments` SharePoint document library
 * via Graph (app-only client-credentials auth).
 *
 *   POST   — upload a file. Raw file bytes in the body; metadata in headers:
 *              x-file-name    (URL-encoded filename, required)
 *              x-our-ref      (URL-encoded case Our Reference, required)
 *              x-doc-type     (URL-encoded Type/choice value, optional)
 *              x-content-type (original MIME type, optional)
 *            Flow: PUT bytes to the library, then PATCH the OurRef/docType columns.
 *
 *   DELETE — remove a file by driveItem id (?id=<driveItemId>). Management only, irreversible.
 *
 * ⚠ Requires the app registration to hold a WRITE scope (Sites.ReadWrite.All /
 *   Files.ReadWrite.All). If it is read-only, both POST and DELETE return Graph 403.
 *
 * Library: caseAttachments
 * GUID: 710dea64-11ae-4ae7-8fde-d4508206e1c1
 * Columns: OurRef (Graph key, capital O) / docType. (See /api/attachments read fn for the casing note.)
 */

const https   = require('https');
const { URL } = require('url');

const LIST_GUID      = '710dea64-11ae-4ae7-8fde-d4508206e1c1';
const SITE_PATH      = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';
const ALLOWED_DOMAIN = '@tmclegal.co.uk';

// Deletion is irreversible (real Graph DELETE) — Management only, deliberately narrower
// than upload (any signed-in TMC user). Mirrors caseupdate's DELETE_ALLOWED_EMAILS.
const DELETE_ALLOWED_EMAILS = [
  'toby@tmclegal.co.uk',
  'danielle@tmclegal.co.uk',
];

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
  context.log('P124 /api/attachmentswrite called: ' + req.method);

  const callerEmail = getCallerEmail(req);

  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    context.res = { status: 500, body: 'Missing required app settings.' };
    return;
  }

  // ── DELETE: remove a library file by driveItem id. Management only, irreversible. ──
  if (req.method === 'DELETE') {
    if (!callerEmail || !DELETE_ALLOWED_EMAILS.includes(callerEmail)) {
      context.res = { status: 403, body: 'Forbidden — deleting attachments is restricted to Management.' };
      return;
    }
    const id = (req.query && req.query.id || '').trim();
    if (!id) {
      context.res = { status: 400, body: 'Missing required query param: id (driveItem id).' };
      return;
    }
    try {
      const token = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
      // Resolve the library's drive id, then delete via the canonical /drives/{id}/items/{id}.
      // (/lists/{guid}/drive/items/{id} is NOT a valid Graph address — returns 400 BadRequest.)
      const driveResp = await graphGet('https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
                      + '/lists/' + LIST_GUID + '/drive?$select=id', token);
      const driveId = driveResp && driveResp.id;
      if (!driveId) throw new Error('Could not resolve the caseAttachments drive id.');
      const url = 'https://graph.microsoft.com/v1.0/drives/' + driveId + '/items/' + encodeURIComponent(id);
      await graphDelete(url, token);
      context.log('AUDIT attachment DELETE id=' + id + ' deletedBy=' + callerEmail);
      context.res = {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ deleted: true }),
      };
    } catch (err) {
      context.log.error('Error deleting attachment:', err.message);
      context.res = { status: 500, body: 'Error: ' + err.message };
    }
    return;
  }

  // ── POST: upload a file, then stamp its OurRef/docType columns. ──
  if (!callerEmail || callerEmail.indexOf(ALLOWED_DOMAIN) === -1) {
    context.res = { status: 403, body: 'Forbidden — you must be signed in with a TMC account.' };
    return;
  }

  const fileName = decodeURIComponent((req.headers['x-file-name']    || '').trim());
  const ourRef   = decodeURIComponent((req.headers['x-our-ref']      || '').trim());
  const docType  = decodeURIComponent((req.headers['x-doc-type']     || '').trim());
  const ctype    = decodeURIComponent((req.headers['x-content-type'] || '').trim()) || 'application/octet-stream';

  if (!fileName || !ourRef) {
    context.res = { status: 400, body: 'Missing required headers: x-file-name and x-our-ref.' };
    return;
  }

  // Body is the raw file bytes. Azure Functions gives a Buffer for binary bodies;
  // fall back to Buffer.from if a string arrives.
  let buffer = req.body;
  if (!Buffer.isBuffer(buffer)) {
    if (typeof buffer === 'string') buffer = Buffer.from(buffer, 'binary');
    else if (buffer == null)        buffer = Buffer.alloc(0);
    else                            buffer = Buffer.from(buffer);
  }
  if (!buffer.length) {
    context.res = { status: 400, body: 'Empty file body.' };
    return;
  }

  try {
    const token = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);

    // 1) Upload the bytes. Simple content PUT (suits files up to Graph's simple-upload limit;
    //    very large files would need an upload session — not implemented in v1).
    //    Same-name upload replaces + versions the existing file (SharePoint keeps history).
    const putUrl = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
                 + '/lists/' + LIST_GUID + '/drive/root:/' + encodeURIComponent(fileName) + ':/content';
    const item = await graphPut(putUrl, token, buffer, ctype);
    const driveItemId = item && item.id;
    if (!driveItemId) throw new Error('Upload succeeded but no driveItem id was returned.');

    // 2) Resolve the backing listItem id so we can stamp the columns.
    const liUrl = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
                + '/lists/' + LIST_GUID + '/drive/items/' + encodeURIComponent(driveItemId)
                + '/listItem?$select=id';
    const li = await graphGet(liUrl, token);
    const listItemId = li && li.id;

    // 3) Stamp OurRef + docType (best-effort; upload already succeeded even if this fails).
    let stamped = false;
    if (listItemId) {
      const patchUrl = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
                     + '/lists/' + LIST_GUID + '/items/' + encodeURIComponent(listItemId) + '/fields';
      const fields = { OurRef: ourRef };
      if (docType) fields.docType = docType;
      await graphPatch(patchUrl, token, fields);
      stamped = true;
    }

    context.log('AUDIT attachment POST name=' + fileName + ' ref=' + ourRef + ' by=' + callerEmail);
    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ uploaded: true, id: driveItemId, name: item.name || fileName, stamped }),
    };
  } catch (err) {
    context.log.error('Error uploading attachment:', err.message);
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

function graphDelete(url, token) {
  return new Promise(function (resolve, reject) {
    const u = new URL(url);
    const options = {
      hostname: u.hostname, path: u.pathname + u.search, method: 'DELETE',
      headers: { Authorization: 'Bearer ' + token },
    };
    const req = https.request(options, function (res) {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        if (res.statusCode >= 400) { reject(new Error('Graph DELETE ' + res.statusCode + ': ' + data.slice(0, 300))); return; }
        resolve({});
      });
    });
    req.on('error', reject);
    req.end();
  });
}

function graphPatch(url, token, body) {
  return new Promise(function (resolve, reject) {
    const payload = JSON.stringify(body);
    const u       = new URL(url);
    const options = {
      hostname: u.hostname, path: u.pathname + u.search, method: 'PATCH',
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
        if (res.statusCode >= 400) { reject(new Error('Graph PATCH ' + res.statusCode + ': ' + data.slice(0, 300))); return; }
        try { resolve(JSON.parse(data)); } catch (e) { resolve({}); }
      });
    });
    req.on('error', reject);
    req.write(payload);
    req.end();
  });
}
