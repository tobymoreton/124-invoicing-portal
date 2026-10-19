/**
 * P124 — Invoicing Portal
 * Azure Function: /api/caseupdate
 *
 * PATCHes fields on a Cases list item via Graph (app-only auth).
 * All authenticated users can update case fields relevant to their role.
 *
 * POST body (JSON) — update:
 *   itemId   string   — SP list item ID
 *   fields   object   — key/value pairs of SP internal field names to update
 *
 * POST body (JSON) — create (used by admin.html "Book In Case"):
 *   action   'create'
 *   fields   object   — SP internal field names for the new item
 *
 * Returns: { updated: true } | { created: true, id }
 *
 * List: Cases
 * GUID: ae420bda-e550-499c-b337-90e4f33617c1
 */

const https   = require('https');
const { URL } = require('url');

const SITE_PATH = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';
const LIST_GUID = 'ae420bda-e550-499c-b337-90e4f33617c1';

const ALLOWED_EMAILS = [
  'toby@tmclegal.co.uk',
  'danielle@tmclegal.co.uk',
  'lesley@tmclegal.co.uk',
  'joanna@tmclegal.co.uk',
  'tracy@tmclegal.co.uk',
  'kelly@tmclegal.co.uk',
  'tom@tmclegal.co.uk',
  'julie@tmclegal.co.uk',
  'daniel@tmclegal.co.uk',
];

// Deletion is irreversible (real Graph DELETE) and orphans anything pointing at the case.
// Kept narrower than the field-update allowlist above. Lesley added 2026-08-10 at Toby's
// request — she books cases in, so she is the one who creates the duplicates that need
// removing. Admin DOCUMENT deletion (/api/attachmentswrite) is a separate gate and is NOT
// widened by this.
const DELETE_ALLOWED_EMAILS = [
  'toby@tmclegal.co.uk',
  'danielle@tmclegal.co.uk',
  'lesley@tmclegal.co.uk',
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
  context.log('P124 /api/caseupdate called');

  const callerEmail = getCallerEmail(req);

  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    context.res = { status: 500, body: 'Missing required app settings.' };
    return;
  }

  // ── DELETE: remove a Cases item entirely. Management only, irreversible. ────
  if (req.method === 'DELETE') {
    if (!callerEmail || !DELETE_ALLOWED_EMAILS.includes(callerEmail)) {
      context.res = { status: 403, body: 'Forbidden — case deletion is restricted to Management.' };
      return;
    }
    let delBody;
    try {
      delBody = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
    } catch (e) {
      context.res = { status: 400, body: 'Invalid JSON body.' };
      return;
    }
    const delItemId = (delBody && delBody.itemId) || (req.query && req.query.itemId);
    if (!delItemId) {
      context.res = { status: 400, body: 'Missing required field: itemId.' };
      return;
    }
    try {
      const token  = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
      const delUrl = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
                   + '/lists/' + LIST_GUID + '/items/' + delItemId;
      await graphDelete(delUrl, token);
      context.log('AUDIT cases DELETE itemId=' + delItemId + ' deletedBy=' + callerEmail);
      context.res = {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ deleted: true }),
      };
    } catch (err) {
      context.log.error('Error deleting case:', err.message);
      context.res = { status: 500, body: 'Error: ' + err.message };
    }
    return;
  }

  // ── POST: update fields (existing behaviour, unchanged) ──────────────
  if (!callerEmail || !ALLOWED_EMAILS.includes(callerEmail)) {
    context.res = { status: 403, body: 'Forbidden — not authorised to update cases.' };
    return;
  }

  let body;
  try {
    body = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
  } catch (e) {
    context.res = { status: 400, body: 'Invalid JSON body.' };
    return;
  }

  const { itemId, fields } = body || {};

  // ── CREATE: new Cases item (Book In Case). Same allowlist as update. ──────
  if (body && body.action === 'create') {
    if (!fields || typeof fields !== 'object' || !Object.keys(fields).length) {
      context.res = { status: 400, body: 'Missing required field: fields.' };
      return;
    }
    try {
      const token   = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
      const postUrl = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
                    + '/lists/' + LIST_GUID + '/items';
      const created = await graphPost(postUrl, token, { fields: fields });
      context.log('AUDIT cases CREATE id=' + (created && created.id) + ' createdBy=' + callerEmail);
      context.res = {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ created: true, id: created && created.id }),
      };
    } catch (err) {
      context.log.error('Error creating case:', err.message);
      context.res = { status: 500, body: 'Error: ' + err.message };
    }
    return;
  }

  if (!itemId || !fields || typeof fields !== 'object') {
    context.res = { status: 400, body: 'Missing required fields: itemId, fields.' };
    return;
  }

  try {
    const token  = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
    const patchUrl = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
                   + '/lists/' + LIST_GUID + '/items/' + itemId + '/fields';

    await graphPatch(patchUrl, token, fields);

    // S146 (Toby, 19/10/2026): "if a case is closed any actions need to be marked non-billable".
    // When this update closes the case, every unbilled billable action on it is set non-billable.
    // Non-fatal: the case update itself has already succeeded.
    let unbilled = null;
    const statusVal = [fields.StatusMirror, fields.Status]
      .map(v => (v && typeof v === 'object') ? (v.Value || '') : (v || ''))
      .map(v => String(v).trim().toLowerCase());
    if (statusVal.includes('closed')) {
      try { unbilled = await unbillClosedCaseActions(token, itemId); }
      catch (e) { context.log.error('S146 unbill on close failed for case ' + itemId + ':', e.message); unbilled = { error: e.message }; }
    }

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ updated: true, ...(unbilled ? { closedCaseActionsUnbilled: unbilled } : {}) }),
    };
  } catch (err) {
    context.log.error('Error updating case:', err.message);
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
          reject(new Error('Graph DELETE ' + res.statusCode + ': ' + data.slice(0, 300)));
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
    const u       = new URL(url);
    const options = {
      hostname: u.hostname,
      path:     u.pathname + u.search,
      method:   'POST',
      headers: {
        Authorization:  'Bearer ' + token,
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(payload),
      },
    };
    const req = https.request(options, function (res) {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        if (res.statusCode >= 400) {
          reject(new Error('Graph POST ' + res.statusCode + ': ' + data.slice(0, 300)));
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

function graphPatch(url, token, body) {
  return new Promise(function (resolve, reject) {
    const payload = JSON.stringify(body);
    const u       = new URL(url);
    const options = {
      hostname: u.hostname,
      path:     u.pathname + u.search,
      method:   'PATCH',
      headers: {
        Authorization:  'Bearer ' + token,
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(payload),
      },
    };

    const req = https.request(options, function (res) {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        if (res.statusCode >= 400) {
          reject(new Error('Graph PATCH ' + res.statusCode + ': ' + data.slice(0, 300)));
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

// ── S146: unbill a closed case's actions ─────────────────────────────────────
const TT2_GUID = '67db204c-30a5-4f4d-b276-60852d9967e1';
async function unbillClosedCaseActions(token, caseItemId) {
  const c = await graphGetJson('https://graph.microsoft.com/v1.0/sites/' + SITE_PATH + '/lists/' + LIST_GUID
    + '/items/' + caseItemId + '/fields?$select=Ourreference_x0028_text_x0029_', token);
  const ref = String(c.Ourreference_x0028_text_x0029_ || '').trim();
  if (!ref) return { ref: null, count: 0 };
  // No $select: Boolean columns are silently dropped from a $select on this tenant.
  let url = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH + '/lists/' + TT2_GUID
    + "/items?$expand=fields&$top=500&$filter=fields/field_16 eq '" + ref.replace(/'/g, "''") + "'";
  let count = 0, failed = 0;
  while (url) {
    const page = await graphGetJson(url, token);
    for (const it of (page.value || [])) {
      const f = it.fields || {};
      if (f.Billable_x003f_ === true && f.Billed_x003f_ !== true) {
        try {
          await graphPatch('https://graph.microsoft.com/v1.0/sites/' + SITE_PATH + '/lists/' + TT2_GUID
            + '/items/' + it.id + '/fields', token, { Billable_x003f_: false });
          count++;
        } catch (e) { failed++; }
      }
    }
    url = page['@odata.nextLink'] || null;
  }
  return { ref, count, failed };
}

function graphGetJson(url, token) {
  return new Promise(function (resolve, reject) {
    const u = new URL(url);
    const req = https.request({
      hostname: u.hostname, path: u.pathname + u.search, method: 'GET',
      headers: { Authorization: 'Bearer ' + token, Accept: 'application/json',
                 Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' },
    }, function (res) {
      let data = '';
      res.on('data', ch => { data += ch; });
      res.on('end', () => {
        if (res.statusCode >= 400) { reject(new Error('Graph GET ' + res.statusCode + ': ' + data.slice(0, 200))); return; }
        try { resolve(JSON.parse(data)); } catch (e) { reject(e); }
      });
    });
    req.on('error', reject);
    req.end();
  });
}
