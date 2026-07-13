const https = require('https');
const { URL } = require('url');

const SITE_PATH = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';
const ALLOWED_DOMAIN = '@tmclegal.co.uk';
// Draftsmen must be able to ADD and EDIT case-related reference data (opponent firms,
// caseworkers, fee earners, courts, client firms) without waiting on management — Toby's
// call, S73, after Tracy was blocked mid-case. CREATE and EDIT are open to any
// authenticated TMC user. DELETE stays restricted: a deleted reference record is
// unrecoverable from the portal and orphans every case already pointing at it.
const DELETE_EMAILS = ['toby@tmclegal.co.uk','danielle@tmclegal.co.uk','lesley@tmclegal.co.uk'];
const LISTS = {
  feeearners:    '750616ac-5c2e-4a3c-91d9-b0e6cad1e6e9',
  courts:        'e867d355-7eb4-40c7-a790-ee1c591a1361',
  opponentfirms: 'e456c697-bf02-42da-baea-f20b3e653857',
  opponentcws:   '0a86dd15-ef3b-4460-8236-b722d22cdc51',
};

// NOTE: the GET path uses a bare $expand=fields for every list. An earlier
// $select workaround existed to dodge a Graph 503 on feeearners, but the real
// cause was four corrupt "Firm: Address Line" projected-lookup columns on the
// SP Fee Earners list, which crashed the list in SP itself and made Graph 503.
// Those columns were deleted (2026-06-30), so a bare expand works again and
// returns all real fields without per-list field-name maintenance.

function getCallerEmail(req) {
  try {
    const h = req.headers && req.headers['x-ms-client-principal'];
    if (!h) return null;
    const p = JSON.parse(Buffer.from(h, 'base64').toString('utf8'));
    if (p.userDetails) return p.userDetails.toLowerCase();
    const c = (p.claims || []).find(c => ['preferred_username','email','upn'].includes(c.typ));
    return c ? c.val.toLowerCase() : null;
  } catch { return null; }
}

async function getToken(tid, cid, cs) {
  return new Promise((resolve, reject) => {
    const body = 'grant_type=client_credentials&client_id=' + encodeURIComponent(cid) + '&client_secret=' + encodeURIComponent(cs) + '&scope=https%3A%2F%2Fgraph.microsoft.com%2F.default';
    const req = https.request({ hostname: 'login.microsoftonline.com', path: '/' + tid + '/oauth2/v2.0/token', method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(body) } }, res => {
      let d = ''; res.on('data', c => d += c);
      res.on('end', () => { try { const j = JSON.parse(d); j.access_token ? resolve(j.access_token) : reject(new Error(j.error_description)); } catch(e) { reject(e); } });
    });
    req.on('error', reject); req.write(body); req.end();
  });
}

async function graphGet(url, token) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    const req = https.request({ hostname: u.hostname, path: u.pathname + u.search, method: 'GET', headers: { Authorization: 'Bearer ' + token, Accept: 'application/json', Prefer: 'allowthrottleablequeries' } }, res => {
      let d = ''; res.on('data', c => d += c);
      res.on('end', () => { if (res.statusCode >= 400) { reject(new Error('Graph ' + res.statusCode + ': ' + d.slice(0,200))); return; } try { resolve(JSON.parse(d)); } catch(e) { reject(e); } });
    });
    req.on('error', reject); req.end();
  });
}

async function graphPost(url, token, body) {
  return new Promise((resolve, reject) => {
    const bs = JSON.stringify(body); const u = new URL(url);
    const req = https.request({ hostname: u.hostname, path: u.pathname + u.search, method: 'POST', headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json', Accept: 'application/json', 'Content-Length': Buffer.byteLength(bs) } }, res => {
      let d = ''; res.on('data', c => d += c);
      res.on('end', () => { if (res.statusCode >= 400) { reject(new Error('Graph POST ' + res.statusCode + ': ' + d.slice(0,200))); return; } try { resolve(JSON.parse(d)); } catch(e) { resolve({}); } });
    });
    req.on('error', reject); req.write(bs); req.end();
  });
}

async function graphPatch(url, token, body) {
  return new Promise((resolve, reject) => {
    const bs = JSON.stringify(body); const u = new URL(url);
    const req = https.request({ hostname: u.hostname, path: u.pathname + u.search, method: 'PATCH', headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json', Accept: 'application/json', 'Content-Length': Buffer.byteLength(bs) } }, res => {
      let d = ''; res.on('data', c => d += c);
      res.on('end', () => { if (res.statusCode >= 400) { reject(new Error('Graph PATCH ' + res.statusCode + ': ' + d.slice(0,200))); return; } try { resolve(JSON.parse(d)); } catch(e) { resolve({}); } });
    });
    req.on('error', reject); req.write(bs); req.end();
  });
}

async function graphDelete(url, token) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    const req = https.request({ hostname: u.hostname, path: u.pathname + u.search, method: 'DELETE', headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' } }, res => {
      let d = ''; res.on('data', c => d += c);
      res.on('end', () => { if (res.statusCode >= 400) { reject(new Error('Graph DELETE ' + res.statusCode + ': ' + d.slice(0,200))); return; } resolve({}); });
    });
    req.on('error', reject); req.end();
  });
}

function buildFields(list, body) {
  const f = {};
  if (list === 'feeearners') {
    if (body.title      != null) f['Title']                           = body.title;
    // First/Last name are SP Calculated columns on this list (formula-derived
    // from Title, e.g. '=LEFT(Title,FIND(" ",Title)-1)') - Graph returns 403
    // accessDenied on any attempt to write them directly. Do NOT write them;
    // they populate automatically from Title the moment it's set.
    if (body.email      != null) f['Email']                           = body.email;
    if (body.directLine != null) f['Directline']                      = body.directLine;
    if (body.status     != null) f['Status']                          = body.status;
    if (body.notes      != null) f['Notes']                           = body.notes;
    if (body.qualified  != null) f['Dateofqualification']             = body.qualified;
    if (body.firmToLink != null) f['FirmToLink']                       = body.firmToLink;
    // Field names confirmed from live Graph data 2026-06-30 (record 'Abbie Roberts').
    // 'Experience' internal name not yet confirmed (no populated record) - omitted
    // from write until verified.
  } else if (list === 'courts') {
    if (body.title    != null) f['Title']    = body.title;
    if (body.address1 != null) f['Address1'] = body.address1;
    if (body.address2 != null) f['Address2'] = body.address2;
    if (body.address3 != null) f['Address3'] = body.address3;
    if (body.address4 != null) f['Address4'] = body.address4;
    if (body.address5 != null) f['Address5'] = body.address5;
  } else if (list === 'opponentfirms') {
    if (body.title    != null) f['Title']                       = body.title;
    if (body.address1 != null) f['Address_x0020_line_x0020_1'] = body.address1;
    if (body.address2 != null) f['Address_x0020_line_x0020_2'] = body.address2;
    if (body.address3 != null) f['Address_x0020_line_x0020_3'] = body.address3;
    if (body.address4 != null) f['Address_x0020_line_x0020_4'] = body.address4;
    if (body.address5 != null) f['Address_x0020_line_x0020_5'] = body.address5;
    if (body.website  != null) f['Website']                     = body.website;
  } else if (list === 'opponentcws') {
    if (body.title != null) f['Title']              = body.title;
    if (body.phone != null) f['Phone_x0020_number'] = body.phone;
    if (body.email != null) f['Email']              = body.email;
    if (body.firm  != null) f['CaseworkerFirmMirror'] = body.firm;
  }
  return f;
}

module.exports = async function (context, req) {
  const callerEmail = getCallerEmail(req);
  if (!callerEmail || !callerEmail.endsWith(ALLOWED_DOMAIN)) {
    context.res = { status: 403, body: 'Forbidden' }; return;
  }
  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  const listKey = (req.query.list || '').toLowerCase();
  const guid = LISTS[listKey];
  if (!guid) { context.res = { status: 400, body: 'Unknown list: ' + listKey }; return; }

  try {
    const token = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
    const baseUrl = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH + '/lists/' + guid + '/items';

    // GET: return all items sorted by Title
    if (req.method === 'GET') {
      const items = []; let url = baseUrl + '?$expand=fields&$top=999';
      while (url) {
        const data = await graphGet(url, token);
        (data.value || []).forEach(i => items.push({ id: i.id, fields: i.fields || {} }));
        url = data['@odata.nextLink'] || null;
      }
      items.sort((a, b) => (a.fields.Title || '').localeCompare(b.fields.Title || ''));
      context.res = { status: 200, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-cache' }, body: JSON.stringify({ value: items }) };
      return;
    }

    // POST (create) and PATCH (edit) are open to all TMC staff — gated only by the domain
    // check above. DELETE is irreversible and orphans referencing cases: management only.
    if (req.method === 'DELETE' && !DELETE_EMAILS.includes(callerEmail)) {
      context.res = { status: 403, body: 'Deleting a reference record is restricted to Toby, Danielle and Lesley. You can still add and edit entries.' };
      return;
    }

    // POST: create new item
    if (req.method === 'POST') {
      const body = typeof req.body === 'string' ? JSON.parse(req.body) : (req.body || {});
      const fields = buildFields(listKey, body);
      if (!fields['Title']) { context.res = { status: 400, body: 'Title required' }; return; }
      const created = await graphPost(baseUrl, token, { fields });
      context.res = { status: 201, headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ id: created.id }) };
      return;
    }

    // PATCH: update existing item
    if (req.method === 'PATCH') {
      const itemId = req.query.id;
      if (!itemId) { context.res = { status: 400, body: 'id query param required for PATCH' }; return; }
      const body = typeof req.body === 'string' ? JSON.parse(req.body) : (req.body || {});
      const fields = buildFields(listKey, body);
      if (Object.keys(fields).length === 0) { context.res = { status: 400, body: 'No fields to update' }; return; }
      await graphPatch(baseUrl + '/' + itemId + '/fields', token, fields);
      context.res = { status: 200, headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ ok: true }) };
      return;
    }

    // DELETE: permanently remove an item (restricted to DELETE_EMAILS via guard above)
    if (req.method === 'DELETE') {
      const itemId = req.query.id;
      if (!itemId) { context.res = { status: 400, body: 'id query param required for DELETE' }; return; }
      await graphDelete(baseUrl + '/' + itemId, token);
      context.res = { status: 200, headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ ok: true, deleted: itemId }) };
      return;
    }

    context.res = { status: 405, body: 'Method not allowed' };
  } catch (err) {
    context.log.error('/api/reflists error:', err.message);
    context.res = { status: 500, body: 'Error: ' + err.message };
  }
};
