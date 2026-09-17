/**
 * P124 — Invoicing Portal
 * Azure Function: /api/msgpreview   (S131, 2026-09-16)
 *
 * Why: staff upload instruction emails to caseAttachments as Outlook .msg files. SharePoint has no
 * viewer for .msg — the file's webUrl just downloads it into a blank tab, which is what Lesley
 * reported as "unable to open instructions" (16/09/2026). The .html acknowledgments render fine.
 *
 * What: reads the .msg out of the caseAttachments library via Graph (app-only, same auth as
 * /api/attachments), parses it server-side with @kenjiuno/msgreader (pure JS — the API's FIRST
 * npm dependency; SWA's build runs `npm install` in /api) and returns the headers, plain-text
 * body and the list of files attached INSIDE the email, so case.html can show it in a modal.
 *
 * Routes (GET, signed-in @tmclegal.co.uk only):
 *   ?id=<driveItemId>            → JSON { subject, from, to, cc, date, body, attachments[], ... }
 *   ?id=<driveItemId>&att=<n>    → the n-th attachment inside the email, as a download
 *
 * Guards: the item must live in the caseAttachments drive, be a .msg, and be under MAX_BYTES.
 * Read-only — nothing is written anywhere.
 */

const https = require('https');
const { URL } = require('url');
const MsgReader = require('@kenjiuno/msgreader').default;

const API_BUILD      = 'msgpreview-v1-20260916';
const LIST_GUID      = '710dea64-11ae-4ae7-8fde-d4508206e1c1';   // caseAttachments
const SITE_PATH      = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';
const ALLOWED_DOMAIN = '@tmclegal.co.uk';
const MAX_BYTES      = 25 * 1024 * 1024;

function getCallerEmail(req) {
  try {
    const header = req.headers && req.headers['x-ms-client-principal'];
    if (!header) return null;
    const principal = JSON.parse(Buffer.from(header, 'base64').toString('utf8'));
    if (principal.userDetails) return principal.userDetails.toLowerCase();
    const claim = (principal.claims || []).find(
      c => c.typ === 'preferred_username' || c.typ === 'email' || c.typ === 'upn'
        || c.typ === 'http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress'
    );
    return claim ? claim.val.toLowerCase() : null;
  } catch { return null; }
}

function smtpOnly(s) { return (s && s.indexOf('@') !== -1 && s.indexOf('/') === -1) ? s : ''; }

// Pure: .msg bytes + Graph item → the JSON the modal renders. Kept separate so it can be tested
// offline against a real .msg without Graph.
function buildPreview(bytes, item) {
  const reader = new MsgReader(bytes);
  const d = reader.getFileData() || {};
  if (d.error) throw new Error('msg parse: ' + d.error);
  const recips = (d.recipients || []).map(r => ({
    name:  r.name || '',
    email: smtpOnly(r.smtpAddress) || smtpOnly(r.email),
    type:  (r.recipType || '').toLowerCase(),
  }));
  return {
    id:          item.id,
    name:        item.name,
    webUrl:      item.webUrl || '',
    subject:     d.subject || d.normalizedSubject || '',
    from:        { name: d.senderName || '', email: smtpOnly(d.senderSmtpAddress) || smtpOnly(d.sentRepresentingSmtpAddress) || smtpOnly(d.senderEmail) },
    to:          recips.filter(r => r.type === 'to'),
    cc:          recips.filter(r => r.type === 'cc'),
    date:        d.messageDeliveryTime || d.clientSubmitTime || d.creationTime || '',
    body:        d.body || '',
    bodyMissing: !d.body && !!d.compressedRtf,   // RTF/HTML-only body — nothing plain to show
    attachments: (d.attachments || []).map((a, i) => ({
      i, name: a.fileName || a.name || ('attachment-' + i), size: a.contentLength || 0,
      mime: a.attachMimeTag || '', embeddedMessage: !!a.innerMsgContent,
    })),
    build:       API_BUILD,
  };
}

function extractAttachment(bytes, n) {
  const reader = new MsgReader(bytes);
  const d = reader.getFileData() || {};
  const list = d.attachments || [];
  if (n < 0 || n >= list.length) return { status: 404, text: 'No such attachment.' };
  const a = reader.getAttachment(list[n]);
  const content = a && a.content;
  if (!content || !(content instanceof Uint8Array)) return { status: 415, text: 'That attachment is an embedded message — download the .msg to open it in Outlook.' };
  const fname = String(a.fileName || list[n].fileName || ('attachment-' + n)).replace(/[\r\n"]/g, '');
  return { status: 200, fname, mime: list[n].attachMimeTag || 'application/octet-stream', buf: Buffer.from(content) };
}

module.exports = async function (context, req) {
  context.log('P124 /api/msgpreview called');

  const callerEmail = getCallerEmail(req);
  if (!callerEmail || callerEmail.indexOf(ALLOWED_DOMAIN) === -1) {
    context.res = { status: 403, headers: { 'X-Api-Build': API_BUILD }, body: 'Forbidden — you must be signed in with a TMC account.' };
    return;
  }
  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    context.res = { status: 500, headers: { 'X-Api-Build': API_BUILD }, body: 'Missing required app settings.' };
    return;
  }

  const id  = (req.query && req.query.id  || '').trim();
  const att = (req.query && req.query.att !== undefined && req.query.att !== '') ? parseInt(req.query.att, 10) : null;
  if (!id) {
    context.res = { status: 400, headers: { 'X-Api-Build': API_BUILD }, body: 'Missing required query param: id (driveItem id).' };
    return;
  }

  try {
    const token   = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
    const dResp   = await graphGet('https://graph.microsoft.com/v1.0/sites/' + SITE_PATH + '/lists/' + LIST_GUID + '/drive?$select=id', token);
    const driveId = dResp && dResp.id;
    if (!driveId) throw new Error('Could not resolve the caseAttachments drive id.');

    const itemBase = 'https://graph.microsoft.com/v1.0/drives/' + driveId + '/items/' + encodeURIComponent(id);
    const item = await graphGet(itemBase + '?$select=id,name,size,file,webUrl', token);
    if (!item || !item.file) { context.res = { status: 404, headers: { 'X-Api-Build': API_BUILD }, body: 'Not a file in caseAttachments.' }; return; }
    if (!/\.msg$/i.test(item.name || '')) { context.res = { status: 415, headers: { 'X-Api-Build': API_BUILD }, body: 'Only Outlook .msg files can be previewed.' }; return; }
    if ((item.size || 0) > MAX_BYTES) { context.res = { status: 413, headers: { 'X-Api-Build': API_BUILD }, body: 'Message is over 25 MB — download it instead.' }; return; }

    const bytes = await graphGetBinary(itemBase + '/content', token);

    if (att !== null && !isNaN(att)) {
      const r = extractAttachment(bytes, att);
      if (r.status !== 200) { context.res = { status: r.status, headers: { 'X-Api-Build': API_BUILD }, body: r.text }; return; }
      context.log('AUDIT msgpreview ATT id=' + id + ' att=' + att + ' name=' + r.fname + ' by=' + callerEmail);
      context.res = {
        status: 200,
        headers: {
          'Content-Type': r.mime,
          'Content-Disposition': 'attachment; filename="' + r.fname.replace(/[^\x20-\x7e]/g, '_') + '"; filename*=UTF-8\'\'' + encodeURIComponent(r.fname),
          'Cache-Control': 'no-cache',
          'X-Api-Build': API_BUILD,
        },
        isRaw: true,
        body: r.buf,
      };
      return;
    }

    const out = buildPreview(bytes, item);
    context.log('AUDIT msgpreview VIEW id=' + id + ' name=' + item.name + ' by=' + callerEmail);
    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-cache', 'X-Api-Build': API_BUILD },
      body: JSON.stringify(out),
    };
  } catch (err) {
    context.log.error('Error previewing msg:', err.message);
    context.res = { status: 500, headers: { 'X-Api-Build': API_BUILD }, body: 'Error: ' + err.message };
  }
};
module.exports.buildPreview = buildPreview;
module.exports.extractAttachment = extractAttachment;

function getToken(tenantId, clientId, clientSecret) {
  return new Promise(function(resolve, reject) {
    var body = new URLSearchParams({
      grant_type:    'client_credentials',
      client_id:     clientId,
      client_secret: clientSecret,
      scope:         'https://graph.microsoft.com/.default',
    }).toString();
    var options = {
      hostname: 'login.microsoftonline.com',
      path:     '/' + tenantId + '/oauth2/v2.0/token',
      method:   'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(body) },
    };
    var req = https.request(options, function(res) {
      var data = '';
      res.on('data', function(chunk) { data += chunk; });
      res.on('end', function() {
        try {
          var json = JSON.parse(data);
          if (json.access_token) resolve(json.access_token);
          else reject(new Error('Token error: ' + (json.error_description || data)));
        } catch(e) { reject(e); }
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

function graphGet(url, token) {
  return new Promise(function(resolve, reject) {
    var u = new URL(url);
    var req = https.request({ hostname: u.hostname, path: u.pathname + u.search, method: 'GET',
      headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' } }, function(res) {
      var data = '';
      res.on('data', function(chunk) { data += chunk; });
      res.on('end', function() {
        if (res.statusCode >= 400) { reject(new Error('Graph ' + res.statusCode + ': ' + data.slice(0, 300))); return; }
        try { resolve(JSON.parse(data)); } catch(e) { reject(new Error('JSON parse error: ' + e.message)); }
      });
    });
    req.on('error', reject);
    req.end();
  });
}

// /content answers 302 to a pre-authenticated download URL on a different host — follow it WITHOUT
// the bearer (the URL carries its own auth; a bearer on a foreign host is a leak for no gain).
function graphGetBinary(url, token, hops) {
  hops = hops || 0;
  return new Promise(function(resolve, reject) {
    var u = new URL(url);
    var headers = { Accept: '*/*' };
    if (u.hostname === 'graph.microsoft.com') headers.Authorization = 'Bearer ' + token;
    var req = https.request({ hostname: u.hostname, path: u.pathname + u.search, method: 'GET', headers: headers }, function(res) {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        res.resume();
        if (hops >= 3) { reject(new Error('Too many redirects fetching file content.')); return; }
        resolve(graphGetBinary(res.headers.location, token, hops + 1));
        return;
      }
      var chunks = [];
      res.on('data', function(c) { chunks.push(c); });
      res.on('end', function() {
        var buf = Buffer.concat(chunks);
        if (res.statusCode >= 400) { reject(new Error('Graph content ' + res.statusCode + ': ' + buf.toString('utf8').slice(0, 300))); return; }
        if (buf.length > MAX_BYTES) { reject(new Error('File exceeds the 25 MB preview limit.')); return; }
        resolve(buf);
      });
    });
    req.on('error', reject);
    req.end();
  });
}
