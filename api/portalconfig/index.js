const https = require('https');
const { URL } = require('url');

// Generic key/value store on the PortalConfig list. That list has only two custom
// columns - Title (Text) and Value (Text) - confirmed from live Graph field dump
// 2026-07-15. One item per key; Value carries a JSON blob for anything richer than a
// single string. First consumer: the "CasesBanner" key (cases.html rolling news strip).
const SITE_PATH = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';
const CONFIG_LIST_GUID = '6661b8ba-3f10-436b-b4de-88b370e8160b';
const ALLOWED_DOMAIN = '@tmclegal.co.uk';
// Per-key write permission. Default is Toby+Danielle (the usual Management tier used
// everywhere else in this app - reflists, client firms, attachments delete). The banner
// is the one deliberate exception: Toby asked for a control "only I can see", so
// CasesBannerText/Active stay Toby-only.
const MANAGEMENT_EMAILS = ['toby@tmclegal.co.uk','danielle@tmclegal.co.uk'];
const TOBY_ONLY_KEYS = ['CasesBannerText','CasesBannerActive','AttachmentDocTypes'];
function writeEmailsFor(key){ return TOBY_ONLY_KEYS.includes(key) ? ['toby@tmclegal.co.uk'] : MANAGEMENT_EMAILS; }

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

// SharePoint "enhanced rich text" wraps the value in <div class="ExternalClass...">...</div>
// and uses <div>/<p>/<br> for line breaks. Undo that on read so consumers get clean text whether
// the Value column is plain or rich text — the banner keeps its line breaks (→ star separators)
// and JSON configs like AttachmentDocTypes parse correctly again.
function unwrapRichText(v) {
  if (v == null) return v;
  let s = String(v);
  s = s.replace(/<\/(div|p)>/gi, '\n').replace(/<br\s*\/?>/gi, '\n');
  s = s.replace(/<[^>]+>/g, '');
  s = s.replace(/&nbsp;/gi, ' ').replace(/&lt;/gi, '<').replace(/&gt;/gi, '>').replace(/&quot;/gi, '\"').replace(/&#39;/gi, "'").replace(/&apos;/gi, "'").replace(/&amp;/gi, '&');
  s = s.replace(/\n{3,}/g, '\n\n').replace(/^\n+|\n+$/g, '');
  return s;
}

module.exports = async function (context, req) {
  const callerEmail = getCallerEmail(req);
  if (!callerEmail || !callerEmail.endsWith(ALLOWED_DOMAIN)) {
    context.res = { status: 403, body: 'Forbidden' }; return;
  }
  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  const key = (req.query.key || '').trim();
  if (!key) { context.res = { status: 400, body: 'Missing key' }; return; }

  try {
    const token = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
    const baseUrl = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH + '/lists/' + CONFIG_LIST_GUID + '/items';

    // Small list, one row per key - bare $expand=fields and filter client-side rather
    // than maintaining a $filter (unindexed Title would 400 anyway).
    async function findByKey() {
      let url = baseUrl + '?$expand=fields&$top=999';
      while (url) {
        const data = await graphGet(url, token);
        const hit = (data.value || []).find(i => (i.fields && i.fields.Title) === key);
        if (hit) return hit;
        url = data['@odata.nextLink'] || null;
      }
      return null;
    }

    if (req.method === 'GET') {
      const item = await findByKey();
      context.res = {
        status: 200,
        headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-cache' },
        body: JSON.stringify({ key, value: item ? (unwrapRichText(item.fields.Value) || null) : null }),
      };
      return;
    }

    // POST: upsert. Restricted - deliberately just Toby, not the usual Management tier.
    if (req.method === 'POST') {
      if (!writeEmailsFor(key).includes(callerEmail)) {
        context.res = { status: 403, body: 'Forbidden — you do not have permission to update this setting.' };
        return;
      }
      const body = typeof req.body === 'string' ? JSON.parse(req.body) : (req.body || {});
      if (body.value === undefined) { context.res = { status: 400, body: 'Missing value' }; return; }
      const existing = await findByKey();
      if (existing) {
        await graphPatch(baseUrl + '/' + existing.id + '/fields', token, { Value: body.value });
        context.res = { status: 200, headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ id: existing.id, updated: true }) };
      } else {
        const created = await graphPost(baseUrl, token, { fields: { Title: key, Value: body.value } });
        context.res = { status: 201, headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ id: created.id, created: true }) };
      }
      return;
    }

    context.res = { status: 405, body: 'Method not allowed' };
  } catch (err) {
    context.log.error('/api/portalconfig error:', err.message);
    context.res = { status: 500, body: 'Error: ' + err.message };
  }
};
