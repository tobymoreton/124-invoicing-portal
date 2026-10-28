/**
 * P124 — Invoicing Portal
 * Azure Function: /api/caseask   (28/10/2026)
 *
 * POST { ref, question }  — plain-English question about one case's Case Actions.
 *
 * Design rule: Claude SELECTS the matching entries; this code does the arithmetic.
 * Claude never returns a total. The response carries the selected TT2 item ids, Claude's
 * one-line reading of the question, and count/hours computed here from the rows themselves,
 * so the page can filter the table to exactly what was counted.
 *
 * App settings: ANTHROPIC_API_KEY (required), ANTHROPIC_MODEL (optional, default below),
 *               TENANT_ID / CLIENT_ID / CLIENT_SECRET (existing).
 * The key is read from settings only. The repo is public — never put the key in code.
 */

const https   = require('https');
const { URL } = require('url');

const BUILD          = 'caseask-20261028-2';
const TT2_GUID       = '67db204c-30a5-4f4d-b276-60852d9967e1';
const SITE_PATH      = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';
const ALLOWED_DOMAIN = '@tmclegal.co.uk';
const DEFAULT_MODEL  = 'claude-sonnet-5-5';
const WD_MAX         = 600;   // characters of Work Done sent per row
const Q_MAX          = 500;   // characters of question accepted

function getCallerEmail(req) {
  try {
    const header = req.headers && req.headers['x-ms-client-principal'];
    if (!header) return null;
    const principal = JSON.parse(Buffer.from(header, 'base64').toString('utf8'));
    if (principal.userDetails) return principal.userDetails.toLowerCase();
    const claim = (principal.claims || []).find(c => ['preferred_username', 'email', 'upn',
      'http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress'].includes(c.typ));
    return claim ? claim.val.toLowerCase() : null;
  } catch { return null; }
}

function reply(context, status, obj) {
  context.res = {
    status,
    headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-cache', 'X-Api-Build': BUILD },
    body: JSON.stringify(obj),
  };
}

const ukDate = iso => { const d = new Date(iso); return isNaN(d) ? '' : d.toLocaleDateString('en-GB', { timeZone: 'Europe/London' }); };
const clean  = s => String(s || '').replace(/[\r\n\t|]+/g, ' ').replace(/\s{2,}/g, ' ').trim();

module.exports = async function (context, req) {
  const callerEmail = getCallerEmail(req);
  if (!callerEmail || !callerEmail.endsWith(ALLOWED_DOMAIN)) return reply(context, 403, { error: 'TMC Legal staff only.' });

  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET, ANTHROPIC_API_KEY, ANTHROPIC_MODEL } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) return reply(context, 500, { error: 'Missing Graph app settings.' });
  if (!ANTHROPIC_API_KEY) return reply(context, 500, { error: 'ANTHROPIC_API_KEY app setting is not set.' });

  const b = req.body || {};
  const ref = String(b.ref || '').trim();
  const question = String(b.question || '').trim().slice(0, Q_MAX);
  if (!ref || !question) return reply(context, 400, { error: 'ref and question are required.' });

  try {
    // 1. Load this case's actions (same filter as /api/caseactions GET ?ref=).
    const token = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
    const esc = ref.replace(/'/g, "''");
    let url = `https://graph.microsoft.com/v1.0/sites/${SITE_PATH}/lists/${TT2_GUID}/items`
            + `?$expand=fields&$top=999&$filter=fields/field_16 eq '${esc}'`;
    let rows = [];
    while (url) { const page = await graphGet(url, token); rows = rows.concat(page.value || []); url = page['@odata.nextLink'] || null; }
    if (!rows.length) return reply(context, 200, { ids: [], interpretation: '', note: 'No actions on this case.', count: 0, hours: 0, build: BUILD });

    // 2. Compact table for Claude. Ids are the TT2 item ids the page already holds.
    const byId = new Map();
    const lines = rows.map(r => {
      const f = r.fields || {};
      byId.set(String(r.id), f);
      const hrs = parseFloat(f['TimeSpentMirror'] || f['field_3']) || 0;
      const billable = f['Billable_x003f_'] === false ? 'N' : 'Y';
      const billed = f['Billed_x003f_'] === true ? 'Y' : 'N';
      const wd = clean(f['field_2']);
      return [r.id, ukDate(f['field_12'] || f['field_1']), clean(f['Completedby_x0028_text_x0029_'] || f['field_18']),
              clean(f['field_9']), hrs, billable, billed, wd.length > WD_MAX ? wd.slice(0, WD_MAX) + '…' : wd].join(' | ');
    });

    const system = 'You help staff at a legal costs firm query the time ledger ("Case Actions") for ONE case. '
      + 'Each row is: id | date work done (dd/mm/yyyy) | fee earner | brief (work type) | hours | billable Y/N | billed Y/N | work done narrative. '
      + 'Select every row that answers the question. You MUST answer by calling the select_rows tool exactly once; never answer in plain text. Read narratives for meaning, not just keywords '
      + '(e.g. "preparing for the detailed assessment" covers preparation for the hearing, points of dispute, replies, bundles, skeletons where the context shows it). '
      + 'Never compute totals; the system adds up the rows you select. If the question is ambiguous, choose the most reasonable reading and state it. '
      + 'If no rows match, return an empty list and say why in the note. Use only the rows given; do not invent anything.';

    const user = `Question: ${question}\n\nRows (${lines.length}):\n${lines.join('\n')}`;

    const ai = await anthropic(ANTHROPIC_API_KEY, {
      model: ANTHROPIC_MODEL || DEFAULT_MODEL,
      max_tokens: 4000,
      system,
      tools: [{
        name: 'select_rows',
        description: 'Return the ids of the rows that answer the question.',
        input_schema: {
          type: 'object',
          properties: {
            interpretation: { type: 'string', description: 'One sentence: how you read the question, i.e. what was counted.' },
            ids: { type: 'array', items: { type: 'string' }, description: 'Ids of every matching row.' },
            note: { type: 'string', description: 'Optional: borderline rows included/excluded, or why nothing matched. One or two sentences.' },
          },
          required: ['interpretation', 'ids'],
        },
      }],
      tool_choice: { type: 'auto' }, // b2: 'tool'/'any' rejected by this model (400) — system prompt requires the call instead
      messages: [{ role: 'user', content: user }],
    });

    const use = (ai.content || []).find(c => c.type === 'tool_use' && c.name === 'select_rows');
    if (!use) {
      const txt = (ai.content || []).filter(c => c.type === 'text').map(c => c.text).join(' ').slice(0, 500);
      return reply(context, 502, { error: 'No selection returned by the model.', modelText: txt, stop: ai.stop_reason, build: BUILD });
    }

    // 3. Keep only ids that really are on this case; do the arithmetic here.
    const asked = Array.isArray(use.input.ids) ? use.input.ids.map(String) : [];
    const ids = [...new Set(asked)].filter(id => byId.has(id));
    const dropped = asked.length - ids.length;
    const hours = ids.reduce((s, id) => { const f = byId.get(id); return s + (parseFloat(f['TimeSpentMirror'] || f['field_3']) || 0); }, 0);

    context.log(`AUDIT caseask ref=${ref} by=${callerEmail} rows=${rows.length} selected=${ids.length}`);
    return reply(context, 200, {
      ids, count: ids.length, hours: Math.round(hours * 100) / 100,
      interpretation: String(use.input.interpretation || ''), note: String(use.input.note || ''),
      ...(dropped ? { unknownIdsDropped: dropped } : {}),
      rowsSearched: rows.length, build: BUILD,
    });
  } catch (err) {
    context.log.error('caseask error:', err.message);
    return reply(context, 500, { error: err.message, build: BUILD });
  }
};

function anthropic(key, payload) {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify(payload);
    const req = https.request({
      hostname: 'api.anthropic.com', path: '/v1/messages', method: 'POST',
      headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01', 'content-type': 'application/json', 'content-length': Buffer.byteLength(body) },
    }, res => {
      let data = ''; res.on('data', c => data += c);
      res.on('end', () => {
        if (res.statusCode >= 400) { reject(new Error(`Anthropic ${res.statusCode}: ${data.slice(0, 300)}`)); return; }
        try { resolve(JSON.parse(data)); } catch (e) { reject(new Error('Anthropic JSON parse: ' + e.message)); }
      });
    });
    req.setTimeout(40000, () => req.destroy(new Error('Anthropic timeout (40s)')));
    req.on('error', reject);
    req.write(body); req.end();
  });
}

function getToken(tenantId, clientId, clientSecret) {
  return new Promise((resolve, reject) => {
    const body = new URLSearchParams({ grant_type: 'client_credentials', client_id: clientId, client_secret: clientSecret, scope: 'https://graph.microsoft.com/.default' }).toString();
    const req = https.request({
      hostname: 'login.microsoftonline.com', path: `/${tenantId}/oauth2/v2.0/token`, method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(body) },
    }, res => {
      let data = ''; res.on('data', c => data += c);
      res.on('end', () => {
        try { const j = JSON.parse(data); j.access_token ? resolve(j.access_token) : reject(new Error('Token error: ' + (j.error_description || data))); }
        catch (e) { reject(e); }
      });
    });
    req.setTimeout(20000, () => req.destroy(new Error('Token request timeout (20s)')));
    req.on('error', reject);
    req.write(body); req.end();
  });
}

function graphGet(url, token) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    const req = https.request({
      hostname: u.hostname, path: u.pathname + u.search, method: 'GET',
      headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
    }, res => {
      let data = ''; res.on('data', c => data += c);
      res.on('end', () => {
        if (res.statusCode >= 400) { reject(new Error(`Graph GET ${res.statusCode}: ${data.slice(0, 300)}`)); return; }
        try { resolve(JSON.parse(data)); } catch (e) { reject(new Error('JSON parse: ' + e.message)); }
      });
    });
    req.setTimeout(20000, () => req.destroy(new Error('Graph GET timeout (20s)')));
    req.on('error', reject);
    req.end();
  });
}
