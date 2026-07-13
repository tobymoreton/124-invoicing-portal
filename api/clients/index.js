const https = require('https');

const TENANT_ID     = process.env.TENANT_ID;
const CLIENT_ID     = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const SITE_PATH     = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';
const CLIENTS_GUID  = '901e8cbd-8760-4051-9eb0-0d5c0db1c06d';

const ADMIN_EMAILS  = ['toby@tmclegal.co.uk','danielle@tmclegal.co.uk','lesley@tmclegal.co.uk'];
const ALLOWED_DOMAIN = '@tmclegal.co.uk';
// S73: draftsmen can CREATE a client firm (they hit new instructing firms mid-case and
// must not be blocked). They CANNOT edit an existing one, and a create by a non-admin
// has the PRICING fields stripped — drafting fee basis/%, minimum fee, LAA fee, default
// hourly rate and billing email drive what TMC charges. Those stay with management.
const PRICING_FIELDS = ['feeBasis','feePercent','minFee','laaFee','hourlyRate','billingEmail'];

async function getToken(tenantId, clientId, clientSecret) {
  return new Promise((resolve, reject) => {
    const body = `grant_type=client_credentials&client_id=${encodeURIComponent(clientId)}&client_secret=${encodeURIComponent(clientSecret)}&scope=https%3A%2F%2Fgraph.microsoft.com%2F.default`;
    const req = https.request({
      hostname: 'login.microsoftonline.com',
      path: `/${tenantId}/oauth2/v2.0/token`,
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(body) },
    }, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try { const j = JSON.parse(data); j.access_token ? resolve(j.access_token) : reject(new Error(j.error_description || JSON.stringify(j))); }
        catch (e) { reject(new Error('JSON parse: ' + e.message)); }
      });
    });
    req.on('error', reject); req.write(body); req.end();
  });
}

function graphRequest(url, token, method, body) {
  return new Promise((resolve, reject) => {
    const b   = body ? JSON.stringify(body) : null;
    const u   = new URL(url);
    const req = https.request({
      hostname: u.hostname, path: u.pathname + u.search, method: method || 'GET',
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        Prefer: 'allowthrottleablequeries',
        ...(b ? { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(b) } : {}),
      },
    }, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        if (res.statusCode >= 400) { reject(new Error(`Graph ${method||'GET'} ${res.statusCode}: ${data.slice(0,300)}`)); return; }
        try { resolve(data ? JSON.parse(data) : {}); } catch (e) { resolve({}); }
      });
    });
    req.on('error', reject);
    if (b) req.write(b);
    req.end();
  });
}

module.exports = async function (context, req) {
  const method = req.method;

  // Auth check: all TMC staff can read and create; only admins can edit, and only
  // admins can set pricing on a create.
  const callerEmail = (req.headers['x-ms-client-principal-name'] || '').toLowerCase();
  const isAdmin = ADMIN_EMAILS.includes(callerEmail);
  const isStaff = callerEmail.endsWith(ALLOWED_DOMAIN);

  try {
    const token = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
    const baseUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_PATH}/lists/${CLIENTS_GUID}/items`;

    // ── GET: fetch all client firms ──────────────────────────────────────
    if (method === 'GET') {
      const items = [];
      let url = `${baseUrl}?$expand=fields&$top=999`;
      while (url) {
        const data = await graphRequest(url, token);
        (data.value || []).forEach(item => items.push({
          id: item.id,
          fields: item.fields || {},
        }));
        url = data['@odata.nextLink'] || null;
      }
      // Sort alphabetically by Title
      items.sort((a, b) => (a.fields.Title || '').localeCompare(b.fields.Title || ''));
      context.res = {
        status: 200,
        headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-cache' },
        body: JSON.stringify({ value: items }),
      };
      return;
    }

    // ── POST: create new client firm (any TMC user; pricing = admin only) ─
    if (method === 'POST') {
      if (!isStaff) { context.res = { status: 403, body: 'Forbidden' }; return; }
      let body;
      try { body = typeof req.body === 'string' ? JSON.parse(req.body) : req.body; }
      catch { context.res = { status: 400, body: 'Invalid JSON' }; return; }

      // Strip pricing from a non-admin create rather than rejecting the whole request:
      // the firm still gets added, it just carries no fee terms until management sets them.
      if (!isAdmin) PRICING_FIELDS.forEach(k => { delete body[k]; });

      const fields = buildFields(body);
      if (!fields.Title) { context.res = { status: 400, body: 'Firm name required' }; return; }

      const created = await graphRequest(`${baseUrl}`, token, 'POST', { fields });
      context.res = {
        status: 201,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: created.id, fields: created.fields || {} }),
      };
      return;
    }

    // ── PATCH: update client firm (admin only) ───────────────────────────
    if (method === 'PATCH') {
      if (!isAdmin) { context.res = { status: 403, body: 'Editing a client firm is restricted to Toby, Danielle and Lesley. You can still add a new firm.' }; return; }
      let body;
      try { body = typeof req.body === 'string' ? JSON.parse(req.body) : req.body; }
      catch { context.res = { status: 400, body: 'Invalid JSON' }; return; }

      const { itemId } = body;
      if (!itemId) { context.res = { status: 400, body: 'Missing itemId' }; return; }

      const fields = buildFields(body);
      await graphRequest(`${baseUrl}/${itemId}/fields`, token, 'PATCH', fields);
      context.res = { status: 200, body: 'OK' };
      return;
    }

    context.res = { status: 405, body: 'Method not allowed' };

  } catch (err) {
    context.log.error('Error in /api/clients:', err.message);
    context.res = { status: 500, body: 'Internal error: ' + err.message };
  }
};

function buildFields(body) {
  const f = {};
  if (body.name        !== undefined) f['Title']                          = body.name;
  if (body.address1    !== undefined) f['AddressLine1']                   = body.address1;
  if (body.address2    !== undefined) f['AddressLine2']                   = body.address2;
  if (body.address3    !== undefined) f['AddressLine3']                   = body.address3;
  if (body.address4    !== undefined) f['AddressLine4']                   = body.address4;
  if (body.address5    !== undefined) f['Address_x0020_Line_x0020_5']     = body.address5;
  if (body.website     !== undefined) f['Website']                        = body.website;
  if (body.vatNumber   !== undefined) f['VATNumber']                      = body.vatNumber;
  if (body.feeBasis    !== undefined) f['Draftingfeebasis']               = body.feeBasis;
  if (body.feePercent  !== undefined) f['Draftingfee_x0025_']             = parseFloat(body.feePercent) || null;
  if (body.minFee      !== undefined) f['Minimumdraftingfee_x0025_']      = parseFloat(body.minFee)     || null;
  if (body.laaFee      !== undefined) f['LAA_x0020_Drafting_x0020_fee_x00'] = body.laaFee;
  if (body.billingEmail !== undefined) f['Billingcontactemail']           = body.billingEmail;
  if (body.hourlyRate  !== undefined) f['Default_x0020_hourly_x0020_rate'] = body.hourlyRate;
  return f;
}
