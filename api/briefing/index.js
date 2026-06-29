const https = require('https');

const TENANT_ID     = process.env.TENANT_ID;
const CLIENT_ID     = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const SITE_PATH     = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';
const USERS_GUID    = '79d09d87-bc95-40fe-b924-138886faea05';

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

function graphRequest(url, token) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    const req = https.request({
      hostname: u.hostname, path: u.pathname + u.search, method: 'GET',
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        Prefer: 'allowthrottleablequeries',
      },
    }, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        if (res.statusCode >= 400) { reject(new Error(`Graph GET ${res.statusCode}: ${data.slice(0, 300)}`)); return; }
        try { resolve(data ? JSON.parse(data) : {}); } catch (e) { resolve({}); }
      });
    });
    req.on('error', reject);
    req.end();
  });
}

module.exports = async function (context, req) {
  if (req.method !== 'GET') {
    context.res = { status: 405, body: 'Method not allowed' };
    return;
  }

  const callerEmail = (req.headers['x-ms-client-principal-name'] || '').toLowerCase();
  if (!callerEmail) {
    context.res = { status: 401, body: 'Unauthorised' };
    return;
  }

  try {
    const token = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);

    // Fetch all users (tiny list — never more than ~15 rows) and filter client-side
    // field_2 = Email; Daily_x0020_Briefing_x0020__x261 = Daily Briefing ☕
    const url = `https://graph.microsoft.com/v1.0/sites/${SITE_PATH}/lists/${USERS_GUID}/items?$expand=fields&$top=50`;
    const data = await graphRequest(url, token);
    const items = data.value || [];

    const userRow = items.find(item =>
      (item.fields?.['field_2'] || '').toLowerCase() === callerEmail
    );

    const briefing = userRow?.fields?.['Daily_x0020_Briefing_x0020__x261'] || '';

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-cache' },
      body: JSON.stringify({ briefing }),
    };
  } catch (err) {
    context.log.error('Error in /api/briefing:', err.message);
    context.res = { status: 500, body: 'Internal error: ' + err.message };
  }
};
