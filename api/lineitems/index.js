const https  = require('https');
const { URL } = require('url');

const LIST_GUID = '496468a5-e2ed-48db-8826-58cb08844eee';

const SELECT_FIELDS = [
  'id','field_1','field_2','field_3','ProRataApportionment',
  'CompletedByEmail','ValueMirror','InvoiceIDRef','CaseName',
  'field_5','Completed_x0020_on','BillableYorN_x0020__x2753_',
  'InvoiceType','InvoiceDate',
].join(',');

module.exports = async function (context, req) {
  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET, SITE_ID } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET || !SITE_ID) {
    context.res = { status: 500, body: 'Missing required app settings.' };
    return;
  }
  try {
    const token     = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
    const lineItems = await fetchAll(token, SITE_ID);
    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-cache' },
      body: JSON.stringify(lineItems),
    };
  } catch (err) {
    context.log.error('lineitems error:', err.message);
    context.res = { status: 500, body: `Error: ${err.message}` };
  }
};

function getToken(tenantId, clientId, clientSecret) {
  return new Promise((resolve, reject) => {
    const body = new URLSearchParams({
      grant_type: 'client_credentials',
      client_id: clientId,
      client_secret: clientSecret,
      scope: 'https://graph.microsoft.com/.default',
    }).toString();
    const options = {
      hostname: 'login.microsoftonline.com',
      path: `/${tenantId}/oauth2/v2.0/token`,
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(body) },
    };
    const req = https.request(options, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try {
          const json = JSON.parse(data);
          if (json.access_token) resolve(json.access_token);
          else reject(new Error(`Token error: ${json.error_description || data}`));
        } catch (e) { reject(e); }
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

async function fetchAll(token, siteId) {
  const base = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${LIST_GUID}/items` +
               `?$expand=fields($select=${SELECT_FIELDS})&$top=500`;
  let url = base, all = [];
  while (url) {
    const page = await graphGet(url, token);
    all = all.concat((page.value || []).map(normalise));
    url = page['@odata.nextLink'] || null;
  }
  return all.filter(i => i.Billable === true);
}

function graphGet(url, token) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    const options = {
      hostname: u.hostname,
      path: u.pathname + u.search,
      method: 'GET',
      headers: { Authorization: `Bearer ${token}`, Accept: 'application/json', ConsistencyLevel: 'eventual' },
    };
    const req = https.request(options, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        if (res.statusCode >= 400) { reject(new Error(`Graph ${res.statusCode}: ${data.slice(0,200)}`)); return; }
        try { resolve(JSON.parse(data)); } catch (e) { reject(e); }
      });
    });
    req.on('error', reject);
    req.end();
  });
}

function normalise(item) {
  const f = item.fields || {};
  return {
    _id:             String(item.id),
    WorkDone:        f.field_1                        || null,
    TimeSpent:       toNum(f.field_2),
    Rate:            toNum(f.field_3),
    ProRata:         toNum(f.ProRataApportionment),
    CompletedByEmail: f.CompletedByEmail              || null,
    Value:           toNum(f.ValueMirror),
    InvoiceIDRef:    f.InvoiceIDRef                   || null,
    CaseName:        f.CaseName                       || null,
    OurRef:          f.field_5                        || null,
    CompletedOn:     f.Completed_x0020_on             || null,
    Billable:        f.BillableYorN_x0020__x2753_     === true,
    InvoiceType:     f.InvoiceType                    || null,
    InvoiceDate:     f.InvoiceDate                    || null,
  };
}

function toNum(v) {
  const n = parseFloat(v);
  return isNaN(n) ? null : n;
}
