const https  = require('https');
const { URL } = require('url');

const LIST_GUID = '5c366b19-0da9-4be9-b68f-60e6a0209cdb';
const SITE_PATH = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';

const SELECT_FIELDS = [
  'id','OrderDetails','VendorName','Casename','Ourref',
  'InvoiceDate','DueDate','AmountDue','AmountOutstanding',
  'Status','Invoicetype','LAorIP',
  'PaymentAmount1','PaymentAmount2','PaymentAmount3',
  'PaymentDate1','PaymentDate2','PaymentDate3','Case_x0020_ID',
].join(',');

module.exports = async function (context, req) {
  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    context.res = { status: 500, body: 'Missing required app settings.' };
    return;
  }
  try {
    const token    = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
    const invoices = await fetchAllInvoices(token);
    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-cache' },
      body: JSON.stringify(invoices),
    };
  } catch (err) {
    context.log.error('Error:', err.message);
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

async function fetchAllInvoices(token) {
  const base = `https://graph.microsoft.com/v1.0/sites/${SITE_PATH}/lists/${LIST_GUID}/items` +
               `?$expand=fields&$top=500`;
  let url = base, all = [];
  while (url) {
    const page = await graphGet(url, token);
    all = all.concat((page.value || []).map(normalise));
    url = page['@odata.nextLink'] || null;
  }
  all.sort((a, b) => {
    const da = a.DueDate ? new Date(a.DueDate).getTime() : Infinity;
    const db = b.DueDate ? new Date(b.DueDate).getTime() : Infinity;
    return da - db;
  });
  return all;
}

function graphGet(url, token) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    const options = {
      hostname: u.hostname,
      path: u.pathname + u.search,
      method: 'GET',
      headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
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
    _id: String(item.id),
    OrderDetails: f.OrderDetails || null,
    VendorName: f.VendorName || null,
    Casename: f.Casename || null,
    Ourref: f.Ourref || null,
    InvoiceDate: f.InvoiceDate || null,
    DueDate: f.DueDate || null,
    AmountDue: toNum(f.AmountDue),
    AmountOutstanding: toNum(f.AmountOutstanding),
    Status: f.Status || null,
    Invoicetype: f.Invoicetype || null,
    LAorIP: f.LAorIP || null,
    PaymentAmount1: toNum(f.PaymentAmount1),
    PaymentAmount2: toNum(f.PaymentAmount2),
    PaymentAmount3: toNum(f.PaymentAmount3),
    PaymentDate1: f.PaymentDate1 || null,
    PaymentDate2: f.PaymentDate2 || null,
    PaymentDate3: f.PaymentDate3 || null,
    Case_x0020_ID: f.Case_x0020_ID || null,
  };
}

function toNum(v) { const n = parseFloat(v); return isNaN(n) ? null : n; }
