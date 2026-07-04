const https = require('https');

const TENANT_ID    = process.env.TENANT_ID;
const CLIENT_ID    = process.env.CLIENT_ID;
const CLIENT_SECRET= process.env.CLIENT_SECRET;
const SITE_PATH    = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';
const REMINDERS_LIST_GUID = 'a5005eb2-a503-4a87-a134-6e85619656d9';

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
        catch (e) { reject(new Error(`JSON parse: ${e.message}`)); }
      });
    });
    req.on('error', reject); req.write(body); req.end();
  });
}

function graphGet(url, token) {
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
        if (res.statusCode >= 400) { reject(new Error(`Graph GET ${res.statusCode}: ${data.slice(0,300)}`)); return; }
        try { resolve(JSON.parse(data)); } catch (e) { reject(new Error(`JSON parse: ${e.message}`)); }
      });
    });
    req.on('error', reject); req.end();
  });
}

function graphPost(url, token, body) {
  return new Promise((resolve, reject) => {
    const b = JSON.stringify(body);
    const u = new URL(url);
    const req = https.request({
      hostname: u.hostname, path: u.pathname + u.search, method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(b) },
    }, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        if (res.statusCode >= 400) { reject(new Error(`Graph POST ${res.statusCode}: ${data.slice(0,300)}`)); return; }
        try { resolve(JSON.parse(data)); } catch { resolve({}); }
      });
    });
    req.on('error', reject); req.write(b); req.end();
  });
}

function graphPatch(url, token, fields) {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify(fields);
    const u = new URL(url);
    const req = https.request({
      hostname: u.hostname, path: u.pathname + u.search, method: 'PATCH',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) },
    }, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        if (res.statusCode >= 400) { reject(new Error(`Graph PATCH ${res.statusCode}: ${data.slice(0,300)}`)); return; }
        try { resolve(data ? JSON.parse(data) : {}); } catch { resolve({}); }
      });
    });
    req.on('error', reject); req.write(body); req.end();
  });
}

module.exports = async function (context, req) {
  const method = req.method;

  try {
    const token = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);

    // ── GET: fetch all reminders, filter client-side by case ref ─────────
    if (method === 'GET') {
      const ref = (req.query.ref || '').trim();
      const allMode = req.query.all === '1';
      if (!ref && !allMode) { context.res = { status: 400, body: 'Missing ref (or all=1)' }; return; }

      const openOnly = req.query.open === '1';

      // Fetch all items for this case ref, filter client-side.
      // Completed_x003f_ boolean: Graph silently omits when false; returns true when completed.
      // Server-side $filter on this field requires SP index on Completed_x003f_.
      // Once indexed, re-enable server-side filter for performance.
      const items = [];
      let url = `https://graph.microsoft.com/v1.0/sites/${SITE_PATH}/lists/${REMINDERS_LIST_GUID}/items?$expand=fields&$top=999`;
      while (url) {
        const data = await graphGet(url, token);
        (data.value || []).forEach(item => {
          const f = item.fields || {};
          if (allMode || (f['Ourreference_x0028_text_x0029_'] || '').trim() === ref) items.push(item);
        });
        url = data['@odata.nextLink'] || null;
      }

      // Client-side completed filter.
      // Completed_x003f_ boolean is silently dropped by Graph for both true and false — unreliable.
      // CompletedMirror (plain text) is used instead: 'Yes' = completed, absent/other = open.
      const filtered = openOnly
        ? items.filter(item => (item.fields?.['CompletedMirror'] || '') !== 'Yes')
        : items;

      // Sort: incomplete by due date asc, completed last
      filtered.sort((a, b) => {
        const ac = a.fields?.['Completed_x003f_'] === true;
        const bc = b.fields?.['Completed_x003f_'] === true;
        if (ac !== bc) return ac ? 1 : -1;
        const ad = a.fields?.['Duedate'] || '';
        const bd = b.fields?.['Duedate'] || '';
        return ad.localeCompare(bd);
      });

      context.res = {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ value: filtered }),
      };
      return;
    }

    // ── POST: create reminder ─────────────────────────────────────────────
    if (method === 'POST') {
      let body;
      try { body = typeof req.body === 'string' ? JSON.parse(req.body) : req.body; }
      catch { context.res = { status: 400, body: 'Invalid JSON' }; return; }

      const fields = {
        Title:                             body.title || body.ref || '',
        Ourreference_x0028_text_x0029_:   body.ref   || '',
        Duedate:                           body.duedate || null,
        Event:                             body.event || null,
        Priority:                          body.priority || 'Medium',
        Brief:                             body.detail || null,
        'Completed_x003f_':                body.completed === true,
        CompletedMirror:                   body.completed === true ? 'Yes' : 'No',
        'Admin_x003f_':                    body.admin === true,
        'EmailReminder_x003f_':            body.emailReminder === true,
      };
      if (body.remindee) fields['Add_x0020_a_x0020_remindee_x0020'] = body.remindee;

      const createUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_PATH}/lists/${REMINDERS_LIST_GUID}/items`;
      const created = await graphPost(createUrl, token, { fields });
      context.res = {
        status: 201,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: created.id, fields: created.fields || {} }),
      };
      return;
    }

    // ── PATCH: update reminder ────────────────────────────────────────────
    if (method === 'PATCH') {
      let body;
      try { body = typeof req.body === 'string' ? JSON.parse(req.body) : req.body; }
      catch { context.res = { status: 400, body: 'Invalid JSON' }; return; }

      const { itemId, fields } = body;
      if (!itemId) { context.res = { status: 400, body: 'Missing itemId' }; return; }

      const patchUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_PATH}/lists/${REMINDERS_LIST_GUID}/items/${itemId}/fields`;
      await graphPatch(patchUrl, token, fields);
      context.res = { status: 200, body: 'OK' };
      return;
    }

    context.res = { status: 405, body: 'Method not allowed' };

  } catch (err) {
    context.log.error('Error in /api/reminders:', err.message);
    context.res = { status: 500, body: 'Internal error: ' + err.message };
  }
};
