/**
 * P124 — Invoicing Portal
 * Azure Function: /api/billfigures
 *
 * Writes bill figures parsed from an eBill (browser-side, in case.html) onto a
 * Cases list item. Replaces the route: Claude extraction → JSON emailed to
 * automation@ → PA745 → SharePoint.
 *
 * It exists as its own Function rather than reusing /api/caseupdate because the
 * three defects found in PA745 all came from writing an unvalidated bag of keys:
 *   1. PA745's Parse_JSON held a sample object, not a schema, so a missing key
 *      resolved to null and Update_item wrote it — BLANKING a figure that was
 *      already correct on the case. Here, an unknown key is rejected outright and
 *      a null is only accepted on the three fields where blank is meaningful.
 *   2. PA745 triggered on any mail to a shared mailbox with a matching subject,
 *      with no sender check. Here the caller must be a named member of staff.
 *   3. PA745's Get_items had no $top and took first(), so on two cases sharing a
 *      text ref the figures landed on whichever SharePoint returned first. Here
 *      the item is fetched by id and its own reference is compared to the ref the
 *      page sent; a mismatch is refused.
 *
 * POST body (JSON):
 *   ref            string — case reference as shown on the page (?ref=)
 *   itemId         string — Cases list item id
 *   sourceFilename string — the eBill filename, for the provenance stamp
 *   fields         object — the twelve figures below, numbers or null
 *
 * Returns: { updated: true, stamp: "<file> | <who> | <utc>" }
 *
 * List: Cases  GUID: ae420bda-e550-499c-b337-90e4f33617c1
 *
 * ⚠️ PREREQUISITE: the Cases list must carry `BillFiguresImport` as a
 * SINGLE LINE OF TEXT column. It is PATCHed in the same call as the figures, so
 * if it does not exist Graph 400s and the whole write fails — deliberately, so a
 * set of figures can never land with no record of where it came from.
 * Never Yes/No: Graph drops Yes/No columns from $expand=fields($select=...)
 * (S76 ExcludeFromMSR, S85 Bespoke, and the OtherTMCPCVatable conversion).
 */

const https   = require('https');
const { URL } = require('url');

const SITE_PATH  = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';
const LIST_GUID  = 'ae420bda-e550-499c-b337-90e4f33617c1';
const REF_FIELD  = 'Ourreference_x0028_text_x0029_';
const STAMP_FIELD = 'BillFiguresImport';
const API_BUILD  = 'S129-billfigures-v1';

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

// The only keys this endpoint will write. Anything else in the body is refused
// rather than ignored — an endpoint that silently drops what it does not
// recognise cannot be relied on to have written what you sent it.
//
// nullable: blank carries meaning and must survive.
//   NonVatableLAAPC / LegalAidOnlyProfitCosts — blank means no legal aid figure
//     on this bill. The LAA row is present in every Precedent S template and
//     carries zero on an inter partes matter, so a zero would be stamped over
//     whatever a human had typed.
//   OtherTMCPCVatable — blank means no bill has been imported yet, and must NOT
//     read as "No". Text column holding 'Yes' / 'No' (normalised 2026-09-14).
const NUMERIC_FIELDS = {
  'ProfitCostsClaimed_x0028_Ex_x002': { nullable: false },
  'DraftingTimeClaimed':              { nullable: false },
  'CounselsFeesClaimed':              { nullable: false },
  'DisbursementsClaimed':             { nullable: false },
  'VATonDisbursements':               { nullable: false },
  'Other_x0020_TMC_x0020_PC_x0020__': { nullable: false },
  'NonVatableProfitCosts':            { nullable: false },
  'NonVatableCounselFees':            { nullable: false },
  'NonVatableTMCDraftingTime':        { nullable: false },
  'NonVatableLAAPC':                  { nullable: true  },
  'LegalAidOnlyProfitCosts':          { nullable: true  },
};
const YESNO_FIELD = 'OtherTMCPCVatable';

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

module.exports = async function (context, req) {
  context.log('P124 /api/billfigures called');

  const headers = { 'Content-Type': 'application/json', 'X-Api-Build': API_BUILD };
  const bad = (status, body) => { context.res = { status, headers, body: JSON.stringify({ error: body }) }; };

  const callerEmail = getCallerEmail(req);
  if (!callerEmail || !ALLOWED_EMAILS.includes(callerEmail)) {
    return bad(403, 'Forbidden — not authorised to import bill figures.');
  }

  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) return bad(500, 'Missing required app settings.');

  let body;
  try { body = typeof req.body === 'string' ? JSON.parse(req.body) : req.body; }
  catch (e) { return bad(400, 'Invalid JSON body.'); }

  const ref            = body && String(body.ref || '').trim();
  const itemId         = body && String(body.itemId || '').trim();
  const sourceFilename = body && String(body.sourceFilename || '').trim();
  const fields         = body && body.fields;

  if (!ref)            return bad(400, 'Missing ref.');
  if (!itemId)         return bad(400, 'Missing itemId.');
  if (!sourceFilename) return bad(400, 'Missing sourceFilename — figures are never written without a provenance stamp.');
  if (!fields || typeof fields !== 'object') return bad(400, 'Missing fields.');

  // ── Validate every key before anything is sent to Graph. A partially valid
  //    payload is refused whole: half a bill is worse than none.
  const patch = {};
  const problems = [];

  for (const key of Object.keys(fields)) {
    if (!NUMERIC_FIELDS[key] && key !== YESNO_FIELD) problems.push('Unknown field: ' + key);
  }
  for (const [key, rule] of Object.entries(NUMERIC_FIELDS)) {
    if (!(key in fields)) { problems.push('Missing field: ' + key); continue; }
    const v = fields[key];
    if (v === null) {
      if (!rule.nullable) problems.push(key + ' is null, but blank is not a meaningful value for it.');
      else patch[key] = null;
      continue;
    }
    if (typeof v !== 'number' || !isFinite(v)) { problems.push(key + ' is not a number.'); continue; }
    if (v < 0) { problems.push(key + ' is negative (' + v + ').'); continue; }
    patch[key] = Math.round((v + Number.EPSILON) * 100) / 100;
  }
  if (!(YESNO_FIELD in fields)) {
    problems.push('Missing field: ' + YESNO_FIELD);
  } else {
    const yv = fields[YESNO_FIELD];
    if (yv === null) patch[YESNO_FIELD] = null;
    else if (yv === 'Yes' || yv === 'No') patch[YESNO_FIELD] = yv;
    else problems.push(YESNO_FIELD + " must be 'Yes', 'No' or null — got " + JSON.stringify(yv) + '.');
  }
  if (problems.length) return bad(400, 'Payload refused, nothing written: ' + problems.join(' · '));

  try {
    const token = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
    const itemUrl = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH
                  + '/lists/' + LIST_GUID + '/items/' + encodeURIComponent(itemId);

    // ── The ref on the page must be the ref on the item. PA745 resolved the case
    //    by filtering on the text ref and taking first(), so duplicate refs sent
    //    figures to whichever row SharePoint happened to return.
    const existing = await graphGet(itemUrl + '?$expand=fields($select=' + REF_FIELD + ')', token);
    const storedRef = String((existing.fields || {})[REF_FIELD] || '').trim();
    if (storedRef !== ref) {
      return bad(409, 'Refusing to write: item ' + itemId + ' carries reference "' + storedRef
                    + '" but the page sent "' + ref + '".');
    }

    const stamp = sourceFilename.slice(0, 120) + ' | ' + callerEmail + ' | ' + new Date().toISOString();
    patch[STAMP_FIELD] = stamp.slice(0, 250);

    await graphPatch(itemUrl + '/fields', token, patch);

    context.log('AUDIT billfigures ref=' + ref + ' itemId=' + itemId + ' by=' + callerEmail + ' file=' + sourceFilename);
    context.res = { status: 200, headers, body: JSON.stringify({ updated: true, stamp }) };
  } catch (err) {
    context.log.error('Error in /api/billfigures:', err.message);
    context.res = { status: 500, headers, body: JSON.stringify({ error: 'Error: ' + err.message }) };
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
    const req = https.request({
      hostname: 'login.microsoftonline.com',
      path:     '/' + tenantId + '/oauth2/v2.0/token',
      method:   'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(body) },
    }, function (res) {
      let data = '';
      res.on('data', c => { data += c; });
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

function graphGet(url, token) {
  return new Promise(function (resolve, reject) {
    const u = new URL(url);
    const req = https.request({
      hostname: u.hostname, path: u.pathname + u.search, method: 'GET',
      headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' },
    }, function (res) {
      let data = '';
      res.on('data', c => { data += c; });
      res.on('end', () => {
        if (res.statusCode >= 400) { reject(new Error('Graph GET ' + res.statusCode + ': ' + data.slice(0, 300))); return; }
        try { resolve(data ? JSON.parse(data) : {}); } catch (e) { resolve({}); }
      });
    });
    req.on('error', reject);
    req.end();
  });
}

function graphPatch(url, token, body) {
  return new Promise(function (resolve, reject) {
    const payload = JSON.stringify(body);
    const u = new URL(url);
    const req = https.request({
      hostname: u.hostname, path: u.pathname + u.search, method: 'PATCH',
      headers: {
        Authorization:    'Bearer ' + token,
        'Content-Type':   'application/json',
        'Content-Length': Buffer.byteLength(payload),
      },
    }, function (res) {
      let data = '';
      res.on('data', c => { data += c; });
      res.on('end', () => {
        if (res.statusCode >= 400) { reject(new Error('Graph PATCH ' + res.statusCode + ': ' + data.slice(0, 300))); return; }
        try { resolve(data ? JSON.parse(data) : {}); } catch (e) { resolve({}); }
      });
    });
    req.on('error', reject);
    req.write(payload);
    req.end();
  });
}
