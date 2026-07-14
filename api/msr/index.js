/**
 * P124 — Case Portal
 * Azure Function: /api/msr  (Matter Status Report)
 *
 * Reads the SP Cases list LIVE via Graph (app-only client-credentials auth) and returns the
 * matters that belong on a client-facing Matter Status Report for one instructing firm.
 *
 * THIS FUNCTION DELIBERATELY HAS NO CACHE AND NO FILE EXPORT.
 * It replaces an offline pipeline (PA999.2 -> ~2,000 .md files on OneDrive -> preprocessor ->
 * Excel) which was found on 2026-07-13 to be serving PHANTOM MATTERS. Root cause: that pipeline
 * named each file `/{Firm}/{OurReference}.md` — a path built from two MUTABLE fields — and only
 * ever deleted the file at the CURRENT path. Change a case's reference or firm and the old file
 * is orphaned on disk forever, and keeps being read back as a live matter. `1741120.md`
 * (Haywood) sat on the HJA report for four months as a matter that does not exist.
 * A view that reads SharePoint directly cannot have that class of bug. Do not reintroduce an
 * intermediate cache or file export here.
 *
 * Query params:
 *   firm — instructing firm name, matched against Firm_x0028_text_x0029_.
 *          Defaults to 'Hodge Jones & Allen LLP'.
 *
 * Access: all authenticated @tmclegal.co.uk staff.
 *
 * List: Cases
 * GUID: ae420bda-e550-499c-b337-90e4f33617c1
 */

const https   = require('https');
const { URL } = require('url');

const LIST_GUID   = 'ae420bda-e550-499c-b337-90e4f33617c1';
const SITE_PATH   = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';
const DEFAULT_FIRM = 'Hodge Jones & Allen LLP';

// Fields this report needs. Kept separate from api/cases SELECT_FIELDS on purpose so the main
// preload payload stays lean.
//
// TotalIPCostsMirror is the AUTHORITATIVE bill total. It is a plain Number column mirrored from
// the SP CALCULATED column Calc_TotalInterPartesCosts by PA124.9 (ongoing) / PA124.10 (backfill).
// Graph app-only returns NULL for SP Calculated columns, so Calc_TotalInterPartesCosts itself is
// unreadable here. NEVER recompute the bill total from its components — the mirror or nothing.
//
// FeeEarner is a SP LOOKUP (confirmed: it returns as SPListExpandedReference with FeeEarner#Id in
// the SP connector payload). Graph app-only returns null for Lookups, so the FM text mirror is
// selected alongside it and the two are coalesced on read.
const SELECT_FIELDS = [
  'Title',
  'Ourreference_x0028_text_x0029_',
  'Firm_x0028_text_x0029_',
  'FeeEarner',
  'FeeEarner_x0028_fromFM_x0029_',
  'FeeEarnerTextValue',
  'ClientCaseReference',
  'assignedToTextValue',
  'AssignedToMirror',
  'DateServedInformally0',
  'DateServedFormally0',
  'TotalIPCostsMirror',
  'CounselsFeesClaimed',
  'DisbursementsClaimed',
  'InterimPayment1',
  'InterimPayment2',
  'InterimPayment3',
  'PayingPartysLastOffer',
  'Current_x0020_Position',
  'LastAction',
  // S76 — sign-off and exclusion. Plain text (email / ISO date) + one Yes/No.
  'MSRVerifiedBy',
  'MSRVerifiedUTC',
  'ExcludeFromMSR',
  // filtering only — not displayed
  'Status',
  'StatusMirror',
  'InterPartesorLegalAid',
  'DateClosed0',
  'DateSettled0',
].join(',');

function getCallerEmail(req) {
  try {
    const header = req.headers && req.headers['x-ms-client-principal'];
    if (!header) return null;
    const decoded = Buffer.from(header, 'base64').toString('utf8');
    const principal = JSON.parse(decoded);
    if (principal.userDetails) return principal.userDetails.toLowerCase();
    const claim = (principal.claims || []).find(
      c => c.typ === 'preferred_username' || c.typ === 'email' || c.typ === 'upn'
        || c.typ === 'http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress'
    );
    return claim ? claim.val.toLowerCase() : null;
  } catch { return null; }
}

module.exports = async function (context, req) {
  context.log('P124 /api/msr called');

  const callerEmail = getCallerEmail(req);
  if (!callerEmail || !callerEmail.endsWith('@tmclegal.co.uk')) {
    context.res = { status: 403, body: 'Forbidden — you must be signed in.' };
    return;
  }

  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    context.res = { status: 500, body: 'Missing required app settings.' };
    return;
  }

  const firm = (req.query.firm || DEFAULT_FIRM).trim();

  try {
    const token  = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
    const all    = await getAllCases(token);
    const shaped = all.map(item => shape(item));
    const onReport = shaped.filter(r => qualifies(r, firm));

    // S76 — excluded matters are NOT deleted from the report, they are moved to a separate,
    // collapsed list on the page with an Undo. A matter that silently disappears from a client
    // report is the exact bug this whole report exists to prevent; it just points the other way.
    const rows     = onReport.filter(r => !r.excluded).sort(sortByDateServed);
    const excluded = onReport.filter(r =>  r.excluded).sort(sortByDateServed);

    // Firm picker options, derived from the same fetch — no second call, and it can only ever
    // offer firms that actually exist in the list.
    const firms = Array.from(new Set(shaped.map(r => r.firm).filter(Boolean))).sort();

    // Draftsman filter options — only people who actually have a matter ON THIS REPORT, so the
    // dropdown can never offer a name that yields an empty table.
    const assignees = Array.from(new Set(rows.map(r => r.assignee).filter(Boolean))).sort();

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: JSON.stringify({
        firm,
        firms,
        assignees,
        generated: new Date().toISOString(),
        count: rows.length,
        value: rows,
        excluded,
        totals: totalsOf(rows),
      }),
    };
  } catch (err) {
    context.log.error('Error building MSR:', err.message);
    context.res = { status: 500, body: 'Error: ' + err.message };
  }
};

// ---------------------------------------------------------------------------
// FILTER — the complete rule. Enforced HERE, in code, on the server.
//
// It lives in code and not in the page (and emphatically not in a prompt) because a filter that
// lives anywhere else gets silently re-decided every time someone runs the report.
//
// NOTE ON THE 'FILEMAKER MIGRATION COHORT' — the build prompt called for a TEMPORARY extra rule
// excluding rows with DateClosed0 in November 2025, on the basis that the FileMaker import left
// ~1,611 historic matters carrying Status = 'Pending assignment' (the default, never set by a
// human), which would therefore stream onto a client-facing report.
//
// THAT PREMISE IS FALSE. Measured directly against SharePoint, 2026-07-13 (PA124.11, report-only):
//
//     Cases list total ........................................... 2,034
//     rows with DateClosed0 in Nov 2025 .......................... 1,635
//     ... of those, rows with Status = 'Pending assignment' ....       0
//     distinct statuses on those 1,635 rows ...... 'Closed' and
//                                                  'In negotiation (served formally)' — nothing else
//     rows with DateClosed0 in Nov 2025 and NOT Closed ..........       2
//
// The cohort is ALREADY Status = 'Closed', so rule 2 below removes it without any help. The
// exclusion was dead code where it was meant to matter (HJA: 29 rows with it, 29 rows without it)
// and ACTIVELY HARMFUL everywhere else — the only two rows it removed were live Bhatt Murphy
// matters, both Inter Partes, both in negotiation, worth 46,423.50 between them (items 68605 and
// 68636). Hiding a live matter from a client report is the same class of bug as the phantom
// matters this report was built to eliminate, just pointing the other way.
//
// So it is gone, and DateClosed0 plays NO part in this filter. Toby's rule, 2026-07-13: "Date
// Closed is kinda irrelevant in every scenario. Status = Closed is the only definer." Do not
// reintroduce a DateClosed0 test here without measuring against SharePoint first.
// ---------------------------------------------------------------------------
function qualifies(r, firm) {
  // 1. Instructing firm
  if ((r.firm || '').trim().toLowerCase() !== firm.toLowerCase()) return false;

  const status = (r.status || '').toLowerCase();

  // 2/3/4. Closed, Settled and Budget matters are all off the report.
  //        Budgets and costs schedules are not bills and do not belong on it. (The old report
  //        included budgets shaded red for review; that is RETIRED — they are excluded outright.)
  if (status.includes('closed'))  return false;
  if (status.includes('settled')) return false;
  if (status.includes('budget'))  return false;

  // 5. Legal Aid only.
  //
  //    *** USE InterPartesorLegalAid. DO NOT USE PrimaryFundingType. ***
  //
  //    They are different fields and the distinction is load-bearing. A legally-aided CLIENT can
  //    still have a genuine inter partes costs RECOVERY against the paying party. Those matters
  //    carry PrimaryFundingType = 'Legal Aid' but InterPartesorLegalAid = 'Inter Partes', and they
  //    DO belong on the report. Measured 2026-07-13: 7 HJA matters had PrimaryFundingType =
  //    'Legal Aid' but only 2 had InterPartesorLegalAid = 'Legal Aid'. Filtering on funding type
  //    would have silently dropped 5 live inter partes bills off a client report — no error, no
  //    warning, no way to notice.
  if ((r.funding || '').trim().toLowerCase() === 'legal aid') return false;

  // 6. Settled matters (belt and braces alongside the Status test).
  if (r.dateSettled) return false;

  return true;
}

// ---------------------------------------------------------------------------
// Shaping
// ---------------------------------------------------------------------------
function shape(item) {
  const f = item.fields || {};
  const interim = num(f.InterimPayment1) + num(f.InterimPayment2) + num(f.InterimPayment3);

  return {
    id:            item.id,
    ourRef:        str(f['Ourreference_x0028_text_x0029_']),  // identification only — NOT displayed
    title:         str(f.Title),
    firm:          str(f['Firm_x0028_text_x0029_']),
    // FeeEarner is a Lookup -> null via Graph app-only. Fall back to the text mirrors.
    feeEarner:     str(choiceVal(f.FeeEarner))
                     || str(f['FeeEarner_x0028_fromFM_x0029_'])
                     || str(f.FeeEarnerTextValue),
    clientRef:     str(f.ClientCaseReference),
    // AssignedTo is a Person field -> null via Graph app-only. The mirrors are the only route.
    // Frequently blank. Show as unassigned; never hide the row for it.
    assignedTo:    str(f.AssignedToMirror) || str(f.assignedToTextValue),
    assignee:      canonicalAssignee(str(f.AssignedToMirror) || str(f.assignedToTextValue)),
    // Informal service if present, otherwise formal.
    dateServed:    str(f.DateServedInformally0) || str(f.DateServedFormally0),
    // AUTHORITATIVE bill total. Mirror of the SP calculated column. Never recomputed.
    totalIPCosts:  num(f.TotalIPCostsMirror),
    counselFees:   num(f.CounselsFeesClaimed),
    disbursements: num(f.DisbursementsClaimed),
    // Held individually as well as summed: the report column is a TOTAL, but there is no single SP
    // field behind it, so an inline edit has to write back to the three real columns.
    interim1:      num(f.InterimPayment1),
    interim2:      num(f.InterimPayment2),
    interim3:      num(f.InterimPayment3),
    interimTotal:  interim,
    lastOffer:     num(f.PayingPartysLastOffer),
    // Which of the two service-date columns the displayed date actually came from, so an inline
    // edit writes back to the right one instead of guessing. Informal wins when populated.
    dateServedField: str(f.DateServedInformally0) ? 'DateServedInformally0' : 'DateServedFormally0',
    currentPosition: stripHtml(f['Current_x0020_Position']),
    lastAction:      stripHtml(f.LastAction),
    // S76 — draftsman sign-off. Plain text, written by the portal only. A stamp, not a lock:
    // it records who said the row was right and when. It does not stop anyone editing after.
    verifiedBy:  str(f.MSRVerifiedBy),
    verifiedUTC: str(f.MSRVerifiedUTC),
    // S76 — removed from the report by Management. Yes/No.
    excluded:    bool(f.ExcludeFromMSR),
    // filter-only
    status:      str(choiceVal(f.StatusMirror) || choiceVal(f.Status)),
    funding:     str(choiceVal(f.InterPartesorLegalAid)),
    dateClosed:  str(f.DateClosed0),
    dateSettled: str(f.DateSettled0),
  };
}

// ---------------------------------------------------------------------------
// Assignee normalisation
//
// assignedToTextValue / AssignedToMirror are PA-maintained mirrors of a Person column and the SAME
// person is stored three different ways across the list: email (tom@tmclegal.co.uk), display name
// (Tom Winyard) and first name only (Tom). Normalise on READ — do NOT clean the data (S66: it is a
// PA-maintained mirror, the flow would write it straight back, and other consumers are unknown).
//
// Matching is by EXACT token, never substring: 'daniel' is a substring of 'danielle' and Daniel is
// a draftsman while Danielle is the hand-back bucket. A substring match would merge the two.
const ROSTER = {
  toby: 'Toby', tom: 'Tom', tracy: 'Tracy', joanna: 'Joanna', kelly: 'Kelly',
  julie: 'Julie', daniel: 'Daniel', danielle: 'Danielle', lesley: 'Lesley', david: 'David',
};

function canonicalAssignee(v) {
  const s = str(v).toLowerCase();
  if (!s) return '';
  // email -> local part; otherwise the first word of a display name, or the bare first name
  const token = s.indexOf('@') !== -1 ? s.split('@')[0] : s.split(/\s+/)[0];
  return ROSTER[token] || '';   // unknown token -> blank rather than a guess
}

// SP Choice columns come back from Graph as a plain string, but be tolerant of the
// {Value: ...} shape in case a column is ever swapped to a Lookup.
function choiceVal(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') return v.Value || v.value || '';
  return v;
}

function str(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

function num(v) {
  const n = parseFloat(v);
  return isFinite(n) ? n : 0;
}

// SP Yes/No. Graph returns a real boolean, but tolerate the string forms SharePoint and
// Power Automate have both been seen to emit.
function bool(v) {
  if (v === true) return true;
  if (typeof v === 'string') {
    const s = v.trim().toLowerCase();
    return s === 'true' || s === 'yes' || s === '1';
  }
  return v === 1;
}

// ---------------------------------------------------------------------------
// Rich text
//
// Current_x0020_Position and LastAction are SharePoint rich-text columns. They arrive wrapped in
// markup, e.g. <div class="ExternalClassC6F3..."><p class="editor-paragraph">...</p></div>.
// Strip it here, once, on the server — so display and export cannot diverge and neither can ever
// show raw HTML.
// ---------------------------------------------------------------------------
function stripHtml(v) {
  if (!v) return '';
  let s = String(v);
  s = s.replace(/<br\s*\/?>/gi, '\n');
  s = s.replace(/<\/(p|div|li|tr|h[1-6])>/gi, '\n');
  s = s.replace(/<[^>]+>/g, '');
  s = s.replace(/&nbsp;/gi, ' ')
       .replace(/&amp;/gi, '&')      // leaks through visibly if missed
       .replace(/&lt;/gi, '<')
       .replace(/&gt;/gi, '>')
       .replace(/&quot;/gi, '"')
       .replace(/&#39;/g, "'")
       .replace(/&pound;/gi, '\u00a3');
  s = s.replace(/\r/g, '');
  s = s.replace(/\n{3,}/g, '\n\n');
  return s.trim();
}

// Date Served ascending, earliest first. Matters with no Date Served (billing / pre-service
// stage) group at the BOTTOM rather than sorting to the top as empty strings would.
function sortByDateServed(a, b) {
  const A = a.dateServed, B = b.dateServed;
  if (!A && !B) return (a.title || '').localeCompare(b.title || '');
  if (!A) return 1;
  if (!B) return -1;
  return A < B ? -1 : A > B ? 1 : 0;
}

function totalsOf(rows) {
  const t = { totalIPCosts: 0, counselFees: 0, disbursements: 0, interimTotal: 0, lastOffer: 0 };
  rows.forEach(r => {
    t.totalIPCosts  += r.totalIPCosts;
    t.counselFees   += r.counselFees;
    t.disbursements += r.disbursements;
    t.interimTotal  += r.interimTotal;
    t.lastOffer     += r.lastOffer;
  });
  return t;
}

// ---------------------------------------------------------------------------
// Graph
//
// ⚠ NO $select ON fields. This was a deliberate change in S76 and must not be reverted to a
//   $select for tidiness. SharePoint Yes/No columns are SILENTLY DROPPED from the Graph response
//   when they are named in $expand=fields($select=...) — they come back absent even when true
//   (proven on Billable_x003f_ / Billed_x003f_ / Completed_x003f_). ExcludeFromMSR is a Yes/No
//   column, so a $select would return it as undefined on every row and every excluded matter
//   would quietly reappear on the client report. Fetch the whole fields bag and shape it here.
//   SELECT_FIELDS above is retained as the documented field list for this report.
// ---------------------------------------------------------------------------
async function getAllCases(token) {
  let url = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH + '/lists/' + LIST_GUID + '/items'
          + '?$expand=fields&$top=500';
  let all = [];
  while (url) {
    const page = await graphGet(url, token);
    all = all.concat(page.value || []);
    url = page['@odata.nextLink'] || null;
  }
  return all;
}

function getToken(tenantId, clientId, clientSecret) {
  return new Promise(function(resolve, reject) {
    const body = new URLSearchParams({
      grant_type:    'client_credentials',
      client_id:     clientId,
      client_secret: clientSecret,
      scope:         'https://graph.microsoft.com/.default',
    }).toString();

    const options = {
      hostname: 'login.microsoftonline.com',
      path:     '/' + tenantId + '/oauth2/v2.0/token',
      method:   'POST',
      headers: {
        'Content-Type':   'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(body),
      },
    };

    const req = https.request(options, function(res) {
      let data = '';
      res.on('data', function(chunk) { data += chunk; });
      res.on('end', function() {
        try {
          const json = JSON.parse(data);
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
    const u = new URL(url);
    const options = {
      hostname: u.hostname,
      path:     u.pathname + u.search,
      method:   'GET',
      headers: {
        Authorization: 'Bearer ' + token,
        Accept:        'application/json',
      },
    };

    const req = https.request(options, function(res) {
      let data = '';
      res.on('data', function(chunk) { data += chunk; });
      res.on('end', function() {
        if (res.statusCode >= 400) {
          reject(new Error('Graph ' + res.statusCode + ': ' + data.slice(0, 300)));
          return;
        }
        try { resolve(JSON.parse(data)); }
        catch(e) { reject(new Error('JSON parse error: ' + e.message)); }
      });
    });
    req.on('error', reject);
    req.end();
  });
}
