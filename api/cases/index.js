/**
 * P124 — Invoicing Portal
 * Azure Function: /api/cases
 *
 * Searches the SP Cases list via Graph (app-only client-credentials auth).
 * Admin-only — returns 403 for non-admins.
 *
 * Query params:
 *   q       — free-text search string (matches case name or our ref, client-side)
 *   top     — max results (default 10, max 20)
 *   preload — if '1', returns all non-closed cases with no top cap (used for client-side cache)
 *
 * List: Cases
 * GUID: ae420bda-e550-499c-b337-90e4f33617c1
 */

const https   = require('https');
const { URL } = require('url');

const LIST_GUID = 'ae420bda-e550-499c-b337-90e4f33617c1';
const SITE_PATH = 'tmcostings.sharepoint.com:/sites/TMCLegalLimited:';
const CLIENT_FIRMS_GUID = '901e8cbd-8760-4051-9eb0-0d5c0db1c06d';

const ADMIN_EMAILS = [
  'toby@tmclegal.co.uk',
  'danielle@tmclegal.co.uk',
];

// All authenticated users can search cases (draftsmen need this to create drafts)
const FINANCE_EMAILS = [
  'lesley@tmclegal.co.uk',
];

const SELECT_FIELDS = [
  'Title',
  'Ourreference_x0028_text_x0029_',
  'Net_x0020_Drafting_x0020_Fee',
  'Drafting_x0020_fee_x0020_line',
  'VatOnDraftingFee_x003f_',
  'TimedWorkLineOverride',
  'Firm_x0028_text_x0029_',
  'FirmLookupId',
  'Address1_x0028_text_x0029_',
  'Address2_x0028_text_x0029_',
  'Address3_x0028_text_x0029_',
  'Address4_x0028_text_x0029_',
  'Address5_x0028_text_x0029_',
  'ClientCaseReference',
  'DraftingYetToBeInvoiced',
  'caseID_text',
  'Fee_x0025__x0028_number_x0029_',
  'LAA_x0020_Drafting_x0020_Fee_x00',
  'Minfee_x0025__x0028_number_x0029',
  'RateForCase',
  'ProvisionalDraftingFee',
  'ProvisionalMinDraftingFee',
  'LegalAidOnlyProfitCosts',
  'TimedWorkBillableFromOverride',
  'StatusMirror',
  'ProfitCostsClaimed_x0028_Ex_x002',
  'GrossProfitCostsRecoveredAf',
  'VAT_x0020__x0025__x0020_Claimed',
  'Drafting_x0020_fee_x0020_basis',
  'InterPartesorLegalAid',
  'MinimumFee',
  // Full case name fields
  'fullCaseNameMirror',
  'Fullcasenameoveride',
  'CaseNameForOpenCases',
  'OurPartyFirstName',
  'OurPartySurname',
  'OpponentPartyName',
  'Morethanonedefendant_x003f_',
  'Andothers_x003f_',
  'MultiClaimants',
  'MultiDefendants',
  // Settlement & offers
  'Settlementamount',
  'PayingPartysLastOffer',
  'FigureForSettlementSheet',
  'DateSettled0',
  'Offerincludescostsofsssessment_x',
  'Offerincludesinterest_x003f_',
  'BottomLine',
  'LikelyTopEnd',
  'ProfitCostsatAdvisedRates',
  // Costs claimed
  'DraftingTimeClaimed',
  'CounselsFeesClaimed',
  'DisbursementsClaimed',
  'VATonDisbursements',
  'Costsofassessment',
  'OtherTMCInvoices_ex_x0020_bill',
  'NonVatableProfitCosts',
  'OtherTMCPCVatable',
  'OtherTMCFeesVatable',
  'NonVatableCounselFees',
  'NonVatableLAAPC',
  'ClientPaidExpenses',
  'BillDraftedByText',
  'Other_x0020_TMC_x0020_PC_x0020__',
  'TMC_x0020_drafting_x0020_time_x0',
  // Recovery
  'Counselsfeespayable',
  'NetProfitCostsRecoveredBeforeDra',
  'Limit_x0020_Costs_x0020_of_x0020',
  'PreLimitedCostsOfAssessment',
  // Interest — plain stored fields only (all calculated fields return null via Graph)
  'SumForInterestCalculationOverrid',
  'TotalPreAuthorityPayments',
  'DateofAuthoritytoAssess',
  'InterimPayment1',
  'InterimPayment2',
  'InterimPayment3',
  'InterimPaymentDate1',
  'InterimPaymentDate2',
  'InterimPaymentDate3',
  'CostsOfAssessmentForInterestTabO',
  'TotalInterestToDate_Text',
  'DailyInterestRate_Text',
  // Booking In tab
  'DateReceived',
  'Turnaround_x0020__x0001f3af_',
  'FeeEarner_x0028_fromFM_x0029_',
  'Work_x0020_In_x0020_Progress',
  'Add_x0020_To_x0020_Title',
  'Notes',
  'draftingNotes',
  'Acknowledgmentemail',
  'AcknowledgmentSent',
  'PrimaryFundingType',
  'CaseTrack',
  'Task',
  'Casetype',
  // Opponent & Court tab
  'Opponentreference',
  'opponentReference2',
  'opponentFirmMirror',
  'OpponentFirm2Mirror',
  'ServiceEmail',
  'CourtName_Text',
  'Division',
  'ClaimNo',
  'Litigant_x0020_in_x0020_Person_x',
  'Opponent_x0020_Address_x0020_1',
  'Opponent_x0020_Address_x0020_2',
  'Opponent_x0020_Address_x0020_3',
  'Opponent_x0020_Address_x0020_4',
  'Opponent_x0020_Address_x0020_5',
  // Case Overview tab
  'assignedToTextValue',
  'AssignedToMirror',
  'coAssignedMirror',
  'Draftsmanemail',
  'Current_x0020_Position',
  'CP_x0020_Last_x0020_Updated',
  'Flags',
  'LastAction',
  'Last_x0020_Reminder',
  'DateServedInformally0',
  'DatePart8Issued',
  'DateServedFormally0',
  'PODS_x0020_Due',
  'PODS_x0020_Received',
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
  context.log('P124 /api/cases called');

  // All authenticated users can search cases
  const callerEmail = getCallerEmail(req);
  if (!callerEmail) {
    context.res = { status: 403, body: 'Forbidden — you must be signed in.' };
    return;
  }

  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    context.res = { status: 500, body: 'Missing required app settings.' };
    return;
  }

  const q          = (req.query.q || '').trim().toLowerCase();
  const top        = Math.min(parseInt(req.query.top, 10) || 10, 20);
  const includeAll = req.query.all === '1';
  const preload    = req.query.preload === '1';

  // Preload mode: return all non-closed cases for client-side cache — no q required
  if (preload) {
    try {
      const token  = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
      const result = await getAllOpenCases(token);
      context.res = {
        status: 200,
        headers: { 'Content-Type': 'application/json', 'Cache-Control': 'private, max-age=300' },
        body: JSON.stringify(result),
      };
    } catch (err) {
      context.log.error('Error preloading cases:', err.message);
      context.res = { status: 500, body: 'Error: ' + err.message };
    }
    return;
  }

  if (!q || q.length < 2) {
    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-cache' },
      body: JSON.stringify({ value: [] }),
    };
    return;
  }

  try {
    const token  = await getToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
    const result = await searchCases(token, q, top, includeAll);

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-cache' },
      body: JSON.stringify(result),
    };
  } catch (err) {
    context.log.error('Error searching cases:', err.message);
    context.res = { status: 500, body: 'Error: ' + err.message };
  }
};

async function getAllOpenCases(token) {
  var base = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH + '/lists/' + LIST_GUID + '/items'
           + '?$expand=fields($select=' + encodeURIComponent(SELECT_FIELDS) + ')&$top=500';

  var url = base;
  var all = [];
  while (url) {
    var page = await graphGet(url, token);
    all = all.concat(page.value || []);
    url = page['@odata.nextLink'] || null;
  }

  const open = all.filter(function(item) {
    var status = ((item.fields || {}).StatusMirror || '').toLowerCase();
    return status !== 'closed';
  });

  await resolveMissingFirms(token, open);

  return {
    value: open.map(function(item) {
      return { id: item.id, fields: item.fields || {} };
    }),
  };
}

async function searchCases(token, q, top, includeAll) {
  var base = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH + '/lists/' + LIST_GUID + '/items'
           + '?$expand=fields($select=' + encodeURIComponent(SELECT_FIELDS) + ')&$top=500';

  var url = base;
  var all = [];
  while (url) {
    var page = await graphGet(url, token);
    all = all.concat(page.value || []);
    url = page['@odata.nextLink'] || null;
  }

  const matches = all.filter(function(item) {
    var f      = item.fields || {};
    var title  = (f.Title || '').toLowerCase();
    var ref    = (f['Ourreference_x0028_text_x0029_'] || '').toLowerCase();
    var status = (f.StatusMirror || '').toLowerCase();
    // Exclude closed cases unless caller requests all (e.g. direct ref lookup from case.html)
    if (!includeAll && status === 'closed') return false;
    return title.indexOf(q) !== -1 || ref.indexOf(q) !== -1;
  });

  var sliced = matches.slice(0, top);
  await resolveMissingFirms(token, sliced);

  return {
    value: sliced.map(function(item) {
      return { id: item.id, fields: item.fields || {} };
    }),
  };
}

// Resolve blank Firm text mirrors directly from the Firm Lookup id via the Client Firms list.
// Infowise-created cases do not get Firm_x0028_text_x0029_ populated until the item is re-saved
// (the PA mirror flow fires on save). This fills the firm name at read time from the raw Firm
// Lookup id (FirmLookupId), removing that re-save dependency for the firm name specifically.
// Safe by construction: only fills where the mirror is blank; never overwrites an existing value.
// Silent no-op if FirmLookupId is absent or the Client Firms fetch fails (degrades to blank firm).
async function resolveMissingFirms(token, items) {
  try {
    var needed = false;
    (items || []).forEach(function(item) {
      var f = item.fields || {};
      var mirror = (f['Firm_x0028_text_x0029_'] || '').trim();
      if (!mirror && f.FirmLookupId) needed = true;
    });
    if (!needed) return;

    var url = 'https://graph.microsoft.com/v1.0/sites/' + SITE_PATH + '/lists/' + CLIENT_FIRMS_GUID + '/items'
            + '?$expand=fields($select=Title)&$top=999';
    var byId = {};
    while (url) {
      var page = await graphGet(url, token);
      (page.value || []).forEach(function(fi) {
        byId[String(fi.id)] = ((fi.fields || {}).Title || '').trim();
      });
      url = page['@odata.nextLink'] || null;
    }

    (items || []).forEach(function(item) {
      var f = item.fields || {};
      var mirror = (f['Firm_x0028_text_x0029_'] || '').trim();
      if (!mirror && f.FirmLookupId) {
        var name = byId[String(f.FirmLookupId)];
        if (name) f['Firm_x0028_text_x0029_'] = name;
      }
    });
  } catch (e) {
    // Silent no-op - keep the request succeeding with a blank firm rather than erroring.
  }
}

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
      headers: {
        'Content-Type':   'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(body),
      },
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
    var options = {
      hostname: u.hostname,
      path:     u.pathname + u.search,
      method:   'GET',
      headers: {
        Authorization: 'Bearer ' + token,
        Accept:        'application/json',
      },
    };

    var req = https.request(options, function(res) {
      var data = '';
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
