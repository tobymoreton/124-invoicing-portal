/**
 * invoice-generator.js
 * P124 Invoicing Portal — Invoice HTML/PDF generator
 *
 * generateInvoiceHTML(data) -> HTML string (A4, print-ready)
 *
 * data shape:
 *   invoiceNumber    string       'DRAFT' at draft stage; real number at issue
 *   invoiceDate      Date|string|null  null = draft (shows 'DRAFT' instead of date)
 *   dueDate          Date|string|null  optional (defaults to invoiceDate + 30 days)
 *   yourRef          string       ClientCaseReference
 *   ourRef           string       Ourreference_x0028_text_x0029_
 *   caseName         string
 *   firmName         string
 *   address1-5       string
 *   draftingFee      number       IP drafting fee (0 if none)
 *   draftingFeeLine  string       e.g. 'Preparing Draft Bill of Costs | 5.5% of ...'
 *   laFee            number       LA-only drafting fee (0 / omit if none)
 *   laFeeLine        string       e.g. 'Preparing Draft Bill of Costs (LA) | 6% of ...'
 *   timedFee         number       0 if no timed element
 *   timedHours       number
 *   timedRate        number       ignored when timedWorkLine is supplied
 *   timedWorkLine    string|null  override text (already includes hours — don't append again)
 *   bespokeAmt        number       bespoke invoice amount (0 if none); VAT at 20% applied
 *   bespokeInvoiceLine string      bespoke line description
 *   expenses         number       default 0
 *   vatOnDrafting    boolean      if true, VAT applies to draftingFee AND laFee
 *   logoDataUrl      string|null  base64 data URL for logo (inline); falls back to filename
 *   scheduleLines    Array|null   [{date, workDone, hours, rate, amount}] — appended as Schedule of Work
 */

const LOGO_FILENAME = 'New Logo Gold, Silver and Navy.png';
const FIRM_ADDR = 'Fennels Lodge, St Peters Close, Loudwater, Buckinghamshire HP11 1JT.';
const FIRM_CO   = 'Company No: 12348782';
const FIRM_VAT  = 'VAT Registration No: 398 1109 74';
const BACS_SORT = '20-03-84';
const BACS_ACCT = '33605868';

// Navy/gold brand colours (used in schedule)
const BRAND_NAVY = '#1B2A4A';
const BRAND_GOLD = '#C9A84C';

function _esc(s) {
  if (!s) return '';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function _fmtDate(d) {
  if (!d) return '';
  const dt = (d instanceof Date) ? d : new Date(d);
  if (isNaN(dt.getTime())) return '';
  return dt.toLocaleDateString('en-GB', { day:'2-digit', month:'2-digit', year:'numeric' });
}

function _fmtGBP(n) {
  const v = parseFloat(n) || 0;
  return '\u00a3' + v.toLocaleString('en-GB', { minimumFractionDigits:2, maximumFractionDigits:2 });
}

function _addDays(date, n) {
  const d = new Date(date); d.setDate(d.getDate() + n); return d;
}

function _addrLines(data) {
  return [data.firmName, data.address1, data.address2, data.address3, data.address4, data.address5]
    .filter(Boolean).map(l => _esc(l) + '<br>').join('');
}

const INVOICE_CSS = [
  '* { margin:0; padding:0; box-sizing:border-box; }',
  'body { font-family:\'Calibri\',\'Segoe UI\',Arial,sans-serif; font-size:11pt; color:#000; background:#fff; }',
  '.page { width:210mm; min-height:297mm; margin:0 auto; padding:18mm 20mm 20mm 20mm; display:flex; flex-direction:column; page-break-after:always; }',
  '.header { display:flex; justify-content:flex-end; margin-bottom:12mm; }',
  '.logo { width:52mm; height:auto; }',
  '.invoice-title { font-size:22pt; font-weight:bold; margin-bottom:7mm; }',
  '.draft-banner { background:#FEF3C7; border:1px solid #F59E0B; border-radius:4px; padding:6px 14px; font-size:9pt; font-weight:bold; color:#92400E; letter-spacing:.06em; text-align:center; margin-bottom:6mm; }',
  '.meta-table { width:100%; border-collapse:collapse; margin-bottom:10mm; }',
  '.meta-table td { padding:1.2mm 0; font-size:10.5pt; vertical-align:top; }',
  '.meta-table td:first-child { width:45mm; }',
  '.meta-table td:last-child { text-align:right; }',
  '.meta-spacer { height:3mm; }',
  '.case-name { text-align:center; font-weight:bold; text-decoration:underline; font-size:11pt; margin-bottom:7mm; }',
  '.invoice-to { margin-bottom:10mm; font-size:10.5pt; line-height:1.5; }',
  '.items-table { width:100%; border-collapse:collapse; margin-bottom:5mm; }',
  '.items-table thead tr th { border-top:1px solid #000; border-bottom:1px solid #000; text-align:right; font-weight:bold; padding:2mm 0; font-size:10.5pt; }',
  '.items-table tbody tr td { padding:2.5mm 0; font-size:10.5pt; vertical-align:top; }',
  '.items-table tbody tr td:last-child { text-align:right; white-space:nowrap; }',
  '.totals-table { width:100%; border-collapse:collapse; margin-top:4mm; }',
  '.totals-table td { padding:1.5mm 0; font-size:10.5pt; text-align:right; }',
  '.totals-table td:first-child { padding-right:8mm; }',
  '.totals-table .grand-total td { font-weight:bold; font-size:11pt; padding-top:2mm; }',
  '.footer { margin-top:auto; padding-top:12mm; text-align:center; font-size:9pt; color:#000; line-height:1.6; }',
  '.footer .bacs { text-decoration:underline; }',
  '.footer .firm-details { margin-top:3mm; color:#444; font-size:8.5pt; }',
  /* Schedule of Work styles */
  '.schedule-page { width:210mm; min-height:297mm; margin:0 auto; padding:18mm 20mm 20mm 20mm; display:flex; flex-direction:column; }',
  '.schedule-header { background:linear-gradient(135deg,' + BRAND_GOLD + ' 0%,#E8D08A 40%,#F5EAB8 55%,#D4B86A 75%,' + BRAND_GOLD + ' 100%); padding:14px 20px; margin-bottom:0; }',
  '.schedule-header h2 { color:' + BRAND_NAVY + '; font-size:16pt; font-weight:bold; margin:0; letter-spacing:.04em; text-shadow:0 1px 2px rgba(255,255,255,0.4); }',
  '.schedule-header-sub { font-size:9pt; color:' + BRAND_NAVY + '; opacity:0.7; margin-top:2px; letter-spacing:.06em; font-style:italic; }',
  '.schedule-case { padding:14px 20px 0 20px; font-size:12pt; font-weight:bold; color:' + BRAND_NAVY + '; }',
  '.schedule-table { width:100%; border-collapse:collapse; margin-top:12px; font-size:10pt; }',
  '.schedule-table th { background:' + BRAND_NAVY + '; color:#fff; padding:8px 10px; text-align:left; font-weight:600; font-size:9.5pt; }',
  '.schedule-table th.r { text-align:right; }',
  '.schedule-table td { border-bottom:1px solid #E5E1D6; padding:7px 10px; color:#1a1a1a; vertical-align:top; }',
  '.schedule-table td.r { text-align:right; font-family:monospace; white-space:nowrap; }',
  '.schedule-table tr:nth-child(even) td { background:#F5F2EB; }',
  '.schedule-table tr:last-child td { border-bottom:none; }',
  '.schedule-totals { border-top:2px solid ' + BRAND_NAVY + '; margin-top:0; }',
  '.schedule-totals td { padding:8px 10px; font-weight:bold; font-size:10.5pt; color:' + BRAND_NAVY + '; }',
  '.schedule-totals td.r { text-align:right; font-family:monospace; }',
  '@page { size:A4; margin:0; }',
  '@media print { body { background:#fff; } .page, .schedule-page { margin:0; width:100%; } }'
].join(' ');

function generateInvoiceHTML(data) {
  // invoiceDate: null = draft (show 'DRAFT' placeholder)
  const isDraft     = !data.invoiceDate;
  const invoiceDate = isDraft ? null
    : (data.invoiceDate instanceof Date) ? data.invoiceDate : new Date(data.invoiceDate);
  const dueDate     = isDraft ? null
    : data.dueDate ? ((data.dueDate instanceof Date) ? data.dueDate : new Date(data.dueDate))
                   : _addDays(invoiceDate, 30);

  const draftingFee      = parseFloat(data.draftingFee)  || 0;
  const laFee            = parseFloat(data.laFee)          || 0;
  const timedFee         = parseFloat(data.timedFee)       || 0;
  const bespokeAmt       = parseFloat(data.bespokeAmt)     || 0;
  const expenses         = parseFloat(data.expenses)       || 0;
  const hrs              = parseFloat(data.timedHours)     || 0;
  const rate             = parseFloat(data.timedRate)      || 0;

  const hasDrafting = draftingFee > 0;
  const hasLaFee    = laFee       > 0;
  const hasTimed    = timedFee    > 0;
  const hasBespoke  = bespokeAmt  > 0;
  const hasExpenses = expenses    > 0;

  // VAT: timed always; drafting fees when vatOnDrafting true; bespoke always
  const vatBase  = timedFee + (data.vatOnDrafting ? (draftingFee + laFee) : 0) + bespokeAmt;
  const vat      = Math.round(vatBase * 0.2 * 100) / 100;
  const subTotal = draftingFee + laFee + timedFee + bespokeAmt;
  const grand    = subTotal + vat + expenses;

  // Sub-total row shown when more than one line item is present
  const lineCount = (hasDrafting ? 1 : 0) + (hasLaFee ? 1 : 0) + (hasTimed ? 1 : 0) + (hasBespoke ? 1 : 0);
  const hasBoth   = lineCount > 1;

  // Timed work description line — do NOT append hours when timedWorkLine is supplied
  let timedLine = '';
  if (hasTimed) {
    if (data.timedWorkLine) {
      timedLine = _esc(data.timedWorkLine);
    } else {
      timedLine = 'Work done per attached schedule | ' + hrs + ' hrs @ ' + _fmtGBP(rate) + ' / hour';
    }
  }

  // Logo src: prefer inline data URL (self-contained); fall back to filename
  const logoSrc = data.logoDataUrl || LOGO_FILENAME;

  // Line item rows
  const rows = [];
  if (hasDrafting) {
    rows.push('<tr><td style="text-align:left;padding-right:8mm">' + _esc(data.draftingFeeLine || 'Preparing Draft Bill of Costs') + '</td><td>' + _fmtGBP(draftingFee) + '</td></tr>');
  }
  if (hasLaFee) {
    rows.push('<tr><td style="text-align:left;padding-right:8mm">' + _esc(data.laFeeLine || 'Preparing Draft Bill of Costs (Legal Aid)') + '</td><td>' + _fmtGBP(laFee) + '</td></tr>');
  }
  if (hasTimed) {
    rows.push('<tr><td style="text-align:left;padding-right:8mm">' + timedLine + '</td><td>' + _fmtGBP(timedFee) + '</td></tr>');
  }
  if (hasBespoke) {
    rows.push('<tr><td style="text-align:left;padding-right:8mm">' + _esc(data.bespokeInvoiceLine || 'Additional work') + '</td><td>' + _fmtGBP(bespokeAmt) + '</td></tr>');
  }
  const lineRows = rows.join('');

  // Totals rows
  const tots = [];
  if (hasBoth) tots.push('<tr><td>Sub Total</td><td style="white-space:nowrap">' + _fmtGBP(subTotal) + '</td></tr>');
  tots.push('<tr><td>VAT @ 20%</td><td>' + _fmtGBP(vat) + '</td></tr>');
  if (hasExpenses) tots.push('<tr><td>Expenses</td><td>' + _fmtGBP(expenses) + '</td></tr>');
  tots.push('<tr class="grand-total"><td>Grand Total</td><td>' + _fmtGBP(grand) + '</td></tr>');
  const totalsRows = tots.join('');

  const yourRefRow = data.yourRef ? '<tr><td>Your Reference</td><td>' + _esc(data.yourRef) + '</td></tr>' : '';

  // Invoice date / due date cells — show 'DRAFT' placeholder when no date
  const invoiceDateCell = isDraft ? '<em style="color:#92400E">DRAFT — date assigned on issue</em>' : _fmtDate(invoiceDate);
  const dueDateCell     = isDraft ? '—' : _fmtDate(dueDate);

  // Draft banner (shown only on drafts)
  const draftBanner = isDraft
    ? '<div class="draft-banner">⚠ DRAFT — For Review Only — Not Yet Issued</div>'
    : '';

  // Invoice page HTML
  let html = '<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">'
    + '<title>Invoice ' + _esc(data.invoiceNumber) + '</title>'
    + '<style>' + INVOICE_CSS + '</style></head><body>'
    + '<div class="page">'
    + '<div class="header"><img class="logo" src="' + logoSrc + '" alt="TMC Legal"></div>'
    + draftBanner
    + '<div class="invoice-title">INVOICE</div>'
    + '<table class="meta-table">'
    + '<tr><td>Invoice Number</td><td>' + _esc(data.invoiceNumber) + '</td></tr>'
    + '<tr><td>Invoice Date</td><td>' + invoiceDateCell + '</td></tr>'
    + '<tr><td>Due Date</td><td>' + dueDateCell + '</td></tr>'
    + '<tr class="meta-spacer"><td colspan="2"></td></tr>'
    + yourRefRow
    + '<tr><td>Our Reference</td><td>' + _esc(data.ourRef) + '</td></tr>'
    + '</table>'
    + '<div class="case-name">' + _esc(data.caseName) + '</div>'
    + '<div class="invoice-to">Invoice to:<br>' + _addrLines(data) + '</div>'
    + '<table class="items-table">'
    + '<thead><tr><th style="text-align:left;width:100%"></th><th style="white-space:nowrap">Amount</th></tr></thead>'
    + '<tbody>' + lineRows + '</tbody>'
    + '</table>'
    + '<table class="totals-table">' + totalsRows + '</table>'
    + '<div class="footer">'
    + '<span class="bacs">For BACS Payments</span><br>'
    + 'Sort Code: ' + BACS_SORT + '<br>'
    + 'Account No: ' + BACS_ACCT
    + '<div class="firm-details">' + FIRM_ADDR + '<br>' + FIRM_CO + ' &nbsp; ' + FIRM_VAT + '</div>'
    + '</div>'
    + '</div>';

  // Append Schedule of Work (only when there are timed entries to list)
  const scheduleLines = data.scheduleLines;
  if (hasTimed && scheduleLines && scheduleLines.length > 0) {
    html += generateScheduleHTML(scheduleLines, data.caseName, data.ourRef);
  }

  html += '</body></html>';
  return html;
}

/**
 * generateScheduleHTML(lines, caseName, ourRef)
 *
 * Builds a Schedule of Work page in navy/gold branding.
 * Appended to the invoice HTML — same document, new page.
 *
 * lines: [{date, workDone, hours, rate, amount}]
 *   date      string|Date  displayed as DD/MM/YYYY
 *   workDone  string
 *   hours     number
 *   rate      number       £/hr
 *   amount    number       hours × rate
 */
function generateScheduleHTML(lines, caseName, ourRef) {
  if (!lines || !lines.length) return '';

  let totalHours  = 0;
  let totalAmount = 0;
  const rows = lines.map(function(l) {
    const hrs = parseFloat(l.hours)  || 0;
    const amt = parseFloat(l.amount) || 0;
    totalHours  += hrs;
    totalAmount += amt;
    const dateStr = l.date ? _fmtDate(l.date) : '—';
    return '<tr>'
      + '<td style="white-space:nowrap">' + _esc(dateStr) + '</td>'
      + '<td>' + _esc(l.workDone || '') + '</td>'
      + '<td class="r">' + hrs + '</td>'
      + '<td class="r">' + (l.rate ? _fmtGBP(l.rate) : '—') + '</td>'
      + '<td class="r">' + _fmtGBP(amt) + '</td>'
      + '</tr>';
  }).join('');

  totalHours  = Math.round(totalHours  * 100) / 100;
  totalAmount = Math.round(totalAmount * 100) / 100;

  const nameStr = caseName ? _esc(caseName) : '';
  const refStr  = ourRef   ? _esc(ourRef)   : '';

  return '<div class="schedule-page">'
    + '<div class="schedule-header"><h2>Schedule of Work</h2></div>'
    + ((nameStr || refStr) ? '<div class="schedule-case">'
        + (nameStr ? nameStr : '')
        + (nameStr && refStr ? '<div style="font-size:11pt;font-weight:normal;color:#555;margin-top:3px">' + refStr + '</div>' : refStr)
        + '</div>' : '')
    + '<table class="schedule-table">'
    + '<thead><tr>'
    + '<th style="min-width:80px">Date</th>'
    + '<th>Work Done</th>'
    + '<th class="r" style="min-width:50px">Time</th>'
    + '<th class="r" style="min-width:60px">Rate</th>'
    + '<th class="r" style="min-width:75px">Amount</th>'
    + '</tr></thead>'
    + '<tbody>' + rows + '</tbody>'
    + '</table>'
    + '<table class="schedule-table schedule-totals">'
    + '<tr>'
    + '<td colspan="2" style="color:' + BRAND_NAVY + ';font-weight:bold">Totals</td>'
    + '<td class="r">' + totalHours + '</td>'
    + '<td class="r"></td>'
    + '<td class="r">' + _fmtGBP(totalAmount) + '</td>'
    + '</tr>'
    + '</table>'
    + '</div>';
}

/**
 * Opens the invoice in a new print window.
 * Returns the HTML string (caller POSTs it to the Azure Function for SP upload).
 */
function printInvoice(data) {
  const html = generateInvoiceHTML(data);
  const win = window.open('', '_blank');
  win.document.write(html);
  win.document.close();
  win.focus();
  setTimeout(() => win.print(), 600);
  return html;
}

// Node/CommonJS export (for testing); ignored in browser
if (typeof module !== 'undefined') module.exports = { generateInvoiceHTML, generateScheduleHTML, printInvoice };
