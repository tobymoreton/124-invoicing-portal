/**
 * invoice-generator.js
 * P124 Invoicing Portal — Invoice HTML/PDF generator
 *
 * generateInvoiceHTML(data) -> HTML string (A4, print-ready)
 *
 * data shape:
 *   invoiceNumber    string
 *   invoiceDate      Date|string
 *   dueDate          Date|string  optional (defaults to invoiceDate + 30 days)
 *   yourRef          string       ClientCaseReference
 *   ourRef           string       Ourreference_x0028_text_x0029_
 *   caseName         string
 *   firmName         string
 *   address1-5       string
 *   draftingFee      number       0 if no drafting element
 *   draftingFeeLine  string       e.g. 'Preparing Draft Bill of Costs | 5.5% of ...'
 *   timedFee         number       0 if no timed element
 *   timedHours       number
 *   timedRate        number
 *   timedWorkLine    string|null  override (null = default text)
 *   expenses         number       default 0
 *   vatOnDrafting    boolean      if true, VAT applies to drafting fee too
 */

const LOGO_URL  = 'New Logo Gold, Silver and Navy.png';
const FIRM_ADDR = 'Fennels Lodge, St Peters Close, Loudwater, Buckinghamshire HP11 1JT.';
const FIRM_CO   = 'Company No: 12348782';
const FIRM_VAT  = 'VAT Registration No: 398 1109 74';
const BACS_SORT = '20-03-84';
const BACS_ACCT = '33605868';

function _esc(s) {
  if (!s) return '';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function _fmtDate(d) {
  if (!d) return '';
  const dt = (d instanceof Date) ? d : new Date(d);
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
  '.page { width:210mm; min-height:297mm; margin:0 auto; padding:18mm 20mm 20mm 20mm; display:flex; flex-direction:column; }',
  '.header { display:flex; justify-content:flex-end; margin-bottom:12mm; }',
  '.logo { width:52mm; height:auto; }',
  '.invoice-title { font-size:22pt; font-weight:bold; margin-bottom:7mm; }',
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
  '@page { size:A4; margin:0; }',
  '@media print { body { background:#fff; } .page { margin:0; width:100%; } }'
].join(' ');

function generateInvoiceHTML(data) {
  const invoiceDate = (data.invoiceDate instanceof Date) ? data.invoiceDate : new Date(data.invoiceDate || Date.now());
  const dueDate     = data.dueDate ? ((data.dueDate instanceof Date) ? data.dueDate : new Date(data.dueDate))
                                   : _addDays(invoiceDate, 30);

  const draftingFee = parseFloat(data.draftingFee) || 0;
  const timedFee    = parseFloat(data.timedFee)    || 0;
  const expenses    = parseFloat(data.expenses)    || 0;
  const hrs         = parseFloat(data.timedHours)  || 0;
  const rate        = parseFloat(data.timedRate)   || 0;

  const hasDrafting = draftingFee > 0;
  const hasTimed    = timedFee    > 0;
  const hasBoth     = hasDrafting && hasTimed;
  const hasExpenses = expenses    > 0;

  const vatBase  = timedFee + (data.vatOnDrafting ? draftingFee : 0);
  const vat      = Math.round(vatBase * 0.2 * 100) / 100;
  const subTotal = draftingFee + timedFee;
  const grand    = subTotal + vat + expenses;

  // Timed work description line
  let timedLine = '';
  if (hasTimed) {
    if (data.timedWorkLine) {
      timedLine = _esc(data.timedWorkLine) + (hrs ? ' | ' + hrs + ' hrs' : '');
    } else {
      timedLine = 'Work done per attached schedule | ' + hrs + ' hrs @ ' + _fmtGBP(rate) + ' / hour';
    }
  }

  // Line item rows
  const rows = [];
  if (hasDrafting) {
    rows.push('<tr><td style="text-align:left;padding-right:8mm">' + _esc(data.draftingFeeLine || 'Preparing Draft Bill of Costs') + '</td><td>' + _fmtGBP(draftingFee) + '</td></tr>');
  }
  if (hasTimed) {
    rows.push('<tr><td style="text-align:left;padding-right:8mm">' + timedLine + '</td><td>' + _fmtGBP(timedFee) + '</td></tr>');
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

  return '<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">'
    + '<title>Invoice ' + _esc(data.invoiceNumber) + '</title>'
    + '<style>' + INVOICE_CSS + '</style></head><body>'
    + '<div class="page">'
    + '<div class="header"><img class="logo" src="' + LOGO_URL + '" alt="TMC Legal"></div>'
    + '<div class="invoice-title">INVOICE</div>'
    + '<table class="meta-table">'
    + '<tr><td>Invoice Number</td><td>' + _esc(data.invoiceNumber) + '</td></tr>'
    + '<tr><td>Invoice Date</td><td>' + _fmtDate(invoiceDate) + '</td></tr>'
    + '<tr><td>Due Date</td><td>' + _fmtDate(dueDate) + '</td></tr>'
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
    + '</div></body></html>';
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
if (typeof module !== 'undefined') module.exports = { generateInvoiceHTML, printInvoice };
