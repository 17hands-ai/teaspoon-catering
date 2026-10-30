/**
 * Teaspoon catering — interest form (Google Apps Script)
 *
 * Saves each request from teaspoon-edh.com/catering as a row in this spreadsheet,
 * emails the shop, and (optionally) emails the customer a short confirmation.
 *
 * Setup (once):
 *   1. Fill in NOTIFY_EMAIL below and click Save.
 *   2. Choose "setup" in the function menu and click Run. Approve the permissions.
 *      Copy the secret it prints in the Execution log.
 *   3. Deploy → New deployment → Web app. Execute as: Me. Who has access: Anyone.
 *      Copy the Web app URL (ends in /exec).
 *   4. In Vercel, set INTEREST_SCRIPT_URL (the URL) and INTEREST_SECRET (the secret),
 *      then redeploy.
 */

// ← Who gets an email for each new request (comma-separate several addresses).
const NOTIFY_EMAIL = 'you@example.com';
// Send the customer a "we got your request" email? true or false.
const SEND_CUSTOMER_CONFIRMATION = true;
// Only needed if this script was NOT opened from the sheet (Extensions → Apps Script):
// paste the long ID from the sheet's address, between /d/ and /edit.
const SPREADSHEET_ID = '';

// Request numbers look like TS-MMDD-01, counting up each day in this time zone.
const TIME_ZONE = 'America/Los_Angeles';
const SHEET_NAME = 'Interest';
const FIELDS = ['First name', 'Package', 'Event date', 'Time of day', 'Email', 'Contact via', 'Phone',
  'Heard via', 'Notes', 'Email specials', 'Text specials', 'Agreed to terms'];
const HEADERS = ['Received', 'Request ID'].concat(FIELDS, ['Status']);

function setup() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('SECRET');
  if (!secret) {
    secret = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
    props.setProperty('SECRET', secret);
  }
  sheet_();
  // Touch MailApp so the permission prompt asks for email access now.
  MailApp.getRemainingDailyQuota();
  Logger.log('Your INTEREST_SECRET for Vercel is:\n' + secret);
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const secret = PropertiesService.getScriptProperties().getProperty('SECRET');
    if (!secret || body.secret !== secret) return json_({ ok: false, error: 'Unauthorized' });
    const data = body.data || {};

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    let id;
    try {
      const sh = sheet_();
      id = nextId_();
      const values = { 'Received': new Date(), 'Request ID': id, 'Status': 'New' };
      FIELDS.forEach(k => {
        const v = String(data[k] == null ? '' : data[k]);
        // A leading ' stops Sheets from treating text like "=..." as a formula.
        values[k] = /^[=+\-@]/.test(v) ? "'" + v : v;
      });
      // Values go under their column names, so columns can be moved or added in the sheet.
      const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
      sh.appendRow(head.map(h => h in values ? values[h] : ''));
    } finally {
      lock.releaseLock();
    }

    notify_(id, data);
    return json_({ ok: true, id: id });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'Server error' });
  }
}

function nextId_() {
  const props = PropertiesService.getScriptProperties();
  const day = Utilities.formatDate(new Date(), TIME_ZONE, 'yyyyMMdd');
  const n = props.getProperty('SEQ_DAY') === day ? Number(props.getProperty('SEQ_N')) + 1 : 1;
  props.setProperties({ SEQ_DAY: day, SEQ_N: String(n) });
  return 'TS-' + day.slice(4) + '-' + String(n).padStart(2, '0');
}

function sheet_() {
  const ss = spreadsheet_();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else {
    // Older sheets: rename "Request #" and add any new columns at the end.
    const cols = sh.getLastColumn();
    const head = sh.getRange(1, 1, 1, cols).getValues()[0].map(String);
    const old = head.indexOf('Request #');
    if (old >= 0) { sh.getRange(1, old + 1).setValue('Request ID'); head[old] = 'Request ID'; }
    const missing = HEADERS.filter(h => head.indexOf(h) < 0);
    if (missing.length) sh.getRange(1, cols + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
  }
  return sh;
}

function notify_(id, d) {
  const lines = FIELDS.filter(k => d[k]).map(k => k + ': ' + d[k]).join('\n');
  try {
    if (NOTIFY_EMAIL && NOTIFY_EMAIL !== 'you@example.com') {
      MailApp.sendEmail({
        to: NOTIFY_EMAIL,
        replyTo: d.Email,
        subject: 'New catering request ' + id + ' from ' + d['First name'] + ': ' + d.Package + ' on ' + d['Event date'],
        body: 'A new catering request came in from teaspoon-edh.com/catering.\n\n' + lines +
              '\n\nOpen the sheet: ' + spreadsheet_().getUrl()
      });
    }
    if (SEND_CUSTOMER_CONFIRMATION && d.Email) {
      MailApp.sendEmail({
        to: d.Email,
        subject: 'We got your Teaspoon catering request (' + id + ')',
        name: 'Teaspoon El Dorado Hills',
        body: 'Hi ' + (d['First name'] || 'there') + ',\n\n' +
              'Thank you for your interest in Teaspoon El Dorado Hills catering! ' +
              "We'll get back to you within the next business day with additional details regarding your order.\n\n" +
              'Your request: ' + d.Package + ' · ' + d['Event date'] + ' · ' + d['Time of day'] + '\n' +
              'Request ID: ' + id + '\n\n' +
              'Teaspoon El Dorado Hills\n' +
              '4420 Town Center Blvd, Ste 110, El Dorado Hills, CA 95762'
      });
    }
  } catch (err) {
    // The row is already saved; a mail problem shouldn't fail the request.
    console.error('Email failed: ' + err);
  }
}

function spreadsheet_() {
  const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No spreadsheet found. Open this script from the sheet (Extensions → Apps Script), or fill in SPREADSHEET_ID.');
  return ss;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
