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

const SHEET_NAME = 'Interest';
const FIELDS = ['Package', 'Event date', 'Time of day', 'Email', 'Contact via', 'Phone',
  'Heard via', 'Notes', 'Email specials', 'Text specials', 'Agreed to terms'];
const HEADERS = ['Received', 'Request #'].concat(FIELDS, ['Status']);

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
      id = 'TS-' + (1000 + sh.getLastRow());
      // A leading ' stops Sheets from treating text like "=..." as a formula.
      const row = [new Date(), id].concat(FIELDS.map(k => {
        const v = String(data[k] == null ? '' : data[k]);
        return /^[=+\-@]/.test(v) ? "'" + v : v;
      }), ['New']);
      sh.appendRow(row);
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

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
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
        subject: 'New catering request ' + id + ': ' + d.Package + ' on ' + d['Event date'],
        body: 'A new catering request came in from teaspoon-edh.com/catering.\n\n' + lines +
              '\n\nOpen the sheet: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
      });
    }
    if (SEND_CUSTOMER_CONFIRMATION && d.Email) {
      MailApp.sendEmail({
        to: d.Email,
        subject: 'We got your Teaspoon catering request (' + id + ')',
        name: 'Teaspoon El Dorado Hills',
        body: 'Hi there,\n\n' +
              'Thank you for your interest in Teaspoon El Dorado Hills catering! ' +
              "We'll get back to you within the next business day with additional details regarding your order.\n\n" +
              'Your request: ' + d.Package + ' · ' + d['Event date'] + ' · ' + d['Time of day'] + '\n' +
              'Request number: ' + id + '\n\n' +
              'Teaspoon El Dorado Hills\n' +
              '4420 Town Center Blvd, Ste 110, El Dorado Hills, CA 95762'
      });
    }
  } catch (err) {
    // The row is already saved; a mail problem shouldn't fail the request.
    console.error('Email failed: ' + err);
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
