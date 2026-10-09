/**
 * SETUP_EMAILLOG.gs
 * One shared entry point for every automated email in this system.
 * Nothing should call MailApp/GmailApp directly anywhere else —
 * always call sendLoggedEmail_() so every send is logged automatically.
 */

const EMAIL_LOG_SHEET = 'Email Log';
const EMAIL_LOG_HEADERS = ['Timestamp', 'To', 'Subject', 'Type', 'Related To (Student/Teacher)', 'Status', 'Error (if any)'];

function buildEmailLogTab() {
  const sh = getOrCreateSheet_(EMAIL_LOG_SHEET);
  sh.clear();
  sh.getRange(1, 1, 1, EMAIL_LOG_HEADERS.length)
    .setValues([EMAIL_LOG_HEADERS])
    .setFontWeight('bold').setBackground('#1c4587').setFontColor('#ffffff');
  sh.setFrozenRows(1);

  // Defensive: the MoM Log "Email Sent" formulas reference this sheet's
  // columns out to row 2000. If this sheet ever has fewer than 2000 rows
  // (e.g. Google Sheets' default 1000-row new-sheet size), those formulas
  // silently error out (#N/A inside SUMPRODUCT/MAX) and IFERROR masks it
  // as "Not Sent" even when a matching Email Log row genuinely exists.
  // Keep at least 2000 rows here so that never happens.
  if (sh.getMaxRows() < 2000) {
    sh.insertRowsAfter(sh.getMaxRows(), 2000 - sh.getMaxRows());
  }
  const widths = [150, 220, 250, 180, 200, 90, 250];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  SpreadsheetApp.getActiveSpreadsheet().toast('Email Log built.', 'Done', 4);
}

/**
 * Send an email and log the attempt (success or failure) in Email Log.
 * @param {string} to        recipient email
 * @param {string} subject
 * @param {string} htmlBody
 * @param {string} type       e.g. 'Missed Action Plan', 'MoM Reminder', 'Missed Verification'
 * @param {string} relatedTo  e.g. student name/ID or teacher name, for traceability
 */
function sendLoggedEmail_(to, subject, htmlBody, type, relatedTo, cc) {
  if (!to) return;
  const options = { htmlBody: htmlBody, name: "Academic Dean's Office" };
  if (cc) options.cc = cc;
  let status = 'Sent', error = '';
  try {
    GmailApp.sendEmail(to, subject, '', options);
  } catch (e) {
    status = 'Failed';
    error = e.message;
  }
  logEmailOnly_(to, subject, type, relatedTo, status, error);
}

/**
 * FIX: this function — and sendLoggedEmail_ above — were called
 * everywhere in the project (missed Action Plan emails, remedial
 * punctuality emails, Progress Report ready emails, MoM sends, teacher
 * onboarding emails) but were never actually defined, so every one of
 * those email-sending features was throwing "logEmailOnly_ is not
 * defined" and silently failing. Added both here, matching the exact
 * call signatures already used throughout the project (same fix
 * applied to the British workbook's Setup_EmailLog.gs).
 *
 * Appends one row to Email Log without sending anything — used when the
 * caller already sent the email itself (e.g. one shared batch email with
 * multiple recipients, like MoM) and just needs the record kept, once
 * per recipient.
 */
function logEmailOnly_(to, subject, type, relatedTo, status, error) {
  const sh = getOrCreateSheet_(EMAIL_LOG_SHEET);
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, EMAIL_LOG_HEADERS.length).setValues([EMAIL_LOG_HEADERS])
      .setFontWeight('bold').setBackground('#1c4587').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  sh.appendRow([new Date(), to, subject, type, relatedTo, status, error || '']);
}


