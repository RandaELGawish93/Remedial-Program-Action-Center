/**
 * GenericEntryCheck.gs
 *
 * Flags too-generic Weakness/Steps/Means entries in Week 1 and Week 5
 * action plans (Week 10 has no such fields - it's attendance/quiz only)
 * and lets the Dean review + approve before anything is emailed to a
 * teacher.
 *
 * NOTHING in this file sends email automatically. scanGenericEntries_()
 * only reads teacher workbooks and writes flagged rows into the
 * "Generic Entry Review" sheet in THIS spreadsheet. Sending only happens
 * when sendApprovedGenericEmails_() is run by hand, and even then it only
 * emails rows where the Approve checkbox is ticked, after a confirmation
 * dialog.
 *
 * Run either function from the Apps Script Run dropdown (or wire up your
 * own menu item / trigger later once you're happy with the results).
 *
 * Detection is rules-first: a phrase blocklist (editable in the
 * "Generic Phrases" sheet, no code changes needed) and a length check
 * catch the obvious cases for free. Anything in between "obviously fine"
 * and "obviously generic" gets a real judgment call from the Gemini API
 * free tier (needs a GEMINI_API_KEY script property - if it's not set,
 * borderline entries are just left alone rather than flagged).
 */

const GENERIC_PHRASES_SHEET = 'Generic Phrases';
const GENERIC_REVIEW_SHEET = 'Generic Entry Review';

const GENERIC_WEEK_CONFIGS = [
  {
    weekNum: 1,
    tabName: 'Week 1',
    cols: WEEK1_COLS,
    hardFailLen: 12,
    clearLen: 24,
    fields: [
      { key: 'WEAK1', label: 'Weakness 1' },
      { key: 'STEPS1', label: 'Steps 1' },
      { key: 'WEAK2', label: 'Weakness 2' },
      { key: 'STEPS2', label: 'Steps 2' },
      { key: 'MEANS', label: 'Means' }
    ]
  },
  {
    weekNum: 5,
    tabName: 'Week 5',
    cols: WEEK5_COLS,
    hardFailLen: 25,
    clearLen: 50,
    fields: [
      { key: 'WEAK1_AP2', label: 'Weakness 1' },
      { key: 'STEPS1_AP2', label: 'Steps 1' },
      { key: 'WEAK2_AP2', label: 'Weakness 2' },
      { key: 'STEPS2_AP2', label: 'Steps 2' },
      { key: 'MEANS_AP2', label: 'Means' }
    ]
  }
];

const GENERIC_REVIEW_HEADERS = [
  'Flagged Date', 'Teacher', 'Student', 'Subject', 'Week', 'Field',
  'Entry Text', 'Suggested Note', 'Source', 'Approve', 'Ignore', 'Status', 'Sent Date', 'EntryHash'
];

const GENERIC_DEFAULT_PHRASES = [
  'extra practice', 'more practice', 'general weakness', 'general weaknesses',
  'needs improvement', 'needs more practice', 'practice sheets', 'additional support',
  'extra support', 'study more', 'work harder', 'try harder', 'homework', 'review'
];


function safeAlert_(title, message) {
  try {
    const ui = SpreadsheetApp.getUi();
    if (title) { ui.alert(title, message, ui.ButtonSet.OK); } else { ui.alert(message); }
  } catch (e) {
    Logger.log((title ? title + ' - ' : '') + message);
  }
}

function safeConfirm_(title, message) {
  try {
    const ui = SpreadsheetApp.getUi();
    return ui.alert(title, message, ui.ButtonSet.YES_NO) === ui.Button.YES;
  } catch (e) {
    Logger.log('No spreadsheet UI available (running from the Apps Script editor, not a menu click) - refusing to send without an explicit Yes, so nothing was emailed. Run this from the sheet\'s menu to get a confirmation prompt, or approve manually.');
    return false;
  }
}

function ensureGenericSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let phrasesSh = ss.getSheetByName(GENERIC_PHRASES_SHEET);
  if (!phrasesSh) {
    phrasesSh = ss.insertSheet(GENERIC_PHRASES_SHEET);
    phrasesSh.getRange(1, 1).setValue('Phrase (not case-sensitive, matches anywhere inside an entry - edit this list anytime)');
    phrasesSh.getRange(1, 1).setFontWeight('bold');
    phrasesSh.getRange(2, 1, GENERIC_DEFAULT_PHRASES.length, 1).setValues(GENERIC_DEFAULT_PHRASES.map(p => [p]));
    phrasesSh.setColumnWidth(1, 420);
  }

  let reviewSh = ss.getSheetByName(GENERIC_REVIEW_SHEET);
  if (!reviewSh) {
    reviewSh = ss.insertSheet(GENERIC_REVIEW_SHEET);
    reviewSh.getRange(1, 1, 1, GENERIC_REVIEW_HEADERS.length).setValues([GENERIC_REVIEW_HEADERS]);
    reviewSh.getRange(1, 1, 1, GENERIC_REVIEW_HEADERS.length).setFontWeight('bold');
    reviewSh.setFrozenRows(1);
    reviewSh.setColumnWidth(7, 260);
    reviewSh.setColumnWidth(8, 320);
    reviewSh.hideColumns(14);
  }

  return { phrasesSh, reviewSh };
}

function readGenericPhrases_(phrasesSh) {
  const lastRow = phrasesSh.getLastRow();
  if (lastRow < 2) return [];
  return phrasesSh.getRange(2, 1, lastRow - 1, 1).getValues()
    .map(r => String(r[0]).trim().toLowerCase())
    .filter(p => p.length > 0);
}

function normalizeGenericText_(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/[.,!?;:]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function simpleHash_(s) {
  let hash = 0;
  const str = String(s);
  for (let i = 0; i < str.length; i++) {
    hash = ((hash << 5) - hash) + str.charCodeAt(i);
    hash |= 0;
  }
  return String(hash);
}

function classifyGenericEntry_(rawText, phrases, hardFailLen, clearLen) {
  const text = normalizeGenericText_(rawText);
  if (!text) return { status: 'clear' };

  for (const phrase of phrases) {
    if (phrase && text.includes(phrase)) {
      return { status: 'flagged', reason: 'Matches phrase: "' + phrase + '"' };
    }
  }

  if (text.length < hardFailLen) {
    return { status: 'flagged', reason: 'Very short (' + text.length + ' characters)' };
  }

  if (text.length >= clearLen) {
    return { status: 'clear' };
  }

  return { status: 'borderline' };
}

function ruleSuggestion_(fieldLabel) {
  if (fieldLabel.indexOf('Weakness') === 0) {
    return 'Name the specific skill or concept, not a general category (e.g. "identifying main idea in nonfiction passages" rather than "reading").';
  }
  if (fieldLabel.indexOf('Steps') === 0) {
    return 'Describe the specific action or resource, not just "practice" (e.g. "guided annotation of weekly nonfiction passages").';
  }
  return 'Be specific about what this actually involves, not a generic label.';
}

function callGeminiGenericJudge_(entryText, fieldLabel, subject, weekNum) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) return { generic: false, suggestion: '' };

  const prompt = 'You are helping an Academic Dean review remedial action plan entries written by teachers at a school.\n' +
    'This is a Week ' + weekNum + ' "' + fieldLabel + '" entry for a ' + subject + ' student:\n' +
    '"' + entryText + '"\n\n' +
    'Judge whether this entry is too generic/vague to be useful (it should name a specific skill, ' +
    'strategy, or resource rather than a broad category). It is normal for Week 1 entries to be ' +
    'somewhat broad, but by Week 5 entries should be concrete.\n\n' +
    'Respond with ONLY compact JSON, no markdown fences, no explanation outside the JSON:\n' +
    '{"generic": true or false, "suggestion": "if generic, one short sentence suggesting a more specific version tailored to the subject; empty string if not generic"}';

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.2, maxOutputTokens: 200 }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const resp = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent',
      options
    );
    if (resp.getResponseCode() !== 200) {
      Logger.log('Gemini API error ' + resp.getResponseCode() + ': ' + resp.getContentText());
      return { generic: false, suggestion: '' };
    }
    const json = JSON.parse(resp.getContentText());
    let text = json.candidates && json.candidates[0] && json.candidates[0].content &&
      json.candidates[0].content.parts && json.candidates[0].content.parts[0] &&
      json.candidates[0].content.parts[0].text;
    if (!text) return { generic: false, suggestion: '' };

    text = text.trim().replace(/^```json/i, '').replace(/^```/, '').replace(/```$/, '').trim();
    const parsed = JSON.parse(text);
    return { generic: !!parsed.generic, suggestion: String(parsed.suggestion || '') };
  } catch (e) {
    Logger.log('Gemini judge failed: ' + e.message);
    return { generic: false, suggestion: '' };
  }
}

function scanGenericEntries_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const { phrasesSh, reviewSh } = ensureGenericSheets_();
  const phrases = readGenericPhrases_(phrasesSh);

  const existingLastRow = reviewSh.getLastRow();
  const existingKeys = {};
  if (existingLastRow > 1) {
    const existing = reviewSh.getRange(2, 1, existingLastRow - 1, GENERIC_REVIEW_HEADERS.length).getValues();
    existing.forEach((row, i) => {
      const key = [row[1], row[2], row[4], row[5]].join('||');
      const rowIndex = i + 2;
      if (row[10] === true && row[11] === 'Pending') {
        reviewSh.getRange(rowIndex, 12).setValue('Ignored');
        row[11] = 'Ignored';
      }
      existingKeys[key] = { hash: row[13], status: row[11], rowIndex: rowIndex };
    });
  }

  const trSh = ss.getSheetByName(SHEETS.TEACHER_ROSTER);
  const lastRow = trSh.getLastRow();
  if (lastRow < 2) { safeAlert_(null, 'No teachers in roster.'); return; }
  const rows = trSh.getRange(2, 1, lastRow - 1, TEACHER_ROSTER_HEADERS.length).getValues();

  const newRows = [];
  let scannedTeachers = 0, aiCalls = 0, resolvedStale = 0;

  rows.forEach(row => {
    const teacherName = row[TEACHER_ROSTER_COL.NAME - 1];
    const subject = row[TEACHER_ROSTER_COL.SUBJECT - 1];
    const workbookUrl = row[TEACHER_ROSTER_COL.WORKBOOK_LINK - 1];
    if (!teacherName || !workbookUrl) return;

    let tws;
    try {
      const fileId = extractFileIdFromUrl_(workbookUrl);
      tws = SpreadsheetApp.openById(fileId);
    } catch (e) {
      return;
    }
    scannedTeachers++;

    GENERIC_WEEK_CONFIGS.forEach(cfg => {
      const sh = tws.getSheetByName(cfg.tabName);
      if (!sh) return;
      const lr = sh.getLastRow();
      if (lr < 4) return;

      for (let r = 4; r <= lr; r++) {
        const studentName = sh.getRange(r, cfg.cols.NAME).getValue();
        if (!studentName) continue;

        cfg.fields.forEach(field => {
          const colIdx = cfg.cols[field.key];
          const val = sh.getRange(r, colIdx).getValue();
          const text = String(val || '').trim();
          if (!text) return;

          const key = [teacherName, studentName, cfg.weekNum, field.label].join('||');
          const hash = simpleHash_(text);
          const prior = existingKeys[key];
          if (prior && String(prior.hash) === hash) return;

          if (prior && prior.status === 'Pending') {
            reviewSh.getRange(prior.rowIndex, 12).setValue('Resolved (entry changed)');
            resolvedStale++;
          }

          const result = classifyGenericEntry_(text, phrases, cfg.hardFailLen, cfg.clearLen);
          let source = null, suggestion = null;

          if (result.status === 'flagged') {
            source = 'Rule';
            suggestion = ruleSuggestion_(field.label);
          } else if (result.status === 'borderline') {
            aiCalls++;
            const judged = callGeminiGenericJudge_(text, field.label, subject, cfg.weekNum);
            if (judged.generic) {
              source = 'AI';
              suggestion = judged.suggestion || ruleSuggestion_(field.label);
            }
          }

          if (source) {
            newRows.push([
              new Date(), teacherName, studentName, subject, cfg.weekNum, field.label,
              text, suggestion, source, false, false, 'Pending', '', hash
            ]);
          }
        });
      }
    });
  });

  if (newRows.length) {
    const startRow = reviewSh.getLastRow() + 1;
    reviewSh.getRange(startRow, 1, newRows.length, GENERIC_REVIEW_HEADERS.length).setValues(newRows);
    reviewSh.getRange(startRow, 10, newRows.length, 2).insertCheckboxes();
  }

  safeAlert_('Generic entry scan complete', 
    'Teachers scanned: ' + scannedTeachers + '\n' +
    'AI judgment calls: ' + aiCalls + '\n' +
    'New entries flagged: ' + newRows.length + '\n' +
    'Stale flags resolved (entry text changed since): ' + resolvedStale + '\n\n' +
    'Review them in the "' + GENERIC_REVIEW_SHEET + '" tab, tick Approve for the ones you want sent, ' +
    'then run sendApprovedGenericEmails_.'
  );
}

function genEscapeHtml_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function sendApprovedGenericEmails_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reviewSh = ss.getSheetByName(GENERIC_REVIEW_SHEET);
  if (!reviewSh) { safeAlert_(null, 'No "' + GENERIC_REVIEW_SHEET + '" sheet found. Run scanGenericEntries_ first.'); return; }
  const lastRow = reviewSh.getLastRow();
  if (lastRow < 2) { safeAlert_(null, 'Nothing to send.'); return; }
  const data = reviewSh.getRange(2, 1, lastRow - 1, GENERIC_REVIEW_HEADERS.length).getValues();

  const trSh = ss.getSheetByName(SHEETS.TEACHER_ROSTER);
  const trLastRow = trSh.getLastRow();
  const trRows = trSh.getRange(2, 1, trLastRow - 1, TEACHER_ROSTER_HEADERS.length).getValues();
  const teacherEmail = {};
  trRows.forEach(r => {
    const name = r[TEACHER_ROSTER_COL.NAME - 1];
    const email = r[TEACHER_ROSTER_COL.EMAIL - 1];
    if (name) teacherEmail[name] = email;
  });

  const byTeacher = {};
  data.forEach((row, i) => {
    const teacher = row[1], student = row[2], subject = row[3], week = row[4],
      field = row[5], entryText = row[6], suggestion = row[7], approve = row[9], ignore = row[10], status = row[11];
    if (approve !== true || ignore === true) return;
    if (status === 'Sent') return;
    if (!byTeacher[teacher]) byTeacher[teacher] = [];
    byTeacher[teacher].push({ rowIndex: i + 2, student, subject, week, field, entryText, suggestion });
  });

  const teacherNames = Object.keys(byTeacher);
  if (!teacherNames.length) {
    safeAlert_(null, 'No approved, unsent rows found. Tick the Approve checkbox on the rows you want sent first.');
    return;
  }

  const preview = teacherNames.map(t => t + ' (' + byTeacher[t].length + ' entries)' + (teacherEmail[t] ? '' : ' [NO EMAIL ON FILE - will be skipped]')).join('\n');
  const confirmed = safeConfirm_(
    'Send approved emails?',
    'This will email the following teachers:\n\n' + preview + '\n\nContinue?'
  );
  if (!confirmed) return;

  let sentCount = 0;
  const skippedNoEmail = [];

  teacherNames.forEach(teacher => {
    const email = teacherEmail[teacher];
    if (!email) { skippedNoEmail.push(teacher); return; }

    const items = byTeacher[teacher];
    const byStudent = {};
    items.forEach(it => {
      const k = it.student + ' (Week ' + it.week + ', ' + it.subject + ')';
      if (!byStudent[k]) byStudent[k] = [];
      byStudent[k].push(it);
    });

    let cardsHtml = '';
    Object.keys(byStudent).forEach(studentKey => {
      let rowsHtml = '';
      byStudent[studentKey].forEach(it => {
        rowsHtml +=
          '<div style="background:#fef7e0;border-left:3px solid #f9ab00;border-radius:4px;padding:8px 12px;margin:4px 0 6px;">' +
            '<div style="font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:#b06000;font-weight:bold;margin-bottom:2px;">' + genEscapeHtml_(it.field) + '</div>' +
            '<div style="font-size:13px;color:#3c4043;font-style:italic;">&ldquo;' + genEscapeHtml_(it.entryText) + '&rdquo;</div>' +
          '</div>' +
          '<div style="background:#e6f4ea;border-left:3px solid #34a853;border-radius:4px;padding:8px 12px;margin:0 0 14px;">' +
            '<div style="font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:#137333;font-weight:bold;margin-bottom:2px;">Suggested</div>' +
            '<div style="font-size:13px;color:#3c4043;">' + genEscapeHtml_(it.suggestion) + '</div>' +
          '</div>';
      });
      cardsHtml +=
        '<div style="margin:16px 0;padding:14px 16px;background:#f8f9fa;border-radius:6px;">' +
          '<div style="font-weight:bold;color:#1c4587;margin-bottom:8px;font-size:14px;">' + genEscapeHtml_(studentKey) + '</div>' +
          rowsHtml +
        '</div>';
    });

    const bodyHtml = emailWrapper_('Action Plan Entries - A Few Could Use More Detail',
      '<p>Hi ' + genEscapeHtml_(teacher) + ',</p>' +
      '<p>While reviewing this week\'s action plan entries, a few could use more specific detail so they\'re most useful for the students:</p>' +
      cardsHtml +
      '<p>Thanks for all you do for these students!</p>'
    );

    sendLoggedEmail_(email, 'Action Plan Entries - A Few Could Use More Detail', bodyHtml, 'Generic Entry Review', teacher);

    const now = new Date();
    items.forEach(it => {
      reviewSh.getRange(it.rowIndex, 12).setValue('Sent');
      reviewSh.getRange(it.rowIndex, 13).setValue(now);
    });

    sentCount += items.length;
  });

  let msg = 'Sent digest emails to ' + teacherNames.filter(t => teacherEmail[t]).length + ' teacher(s), ' +
    sentCount + ' entries total.';
  if (skippedNoEmail.length) {
    msg += '\n\nSkipped (no email on file): ' + skippedNoEmail.join(', ');
  }
  safeAlert_(null, msg);
}

// Public entry points (Run dropdown only lists non-underscore top-level functions)
function runGenericEntryScan() {
  scanGenericEntries_();
}

function runSendApprovedGenericEmails() {
  sendApprovedGenericEmails_();
}

// Live sync: instant Ignore -> Status update, no need to re-run the scan.
// Installable trigger (not a simple trigger) so it needs a one-time
// authorization - run installGenericReviewEditTrigger() once from the
// Run dropdown and approve the permissions prompt.
function applyGenericReviewRowFormat_(sh, row) {
  const approve = sh.getRange(row, 10).getValue() === true;
  const ignore = sh.getRange(row, 11).getValue() === true;
  const statusCell = sh.getRange(row, 12);
  const currentStatus = statusCell.getValue();
  const rowRange = sh.getRange(row, 1, 1, GENERIC_REVIEW_HEADERS.length - 1); // A:M, leave hidden EntryHash alone

  if (ignore) {
    if (currentStatus === 'Pending') statusCell.setValue('Ignored');
    rowRange.setBackground('#f3f3f3'); // gray = dismissed
  } else {
    if (currentStatus === 'Ignored') statusCell.setValue('Pending');
    rowRange.setBackground(approve ? '#d9ead3' : null); // green = ready to send, else clear
  }
}

function onEditGenericReview(e) {
  try {
    const range = e.range;
    const sh = range.getSheet();
    if (sh.getName() !== GENERIC_REVIEW_SHEET) return;

    const startCol = range.getColumn();
    const endCol = startCol + range.getNumColumns() - 1;
    // Only react when the edit touches Approve (10) or Ignore (11)
    if (endCol < 10 || startCol > 11) return;

    const startRow = Math.max(range.getRow(), 2); // skip header row
    const endRow = range.getRow() + range.getNumRows() - 1;
    if (endRow < 2) return;

    for (let row = startRow; row <= endRow; row++) applyGenericReviewRowFormat_(sh, row);
  } catch (err) {
    Logger.log('onEditGenericReview error: ' + err);
  }
}

// One-time / on-demand cleanup: repaints every row's background from its
// current Approve/Ignore checkbox state. Useful after rows were ticked
// before the live trigger existed, or if formatting ever drifts.
function repaintGenericReviewFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(GENERIC_REVIEW_SHEET);
  if (!sh) { safeAlert_(null, 'No "' + GENERIC_REVIEW_SHEET + '" sheet found.'); return; }
  const lastRow = sh.getLastRow();
  if (lastRow < 2) { safeAlert_(null, 'Nothing to repaint.'); return; }
  for (let row = 2; row <= lastRow; row++) applyGenericReviewRowFormat_(sh, row);
  safeAlert_(null, 'Refreshed formatting on ' + (lastRow - 1) + ' row(s).');
}

function installGenericReviewEditTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onEditGenericReview') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onEditGenericReview')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
  safeAlert_('Trigger installed', 'Ticking Ignore on the "' + GENERIC_REVIEW_SHEET + '" tab now updates Status instantly.');
}
