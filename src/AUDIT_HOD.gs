/**
 * AUDIT_HOD.gs
 * Menus, audits, HoD / teacher / principal communication, generation guard.
 *
 * EVERY email built here is shown in a PREVIEW window first (recipient, CC,
 * subject and the full message). Nothing is sent until you click "Send selected".
 *
 * Menus (built by buildAllMenus_, called from onOpen in CODE.gs):
 *   1. Setup                       – build tabs + HoD Directory
 *   2. Teachers & Principals       – workbooks, roster, profiles + scores, onboarding, principal workbooks
 *   3. Audit & HoD Communication   – department status, audits per tab group, template check, thank-yous
 *   4. MoM Log                     – unchanged
 *   5. Send Deadlines              – deadline per tab group to teachers, HoDs in CC
 *   6. Student Roster & Probation  – unchanged
 *   7. System & End of Term        – unchanged
 */

// ---------------------------------------------------------------- config

const AUDIT_SHEETS = {
  HOD: 'HoD Directory',
  AUDIT: 'Action Plan Audit',
  STATUS: 'HoD Status Report',
  TEMPLATE: 'AP Template Check'
};

/** Entries that count as EMPTY even though something is typed (not case-sensitive). */
const AP_BLANK_EQUIVALENTS = [
  '-', '--', '.', 'x', 'n/a', 'na', 'none', 'nil', 'tbd', 'nothing',
  'not spotted', 'not spotted yet', 'no weakness', 'no weaknesses'
];

/** CC the HoD(s) and the division principal(s) on teacher thank-you emails. */
const THANK_YOU_CC_HOD = true;

const HOD_HEADERS = ['Subject', 'Division (ES / MS / HS / All)', 'HoD Name', 'HoD Email', 'Active (Yes/No)'];

// Starting rows only — fill in names and emails, add rows freely.
// A subject can have several rows (e.g. one HoD per division).
const HOD_SEED = [
  ['English', 'ES', '', '', 'Yes'],
  ['English', 'MS, HS', '', '', 'Yes'],
  ['Mathematics', 'ES', '', '', 'Yes'],
  ['Mathematics', 'MS, HS', '', '', 'Yes'],
  ['Science', 'All', '', '', 'Yes'],
  ['Arabic', 'All', '', '', 'Yes'],
  ['Arabic Social Studies', 'All', '', '', 'Yes'],
  ['Social Studies', 'All', '', '', 'Yes'],
  ['Islamic Religion', 'All', '', '', 'Yes']
];

const AUDIT_TABS = ['Week 1', 'Week 5', 'Progress Report Entries', 'Week 10'];
/** Tab groups used by audits, HoD notices, deadlines and thank-yous. */
const PHASES = {
  1: { key: 'W1', label: 'Week 1 (Action Plan 1)', tabs: ['Week 1'], sheet: 'Audit — Week 1' },
  2: { key: 'W5', label: 'Week 5 (Action Plan 2) + Progress Report Entries', tabs: ['Week 5', 'Progress Report Entries'], sheet: 'Audit — Week 5 & PR' },
  3: { key: 'W10', label: 'Week 10 (Action Plan 2 Closeout)', tabs: ['Week 10'], sheet: 'Audit — Week 10' }
};

/** What teachers must complete in each tab group (used in deadline emails). */
const PHASE_REQUIREMENTS = {
  1: ['Week 1 tab — for every student: Spotted Weakness 1 & 2, Steps to Tackle Weakness 1 & 2, Means of Assessment, Evidence of Improvement',
      'Tick “Generate?” on each row so the Action Plan is created'],
  2: ['Week 5 tab — Cumulative Score 1, weekly Exit Tickets & attendance (W1–W5), both quizzes, Written Summary, and the six Action Plan 2 fields',
      'Tick “Generate?” on each row so Action Plan 2 is added',
      'Progress Report Entries tab — answer all 12 statements for every student and set Status to “Done”'],
  3: ['Week 10 tab — Cumulative Score 2, weekly Exit Tickets & attendance (W6–W10), both quizzes and the Written Summary',
      'Tick “Generate?” on each row so the final Action Plan PDF is produced']
};

/** Free-text columns per tab (checked against the Generic Phrases list for thank-yous). */
function textColsFor_(tabName) {
  if (tabName === 'Week 1') return requiredColsFor_('Week 1');
  if (tabName === 'Week 5') { const c = WEEK5_COLS; return [c.WRITTEN_SUMMARY, c.WEAK1_AP2, c.STEPS1_AP2, c.WEAK2_AP2, c.STEPS2_AP2, c.MEANS_AP2, c.EVIDENCE_AP2]; }
  if (tabName === 'Week 10') return [WEEK10_COLS.WRITTEN_SUMMARY];
  return [];
}

/** Asks which tab group. Returns 1, 2, 3 or null. */
function askPhase_(title) {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt(title,
    'Which tabs?\n\n1 = Week 1\n2 = Week 5 + Progress Report Entries\n3 = Week 10\n\nType 1, 2 or 3:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return null;
  const n = parseInt(r.getResponseText(), 10);
  if (!PHASES[n]) { ui.alert('Please type 1, 2 or 3.'); return null; }
  return n;
}

function askDeadline_(optional) {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Deadline', optional
    ? 'Deadline for outstanding items (e.g. Thursday 8 October 2026) — leave blank for none:'
    : 'Deadline (e.g. Thursday 8 October 2026):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return null;
  const d = r.getResponseText().trim();
  if (!d && !optional) { ui.alert('No deadline entered — nothing prepared.'); return null; }
  return d;
}

// ---------------------------------------------------------------- menus

function buildAllMenus_(ui) {
  ui.createMenu('1. Setup')
    .addItem('Build Settings Tab', 'buildSettingsTab')
    .addItem('Build Student Roster', 'buildStudentRoster')
    .addItem('Build Scores Progress', 'buildScoresProgress')
    .addItem('Build Dashboard Shell', 'buildDashboardShell')
    .addItem('Build Email Log', 'buildEmailLogTab')
    .addItem('Build Teacher Roster', 'buildTeacherRoster')
    .addItem('Build Remedial Attendance (Teachers)', 'buildRemedialAttendance')
    .addItem('Build MoM Log', 'buildMoMLog')
    .addItem('Build HoD Directory', 'buildHodDirectory')
    .addToUi();

  ui.createMenu('2. Teachers & Principals')
    .addItem('Create Teacher Workbooks', 'createTeacherWorkbooks')
    .addItem('Fill Teacher Folder Links', 'fillTeacherFolderLinks')
    .addItem('Refresh Teacher Roster (if roster changed)', 'refreshTeacherRoster')
    .addItem('Sync Teacher Profiles + Refresh Scores Progress', 'syncTeacherProfilesAndScores')
    .addSeparator()
    .addItem('Send: Workbook + Guide (all teachers)…', 'sendTeacherWorkbookLinksAndGuide')
    .addItem('Send: Workbook + Guide (selected teacher)…', 'sendWorkbookAndGuideToSelectedTeacher')
    .addSeparator()
    .addItem('Send: Principal Workbooks…', 'sendPrincipalWorkbooks')
    .addToUi();

  ui.createMenu('3. Audit & HoD Communication')
    .addItem('1. Send Department Workbooks Status to HoDs…', 'sendDepartmentStatusToHods')
    .addSeparator()
    .addItem('2. Audit Week 1 Tab', 'auditWeek1')
    .addItem('3. Send Week 1 Missing Entries to HoDs (deadline)…', 'sendWeek1MissingToHods')
    .addSeparator()
    .addItem('4. Audit Week 5 + Progress Report Entries Tabs', 'auditWeek5AndPR')
    .addItem('5. Send Week 5 + Progress Report Missing Entries to HoDs (deadline)…', 'sendWeek5MissingToHods')
    .addSeparator()
    .addItem('6. Audit Week 10 Tab', 'auditWeek10')
    .addItem('7. Send Week 10 Missing Entries to HoDs (deadline)…', 'sendWeek10MissingToHods')
    .addSeparator()
    .addItem('Check Action Plans vs Template', 'checkActionPlanTemplateAlignment')
    .addItem('Reset Template Check (start over)', 'resetTemplateCheckProgress')
    .addSeparator()
    .addItem('8. Send Thank-You for Flawless Tabs…', 'sendFlawlessThankYous')
    .addToUi();

  ui.createMenu('4. MoM Log')
    .addItem('Generate MoM — Round 1 (mark attendance first)', 'generateMoMRound1')
    .addItem('Send MoM — Round 1', 'openSendMoMDialogRound1')
    .addItem('Generate MoM — Round 2 (mark attendance first)', 'generateMoMRound2')
    .addItem('Send MoM — Round 2', 'openSendMoMDialogRound2')
    .addItem('Generate MoM — Round 3 (mark attendance first)', 'generateMoMRound3')
    .addItem('Send MoM — Round 3', 'openSendMoMDialogRound3')
    .addToUi();

  ui.createMenu('5. Send Deadlines')
    .addItem('Send Deadline to Teachers (CC HoDs + Principals)…', 'sendDeadlinesToTeachers')
    .addToUi();

  ui.createMenu('6. Student Roster & Probation')
    .addItem('Upload Probation Letter (selected row)', 'openProbationLetterUpload')
    .addSeparator()
    .addItem('Protect Student Data (owner-only)', 'protectStudentData')
    .addToUi();

  ui.createMenu('7. System & End of Term')
    .addItem('Install Periodic Sync (run once)', 'installPeriodicSync')
    .addItem('Recovery: Delete All Triggers', 'deleteAllTriggers')
    .addSeparator()
    .addItem('Generate Division Reports (Week 11 only)', 'generateDivisionReports')
    .addItem('Archive Trimester...', 'archiveTrimester')
    .addToUi();
}

function refreshTeacherRoster() { populateTeacherRosterFromStudentRoster_(); }

function syncTeacherProfilesAndScores() {
  syncTeacherProfiles();
  refreshScoresProgress();
  SpreadsheetApp.getActiveSpreadsheet().toast('Teacher profiles synced and Scores Progress refreshed.', 'Done', 6);
}

// ================================================================ Email preview (every send goes through here)

/**
 * items: [{ to, cc, subject, html, type, related, attachIds:[fileId], marks:[{sheet,row,statusCol,dateCol}] }]
 * Opens a preview window listing every email. Nothing is sent until you click "Send selected".
 */
function previewEmails_(title, items, emptyMsg) {
  const ui = SpreadsheetApp.getUi();
  if (!items.length) { ui.alert(emptyMsg || 'Nothing to send.'); return; }
  items.forEach(it => {
    it.attachNames = (it.attachIds || []).map(id => { try { return DriveApp.getFileById(id).getName(); } catch (e) { return '(file not found)'; } });
  });
  const id = 'EMQ_' + Utilities.getUuid();
  const json = JSON.stringify(items);
  const cache = CacheService.getUserCache();
  const CH = 90000, parts = Math.ceil(json.length / CH);
  const put = { [id + '_n']: String(parts) };
  for (let i = 0; i < parts; i++) put[id + '_' + i] = json.slice(i * CH, (i + 1) * CH);
  cache.putAll(put, 21600);

  const data = json.replace(/</g, '\\u003c').replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
  const html = `<!DOCTYPE html><html><head><base target="_blank"><style>
    body{font-family:Arial,sans-serif;margin:0;color:#333}
    .wrap{display:flex;height:600px}
    .list{width:300px;border-right:1px solid #ddd;overflow:auto;background:#fafafa}
    .item{padding:8px 10px;border-bottom:1px solid #eee;cursor:pointer;font-size:12px;display:flex;gap:6px}
    .item.active{background:#cfe2f3}.item b{display:block}
    .main{flex:1;overflow:auto;padding:12px 16px}
    .meta{font-size:13px;background:#f3f3f3;padding:8px 10px;border-radius:4px;margin-bottom:12px;line-height:1.6}
    .bar{padding:10px 14px;border-top:1px solid #ddd;display:flex;gap:10px;align-items:center;font-size:13px}
    button{background:#1c4587;color:#fff;border:0;padding:8px 16px;border-radius:4px;cursor:pointer;font-size:13px}
    button.sec{background:#888}button:disabled{opacity:.5;cursor:default}
  </style></head><body>
  <div class="wrap"><div class="list" id="list"></div><div class="main" id="main"></div></div>
  <div class="bar" id="bar"><label><input type="checkbox" id="all" checked> Select all</label>
    <span id="count"></span><span style="flex:1"></span>
    <button class="sec" onclick="google.script.host.close()">Cancel</button>
    <button class="sec" id="test" style="background:#0b5394">Send this one to me (test)</button>
    <button id="send">Send selected</button></div>
  <script>
    const Q = ${data}; const ID = '${id}'; const sel = Q.map(() => true); let cur = 0;
    const esc = s => String(s || '').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
    function count(){ const n = sel.filter(Boolean).length; document.getElementById('count').textContent = n + ' of ' + Q.length + ' selected'; document.getElementById('send').disabled = !n; }
    function list(){ document.getElementById('list').innerHTML = Q.map((q,i) =>
      '<div class="item'+(i===cur?' active':'')+'" onclick="show('+i+')"><input type="checkbox" '+(sel[i]?'checked':'')+' onclick="event.stopPropagation();sel['+i+']=this.checked;count()">'+
      '<div><b>'+esc(q.to)+'</b>'+esc(q.subject)+'</div></div>').join(''); }
    function show(i){ cur = i; list(); const q = Q[i];
      document.getElementById('main').innerHTML = '<div class="meta"><b>To:</b> '+esc(q.to)+(q.cc?'<br><b>CC:</b> '+esc(q.cc):'')+
        '<br><b>Subject:</b> '+esc(q.subject)+((q.attachNames||[]).length?'<br><b>Attachment:</b> '+esc(q.attachNames.join(', ')):'')+
        '<br><span style="color:#888">Email '+(i+1)+' of '+Q.length+'</span></div>' + q.html; }
    document.getElementById('all').onchange = e => { sel.fill(e.target.checked); list(); count(); };
    document.getElementById('send').onclick = () => {
      const idx = sel.map((s,i) => s ? i : -1).filter(i => i >= 0);
      if (!confirm('Send ' + idx.length + ' email(s) now?')) return;
      document.getElementById('send').disabled = true; document.getElementById('send').textContent = 'Sending…';
      google.script.run.withSuccessHandler(r => {
        document.getElementById('main').innerHTML = '<h3>Done</h3><p>' + r.sent + ' email(s) sent and logged in Email Log.</p>' +
          (r.failed.length ? '<p style="color:#990000">Failed:<br>' + r.failed.map(esc).join('<br>') + '</p>' : '');
        document.getElementById('bar').innerHTML = '<span style="flex:1"></span><button onclick="google.script.host.close()">Close</button>';
      }).withFailureHandler(e => { alert(e.message); document.getElementById('send').disabled = false; document.getElementById('send').textContent = 'Send selected'; })
        .sendPreviewedEmails(ID, idx);
    };
    document.getElementById('test').onclick = () => {
      const b = document.getElementById('test'); b.disabled = true; b.textContent = 'Sending test…';
      google.script.run.withSuccessHandler(to => { b.disabled = false; b.textContent = 'Send this one to me (test)';
          alert('Test copy of email ' + (cur + 1) + ' sent to ' + to + '. Nobody else received it, and nothing was marked as sent.'); })
        .withFailureHandler(e => { b.disabled = false; b.textContent = 'Send this one to me (test)'; alert(e.message); })
        .sendPreviewTestEmail(ID, cur);
    };
    show(0); count();
  </script></body></html>`;
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(1050).setHeight(680), title + ' — preview (' + items.length + ')');
}

/** Called from the preview window: sends ONE email only to you, marked [TEST]. Does not mark anything as sent. */
function sendPreviewTestEmail(id, idx) {
  const cache = CacheService.getUserCache();
  const n = parseInt(cache.get(id + '_n') || '0', 10);
  if (!n) throw new Error('This preview has expired or was already sent. Please run the menu item again.');
  let json = '';
  for (let i = 0; i < n; i++) json += cache.get(id + '_' + i) || '';
  const it = JSON.parse(json)[idx];
  if (!it) throw new Error('Email not found.');
  const me = Session.getEffectiveUser().getEmail();
  const banner = `<div style="font-family:Arial,sans-serif;background:#fff2cc;border:1px solid #f1c232;padding:8px 12px;margin-bottom:12px;font-size:13px;">
    <b>TEST COPY</b> — this would go to: <b>${auditEsc_(it.to)}</b>${it.cc ? ' &nbsp;|&nbsp; CC: ' + auditEsc_(it.cc) : ''}</div>`;
  const options = { htmlBody: banner + it.html, name: "Academic Dean's Office" };
  if (it.attachIds && it.attachIds.length) {
    options.attachments = it.attachIds.map(fid => { try { return DriveApp.getFileById(fid).getBlob(); } catch (e) { return null; } }).filter(Boolean);
  }
  GmailApp.sendEmail(me, '[TEST] ' + it.subject, '', options);
  return me;
}

/** Called from the preview window. Sends the selected emails, logs them, and marks review rows as Sent. */
function sendPreviewedEmails(id, idxs) {
  const cache = CacheService.getUserCache();
  const n = parseInt(cache.get(id + '_n') || '0', 10);
  if (!n) throw new Error('This preview has expired or was already sent. Please run the menu item again.');
  let json = '';
  for (let i = 0; i < n; i++) json += cache.get(id + '_' + i) || '';
  const keys = [id + '_n']; for (let i = 0; i < n; i++) keys.push(id + '_' + i);
  cache.removeAll(keys); // one-shot: prevents double sending
  const items = JSON.parse(json);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  let sent = 0; const failed = [];
  const toSort = {};
  (idxs || []).forEach(i => {
    const it = items[i];
    if (!it) return;
    const options = { htmlBody: it.html, name: "Academic Dean's Office" };
    if (it.cc) options.cc = it.cc;
    if (it.attachIds && it.attachIds.length) {
      options.attachments = it.attachIds.map(fid => { try { return DriveApp.getFileById(fid).getBlob(); } catch (e) { return null; } }).filter(Boolean);
    }
    let status = 'Sent', error = '';
    try { GmailApp.sendEmail(it.to, it.subject, '', options); sent++; }
    catch (e) { status = 'Failed'; error = e.message; failed.push(`${it.to}: ${e.message}`); }
    logEmailOnly_(it.to + (it.cc ? ' (cc: ' + it.cc + ')' : ''), it.subject, it.type, it.related, status, error);
    if (status === 'Sent') (it.marks || []).forEach(m => {
      const sh = ss.getSheetByName(m.sheet);
      if (!sh) return;
      sh.getRange(m.row, m.statusCol).setValue('Sent');
      if (m.dateCol) sh.getRange(m.row, m.dateCol).setValue(now).setNumberFormat('dd/MM/yyyy');
      if (m.sendCol) sh.getRange(m.row, m.sendCol).setValue(false);   // untick "Send to HoD?"
      if (m.sortSent) toSort[m.sheet] = m.statusCol;
    });
  });
  // Sent rows turn green and drop below the unsent ones (unsent stay at the top).
  Object.keys(toSort).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) moveSentRowsToBottom_(sh, toSort[name]);
  });
  return { sent: sent, failed: failed };
}

/** Moves every "Sent" row to the bottom of the sheet (keeping their order) and makes sure Sent rows show green. */
function moveSentRowsToBottom_(sh, statusCol) {
  const last = sh.getLastRow();
  if (last >= 3) {
    const status = sh.getRange(2, statusCol, last - 1, 1).getValues().map(r => r[0]);
    const sentRows = [];
    status.forEach((v, i) => { if (v === 'Sent') sentRows.push(i + 2); });
    const alreadyAtBottom = sentRows.every((r, i) => r === last - sentRows.length + 1 + i);
    if (!alreadyAtBottom) {
      let moved = 0;
      sentRows.forEach(r => {
        const cur = r - moved;
        sh.moveRows(sh.getRange(cur + ':' + cur), last + 1);
        moved++;
      });
    }
  }
  ensureSentRowRule_(sh, statusCol);
}

/** Conditional-format rule (placed first so it wins): whole row green when Status = "Sent". */
function ensureSentRowRule_(sh, statusCol) {
  const colL = String.fromCharCode(64 + statusCol);
  const formula = '=$' + colL + '2="Sent"';
  const range = sh.getRange(2, 1, Math.max(sh.getMaxRows() - 1, 1), Math.max(sh.getLastColumn(), statusCol + 1));
  const rules = sh.getConditionalFormatRules().filter(r => {
    const c = r.getBooleanCondition();
    return !(c && c.getCriteriaType() === SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA && c.getCriteriaValues()[0] === formula);
  });
  rules.unshift(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formula)
    .setBackground(THEME.statusGreen).setRanges([range]).build());
  sh.setConditionalFormatRules(rules);
}

// ---------------------------------------------------------------- shared helpers

function auditEsc_(s) {
  return String(s === null || s === undefined ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function auditNorm_(v) {
  return String(v === null || v === undefined ? '' : v).replace(/\s+/g, ' ').trim().toLowerCase();
}

/** True if a cell is empty or only holds a placeholder such as "N/A" or "not spotted yet". */
function isBlankEntry_(v) {
  if (v === null || v === undefined) return true;
  if (v instanceof Date) return false;
  if (typeof v === 'number') return false;
  const s = auditNorm_(v);
  return s === '' || AP_BLANK_EQUIVALENTS.indexOf(s) !== -1;
}

/** Same rule as isBlankEntry_, as an RE2 pattern for use inside sheet formulas. */
function blankPatternRe2_() {
  const items = AP_BLANK_EQUIVALENTS.map(s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
  return '^(' + items.join('|') + ')?$';
}

/** "4" + "B" -> "4B"; also copes with classes already stored as "4B". */
function gradeClassLabel_(grade, cls) {
  const letter = String(cls === null || cls === undefined ? '' : cls).trim().replace(/^\d+\s*/, '');
  const g = String(grade === null || grade === undefined ? '' : grade).trim();
  return g ? `${g}${letter}` : letter;
}

function auditCurrentWeek_() {
  const settings = readSettingsMap_();
  const termStart = TERM_START_DATES[String(settings['Term # (1, 2, or 3)'])];
  const numWeeks = parseInt(settings['Number of Weeks in Term'], 10) || 10;
  if (!termStart) return 1;
  return Math.max(1, Math.min(numWeeks, Math.floor((new Date() - termStart) / (7 * 86400000)) + 1));
}

function teacherMatches_(teacherField, teacherName) {
  return !!teacherField && !!teacherName &&
    String(teacherField).trim().indexOf(String(teacherName).trim()) === 0;
}

function auditTeacherList_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.TEACHER_ROSTER);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  return sh.getRange(2, 1, lastRow - 1, TEACHER_ROSTER_HEADERS.length).getValues()
    .filter(r => String(r[0]).trim())
    .map(r => ({
      name: String(r[0]).trim(),
      email: String(r[TEACHER_ROSTER_COL.EMAIL - 1] || '').trim(),
      subject: String(r[TEACHER_ROSTER_COL.SUBJECT - 1] || '').trim(),
      classes: String(r[TEACHER_ROSTER_COL.CLASSES - 1] || '').trim(),
      workbookUrl: String(r[TEACHER_ROSTER_COL.WORKBOOK_LINK - 1] || '').trim()
    }));
}

function rosterIndex_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ROSTER);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  return sh.getRange(2, 1, lastRow - 1, ROSTER_HEADERS.length).getValues()
    .filter(r => String(r[ROSTER_COL.ID - 1]).trim())
    .map(r => ({
      id: String(r[ROSTER_COL.ID - 1]).trim(),
      name: r[ROSTER_COL.NAME - 1],
      grade: r[ROSTER_COL.GRADE - 1],
      cls: r[ROSTER_COL.CLASS - 1],
      subject: String(r[ROSTER_COL.SUBJECT - 1] || '').trim(),
      teacherField: r[ROSTER_COL.TEACHER - 1],
      apLink: String(r[ROSTER_COL.ACTION_PLAN - 1] || '').trim(),
      division: divisionFromGrade_(r[ROSTER_COL.GRADE - 1]) || ''
    }));
}

function openWorkbookSafe_(url) {
  if (!url) return null;
  try { return SpreadsheetApp.openById(extractFileIdFromUrl_(url)); } catch (e) {}
  try { return SpreadsheetApp.openByUrl(url); } catch (e) {}
  return null;
}

/** Reads a teacher tab into {sh, headers, rows:[{rowNum, v}]} (only rows with a Student ID). */
function readTeacherTab_(ss, tabName) {
  const sh = ss.getSheetByName(tabName);
  if (!sh) return null;
  const headerRow = tabName === 'Progress Report Entries' ? 1 : 3;
  const lastRow = sh.getLastRow(), lastCol = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const rows = lastRow > headerRow
    ? sh.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues()
      .map((v, i) => ({ rowNum: headerRow + 1 + i, v: v }))
      .filter(r => String(r.v[0]).trim() !== '')
    : [];
  return { sh: sh, headers: headers, rows: rows };
}

/** Columns that must be filled before a row may be generated / counted as complete. */
function requiredColsFor_(tabName) {
  if (tabName === 'Week 1') {
    const c = WEEK1_COLS;
    return [c.WEAK1, c.STEPS1, c.WEAK2, c.STEPS2, c.MEANS, c.EVIDENCE];
  }
  if (tabName === 'Week 5') {
    const c = WEEK5_COLS;
    return [c.CUMULATIVE_1, c.QZ1, c.QZ2, c.WRITTEN_SUMMARY, c.WEAK1_AP2, c.STEPS1_AP2,
      c.WEAK2_AP2, c.STEPS2_AP2, c.MEANS_AP2, c.EVIDENCE_AP2];
  }
  if (tabName === 'Week 10') {
    const c = WEEK10_COLS;
    return [c.CUMULATIVE_2, c.QZ3, c.QZ4, c.WRITTEN_SUMMARY];
  }
  if (tabName === 'Progress Report Entries') return PR_STATEMENTS.map((_, i) => 3 + i);
  return [];
}

/** Has the output for this row been produced? (link generated / progress report Done) */
function rowOutputDone_(tabName, v) {
  if (tabName === 'Week 1') return String(v[WEEK1_COLS.ACTION_PLAN_LINK - 1]).trim() !== '';
  if (tabName === 'Week 5') return String(v[WEEK5_COLS.ACTION_PLAN_LINK - 1]).trim() !== '';
  if (tabName === 'Week 10') return String(v[WEEK10_COLS.ACTION_PLAN_PDF_LINK - 1]).trim() !== '';
  if (tabName === 'Progress Report Entries') return String(v[3 + PR_STATEMENTS.length - 1]).trim() === 'Done';
  return false;
}

/** Evaluates one row: which required fields are empty, and which only hold a placeholder. */
function evaluateRow_(tabName, headers, v) {
  const cols = requiredColsFor_(tabName);
  const label = c => headers[c - 1] || ('Column ' + c);
  const bad = cols.filter(c => isBlankEntry_(v[c - 1]));
  const empty = bad.filter(c => String(v[c - 1] === null || v[c - 1] === undefined ? '' : v[c - 1]).trim() === '').map(label);
  const placeholders = bad.filter(c => String(v[c - 1] === null || v[c - 1] === undefined ? '' : v[c - 1]).trim() !== '')
    .map(c => ({ label: label(c), text: String(v[c - 1]).trim() }));
  return { missing: bad.map(label), empty: empty, placeholders: placeholders,
    fullyEmpty: bad.length === cols.length, total: cols.length };
}

/** "Empty: A, B | Rewrite (not a real entry): C — “not spotted yet”" */
function formatMissing_(ev) {
  const parts = [];
  if (ev.empty.length) parts.push('Empty: ' + ev.empty.join(', '));
  if (ev.placeholders.length) parts.push('Rewrite (not a real entry): ' +
    ev.placeholders.map(p => `${p.label} — “${p.text}”`).join('; '));
  return parts.join(' | ');
}

/** Same text for emails: each part on its own line, labels in bold. */
function missingHtml_(s) {
  return auditEsc_(s).split(' | ').map(p =>
    p.replace(/^(Empty:|Rewrite \(not a real entry\):)/, '<b>$1</b>')).join('<br>');
}

const PLACEHOLDER_NOTE_HTML =
  '<p style="margin-top:16px;"><b>Please note:</b> entries such as “not spotted yet”, “not spotted”, “N/A”, “none” or “-” ' +
  'are <b>not real entries</b>. They must be rewritten with the actual weakness, the specific steps to tackle it, and how ' +
  'improvement will be assessed — otherwise the action plan cannot be generated.</p>';

// ---------------------------------------------------------------- HoD directory

function buildHodDirectory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(AUDIT_SHEETS.HOD);
  if (sh && sh.getLastRow() > 1) {
    SpreadsheetApp.getUi().alert('HoD Directory already exists — edit it directly. (Nothing was overwritten.)');
    return;
  }
  if (!sh) sh = ss.insertSheet(AUDIT_SHEETS.HOD);
  sh.clear();
  sh.getRange(1, 1, 1, HOD_HEADERS.length).setValues([HOD_HEADERS])
    .setFontWeight('bold').setBackground(THEME.darkBlue).setFontColor(THEME.white).setHorizontalAlignment('center');
  sh.getRange(2, 1, HOD_SEED.length, HOD_HEADERS.length).setValues(HOD_SEED);
  sh.getRange(2, 5, 60, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No']).build());
  sh.setFrozenRows(1);
  [180, 180, 200, 260, 120].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange(1, 1, HOD_SEED.length + 1, HOD_HEADERS.length).setBorder(true, true, true, true, true, true);
  sh.getRange('G1').setValue('Subject must match the Student Roster spelling (e.g. "Mathematics"). ' +
    'Division: ES, MS, HS — comma-separated — or All. Rows with no email are skipped when sending.')
    .setFontStyle('italic').setFontColor('#666666');
  ss.toast('HoD Directory created — please fill in the missing emails.', 'Done', 6);
}

function readHodDirectory_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AUDIT_SHEETS.HOD);
  if (!sh || sh.getLastRow() < 2) {
    throw new Error('HoD Directory is missing — run "9. Audits & HoD Reports > Build HoD Directory" first.');
  }
  return sh.getRange(2, 1, sh.getLastRow() - 1, HOD_HEADERS.length).getValues()
    .filter(r => String(r[0]).trim() && String(r[2]).trim())
    .map(r => ({
      subject: auditNorm_(r[0]),
      divisions: String(r[1] || 'All').toUpperCase().split(/[,\/]/).map(s => s.trim()).filter(Boolean),
      name: String(r[2]).trim(),
      email: String(r[3] || '').trim(),
      active: String(r[4] || 'Yes').trim().toLowerCase() !== 'no'
    }));
}

function canonicalSubject_(s) {
  const n = auditNorm_(s);
  if (n === 'math' || n === 'maths') return 'mathematics';
  if (n === 'religion') return 'islamic religion';
  if (n === 'social studies (history)' || n === 'history') return 'social studies';
  return n;
}

function hodsFor_(hods, subject, division) {
  const subj = canonicalSubject_(subject);
  return hods.filter(h => h.active && canonicalSubject_(h.subject) === subj &&
    (h.divisions.indexOf('ALL') !== -1 || !division || h.divisions.indexOf(String(division).toUpperCase()) !== -1));
}

/** Teacher-side context: roster rows for this teacher + per-row routing (subject, division, HoDs). */
function teacherContext_(teacher, roster, hods) {
  const entries = roster.filter(e => teacherMatches_(e.teacherField, teacher.name));
  const occ = {};
  return {
    entries: entries,
    /** Call once per workbook row, in row order, per tab. */
    route: function (studentId, tabKey) {
      const key = tabKey + '|' + String(studentId).trim();
      const n = occ[key] = (occ[key] === undefined ? 0 : occ[key] + 1);
      const matches = entries.filter(e => e.id === String(studentId).trim());
      const e = matches[n] || matches[0] || null;
      const subject = e ? e.subject : teacher.subject.split(',')[0].trim();
      const division = e ? e.division : '';
      const hs = hodsFor_(hods, subject, division);
      return {
        subject: subject, division: division,
        cls: e ? gradeClassLabel_(e.grade, e.cls) : '',
        hodNames: hs.map(h => h.name).join(', '),
        hodEmails: hs.map(h => h.email).filter(Boolean).join(', ')
      };
    },
    /** All HoD groups this teacher reports to (for teacher-level issues). */
    allRoutes: function () {
      const seen = {}, out = [];
      (entries.length ? entries : [{ subject: teacher.subject.split(',')[0].trim(), division: '' }]).forEach(e => {
        const hs = hodsFor_(hods, e.subject, e.division);
        const r = {
          subject: e.subject, division: e.division,
          hodNames: hs.map(h => h.name).join(', '),
          hodEmails: hs.map(h => h.email).filter(Boolean).join(', ')
        };
        const k = r.hodNames + '|' + r.subject;
        if (!seen[k]) { seen[k] = 1; out.push(r); }
      });
      return out;
    }
  };
}

/** Writes a review table (header + rows) to a tab, replacing what was there. */
function writeReviewSheet_(name, headers, rows, sendCol) {
  const sh = getOrCreateSheet_(name);
  sh.clear();
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground(THEME.darkBlue).setFontColor(THEME.white)
    .setHorizontalAlignment('center').setWrap(true);
  sh.setFrozenRows(1);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows).setVerticalAlignment('top');
    if (sendCol) sh.getRange(2, sendCol, rows.length, 1).insertCheckboxes();
    sh.getRange(1, 1, rows.length + 1, headers.length).setBorder(true, true, true, true, true, true);
  }
  return sh;
}

/** HoD display name for an email, from the HoD Directory (falls back to the email). */
function hodNameByEmail_(email) {
  const h = readHodDirectory_().find(x => x.email.toLowerCase() === String(email).toLowerCase());
  return h ? h.name : email;
}

/** Splits "a@x, b@x" into an array. */
function splitEmails_(s) {
  return String(s || '').split(/[,;]/).map(x => x.trim()).filter(Boolean);
}

// ================================================================ Grades 1-2: no Exit Ticket

/** Grades with no Exit Ticket — ET is N/A in their workbooks and action plans, and left out of the graph. */
const NO_ET_GRADES = [1, 2];

function isNoETGrade_(grade) {
  return NO_ET_GRADES.indexOf(parseInt(grade, 10)) !== -1;
}

/** True when the ET values for a 5-week block are all N/A (or blank with at least one N/A). */
function etNotApplicable_(weeklyData) {
  const vals = (weeklyData || []).map(d => (d && d.et !== undefined && d.et !== null) ? String(d.et).trim().toUpperCase() : '');
  return vals.some(v => v === 'N/A') && vals.every(v => v === 'N/A' || v === '');
}

// ================================================================ Generation guard

const MANDATORY_NOTE_MARK = '⚠️ Mandatory field';

/**
 * Called at the top of each generate handler in ACTIONPLANFILL.gs (Week 1 / 5 / 10).
 * If any required field is empty or only a placeholder, those cells turn RED with a
 * note telling the teacher it is a mandatory field, and nothing is generated.
 * Cells that have since been fixed lose the red + note automatically.
 */
function assertRowComplete_(sh, row, tabName) {
  const headerRow = 3;
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const rowRange = sh.getRange(row, 1, 1, lastCol);
  const v = rowRange.getValues()[0];
  const notes = rowRange.getNotes();
  const bgs = rowRange.getBackgrounds();
  const res = evaluateRow_(tabName, headers, v);

  requiredColsFor_(tabName).forEach(c => {
    const i = c - 1;
    const raw = String(v[i] === null || v[i] === undefined ? '' : v[i]).trim();
    if (isBlankEntry_(v[i])) {
      notes[0][i] = raw
        ? `${MANDATORY_NOTE_MARK} — "${raw}" is not an acceptable entry. Please write the actual ${headers[i]}, then tick Generate? again.`
        : `${MANDATORY_NOTE_MARK} — please fill in ${headers[i]}, then tick Generate? again.`;
      bgs[0][i] = THEME.statusRed;
    } else if (String(notes[0][i]).indexOf(MANDATORY_NOTE_MARK) === 0) {
      notes[0][i] = '';
      bgs[0][i] = null;
    }
  });
  rowRange.setNotes(notes);
  rowRange.setBackgrounds(bgs);

  if (res.missing.length) {
    throw new Error(`Not generated — ${formatMissing_(res)}. ` +
      `These are mandatory fields (see the red cells). Placeholders such as "N/A", "-" or "not spotted yet" are not acceptable. Fix them, then tick Generate? again.`);
  }
}

/** Groups sorted column numbers into contiguous [start, end] runs. */
function contiguousGroups_(cols) {
  const sorted = cols.slice().sort((a, b) => a - b);
  const groups = [];
  sorted.forEach(c => {
    const last = groups[groups.length - 1];
    if (last && c === last[1] + 1) last[1] = c; else groups.push([c, c]);
  });
  return groups;
}

/**
 * Live Status column for a Week 1 / 5 / 10 tab ("⚠️ 3 field(s) empty — fill before ticking").
 * Called automatically when a teacher workbook tab is built. Cells only turn red when the
 * teacher ticks Generate? with a mandatory field missing (see assertRowComplete_).
 */
function applyGuardToTab_(sh, tab, pattern) {
  if (!sh) return;
  const c = tab === 'Week 1' ? WEEK1_COLS : tab === 'Week 5' ? WEEK5_COLS : WEEK10_COLS;
  const linkCol = tab === 'Week 10' ? c.ACTION_PLAN_PDF_LINK : c.ACTION_PLAN_LINK;
  const L = columnToLetter_;
  const groups = contiguousGroups_(requiredColsFor_(tab));
  const lastRow = sh.getLastRow();

  if (lastRow >= 4) {
    const formulas = [];
    for (let r = 4; r <= lastRow; r++) {
      const miss = '(' + groups.map(g =>
        `SUMPRODUCT(--REGEXMATCH(LOWER(TRIM(${L(g[0])}${r}:${L(g[1])}${r}&"")),"${pattern}"))`).join('+') + ')';
      const link = `${L(linkCol)}${r}`, gen = `${L(c.GENERATE)}${r}`;
      formulas.push([
        `=IF($A${r}="","",IF(${miss}>0,` +
        `IF(${link}<>"","⚠️ Generated with "&${miss}&" empty field(s) — complete & tick again",` +
        `"⚠️ "&${miss}&" mandatory field(s) still empty"),` +
        `IF(${link}<>"","✅ Done",IF(${gen}=TRUE,"⏳ Generating... (usually within 1 min)",""))))`
      ]);
    }
    sh.getRange(4, c.STATUS, formulas.length, 1).setFormulas(formulas);
  }

  // Drop any earlier always-on red rules; keep only the orange highlight on ⚠️ status text.
  const rules = sh.getConditionalFormatRules().filter(rule => {
    const bc = rule.getBooleanCondition();
    if (!bc) return true;
    const v0 = String((bc.getCriteriaValues() || [])[0] || '');
    if (v0.indexOf(pattern) !== -1) return false;
    if (bc.getCriteriaType() === SpreadsheetApp.BooleanCriteria.TEXT_STARTS_WITH && v0 === '⚠️') return false;
    return true;
  });
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextStartsWith('⚠️').setBackground(THEME.statusOrange)
    .setRanges([sh.getRange(4, c.STATUS, Math.max(lastRow, 100) - 3, 1)]).build());
  sh.setConditionalFormatRules(rules);
}


// ================================================================ Audits per tab group

const TD_ = 'border:1px solid #d0d0d0;padding:5px 8px;text-align:left;vertical-align:top;';
const TH_ROW_ = 'background:#cfe2f3;';

const AUDIT_HEADERS = ['Run Date', 'Teacher', 'Subject', 'Division', 'Tab', 'Student ID', 'Student Name',
  'Grade & Class', 'Issue', 'Details', 'Output Link', 'Teacher Workbook', 'HoD(s)', 'HoD Email(s)',
  'Send to HoD?', 'Status', 'Sent Date'];
const AUDIT_COL = { TEACHER: 2, SUBJECT: 3, DIVISION: 4, TAB: 5, ID: 6, NAME: 7, CLS: 8, ISSUE: 9, DETAILS: 10,
  LINK: 11, WORKBOOK: 12, HODS: 13, HOD_EMAILS: 14, SEND: 15, STATUS: 16, SENT_DATE: 17 };

const ISSUE = {
  NO_WORKBOOK: 'Workbook not created',
  CANT_OPEN: "Workbook can't be opened",
  TAB_MISSING: 'Tab missing',
  NOT_STARTED: 'Tab not started (nothing entered)',
  MISSING: 'Missing or placeholder entries',
  GAPS_GENERATED: 'Generated with missing entries — must be regenerated',
  GAPS_DONE: 'Marked Done with missing entries',
  OK: '✅ Complete'
};

/** What "not finished" means per tab when all fields are filled. */
function notDoneIssue_(tab) {
  if (tab === 'Week 1') return 'Complete but action plan not generated';
  if (tab === 'Week 5') return 'Complete but Action Plan 2 not generated';
  if (tab === 'Week 10') return 'Complete but closeout PDF not generated';
  return 'Complete but not marked “Done”';
}

/** Output link / status for a row, for the audit sheet. */
function rowOutput_(tab, v) {
  if (tab === 'Week 1') return String(v[WEEK1_COLS.ACTION_PLAN_LINK - 1] || '');
  if (tab === 'Week 5') return String(v[WEEK5_COLS.ACTION_PLAN_LINK - 1] || '');
  if (tab === 'Week 10') return String(v[WEEK10_COLS.ACTION_PLAN_PDF_LINK - 1] || '');
  return String(v[3 + PR_STATEMENTS.length - 1] || '');
}

/**
 * Evaluates one teacher's workbook for a tab group.
 * Returns { state: 'none'|'notStarted'|'gaps'|'complete', rows: [ {tab, v, ev, route, issue} ], tabIssues: [{tab, issue}] }
 */
function evaluateTeacherPhase_(t, ctx, phase) {
  const res = { state: 'complete', rows: [], tabIssues: [], studentCount: 0 };
  if (!t.workbookUrl) { res.state = 'none'; res.tabIssues.push({ tab: '—', issue: ISSUE.NO_WORKBOOK, detail: 'No link in Teacher Roster column H' }); return res; }
  const ss = openWorkbookSafe_(t.workbookUrl);
  if (!ss) { res.state = 'none'; res.tabIssues.push({ tab: '—', issue: ISSUE.CANT_OPEN, detail: 'Check the link / sharing' }); return res; }

  let anyStarted = false, anyIssue = false;
  PHASES[phase].tabs.forEach(tabName => {
    const tab = readTeacherTab_(ss, tabName);
    if (!tab) { res.tabIssues.push({ tab: tabName, issue: ISSUE.TAB_MISSING, detail: '' }); anyIssue = true; return; }
    if (!tab.rows.length) { res.tabIssues.push({ tab: tabName, issue: ISSUE.NOT_STARTED, detail: 'No students listed' }); anyIssue = true; return; }
    res.studentCount = Math.max(res.studentCount, tab.rows.length);
    const evals = tab.rows.map(r => ({ tab: tabName, v: r.v, ev: evaluateRow_(tabName, tab.headers, r.v), route: ctx.route(r.v[0], tabName) }));
    const allEmpty = evals.every(x => x.ev.fullyEmpty);
    if (!allEmpty) anyStarted = true;
    evals.forEach(x => {
      const done = rowOutputDone_(tabName, x.v);
      if (allEmpty) x.issue = ISSUE.NOT_STARTED;
      else if (x.ev.missing.length && done) x.issue = tabName === 'Progress Report Entries' ? ISSUE.GAPS_DONE : ISSUE.GAPS_GENERATED;
      else if (x.ev.missing.length) x.issue = ISSUE.MISSING;
      else if (!done) x.issue = notDoneIssue_(tabName);
      else x.issue = '';
      if (x.issue) anyIssue = true;
      res.rows.push(x);
    });
  });
  res.state = !anyIssue ? 'complete' : (anyStarted ? 'gaps' : 'notStarted');
  return res;
}

function auditWeek1() { runPhaseAudit_(1); }
function auditWeek5AndPR() { runPhaseAudit_(2); }
function auditWeek10() { runPhaseAudit_(3); }

/** Opens every teacher workbook and lists every problem for the tab group on its review tab. */
function runPhaseAudit_(phase) {
  const P = PHASES[phase];
  const roster = rosterIndex_();
  const hods = readHodDirectory_();
  const teachers = auditTeacherList_();
  const today = new Date();
  const out = [];
  const groups = { none: [], notStarted: [], gaps: [], complete: [] };
  const start = Date.now();

  for (let i = 0; i < teachers.length; i++) {
    if (Date.now() - start > 5.2 * 60 * 1000) { SpreadsheetApp.getUi().alert('Stopped early to stay under the time limit — results so far are saved. Run again.'); break; }
    const t = teachers[i];
    const ctx = teacherContext_(t, roster, hods);
    const r = evaluateTeacherPhase_(t, ctx, phase);
    const onlyNotGenerated = r.state === 'gaps' && !r.tabIssues.length && r.rows.every(x => !x.issue || x.issue.indexOf('Complete but') === 0);
    groups[onlyNotGenerated ? 'complete' : r.state].push(t.name);
    r.tabIssues.forEach(ti => ctx.allRoutes().forEach(rt => out.push([today, t.name, rt.subject, rt.division, ti.tab, '', '',
      '', ti.issue, ti.detail, '', t.workbookUrl, rt.hodNames, rt.hodEmails, !!rt.hodEmails, '', ''])));
    r.rows.filter(x => x.issue && x.issue.indexOf('Complete but') !== 0).forEach(x => out.push([today, t.name, x.route.subject, x.route.division, x.tab, x.v[0], x.v[1],
      x.route.cls, x.issue, x.issue === ISSUE.NOT_STARTED ? '' : formatMissing_(x.ev), rowOutput_(x.tab, x.v), t.workbookUrl,
      x.route.hodNames, x.route.hodEmails, !!x.route.hodEmails, '', '']));
  }

  const sh = writeReviewSheet_(P.sheet, AUDIT_HEADERS, out, AUDIT_COL.SEND);
  if (out.length) {
    sh.getRange(2, 1, out.length, 1).setNumberFormat('dd/MM/yyyy');
    sh.getRange(2, AUDIT_COL.DETAILS, out.length, 1).setWrap(true);
    const ir = sh.getRange(2, AUDIT_COL.ISSUE, out.length, 1);
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('Tab').setBackground(THEME.statusRed).setRanges([ir]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('Workbook').setBackground(THEME.statusRed).setRanges([ir]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('Generated').setBackground(THEME.statusOrange).setRanges([ir]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('Marked').setBackground(THEME.statusOrange).setRanges([ir]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('Missing').setBackground(THEME.statusYellow).setRanges([ir]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('Complete').setBackground(THEME.statusPurple).setRanges([ir]).build()
    ]);
    ensureSentRowRule_(sh, AUDIT_COL.STATUS);
  }
  [90, 140, 110, 60, 120, 90, 200, 70, 220, 320, 150, 150, 150, 200, 70, 70, 90].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.activate();

  const noHod = out.filter(r => !r[AUDIT_COL.HOD_EMAILS - 1]).length;
  SpreadsheetApp.getUi().alert(
    `${P.label} audit — ${out.length} row(s) on "${P.sheet}".\n\n` +
    `NOT STARTED (${groups.notStarted.length + groups.none.length}): ${groups.notStarted.concat(groups.none).join(', ') || '—'}\n\n` +
    `MISSING / PLACEHOLDER ENTRIES (${groups.gaps.length}): ${groups.gaps.join(', ') || '—'}\n\n` +
    `COMPLETE (${groups.complete.length}): ${groups.complete.join(', ') || '—'}` +
    (noHod ? `\n\n⚠️ ${noHod} row(s) have no HoD email — complete the HoD Directory and re-run.` : '') +
    `\n\nUntick "Send to HoD?" for anything you want to hold back, then use the matching "Send … to HoDs" item.`);
}

function sendWeek1MissingToHods() { sendPhaseMissingToHods_(1); }
function sendWeek5MissingToHods() { sendPhaseMissingToHods_(2); }
function sendWeek10MissingToHods() { sendPhaseMissingToHods_(3); }

/** One email per HoD from the ticked, unsent rows of that tab group's audit sheet — shown in preview first. */
function sendPhaseMissingToHods_(phase) {
  const P = PHASES[phase];
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(P.sheet);
  if (!sh || sh.getLastRow() < 2) { ui.alert(`Run the ${P.label} audit first.`); return; }
  const deadline = askDeadline_(false);
  if (!deadline) return;

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, AUDIT_HEADERS.length).getValues();
  const byHod = {};
  const noEmail = new Set();
  data.forEach((r, i) => {
    if (r[AUDIT_COL.SEND - 1] !== true || r[AUDIT_COL.STATUS - 1] === 'Sent') return;
    const emails = splitEmails_(r[AUDIT_COL.HOD_EMAILS - 1]);
    if (!emails.length) { noEmail.add(r[AUDIT_COL.TEACHER - 1]); return; }
    emails.forEach(em => {
      const h = byHod[em] = byHod[em] || { name: hodNameByEmail_(em), teachers: {}, rowNums: [] };
      const tn = r[AUDIT_COL.TEACHER - 1];
      const tt = h.teachers[tn] = h.teachers[tn] || { subject: r[AUDIT_COL.SUBJECT - 1], workbook: r[AUDIT_COL.WORKBOOK - 1], rows: [] };
      tt.rows.push(r);
      h.rowNums.push(i + 2);
    });
  });

  const showTab = P.tabs.length > 1;
  const items = Object.keys(byHod).map(em => {
    const h = byHod[em];
    const sections = Object.keys(h.teachers).sort().map(tn => {
      const tt = h.teachers[tn];
      const teacherLevel = tt.rows.filter(r => !r[AUDIT_COL.ID - 1]);
      const studentRows = tt.rows.filter(r => r[AUDIT_COL.ID - 1]);
      let html = `<p style="margin:16px 0 6px;"><b>${auditEsc_(tn)}</b> — ${auditEsc_(tt.subject)}` +
        (tt.workbook ? ` &nbsp;<a href="${auditEsc_(tt.workbook)}">open workbook</a>` : '') + `</p>`;
      if (studentRows.length && studentRows.every(r => r[AUDIT_COL.ISSUE - 1] === ISSUE.NOT_STARTED)) {
        html += `<p style="margin:0 0 6px;color:#990000;"><b>Not started</b> — nothing has been entered for any of the ${new Set(studentRows.map(r => r[AUDIT_COL.ID - 1])).size} student(s).</p>`;
      }
      teacherLevel.forEach(r => { html += `<p style="margin:0 0 6px;color:#990000;"><b>${auditEsc_(r[AUDIT_COL.TAB - 1])}: ${auditEsc_(r[AUDIT_COL.ISSUE - 1])}</b> ${auditEsc_(r[AUDIT_COL.DETAILS - 1])}</p>`; });
      const listRows = studentRows.filter(r => r[AUDIT_COL.ISSUE - 1] !== ISSUE.NOT_STARTED);
      if (listRows.length) {
        html += `<table style="border-collapse:collapse;font-size:13px;width:100%;"><tr style="${TH_ROW_}">` +
          `<th style="${TD_}">Student</th><th style="${TD_}">Grade & Class</th>${showTab ? `<th style="${TD_}">Tab</th>` : ''}<th style="${TD_}">What is needed</th></tr>` +
          listRows.map(r => `<tr><td style="${TD_}">${auditEsc_(r[AUDIT_COL.NAME - 1])}</td><td style="${TD_}">${auditEsc_(r[AUDIT_COL.CLS - 1])}</td>` +
            (showTab ? `<td style="${TD_}">${auditEsc_(r[AUDIT_COL.TAB - 1])}</td>` : '') +
            `<td style="${TD_}">${auditEsc_(r[AUDIT_COL.ISSUE - 1])}${r[AUDIT_COL.DETAILS - 1] ? '<br>' + missingHtml_(r[AUDIT_COL.DETAILS - 1]) : ''}</td></tr>`).join('') +
          `</table>`;
      }
      return html;
    }).join('');
    const hasPlaceholder = Object.keys(h.teachers).some(tn => h.teachers[tn].rows.some(r => String(r[AUDIT_COL.DETAILS - 1]).indexOf('Rewrite') !== -1));
    const action = phase === 2
      ? 'every missing field is completed, the “Generate?” box is ticked again on Week 5, and each Progress Report is set to “Done”'
      : 'every missing field is completed and the “Generate?” box is ticked again so the plan is regenerated';
    const html = emailWrapper_(`${P.label} — Missing Entries`,
      `<p>Dear ${auditEsc_(h.name)},</p>
       <p>I have reviewed the remedial teacher workbooks for <b>${auditEsc_(P.label)}</b>. The items below in your department are incomplete.
       Please follow up so that ${action} by <b>${auditEsc_(deadline)}</b>.</p>
       ${sections}
       ${hasPlaceholder ? PLACEHOLDER_NOTE_HTML : ''}
       <p>Kind regards,</p>`);
    return { to: em, cc: '', subject: `${P.label} — Missing Entries (due ${deadline})`, html: html,
      type: `HoD — Missing Entries ${P.key}`, related: Object.keys(h.teachers).join(', '),
      marks: h.rowNums.map(n => ({ sheet: P.sheet, row: n, statusCol: AUDIT_COL.STATUS, dateCol: AUDIT_COL.SENT_DATE, sendCol: AUDIT_COL.SEND, sortSent: true })) };
  });

  if (noEmail.size) ui.alert(`No HoD email for: ${Array.from(noEmail).join(', ')} — these are not included. Fill the HoD Directory and re-run the audit to include them.`);
  previewEmails_(`${P.label} — HoD notices`, items, 'Nothing to send (no ticked, unsent rows with an HoD email).');
}

// ================================================================ 1. Department workbooks status → HoDs

const STATUS_SHEET = 'HoD Status Report';
const STATUS_HEADERS = ['HoD Name', 'HoD Email(s)', 'Teacher', 'Subject', 'Tab', 'Issue', 'Students / Details',
  'Teacher Workbook', 'Status', 'Sent Date'];

/**
 * Checks every teacher workbook (tabs due so far, or all tabs), writes the
 * "HoD Status Report" tab, and opens one email per HoD in the preview window:
 * a summary table of her department's teachers (with workbook links) and the details.
 */
function sendDepartmentStatusToHods() {
  const ui = SpreadsheetApp.getUi();
  const week = auditCurrentWeek_();
  const ans = ui.alert('Which tabs?',
    `It is Week ${week}.\n\nYES = check ALL tabs (Week 1, Week 5, Progress Report, Week 10), even ones not due yet.\n` +
    `NO = check only tabs that are due by Week ${week}.`, ui.ButtonSet.YES_NO_CANCEL);
  if (ans === ui.Button.CANCEL || ans === ui.Button.CLOSE) return;
  const includeAll = ans === ui.Button.YES;
  const phases = [1, 2, 3].filter(p => includeAll || p === 1 || (p === 2 && week >= 5) || (p === 3 && week >= 10));
  const deadline = askDeadline_(true);
  if (deadline === null) return;

  const roster = rosterIndex_();
  const hods = readHodDirectory_();
  const teachers = auditTeacherList_();
  const byHod = {};   // email -> {name, teachers:{tn:{subject, workbook, lines:[{tab, issue, details:[]}]}}}
  const sheetRows = [];
  const noHod = new Set();
  const start = Date.now();

  const addLine = (route, t, tab, issue, detail) => {
    sheetRows.push([route.hodNames || '(no HoD match)', route.hodEmails, t.name, route.subject, tab, issue, detail || '', t.workbookUrl, '', '']);
    const emails = splitEmails_(route.hodEmails);
    if (!emails.length) { noHod.add(t.name); return; }
    emails.forEach(em => {
      const h = byHod[em] = byHod[em] || { name: hodNameByEmail_(em), teachers: {} };
      const tt = h.teachers[t.name] = h.teachers[t.name] || { subject: route.subject, workbook: t.workbookUrl, lines: {} };
      const k = tab + '|' + issue;
      const ln = tt.lines[k] = tt.lines[k] || { tab: tab, issue: issue, details: [] };
      if (detail) ln.details.push(detail);
    });
  };

  for (let i = 0; i < teachers.length; i++) {
    if (Date.now() - start > 5 * 60 * 1000) { ui.alert('Stopped early (time limit) — the report is partial.'); break; }
    const t = teachers[i];
    const ctx = teacherContext_(t, roster, hods);
    const flagged = {};
    phases.forEach(p => {
      const r = evaluateTeacherPhase_(t, ctx, p);
      r.tabIssues.forEach(ti => ctx.allRoutes().forEach(rt => { addLine(rt, t, ti.tab, ti.issue, ti.detail); flagged[rt.hodEmails + rt.subject] = 1; }));
      r.rows.filter(x => x.issue).forEach(x => {
        const who = `${x.v[1]}${x.route.cls ? ' (' + x.route.cls + ')' : ''}`;
        addLine(x.route, t, x.tab, x.issue, x.issue === ISSUE.NOT_STARTED ? who : `${who}${x.ev.missing.length ? ': ' + formatMissing_(x.ev) : ''}`);
        flagged[x.route.hodEmails + x.route.subject] = 1;
      });
    });
    ctx.allRoutes().forEach(rt => { if (!flagged[rt.hodEmails + rt.subject]) addLine(rt, t, phases.map(p => PHASES[p].tabs.join(' + ')).join(', '), ISSUE.OK, ''); });
  }

  const sh = writeReviewSheet_(STATUS_SHEET, STATUS_HEADERS, sheetRows, null);
  if (sheetRows.length) sh.getRange(2, 7, sheetRows.length, 1).setWrap(true);
  [160, 220, 140, 120, 160, 260, 380, 160, 70, 90].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  const items = Object.keys(byHod).map(em => {
    const h = byHod[em];
    const names = Object.keys(h.teachers).sort();
    const summary = names.map(tn => {
      const tt = h.teachers[tn];
      const ok = Object.keys(tt.lines).every(k => tt.lines[k].issue === ISSUE.OK);
      return `<tr><td style="${TD_}">${auditEsc_(tn)}</td><td style="${TD_}">${auditEsc_(tt.subject)}</td>` +
        `<td style="${TD_}">${tt.workbook ? `<a href="${auditEsc_(tt.workbook)}">Open workbook</a>` : '—'}</td>` +
        `<td style="${TD_}background:${ok ? '#d9ead3' : '#fce5cd'};">${ok ? '✅ On track' : '⚠️ Action needed'}</td></tr>`;
    }).join('');
    const details = names.map(tn => {
      const lines = Object.keys(h.teachers[tn].lines).map(k => h.teachers[tn].lines[k]).filter(l => l.issue !== ISSUE.OK);
      if (!lines.length) return '';
      return `<p style="margin:16px 0 6px;"><b>${auditEsc_(tn)}</b></p>
        <table style="border-collapse:collapse;font-size:13px;width:100%;"><tr style="${TH_ROW_}">
        <th style="${TD_}">Tab</th><th style="${TD_}">Issue</th><th style="${TD_}">Students / details</th></tr>` +
        lines.map(l => `<tr><td style="${TD_}">${auditEsc_(l.tab)}</td><td style="${TD_}">${auditEsc_(l.issue)}</td>
          <td style="${TD_}">${l.details.map(missingHtml_).join('<br><br>')}</td></tr>`).join('') + '</table>';
    }).join('');
    const html = emailWrapper_(`Department Workbooks Status — Week ${week}`,
      `<p>Dear ${auditEsc_(h.name)},</p>
       <p>Here is the current status of the remedial teacher workbooks in your department
       (${auditEsc_(phases.map(p => PHASES[p].label).join('; '))}).</p>
       <table style="border-collapse:collapse;font-size:13px;width:100%;">
         <tr style="background:#1c4587;color:#fff;"><th style="${TD_}">Teacher</th><th style="${TD_}">Subject</th><th style="${TD_}">Workbook</th><th style="${TD_}">Status</th></tr>
         ${summary}</table>
       ${details ? '<p style="margin-top:18px;"><b>Details</b></p>' + details : ''}
       ${details.indexOf('Rewrite') !== -1 ? PLACEHOLDER_NOTE_HTML : ''}
       ${details ? `<p style="margin-top:16px;">Please follow up with the teachers concerned${deadline ? ` so these are completed by <b>${auditEsc_(deadline)}</b>` : ''}.</p>` : '<p>Everything due so far is complete — thank you.</p>'}
       <p>Kind regards,</p>`);
    return { to: em, cc: '', subject: `Department Workbooks Status — Week ${week}${deadline ? ' (due ' + deadline + ')' : ''}`,
      html: html, type: 'HoD — Department Status', related: names.join(', ') };
  });

  if (noHod.size) ui.alert(`No HoD email found for: ${Array.from(noHod).join(', ')} — not included. Complete the HoD Directory to include them.`);
  previewEmails_('Department workbooks status', items, 'No HoD emails to send — check the HoD Directory.');
}

// ================================================================ 8. Thank-you for flawless tabs

function auditGenericPhrases_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Generic Phrases');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => auditNorm_(r[0])).filter(Boolean);
}

/**
 * Thanks every teacher whose chosen tab group is flawless: every row complete (no
 * placeholders), every plan generated / progress report Done, no generic phrases,
 * and (Week 1) no docs flagged on "AP Template Check". Already-thanked teachers are skipped.
 */
function sendFlawlessThankYous() {
  const ui = SpreadsheetApp.getUi();
  const phase = askPhase_('Thank-you for flawless tabs');
  if (!phase) return;
  const P = PHASES[phase];
  const type = `Thank You — ${P.key}`;

  const badTpl = {};
  if (phase === 1) {
    const tplSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AUDIT_SHEETS.TEMPLATE);
    if (!tplSh || tplSh.getLastRow() < 2) {
      if (ui.alert('Template check not run', 'The "Check Action Plans vs Template" has not been run, so doc alignment can\'t be confirmed. Continue anyway?',
        ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
    } else {
      tplSh.getRange(2, 1, tplSh.getLastRow() - 1, TPL_HEADERS.length).getValues()
        .forEach(r => { if (r[7] === 'Not aligned' || r[7] === 'Cannot open') badTpl[r[0]] = true; });
    }
  }

  const thanked = {};
  const logSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EMAIL_LOG_SHEET);
  if (logSh && logSh.getLastRow() > 1) {
    logSh.getRange(2, 1, logSh.getLastRow() - 1, EMAIL_LOG_HEADERS.length).getValues().forEach(r => {
      if (r[5] === 'Sent' && (r[3] === type || (phase === 1 && r[3] === 'Action Plan Thank You'))) thanked[String(r[4]).trim()] = true;
    });
  }

  const phrases = auditGenericPhrases_();
  const roster = rosterIndex_();
  const hods = readHodDirectory_();
  const already = [];
  const items = [];

  auditTeacherList_().forEach(t => {
    if (badTpl[t.name] || !t.email) return;
    const ctx = teacherContext_(t, roster, hods);
    const r = evaluateTeacherPhase_(t, ctx, phase);
    if (r.state !== 'complete' || !r.rows.length) return;
    const generic = r.rows.some(x => textColsFor_(x.tab).some(c => {
      const txt = auditNorm_(x.v[c - 1]);
      return phrases.some(p => txt.indexOf(p) !== -1);
    }));
    if (generic) return;
    if (thanked[t.name]) { already.push(t.name); return; }
    const n = r.studentCount;
    const what = phase === 1
      ? `the Week 1 action plans for all ${n} of your remedial student${n > 1 ? 's' : ''}. Every entry is filled in and specific, and the plans have been generated in line with the template`
      : phase === 2
        ? `Week 5 — the Action Plan 1 reflection, Action Plan 2 and the Progress Report entries — for all ${n} of your remedial student${n > 1 ? 's' : ''}. Every field is complete and specific`
        : `the Week 10 closeout for all ${n} of your remedial student${n > 1 ? 's' : ''}. Every field is complete and the final action plans have been produced`;
    let cc = '';
    if (THANK_YOU_CC_HOD) {
      // Same CC rule as deadline emails: HoD(s) + division principal(s).
      const ccSet = new Set(ctx.allRoutes().flatMap(rt => splitEmails_(rt.hodEmails)));
      ctx.allRoutes().forEach(rt => { const pe = principalEmailForDivision_(rt.division); if (pe) ccSet.add(pe); });
      cc = Array.from(ccSet).join(',');
    }
    items.push({ to: t.email, cc: cc, subject: `Thank You — ${P.label}`, type: type, related: t.name,
      html: emailWrapper_('Thank You',
        `<p>Dear ${auditEsc_(t.name)},</p>
         <p>Thank you for completing ${what} — exactly what we need to follow each student's progress closely.</p>
         <p>Your care and effort are truly appreciated.</p>
         <p>Kind regards,</p>`) });
  });

  if (already.length) ui.alert(`Already thanked for ${P.label} (skipped): ${already.join(', ')}`);
  previewEmails_(`Thank-you — ${P.label}`, items, `No teacher has a flawless ${P.label} yet.`);
}

// ================================================================ 5. Deadlines → teachers (CC HoDs)

/** Principal's email for a division (ES / MS / HS) — only if marked Active in Settings > Contacts. */
function principalEmailForDivision_(division) {
  const role = { ES: 'Elementary School Principal', MS: 'Middle School Principal', HS: 'High School Principal' }[String(division || '').toUpperCase()];
  return role ? activeEmailFor_(role) : null;
}

/** One email per teacher with the deadline for the chosen tab group, her outstanding items; HoD(s) and principal(s) in CC. */
function sendDeadlinesToTeachers() {
  const ui = SpreadsheetApp.getUi();
  const phase = askPhase_('Send deadline to teachers');
  if (!phase) return;
  const P = PHASES[phase];
  const deadline = askDeadline_(false);
  if (!deadline) return;

  const roster = rosterIndex_();
  const hods = readHodDirectory_();
  const skipped = [];
  const items = [];
  const req = '<ul style="margin:6px 0 12px;">' + PHASE_REQUIREMENTS[phase].map(s => `<li>${auditEsc_(s)}</li>`).join('') + '</ul>';

  auditTeacherList_().forEach(t => {
    if (!t.email) { skipped.push(`${t.name} (no email)`); return; }
    const ctx = teacherContext_(t, roster, hods);
    const r = evaluateTeacherPhase_(t, ctx, phase);
    const routes = ctx.allRoutes();
    const ccSet = new Set(routes.flatMap(rt => splitEmails_(rt.hodEmails)));
    routes.forEach(rt => { const pe = principalEmailForDivision_(rt.division); if (pe) ccSet.add(pe); });
    const cc = Array.from(ccSet).join(',');
    const open = r.rows.filter(x => x.issue);
    let statusHtml;
    if (r.state === 'complete') {
      statusHtml = '<p style="color:#137333;"><b>Your entries for these tabs are already complete — thank you.</b></p>';
    } else if (r.state === 'notStarted' || r.state === 'none') {
      statusHtml = `<p style="color:#990000;"><b>These tabs have not been started yet.</b></p>`;
    } else {
      statusHtml = `<p><b>Still outstanding in your workbook:</b></p>
        <table style="border-collapse:collapse;font-size:13px;width:100%;"><tr style="${TH_ROW_}">
        <th style="${TD_}">Student</th>${P.tabs.length > 1 ? `<th style="${TD_}">Tab</th>` : ''}<th style="${TD_}">What is needed</th></tr>` +
        open.map(x => `<tr><td style="${TD_}">${auditEsc_(x.v[1])}${x.route.cls ? ' (' + auditEsc_(x.route.cls) + ')' : ''}</td>` +
          (P.tabs.length > 1 ? `<td style="${TD_}">${auditEsc_(x.tab)}</td>` : '') +
          `<td style="${TD_}">${auditEsc_(x.issue)}${x.ev.missing.length ? '<br>' + missingHtml_(formatMissing_(x.ev)) : ''}</td></tr>`).join('') + '</table>';
    }
    const hasPlaceholder = open.some(x => x.ev.placeholders && x.ev.placeholders.length);
    items.push({ to: t.email, cc: cc, subject: `Deadline: ${P.label} — ${deadline}`, type: `Deadline ${P.key}`, related: t.name,
      html: emailWrapper_(`Deadline — ${P.label}`,
        `<p>Dear ${auditEsc_(t.name)},</p>
         <p>Please make sure the <b>${auditEsc_(P.label)}</b> section of your remedial workbook is completed by <b>${auditEsc_(deadline)}</b>. This includes:</p>
         ${req}
         ${statusHtml}
         ${hasPlaceholder ? PLACEHOLDER_NOTE_HTML : ''}
         ${t.workbookUrl ? `<p>Your workbook: <a href="${auditEsc_(t.workbookUrl)}">open workbook</a></p>` : ''}
         <p>Kind regards,</p>`) });
  });

  if (skipped.length) ui.alert(`Skipped: ${skipped.join(', ')}`);
  previewEmails_(`Deadline — ${P.label}`, items, 'No teachers to email.');
}


// ================================================================ 2. Teachers & Principals — emails (all previewed)

function onboardingHtml_(name, workbookLink, guideLine) {
  return emailWrapper_('Your Remedial Program Workbook',
    `<p>Dear ${auditEsc_(name)},</p>
     <p>Your remedial follow-up workbook for this term is ready — this link is yours alone:</p>
     <p><a href="${auditEsc_(workbookLink)}">${auditEsc_(workbookLink)}</a></p>
     ${guideLine}
     <p>Please reach out to the Academic Dean's office with any questions.</p>
     <p>Kind regards,</p>`);
}

/** Workbook link + guide link to every teacher (CC: active Settings contacts) — preview first. */
function sendTeacherWorkbookLinksAndGuide() {
  const ui = SpreadsheetApp.getUi();
  const settings = readSettingsMap_();
  const guideId = settings['Teacher Guide File ID (Drive)'];
  if (!guideId) { ui.alert('Upload the Teacher Guide to Drive and put its File ID in Settings first.'); return; }
  const guideFile = DriveApp.getFileById(guideId);
  applyLinkSharing_(guideFile, settings);
  const guideLink = guideFile.getUrl();
  const cc = activeContactsCcString_();
  const skipped = [];
  const items = [];
  auditTeacherList_().forEach(t => {
    if (!t.email || !t.workbookUrl) { skipped.push(t.name); return; }
    items.push({ to: t.email, cc: cc, subject: 'Your Remedial Program Workbook & Guide', type: 'Teacher Onboarding', related: t.name,
      html: onboardingHtml_(t.name, t.workbookUrl,
        `<p>A short guide covering what to fill in and where is here:</p><p><a href="${auditEsc_(guideLink)}">${auditEsc_(guideLink)}</a></p>`) });
  });
  if (skipped.length) ui.alert(`Skipped (missing email or workbook not created yet): ${skipped.join(', ')}`);
  previewEmails_('Workbook + guide (all teachers)', items);
}

/** Same email for the teacher on the selected Teacher Roster row, with the guide attached — preview first. */
function sendWorkbookAndGuideToSelectedTeacher() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sh.getName() !== SHEETS.TEACHER_ROSTER) { ui.alert('Click on a teacher\'s row in Teacher Roster first, then run this again.'); return; }
  const row = SpreadsheetApp.getActiveSpreadsheet().getActiveRange().getRow();
  if (row < 2) { ui.alert('Select a teacher\'s row (not the header row) first.'); return; }
  const teacherName = sh.getRange(row, 1).getValue();
  const tr = teacherRosterRowByName_(teacherName);
  if (!tr || !tr.email || !tr.workbookLink) { ui.alert(`Missing email or workbook link for "${teacherName}" in Teacher Roster — fix that first.`); return; }
  const guideId = readSettingsMap_()['Teacher Guide File ID (Drive)'];
  previewEmails_(`Workbook + guide — ${teacherName}`, [{
    to: tr.email, cc: activeContactsCcString_(), subject: 'Your Remedial Program Workbook & Guide',
    type: 'Teacher Onboarding (Individual)', related: teacherName, attachIds: guideId ? [guideId] : [],
    html: onboardingHtml_(teacherName, tr.workbookLink, `<p>${guideId ? 'Your onboarding guide is attached.' : 'A guide will follow separately.'}</p>`)
  }]);
}

/** Refreshes the three division workbooks (optional) and emails each active principal her link — preview first. */
function sendPrincipalWorkbooks() {
  const ui = SpreadsheetApp.getUi();
  const ans = ui.alert('Principal workbooks', 'Refresh the principal workbooks with the latest data before sending?', ui.ButtonSet.YES_NO_CANCEL);
  if (ans === ui.Button.CANCEL || ans === ui.Button.CLOSE) return;
  if (ans === ui.Button.YES) refreshPrincipalWorkbooks();

  const props = PropertiesService.getScriptProperties();
  const contacts = readContacts_();
  const roles = { ES: ['Elementary School Principal', 'Elementary School Office Assistant', 'Elementary'],
                  MS: ['Middle School Principal', 'Middle School Office Assistant', 'Middle School'],
                  HS: ['High School Principal', 'High School Office Assistant', 'High School'] };
  const items = [], notes = [];
  Object.keys(roles).forEach(div => {
    const [pRole, aRole, label] = roles[div];
    const id = props.getProperty('PRINCIPAL_WORKBOOK_ID_' + div);
    const p = contacts[pRole];
    if (!id) { notes.push(`${label}: workbook not created yet`); return; }
    if (!p || !p.active || !p.email) { notes.push(`${label}: principal not marked Active in Settings`); return; }
    const a = contacts[aRole];
    const url = SpreadsheetApp.openById(id).getUrl();
    items.push({ to: p.email, cc: (a && a.active && a.email) ? a.email : '', subject: `${label} — Remedial Follow-Up Workbook`,
      type: 'Principal Workbook', related: label,
      html: emailWrapper_(`${label} — Remedial Follow-Up`,
        `<p>Dear ${auditEsc_(p.name)},</p>
         <p>Your ${auditEsc_(label)} remedial follow-up workbook is up to date: <a href="${auditEsc_(url)}">open workbook</a>.</p>
         <p>It shows today's remedial schedule and attendance, and which teachers still have Action Plans or Progress Reports outstanding.</p>
         <p>Kind regards,</p>`) });
  });
  if (notes.length) ui.alert('Not included:\n' + notes.join('\n'));
  previewEmails_('Principal workbooks', items, 'No principal emails to send.');
}

// ================================================================ #6 Template alignment

const TPL_HEADERS = ['Teacher', 'Teacher Workbook', 'Student', 'Grade & Class', 'Subject',
  'Action Plan Link', 'Found In', 'Result', 'Issues', 'Checked At'];
const TPL_PROGRESS_KEY = 'AP_TEMPLATE_CHECK_NEXT';

function resetTemplateCheckProgress() {
  PropertiesService.getScriptProperties().deleteProperty(TPL_PROGRESS_KEY);
  SpreadsheetApp.getActiveSpreadsheet().toast('Template check will start over on the next run.', 'Reset', 5);
}

function cleanLabel_(s) {
  return String(s || '').replace(/[*#]/g, '').replace(/\s+/g, ' ').trim().toLowerCase();
}

/** Structural fingerprint of a Doc: per table, per row -> cell count + label (first cell, when the row has 2+ cells). */
function docSignature_(doc) {
  return doc.getBody().getTables().map(t => {
    const rows = [];
    for (let r = 0; r < t.getNumRows(); r++) {
      const row = t.getRow(r);
      const n = row.getNumCells();
      rows.push({ n: n, label: n >= 2 ? cleanLabel_(row.getCell(0).getText()) : null });
    }
    return rows;
  });
}

function compareSignatures_(tpl, doc) {
  const issues = [];
  if (tpl.length !== doc.length) issues.push(`Has ${doc.length} tables, template has ${tpl.length} (built from an old/edited template)`);
  const n = Math.min(tpl.length, doc.length);
  for (let i = 0; i < n && issues.length < 6; i++) {
    if (tpl[i].length !== doc[i].length) { issues.push(`Table ${i + 1}: ${doc[i].length} rows, template ${tpl[i].length}`); continue; }
    for (let r = 0; r < tpl[i].length; r++) {
      if (tpl[i][r].n !== doc[i][r].n) { issues.push(`Table ${i + 1} row ${r + 1}: ${doc[i][r].n} cells, template ${tpl[i][r].n}`); break; }
      if (tpl[i][r].label !== null && tpl[i][r].label !== doc[i][r].label) {
        issues.push(`Table ${i + 1} row ${r + 1}: label "${doc[i][r].label}" ≠ template "${tpl[i][r].label}"`); break;
      }
    }
  }
  return issues;
}

/**
 * TEMPLATE ENFORCEMENT ONLY. Opens every Action Plan link (each teacher's Week 1
 * tab and Student Roster "Action Plan Link") and confirms it was produced by the
 * system from the official template: same layout (tables, rows, labels and items)
 * as the template, created by the system (not a teacher's own file), not an
 * uploaded PDF/Word file, not in the Trash. It does NOT check the entries —
 * that is the job of the Week 1 / 5 / 10 audits.
 * Problems are listed on "AP Template Check" with the TEACHER WORKBOOK URL.
 * Stops safely before the time limit and continues where it left off on the next run.
 */
function checkActionPlanTemplateAlignment() {
  const ui = SpreadsheetApp.getUi();
  const settings = readSettingsMap_();
  const enId = settings['Action Plan English Template ID'];
  const arId = settings['Action Plan Arabic Template ID (Arabic, Religion, Arabic Social Studies only)'] ||
    settings['Action Plan Arabic Template ID'] || enId;
  if (!enId) { ui.alert('Action Plan English Template ID is missing in Settings.'); return; }
  const tplSig = { en: docSignature_(DocumentApp.openById(enId)), ar: docSignature_(DocumentApp.openById(arId)) };
  const systemOwner = Session.getEffectiveUser().getEmail();

  // 1. Collect every Action Plan link with its context.
  const roster = rosterIndex_();
  const teachers = auditTeacherList_();
  const items = [], seen = {};
  teachers.forEach(t => {
    const ss = openWorkbookSafe_(t.workbookUrl);
    if (!ss) return;
    const tab = readTeacherTab_(ss, 'Week 1');
    if (!tab) return;
    const entries = roster.filter(e => teacherMatches_(e.teacherField, t.name));
    tab.rows.forEach(r => {
      const link = String(r.v[WEEK1_COLS.ACTION_PLAN_LINK - 1] || '').trim();
      if (!link) return;
      const id = extractFileIdFromUrl_(link);
      if (!id || seen[id]) return;
      seen[id] = 1;
      const e = entries.find(x => x.id === String(r.v[0]).trim());
      items.push({ fileId: id, link: link, teacher: t.name, workbook: t.workbookUrl, student: r.v[1],
        cls: e ? gradeClassLabel_(e.grade, e.cls) : r.v[WEEK1_COLS.CLASS - 1],
        subject: e ? e.subject : t.subject.split(',')[0].trim(), wbRow: r.v, source: 'Teacher workbook (Week 1)' });
    });
  });
  roster.forEach(e => {
    if (!e.apLink) return;
    const id = extractFileIdFromUrl_(e.apLink);
    if (!id || seen[id]) return;
    seen[id] = 1;
    const t = teachers.find(x => teacherMatches_(e.teacherField, x.name)) || { name: String(e.teacherField), workbookUrl: '' };
    items.push({ fileId: id, link: e.apLink, teacher: t.name, workbook: t.workbookUrl, student: e.name,
      cls: gradeClassLabel_(e.grade, e.cls), subject: e.subject, wbRow: null, source: 'Student Roster' });
  });

  // 2. Resume support.
  const props = PropertiesService.getScriptProperties();
  let startIdx = parseInt(props.getProperty(TPL_PROGRESS_KEY) || '0', 10);
  if (startIdx >= items.length) startIdx = 0;
  let sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AUDIT_SHEETS.TEMPLATE);
  if (startIdx === 0 || !sh) {
    sh = writeReviewSheet_(AUDIT_SHEETS.TEMPLATE, TPL_HEADERS, [], null);
    [140, 160, 200, 80, 120, 160, 150, 150, 420, 120].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  }

  // 3. Check each doc.
  const start = Date.now();
  const out = [];
  let i = startIdx;
  for (; i < items.length; i++) {
    if (Date.now() - start > 4.8 * 60 * 1000) break;
    const it = items[i];
    let result = 'Aligned', issues = [];
    try {
      const file = DriveApp.getFileById(it.fileId);
      const mime = file.getMimeType();
      let owner = '';
      try { owner = file.getOwner() ? file.getOwner().getEmail() : ''; } catch (e) {}
      if (owner && owner.toLowerCase() !== systemOwner.toLowerCase()) {
        issues.push(`Not created by the system — the file belongs to ${owner}`);
      }
      if (file.isTrashed()) { result = 'Not aligned'; issues.push('File is in the Trash'); }
      else if (mime !== MimeType.GOOGLE_DOCS) {
        const name = file.getName();
        if (/Action Plan \((Week 5|Final)\)\.pdf$/i.test(name)) { result = 'Skipped (generated PDF snapshot)'; }
        else { result = 'Not aligned'; issues.push(`Uploaded ${mime === MimeType.PDF ? 'PDF' : 'file'} ("${name}") — not generated from the template`); }
      } else {
        const doc = DocumentApp.openById(it.fileId);
        const variant = templateVariantForSubject_(it.subject);
        const sig = docSignature_(doc);
        let structural = compareSignatures_(tplSig[variant], sig);
        const other = variant === 'ar' ? 'en' : 'ar';
        if (structural.length && !compareSignatures_(tplSig[other], sig).length) structural = []; // matches the other official template
        issues = issues.concat(structural);
        if (issues.length) result = 'Not aligned';
      }
    } catch (e) {
      result = 'Cannot open'; issues.push(e.message);
    }
    if (result === 'Aligned') continue; // only problems and skipped files are listed
    out.push([it.teacher, it.workbook, it.student, it.cls, it.subject, it.link, it.source, result, issues.join('\n'), new Date()]);
  }
  if (out.length) {
    const first = sh.getLastRow() + 1;
    sh.getRange(first, 1, out.length, TPL_HEADERS.length).setValues(out).setVerticalAlignment('top');
    sh.getRange(first, 9, out.length, 1).setWrap(true);
    sh.getRange(first, 10, out.length, 1).setNumberFormat('dd/MM/yyyy HH:mm');
    const resRange = sh.getRange(2, 8, Math.max(sh.getLastRow() - 1, 1), 1);
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Aligned').setBackground(THEME.statusGreen).setRanges([resRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Not aligned').setBackground(THEME.statusRed).setRanges([resRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Cannot open').setBackground(THEME.statusOrange).setRanges([resRange]).build()
    ]);
  }

  if (i < items.length) {
    props.setProperty(TPL_PROGRESS_KEY, String(i));
    ui.alert(`Checked ${i} of ${items.length} action plans — stopped before the time limit.\nRun "Check Action Plans vs Template" again to continue.`);
    return;
  }
  props.deleteProperty(TPL_PROGRESS_KEY);

  // 4. Summary: teachers with problems + their workbook URL.
  const all = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, TPL_HEADERS.length).getValues() : [];
  const bad = {};
  all.forEach(r => { if (r[7] === 'Not aligned' || r[7] === 'Cannot open') { (bad[r[0]] = bad[r[0]] || { url: r[1], n: 0 }).n++; } });
  const names = Object.keys(bad).sort();
  sh.activate();
  ui.alert(`Checked ${items.length} action plan(s). Using the official template: ${items.length - all.length}. Listed on the tab (not aligned / skipped): ${all.length}.\n\n` +
    (names.length
      ? `NOT ALIGNED — ${names.length} teacher(s):\n` + names.map(n => `• ${n} (${bad[n].n}): ${bad[n].url || 'no workbook link'}`).join('\n')
      : 'Every action plan uses the official template. ✅') +
    `\n\nFull details are on "${AUDIT_SHEETS.TEMPLATE}".`);
}


/** One-off / any-time tidy: on every audit tab, untick Sent rows, colour them green and move them below the unsent ones. */
function tidySentRowsInAudits() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(PHASES).forEach(k => {
    const sh = ss.getSheetByName(PHASES[k].sheet);
    if (!sh || sh.getLastRow() < 2) return;
    const rows = sh.getLastRow() - 1;
    const status = sh.getRange(2, AUDIT_COL.STATUS, rows, 1).getValues();
    const sendRange = sh.getRange(2, AUDIT_COL.SEND, rows, 1);
    sendRange.setValues(sendRange.getValues().map((r, i) => [status[i][0] === 'Sent' ? false : r[0]]));
    moveSentRowsToBottom_(sh, AUDIT_COL.STATUS);
  });
}
