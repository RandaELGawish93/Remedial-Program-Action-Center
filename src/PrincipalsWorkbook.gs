/**
 * PRINCIPALSWORKBOOK.gs
 * THREE SEPARATE workbooks — one per division (ES/MS/HS), not tabs in one
 * shared file. Each is read-only for that division's principal + office
 * assistant, very limited functionality (view only, no editing).
 *
 * Each workbook has:
 *  - Dashboard: today's remedial schedule (matched by Day), with a
 *    green/red attendance grid for whichever teacher(s) are scheduled
 *    today, plus missing AP1/Progress Report/AP2 teacher counts + list.
 *  - 7 subject tabs (English, Science, Mathematics, Arabic Social Studies,
 *    Arabic, Religion, Social Studies (History)) — every division now shows
 *    the same full Week 2-10 schedule grid, scoped to On Probation students only.
 *
 * "Live" here means: run "Refresh Principal Workbooks" and it pushes
 * current data in. True instant push isn't safe at scale (see the
 * trigger-quota note in TeacherWorkbookTriggers.gs) — this uses the same
 * pull/refresh pattern as Scores Progress and the Dashboard's Cards 6-7.
 */

const DIVISION_GRADE_RANGES = { ES: [1, 5], MS: [6, 8], HS: [9, 12] };
// Roster stores the FULL subject name (confirmed: "Mathematics", not "Math") —
// filtering on the wrong string was the actual reason the Math tab looked empty.
const PRINCIPAL_SUBJECTS = ['English', 'Science', 'Mathematics', 'Arabic Social Studies', 'Arabic', 'Islamic Religion', 'Social Studies'];
const PRINCIPAL_SUBJECT_LABELS = { 'Social Studies': 'Social Studies (History)', 'Mathematics': 'Math', 'Islamic Religion': 'Religion' };

function classToDivision_(cls) {
  const m = String(cls).match(/^(\d{1,2})/);
  if (!m) return null;
  const grade = parseInt(m[1], 10);
  for (const div in DIVISION_GRADE_RANGES) {
    const [lo, hi] = DIVISION_GRADE_RANGES[div];
    if (grade >= lo && grade <= hi) return div;
  }
  return null;
}

/**
 * Prefer this over classToDivision_() wherever a Roster row is available —
 * Roster's Grade column is always a clean number, whereas Class strings
 * aren't guaranteed to consistently include the grade prefix (this was the
 * same root cause behind the Teacher Roster "classes show letters only" bug).
 */
function divisionFromGrade_(grade) {
  const g = parseInt(grade, 10);
  if (isNaN(g)) return null;
  for (const div in DIVISION_GRADE_RANGES) {
    const [lo, hi] = DIVISION_GRADE_RANGES[div];
    if (g >= lo && g <= hi) return div;
  }
  return null;
}

function createOrGetPrincipalWorkbook_(division) {
  const props = PropertiesService.getScriptProperties();
  const propKey = 'PRINCIPAL_WORKBOOK_ID_' + division;
  const existingId = props.getProperty(propKey);
  if (existingId) {
    try { return SpreadsheetApp.openById(existingId); } catch (e) { /* fall through and recreate */ }
  }

  const settings = readSettingsMap_();
  const folderId = settings['Principals Workbook Folder ID'];
  const divisionLabel = division === 'ES' ? 'Elementary' : division === 'MS' ? 'Middle School' : 'High School';
  const ss = SpreadsheetApp.create(`${divisionLabel} Principal — Remedial Follow-Up`);

  if (folderId) {
    const file = DriveApp.getFileById(ss.getId());
    DriveApp.getFolderById(folderId).addFile(file);
    DriveApp.getRootFolder().removeFile(file);
  }

  ss.getSheets()[0].setName('Dashboard');
  PRINCIPAL_SUBJECTS.forEach(subj => {
    ss.insertSheet(PRINCIPAL_SUBJECT_LABELS[subj] || subj);
  });

  // Read-only: no protections needed since we only ever add VIEWERS, never editors.
  const contacts = readContacts_();
  const principalRole = division === 'ES' ? 'Elementary School Principal' : division === 'MS' ? 'Middle School Principal' : 'High School Principal';
  const assistantRole = division === 'ES' ? 'Elementary School Office Assistant' : division === 'MS' ? 'Middle School Office Assistant' : 'High School Office Assistant';
  [principalRole, assistantRole, 'Campus Director', 'ADCI', 'Student Success Center Manager'].forEach(role => {
    const c = contacts[role];
    if (c && c.active && c.email) ss.addViewer(c.email);
  });

  props.setProperty(propKey, ss.getId());
  return ss;
}

function refreshPrincipalWorkbooks() {
  const startTime = new Date().getTime();
  const MAX_RUNTIME_MS = 4.5 * 60 * 1000;
  const done = [], skipped = [];

  ['ES', 'MS', 'HS'].forEach(division => {
    if (new Date().getTime() - startTime > MAX_RUNTIME_MS) { skipped.push(division); return; }
    refreshOnePrincipalWorkbook_(division);
    done.push(division);
  });

  let msg = `Refreshed: ${done.join(', ') || 'none'}.`;
  if (skipped.length) msg += ` Stopped before the time limit — run again for: ${skipped.join(', ')}.`;
  SpreadsheetApp.getUi().alert(msg);
}

function refreshOnePrincipalWorkbook_(division) {
  const ss = createOrGetPrincipalWorkbook_(division);
  const settings = readSettingsMap_();
  const termStart = TERM_START_DATES[String(settings['Term # (1, 2, or 3)'])];
  const currentWeek = termStart ? Math.floor((new Date() - termStart) / (7 * 86400000)) + 1 : 1;

  const rosterSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ROSTER);
  const rosterLastRow = rosterSh.getLastRow();
  const rosterData = rosterLastRow > 1 ? rosterSh.getRange(2, 1, rosterLastRow - 1, ROSTER_HEADERS.length).getValues() : [];
  const divisionRows = rosterData.filter(r => divisionFromGrade_(r[ROSTER_COL.GRADE - 1]) === division);
  
  refreshPrincipalDashboard_(ss, division, divisionRows, currentWeek, settings);
}

// ---------- Dashboard: rectangular Sun-Thu schedule + missing-work summary ----------

const SCHOOL_DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday'];

function refreshPrincipalDashboard_(ss, division, divisionRows, currentWeek, settings) {
  const sh = ss.getSheetByName('Dashboard');
  sh.clear();

  const divisionLabel = division === 'ES' ? 'Elementary' : division === 'MS' ? 'Middle School' : 'High School';
  sh.getRange('A1').setValue(`${divisionLabel} — Remedial Follow-Up Dashboard`).setFontWeight('bold').setFontSize(14);
  const today = new Date();
  const todayName = Utilities.formatDate(today, Session.getScriptTimeZone(), 'EEEE');
  sh.getRange('A2').setValue(`${todayName}, ${Utilities.formatDate(today, Session.getScriptTimeZone(), 'MM/dd/yyyy')} — Week ${currentWeek}`)
    .setFontStyle('italic');

  let r = 4;
  sh.getRange(r, 1, 1, 6).merge().setValue('WEEKLY REMEDIAL SCHEDULE (Sunday–Thursday)').setFontWeight('bold').setBackground('#1c4587').setFontColor('#ffffff');
  r++;

  // Rectangular grid: one row per (Teacher, Subject) in this division,
  // one column per school day. Time + this week's punctuality status
  // (entered by YOU, the AD, in Remedial Attendance — never by the
  // teacher) sits in whichever single day column she's actually
  // scheduled on; the rest of that row's day columns stay blank.
  const headerRow = r;
  sh.getRange(r, 1, 1, 6).setValues([['Teacher / Subject', ...SCHOOL_DAYS]]).setFontWeight('bold');
  SCHOOL_DAYS.forEach((day, i) => {
    sh.getRange(r, 2 + i).setBackground(day === todayName ? '#93c47d' : '#d9d9d9');
  });
  r++;

  const attSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ATTENDANCE_TEACHERS);
  const scheduleRows = [];
  if (attSh && attSh.getLastRow() > 1) {
    const weekCol = 6 + (currentWeek - 2); // Week 2 starts at column 6
    attSh.getRange(2, 1, attSh.getLastRow() - 1, Math.max(weekCol, 5)).getValues().forEach(row => {
      const [teacher, subject, div, day, time] = row;
      if (div !== division) return;
      const status = (currentWeek >= 2 && weekCol <= row.length) ? row[weekCol - 1] : '';
      scheduleRows.push({ teacher, subject, day, time, status });
    });
  }

  if (!scheduleRows.length) {
    sh.getRange(r, 1).setValue('No remedial schedule entered yet.');
    r += 2;
  } else {
    scheduleRows.forEach(s => {
      const rowValues = ['', '', '', '', '', ''];
      rowValues[0] = `${s.teacher} — ${s.subject}`;
      const dayIndex = SCHOOL_DAYS.findIndex(d => d.toLowerCase() === String(s.day).trim().toLowerCase());
      if (dayIndex >= 0) {
        rowValues[1 + dayIndex] = s.time ? `${s.time}${s.status ? ' — ' + s.status : ''}` : (s.status || '');
      }
      sh.getRange(r, 1, 1, 6).setValues([rowValues]);
      r++;
    });
    sh.getRange(headerRow, 1, r - headerRow, 6).setBorder(true, true, true, true, true, true);
    r++;
  }
  // Missing AP1 / Progress Report / AP2 summary for this division — each
  // entry now explicitly labeled with WHICH item is missing, not just a
  // mixed list. Also scoped to officially On Probation students only —
  // Prevention/Recommended students don't belong in this summary.
  sh.getRange(r, 1, 1, 5).merge().setValue('MISSING WORK SUMMARY').setFontWeight('bold').setBackground('#cc0000').setFontColor('#ffffff');
  r++;
  const probationRows = divisionRows.filter(row => row[ROSTER_COL.STATUS - 1] === 'On Probation');
  const missingAP = probationRows.filter(row => !row[ROSTER_COL.ACTION_PLAN - 1]);
  const missingPR = probationRows.filter(row => !row[ROSTER_COL.PROGRESS_REPORT - 1]);
  sh.getRange(r, 1).setValue(`Teachers with missing Action Plan: ${new Set(missingAP.map(x => x[ROSTER_COL.TEACHER - 1])).size}`); r++;
  sh.getRange(r, 1).setValue(`Teachers with missing Progress Report: ${new Set(missingPR.map(x => x[ROSTER_COL.TEACHER - 1])).size}`); r++;
  r++;
  sh.getRange(r, 1, 1, 4).setValues([['Student', 'Subject', 'Teacher', 'Missing']]).setFontWeight('bold');
  const missingTableHeaderRow = r;
  r++;
  const labeledMissing = [
    ...missingAP.map(row => ({ row, item: 'Action Plan' })),
    ...missingPR.map(row => ({ row, item: 'Progress Report' })),
  ];
  labeledMissing.forEach(({ row, item }) => {
    sh.getRange(r, 1, 1, 4).setValues([[row[ROSTER_COL.NAME - 1], row[ROSTER_COL.SUBJECT - 1], row[ROSTER_COL.TEACHER - 1], item]]);
    r++;
  });
  if (labeledMissing.length) {
    sh.getRange(missingTableHeaderRow, 1, r - missingTableHeaderRow, 4).setBorder(true, true, true, true, true, true);
  }

  for (let c = 1; c <= 8; c++) sh.setColumnWidth(c, 110);
  sh.setColumnWidth(1, 170);
  sh.autoResizeColumns(1, sh.getLastColumn()); // covers however many columns actually have content —
                                                 // the P/A grid can stretch past column F depending on
                                                 // how many weeks have passed, so a fixed count missed it
}
