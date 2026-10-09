/**
 * SETUP_TEACHERROSTER.gs
 * One row per remedial teacher, fully derived from Student Roster — never
 * typed into directly except the folder-link fill, which is automatic too.
 *
 * NOTE ON "CHIPS": Google Sheets' pill-shaped smart chips have no public
 * Apps Script API — only creatable by hand via Insert > Smart chip. Links
 * here use =HYPERLINK() with a short label instead: clickable, short text,
 * opens the file — functionally similar, visually a blue link, not a pill.
 */

const TEACHER_ROSTER_HEADERS = [
  'Teacher Name', 'Email', 'Subject', 'Classes', 'Number of Students (auto)',
  'Action Plan Folder', 'Progress Report Folder', 'Teacher Workbook Link'
];

const TEACHER_ROSTER_COL = {
  NAME: 1, EMAIL: 2, SUBJECT: 3, CLASSES: 4, NUM_STUDENTS: 5,
  ACTION_PLAN_FOLDER: 6, PROGRESS_REPORT_FOLDER: 7, WORKBOOK_LINK: 8
};

function buildTeacherRoster() {
  const sh = getOrCreateSheet_(SHEETS.TEACHER_ROSTER);
  const alreadyHasData = sh.getLastRow() > 1;

  if (!alreadyHasData) {
    sh.clear();
    sh.getRange(1, 1, 1, TEACHER_ROSTER_HEADERS.length)
      .setValues([TEACHER_ROSTER_HEADERS])
      .setFontWeight('bold').setBackground('#1c4587').setFontColor('#ffffff');
    sh.setFrozenRows(1);

    const widths = [170, 220, 130, 200, 150, 150, 150, 160];
    widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  } else {
    // Re-run: don't wipe anything here — populateTeacherRosterFromStudentRoster_()
    // below already preserves Email, folder links, and the Workbook link
    // for any teacher it already knows about. Just refresh the header in
    // case columns were added since this sheet was first built.
    sh.getRange(1, 1, 1, TEACHER_ROSTER_HEADERS.length)
      .setValues([TEACHER_ROSTER_HEADERS])
      .setFontWeight('bold').setBackground('#1c4587').setFontColor('#ffffff');
  }
  populateTeacherRosterFromStudentRoster_();
  sh.autoResizeColumns(1, TEACHER_ROSTER_HEADERS.length);
}

/**
 * Derives the teacher list from Student Roster's "Remedial Teacher" column
 * (format: "Name (email)"). Anything you've typed directly into Teacher
 * Roster's Email column, or the folder/workbook links, always wins over
 * what gets parsed here — Student Roster only fills those in when there's
 * nothing there yet. Re-run any time the Student Roster changes.
 */
function populateTeacherRosterFromStudentRoster_() {
  const rosterSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ROSTER);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.TEACHER_ROSTER);
  const lastRow = rosterSh.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Student Roster is empty — nothing to derive teachers from yet.', 'Notice', 5);
    return;
  }

  const data = rosterSh.getRange(2, 1, lastRow - 1, ROSTER_HEADERS.length).getValues();
  const teachers = {};

  data.forEach(row => {
    const teacherField = row[ROSTER_COL.TEACHER - 1];
    const subject = row[ROSTER_COL.SUBJECT - 1];
    const grade = row[ROSTER_COL.GRADE - 1];
    const rawClass = row[ROSTER_COL.CLASS - 1];
    if (!teacherField) return;

    const m = String(teacherField).match(/^(.*?)\s*\(([^)]+)\)\s*$/);
    const name = m ? m[1].trim() : String(teacherField).trim();
    const email = m ? m[2].trim() : '';
    const key = email || name;

    let clsLabel = '';
    if (rawClass) {
      const letterOnly = String(rawClass).trim().replace(/^\d+\s*/, '');
      clsLabel = `${grade}${letterOnly}`;
    }

    if (!teachers[key]) teachers[key] = { name, email, subjects: new Set(), classes: new Set() };
    if (subject) teachers[key].subjects.add(subject);
    if (clsLabel) teachers[key].classes.add(clsLabel);
  });

  // Preserve existing Email, folder links, and Workbook link if this is a
  // re-run — anything you've typed in by hand always wins over what gets
  // parsed from Student Roster. Read by named column, not position.
  const existingLinks = {};
  const oldLastRow = sh.getLastRow();
  if (oldLastRow > 1) {
    sh.getRange(2, 1, oldLastRow - 1, TEACHER_ROSTER_HEADERS.length).getValues().forEach(r => {
      if (r[0]) existingLinks[r[0]] = {
        email: r[TEACHER_ROSTER_COL.EMAIL - 1],
        actionPlan: r[TEACHER_ROSTER_COL.ACTION_PLAN_FOLDER - 1],
        progressReport: r[TEACHER_ROSTER_COL.PROGRESS_REPORT_FOLDER - 1],
        workbook: r[TEACHER_ROSTER_COL.WORKBOOK_LINK - 1]
      };
    });
  }

  const rows = Object.values(teachers)
    .sort((a, b) => a.name.localeCompare(b.name))
    .map(t => {
      const links = existingLinks[t.name] || {};
      return [
        t.name, links.email || t.email,
        Array.from(t.subjects).join(', '),
        Array.from(t.classes).sort().join(', '),
        '',
        links.actionPlan || '', links.progressReport || '', links.workbook || ''
      ];
    });

  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, TEACHER_ROSTER_HEADERS.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, TEACHER_ROSTER_HEADERS.length).setValues(rows);

  if (rows.length) {
    sh.getRange(1, 1, rows.length + 1, TEACHER_ROSTER_HEADERS.length).setBorder(true, true, true, true, true, true);
    sh.getRange(2, 4, rows.length, 2).setHorizontalAlignment('center');
    sh.getRange(1, 1, 1, TEACHER_ROSTER_HEADERS.length).setHorizontalAlignment('center');
  }

  const numRows = Math.max(rows.length, 60);
  const formulas = [];
  for (let i = 0; i < numRows; i++) {
    const row = i + 2;
    formulas.push([`=IF($A${row}="", "", COUNTIF('${SHEETS.ROSTER}'!$I$2:$I$301, "*"&$A${row}&"*"))`]);
  }
  sh.getRange(2, 5, numRows, 1).setFormulas(formulas);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Teacher Roster derived from Student Roster: ${rows.length} teacher(s) found. ` +
    'Run this again any time the Roster changes.', 'Done', 7
  );
}

/**
 * Fills the two folder-link columns. Both teachers see the same single
 * Action Plan Folder and Progress Report Folder — documents aren't split
 * per-teacher. Shown as HYPERLINK labels ("Open Folder"), not raw URLs.
 */
function fillTeacherFolderLinks() {
  const settings = readSettingsMap_();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.TEACHER_ROSTER);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const termFolder = getTermFolder_(settings['Action Plan Folder ID'], settings);
  const prFolder = getTermFolder_(settings['Progress Report Folder ID'], settings);

  const apFormula = `=HYPERLINK("${termFolder.getUrl()}", "Open Action Plan Folder")`;
  const prFormula = `=HYPERLINK("${prFolder.getUrl()}", "Open Progress Report Folder")`;

  const names = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  names.forEach((row, i) => {
    if (!row[0]) return;
    const r = i + 2;
    sh.getRange(r, TEACHER_ROSTER_COL.ACTION_PLAN_FOLDER).setFormula(apFormula);
    sh.getRange(r, TEACHER_ROSTER_COL.PROGRESS_REPORT_FOLDER).setFormula(prFormula);
  });

  sh.autoResizeColumns(TEACHER_ROSTER_COL.ACTION_PLAN_FOLDER, 2);

  SpreadsheetApp.getActiveSpreadsheet().toast('Folder links filled for all teachers.', 'Done', 5);
}

/**
 * Utility used by DocGen and email functions: looks up a teacher's email
 * by name from Teacher Roster. Returns null if not found.
 */
function teacherEmailByName_(name) {
  if (!name) return null;
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.TEACHER_ROSTER);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  const data = sh.getRange(2, 1, lastRow - 1, 2).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(name).trim()) return data[i][1];
  }
  return null;
}

/** Returns the full Teacher Roster row (as an object) for a teacher name, or null. */
function teacherRosterRowByName_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.TEACHER_ROSTER);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  const data = sh.getRange(2, 1, lastRow - 1, TEACHER_ROSTER_HEADERS.length).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(name).trim()) {
      return {
        rowNum: i + 2, name: data[i][0], email: data[i][1], subject: data[i][2],
        classes: data[i][3], workbookLink: data[i][TEACHER_ROSTER_COL.WORKBOOK_LINK - 1]
      };
    }
  }
  return null;
}

/** Writes a value into a specific Teacher Roster column for a given teacher, by header name. */
function setTeacherRosterField_(teacherName, headerName, value) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.TEACHER_ROSTER);
  const col = TEACHER_ROSTER_HEADERS.indexOf(headerName) + 1;
  if (!col) return;
  const row = teacherRosterRowByName_(teacherName);
  if (!row) return;
  sh.getRange(row.rowNum, col).setValue(value);
}

