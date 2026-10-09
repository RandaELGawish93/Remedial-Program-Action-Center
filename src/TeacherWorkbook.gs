/**
 * TEACHERWORKBOOK.gs
 * Each teacher gets their own spreadsheet (same hub-and-spoke pattern as
 * the Advisor Program Manager). All the logic lives HERE, in the master
 * workbook's script project —
 * teacher workbooks have no script of their own.
 *
 * TABS PER TEACHER WORKBOOK:
 *  - Profile             : read-only identity + live status mirrors + current week
 *  - Progress Report Entries
 *  - Week 1              : Action Plan (per student)
 *  - Week 5              : Progress Report status + Action Plan continuation (per student)
 *  - Week 10             : Reflection / Overall (per student)
 *
 * TWO PROTECTED ZONES on every tab that needs them:
 *  - AD-editable       : notes/overrides only you can touch
 *  - Teacher-editable  : their own entries (plan text, scores)
 */





function createTeacherWorkbooks() {
  const settings = readSettingsMap_();
  const folderId = settings["Teacher's Workbook Folder ID"];
  if (!folderId) {
    SpreadsheetApp.getUi().alert("Fill in Teacher's Workbook Folder ID in Settings first.");
    return;
  }
  const folder = DriveApp.getFolderById(folderId);
  const term = settings['Term # (1, 2, or 3)'] || '?';

  const trSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.TEACHER_ROSTER);
  const lastRow = trSh.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('Teacher Roster is empty.'); return; }

  const rows = trSh.getRange(2, 1, lastRow - 1, TEACHER_ROSTER_HEADERS.length).getValues();
  let created = 0;
  let stoppedEarly = false;
  const totalNeeded = rows.filter(r => r[0] && !r[TEACHER_ROSTER_COL.WORKBOOK_LINK - 1]).length;

  const startTime = new Date().getTime();
  const MAX_RUNTIME_MS = 4.5 * 60 * 1000;

  let failed = 0;
  const failedNames = [];

  for (let i = 0; i < rows.length; i++) {
    if (new Date().getTime() - startTime > MAX_RUNTIME_MS) { stoppedEarly = true; break; }

    const row = rows[i];
    const name = row[0];
    if (!name) continue;
    const existingLink = row[TEACHER_ROSTER_COL.WORKBOOK_LINK - 1];
    if (existingLink) continue;

    try {
      const email = row[1];
      const ss = SpreadsheetApp.create(`Teacher Workbook - ${name} - Term ${term}`);
      const file = DriveApp.getFileById(ss.getId());
      folder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);

      buildTeacherProfileTab_(ss, name);
      buildTeacherProgressReportTab_(ss, name);
      buildTeacherWeekTab_(ss, 1, name);
      buildTeacherWeekTab_(ss, 5, name);
      buildTeacherWeekTab_(ss, 10, name);

      orderTeacherWorkbookTabs_(ss); 

      if (email) ss.addEditor(email);

      trSh.getRange(i + 2, TEACHER_ROSTER_COL.WORKBOOK_LINK).setValue(ss.getUrl());
      created++;
    } catch (e) {
      failed++;
      failedNames.push(`${name}: ${e.message}`);
    }
  }

  const remaining = totalNeeded - created;
  let msg = `${created} teacher workbook(s) created this run.`;
  if (failed) msg += `\n${failed} FAILED — see below.\n\n` + failedNames.slice(0, 10).join('\n');
  if (stoppedEarly || remaining > 0) {
    msg += `\n\n${remaining} teacher(s) still remaining — this stopped itself before hitting ` +
      'Google\'s execution time limit, rather than erroring out partway through. ' +
      'Just run "Create Teacher Workbooks" again — it automatically skips everyone ' +
      'already done and continues with the rest.';
  }
  SpreadsheetApp.getUi().alert(msg);
}

// ---------- Profile tab ----------

function buildTeacherProfileTab_(ss, teacherName) {
  const sh = ss.getSheets()[0].setName('Profile');
  sh.getRange('A1').setValue(`Teacher Profile — ${teacherName}`).setFontWeight('bold').setFontSize(14).setFontColor(THEME.darkNavy);

  const rows = [
    ['Teacher Name (read-only)', teacherName],
    ['Current Week (auto)', ''],
    ['MoM #1 (Week 1) Status', ''],
    ['MoM #2 (Week 5) Status', ''],
    ['MoM #3 (Week 10) Status', ''],
    ['Remedial Attendance — This Week', ''],
    ['Action Plans Folder', ''],
  ];
  sh.getRange(3, 1, rows.length, 2).setValues(rows);
  sh.getRange(3, 1, rows.length, 1).setFontWeight('bold');
  sh.setColumnWidth(1, 320);
  sh.setColumnWidth(2, 320);

  rows.forEach((_, i) => sh.getRange(4 + i, 2).setNote('Pushed from the AD Master workbook — do not edit directly.'));
  sh.getRange(3, 1, rows.length, 2).setBackground(THEME.lightGray);

  const dashRow = 3 + rows.length + 2;
  sh.getRange(dashRow, 1, 1, 2).merge().setValue('DASHBOARD')
    .setFontWeight('bold').setBackground(THEME.lightBlue).setFontColor(THEME.darkGray);

  const settings = readSettingsMap_();
  const termStart = TERM_START_DATES[String(settings['Term # (1, 2, or 3)'])];
  const week5Date = termStart ? new Date(termStart.getTime() + 5 * 7 * 86400000) : '';
  const week10Date = termStart ? new Date(termStart.getTime() + 10 * 7 * 86400000) : '';
  const fmt = d => d ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'MM/dd/yyyy') : '—';

  const dashLabels = [
    'Number of students enrolled',
    'Action Plan 1 — still to generate',
    'Progress Report — still to generate',
    'Action Plan 2 — still to generate',
    `Action Plan 1 deadline (Week 5): ${fmt(week5Date)}`,
    `Progress Report deadline (Week 5): ${fmt(week5Date)}`,
    `Action Plan 2 deadline (Week 10): ${fmt(week10Date)}`,
    'Workbook completion',
  ];
  const week1LinkCol = columnToLetter_(WEEK1_COLS.ACTION_PLAN_LINK);
  const week10LinkCol = columnToLetter_(WEEK10_COLS.ACTION_PLAN_PDF_LINK);
  const dashFormulas = [
    `=COUNTA('Week 1'!A4:A203)`,
    `=COUNTIFS('Week 1'!A4:A203,"<>",'Week 1'!${week1LinkCol}4:${week1LinkCol}203,"")`,
    `=COUNTIFS('Progress Report Entries'!A2:A201,"<>",'Progress Report Entries'!O2:O201,"<>Done")`,
    `=COUNTIFS('Week 10'!A4:A203,"<>",'Week 10'!${week10LinkCol}4:${week10LinkCol}203,"")`,
    '', '', '',
    ''
  ];

  sh.getRange(dashRow + 1, 1, dashLabels.length, 1).setValues(dashLabels.map(l => [l]));
  sh.getRange(dashRow + 1, 1, dashLabels.length, 1).setFontWeight('bold');
  dashFormulas.forEach((f, i) => { if (f) sh.getRange(dashRow + 1 + i, 2).setFormula(f); });

  const totalRow = dashRow + 1, ap1Row = dashRow + 2, prRow = dashRow + 3, ap2Row = dashRow + 4, completionRow = dashRow + 8;
  sh.getRange(completionRow, 2).setFormula(
    `=IFERROR(ROUND((1-(B${ap1Row}+B${prRow}+B${ap2Row})/(3*B${totalRow}))*100,0)&"%","—")`
  );
}

// ---------- Progress Report Entries tab (teacher selects, AD assembles) ----------

const PR_SUBJECT_COLUMNS = {
  'English': 1, 'Mathematics': 2, 'Math': 2, 'Arabic': 3, 'Science': 4, 'Social Studies': 5
};

const PR_STATEMENTS = [
  'Arrives on time and attends regularly',
  'Comes prepared for the Flipped Classroom',
  'Submits assignments on time, original & complete',
  'Correctly answers Exit Tickets',
  'Scores well on quizzes',
  'Remains on task, does not distract classmates',
  'Attends the remedial hours',
  'On the right track / shows progress',
  'Needs to exert more effort',
  'Attentive, follows directions and guidance',
  'Treats the teacher with respect',
  'Self-regulates / takes responsibility for learning'
];

const PR_RESPONSE_OPTIONS = ['Always', 'Usually', 'Sometimes', 'Rarely'];

function buildTeacherProgressReportTab_(ss, teacherName) {
  const sh = ss.insertSheet('Progress Report Entries');
  const headers = ['Student ID', 'Student Name', ...PR_STATEMENTS, 'Status', 'Progress Report PDF Link'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground(THEME.darkBlue).setFontColor(THEME.white);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);

  PR_STATEMENTS.forEach((s, i) => sh.getRange(1, 3 + i).setNote(s));
  sh.getRange(1, 3 + PR_STATEMENTS.length).setNote(
    'Set to "Done" once you\'re sure of your answers for that student. No button to click — the ' +
    'Academic Dean\'s Progress Report only pulls in rows marked Done, whenever she runs it.'
  );

  const rosterSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ROSTER);
  const lastRow = rosterSh.getLastRow();
  const students = [];
  if (lastRow > 1) {
    rosterSh.getRange(2, 1, lastRow - 1, ROSTER_HEADERS.length).getValues().forEach(row => {
      const teacherField = row[ROSTER_COL.TEACHER - 1];
      if (!teacherField || !String(teacherField).startsWith(teacherName)) return;
      students.push([row[ROSTER_COL.ID - 1], row[ROSTER_COL.NAME - 1]]);
    });
  }
  if (students.length) {
    const rows = students.map(s => [s[0], s[1], ...PR_STATEMENTS.map(() => ''), 'Pending', '']);
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
    const respRange = sh.getRange(2, 3, rows.length, PR_STATEMENTS.length);
    respRange.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(PR_RESPONSE_OPTIONS).build());
    const statusRange = sh.getRange(2, 3 + PR_STATEMENTS.length, rows.length, 1);
    statusRange.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Pending', 'Done']).build());
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Done').setBackground(THEME.statusGreen)
      .setRanges([statusRange]).build();
    const pendingRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Pending').setBackground(THEME.statusPurple)
      .setRanges([statusRange]).build();
    sh.setConditionalFormatRules([rule, pendingRule]);

    sh.getRange(1, 1, rows.length + 1, headers.length).setBorder(true, true, true, true, true, true);
    sh.getRange(2, 1, rows.length, 1).setHorizontalAlignment('center');
    respRange.setHorizontalAlignment('center');
    statusRange.setHorizontalAlignment('center');
    sh.getRange(1, 1, 1, headers.length).setHorizontalAlignment('center');
  }

  sh.setColumnWidth(1, 90);
  sh.setColumnWidth(2, 160);
  for (let i = 0; i < PR_STATEMENTS.length; i++) sh.setColumnWidth(3 + i, 70);
  sh.setColumnWidth(3 + PR_STATEMENTS.length, 90);
  sh.setColumnWidth(4 + PR_STATEMENTS.length, 240);
  sh.autoResizeColumns(1, headers.length);
}

// ---------- Week 1 / 5 / 10 tabs ----------

function buildTeacherWeekTab_(ss, weekNum, teacherName) {
  const sh = ss.insertSheet(`Week ${weekNum}`);
  const settings = readSettingsMap_();

  const teacherRowForSubject = teacherRosterRowByName_(teacherName);
  const isNoET = teacherRowForSubject && NO_ET_SUBJECTS.indexOf(teacherRowForSubject.subject) !== -1;
  const qWeeks = teacherRowForSubject ? quizWeeksForSubject_(teacherRowForSubject.subject) : QUIZ_WEEKS.hasET;

  let title, columns;
  if (weekNum === 1) {
    title = 'Week 1 — Action Plan 1';
    columns = ['Student ID', 'Student Name', 'Grade & Class', 'Spotted Weakness 1',
               'Steps to Tackle Weakness 1', 'Spotted Weakness 2', 'Steps to Tackle Weakness 2',
               'Means of Assessment', 'Evidence of Improvement',
               'Review Date (Week 5)', 'Action Plan Link', 'Generate?', 'Status'];
  } else if (weekNum === 5) {
    title = 'Week 5 — Action Plan 1 Reflection & Action Plan 2';
    columns = ['Student ID', 'Student Name', 'Grade & Class', 'Starting Score', 'Cumulative Score 1',
               'W1 ET', 'W1 Present?', 'W2 ET', 'W2 Present?', 'W3 ET', 'W3 Present?',
               'W4 ET', 'W4 Present?', 'W5 ET', 'W5 Present?',
               `QZ1 (Week ${qWeeks[0]})`, `QZ2 (Week ${qWeeks[1]})`,
               'Written Summary',
               'Spotted Weakness 1 (AP2)', 'Steps to Tackle Weakness 1 (AP2)',
               'Spotted Weakness 2 (AP2)', 'Steps to Tackle Weakness 2 (AP2)',
               'Means of Assessment (AP2)', 'Evidence of Improvement (AP2)',
               'Progress Report Status', 'Action Plan Link', 'Generate?', 'Status'];
  } else {
    title = 'Week 10 — Action Plan 2 Closeout';
    columns = ['Student ID', 'Student Name', 'Grade & Class', 'Cumulative Score 2',
               'W6 ET', 'W6 Present?', 'W7 ET', 'W7 Present?', 'W8 ET', 'W8 Present?',
               'W9 ET', 'W9 Present?', 'W10 ET', 'W10 Present?',
               `QZ3 (Week ${qWeeks[2]})`, `QZ4 (Week ${qWeeks[3]})`,
               'Written Summary', 'Action Plan PDF Link',
               'Generate?', 'Status'];
  }

  // FIX (2026-09-24): row 1 and row 2 are no longer merged across every
  // column. Merging across the frozen/unfrozen column boundary is blocked
  // by Google Sheets in BOTH directions — freezing across an existing merge
  // ("Sorry, you can't freeze columns which contain only part of a merged
  // cell") AND merging across an existing freeze ("You can't merge frozen
  // and non-frozen columns") — so reordering freeze/merge alone doesn't
  // work either way. The title (A1) and subtitle (A2) text now just
  // overflows naturally across the empty cells to its right instead of
  // being merged, which sidesteps the conflict entirely. Matches the same
  // fix already applied in the British project's TeacherWorkbook.gs (2026-09-21).
  sh.getRange('A1').setValue(title).setFontWeight('bold').setFontSize(13).setFontColor(THEME.darkNavy);
  sh.getRange('A2').setValue('Fill in your fields for a student, then tick the "Generate?" box at the end of ' +
    'that row. Watch the Status column next to it — it updates instantly.')
    .setFontStyle('italic').setFontColor('#666666');
  sh.getRange(3, 1, 1, columns.length).setValues([columns])
    .setFontWeight('bold').setBackground(THEME.darkBlue).setFontColor(THEME.white);
  sh.setFrozenRows(3);
  sh.setFrozenColumns(3);

  const rosterSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ROSTER);
  const lastRow = rosterSh.getLastRow();
  const students = [];
  if (lastRow > 1) {
    rosterSh.getRange(2, 1, lastRow - 1, ROSTER_HEADERS.length).getValues().forEach(row => {
      const teacherField = row[ROSTER_COL.TEACHER - 1];
      if (!teacherField || !String(teacherField).startsWith(teacherName)) return;
      students.push({
        id: row[ROSTER_COL.ID - 1], name: row[ROSTER_COL.NAME - 1], cls: gradeClassLabel_(row[ROSTER_COL.GRADE - 1], row[ROSTER_COL.CLASS - 1]),
        grade: row[ROSTER_COL.GRADE - 1],
          starting: row[ROSTER_COL.STARTING_SCORE - 1]
      });
    });
  }

  if (students.length) {
    const termStart = TERM_START_DATES[String(settings['Term # (1, 2, or 3)'])];
    const reviewDate = termStart ? new Date(termStart.getTime() + 5 * 7 * 86400000) : '';

    const rows = students.map(s => {
      if (weekNum === 1) return [s.id, s.name, s.cls, '', '', '', '', '', '', reviewDate, '', false, ''];
      const weeklyBlanks = Array(10).fill('');
      if (weekNum === 5) {
        return [s.id, s.name, s.cls, s.starting, '', ...weeklyBlanks, '', '', '', '', '', '', '', '', '', 'Pending', '', false, ''];
      }
      return [s.id, s.name, s.cls, '', ...weeklyBlanks, '', '', '', '', false, ''];
    });
    sh.getRange(4, 1, rows.length, columns.length).setValues(rows);

    const checkboxCol = columns.length - 1;
    sh.getRange(4, checkboxCol, rows.length, 1).insertCheckboxes();

    const linkCol = checkboxCol - 1;
    const linkColLetter = columnToLetter_(linkCol);
    const checkboxColLetter = columnToLetter_(checkboxCol);
    const statusFormulas = [];
    for (let i = 0; i < rows.length; i++) {
      const r = 4 + i;
      statusFormulas.push([
        `=IF(${linkColLetter}${r}<>"", "✅ Done", IF(${checkboxColLetter}${r}=TRUE, "⏳ Generating... (usually within 1 min)", ""))`
      ]);
    }
    sh.getRange(4, columns.length, rows.length, 1).setFormulas(statusFormulas);

    if (weekNum === 5 || weekNum === 10) {
      const c = weekNum === 5 ? WEEK5_COLS : WEEK10_COLS;
      const presentCols = weekNum === 5
        ? [c.W1_PRESENT, c.W2_PRESENT, c.W3_PRESENT, c.W4_PRESENT, c.W5_PRESENT]
        : [c.W6_PRESENT, c.W7_PRESENT, c.W8_PRESENT, c.W9_PRESENT, c.W10_PRESENT];
      presentCols.forEach(col => sh.getRange(4, col, rows.length, 1).insertCheckboxes());

      if (isNoET) {
        const etCols = weekNum === 5
          ? [c.W1_ET, c.W2_ET, c.W3_ET, c.W4_ET, c.W5_ET]
          : [c.W6_ET, c.W7_ET, c.W8_ET, c.W9_ET, c.W10_ET];
        etCols.forEach(col => sh.getRange(4, col, rows.length, 1).setBackground(THEME.lightGray));
      }

      // Grades 1-2 have no Exit Ticket: ET cells default to N/A (greyed) for those students only.
      const naEtCols = weekNum === 5
        ? [c.W1_ET, c.W2_ET, c.W3_ET, c.W4_ET, c.W5_ET]
        : [c.W6_ET, c.W7_ET, c.W8_ET, c.W9_ET, c.W10_ET];
      students.forEach((s, i) => {
        if (!isNoETGrade_(s.grade)) return;
        naEtCols.forEach(col => sh.getRange(4 + i, col).setValue('N/A').setBackground(THEME.lightGray));
      });

      if (weekNum === 5) {
        const progressStatusRange = sh.getRange(4, WEEK5_COLS.PROGRESS_REPORT_STATUS, rows.length, 1);
        progressStatusRange.setDataValidation(
          SpreadsheetApp.newDataValidation().requireValueInList(['Pending', 'Filled', 'Reviewed by AD']).build()
        );
        const existingRules = sh.getConditionalFormatRules();
        existingRules.push(
          SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Pending').setBackground(THEME.statusPurple).setRanges([progressStatusRange]).build(),
          SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Filled').setBackground(THEME.statusYellow).setRanges([progressStatusRange]).build(),
          SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Reviewed by AD').setBackground(THEME.statusGreen).setRanges([progressStatusRange]).build()
        );
        sh.setConditionalFormatRules(existingRules);
      }
    }

    sh.getRange(3, 1, rows.length + 1, columns.length).setBorder(true, true, true, true, true, true);
    sh.getRange(4, 1, rows.length, columns.length).setHorizontalAlignment('center');
    sh.getRange(3, 1, 1, columns.length).setHorizontalAlignment('center');
    const leftAlignColumns = { 2: true };
    if (weekNum === 1) { [4, 5, 6, 7, 8, 9, 11].forEach(col => leftAlignColumns[col] = true); }
    if (weekNum === 5) {
      [WEEK5_COLS.WRITTEN_SUMMARY, WEEK5_COLS.WEAK1_AP2, WEEK5_COLS.STEPS1_AP2,
       WEEK5_COLS.WEAK2_AP2, WEEK5_COLS.STEPS2_AP2,
       WEEK5_COLS.MEANS_AP2, WEEK5_COLS.EVIDENCE_AP2, WEEK5_COLS.ACTION_PLAN_LINK]
        .forEach(col => leftAlignColumns[col] = true);
    }
    if (weekNum === 10) {
      [WEEK10_COLS.WRITTEN_SUMMARY, WEEK10_COLS.ACTION_PLAN_PDF_LINK]
        .forEach(col => leftAlignColumns[col] = true);
    }
    Object.keys(leftAlignColumns).forEach(col => sh.getRange(4, Number(col), rows.length, 1).setHorizontalAlignment('left'));
  }

  // Generate guard built in: live "⚠️ x field(s) empty" status + red empty required cells.
  applyGuardToTab_(sh, `Week ${weekNum}`, blankPatternRe2_());

  sh.autoResizeColumns(1, columns.length);
}

// ---------- Protection: AD / Teacher split ----------



/** Puts a teacher workbook's tabs in the standard order: Profile, Week 1,
 * Week 5, Progress Report Entries, Week 10. Safe to call on any workbook —
 * skips any tab that's missing. */
function orderTeacherWorkbookTabs_(ss) {
  const order = ['Profile', 'Week 1', 'Week 5', 'Progress Report Entries', 'Week 10'];
  order.forEach((name, i) => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    sh.activate();
    ss.moveActiveSheet(i + 1);
  });
}



