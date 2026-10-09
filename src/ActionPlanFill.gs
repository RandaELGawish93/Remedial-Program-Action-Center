/**
 * ACTIONPLANFILL.gs
 * Fills the Action Plan template — both the English and Arabic Drive
 * templates now share the IDENTICAL 18-table structure (verified directly
 * against both live template files):
 *
 *  TABLE[0]  header (Student Name / Class / Subject / Teacher)
 *  TABLE[1]  AP1 Spotted Weakness 1 / Steps 1 (only "1" used, per spec)
 *  TABLE[2]  AP1 Spotted Weakness 2 / Steps 2 (left blank — not used)
 *  TABLE[3]  Means of Assessment (AP1)
 *  TABLE[4]  Evidence of Improvement (AP1)
 *  TABLE[5]  Cumulative Result (Week 5 — Recap)
 *  TABLE[6]  Review Date (Week 5)
 *  TABLE[7]  Weekly scores W1-W5 (ET / Quiz / Absences / Presences / Attendance %)
 *  TABLE[8]  Written Summary #1 (real 1x1 table)
 *  TABLE[9]  Graph #1 (real 1x1 table — chart image goes inside)
 *  TABLE[10] AP2 Spotted Weakness 1 / Steps 1
 *  TABLE[11] AP2 Spotted Weakness 2 / Steps 2 (left blank)
 *  TABLE[12] Means of Assessment (AP2)
 *  TABLE[13] Evidence of Improvement (AP2)
 *  TABLE[14] Weekly scores W6-W10
 *  TABLE[15] Written Summary #2 (real 1x1 table)
 *  TABLE[16] Graph #2 (real 1x1 table)
 *  TABLE[17] row0 Cumulative Result (Week 10) / row1 Review Date (Week 10)
 *
 * IMPORTANT: these indices are fragile — if the template's table order or
 * count ever changes again, this breaks silently (wrong cell gets
 * written). Re-map the indices whenever the template changes.
 *
 * Subjects with no Exit Ticket (Arabic Social Studies, Islamic Religion)
 * get a quiz-only score row. Every other subject gets both ET and Quiz.
 * Quiz is biweekly, not weekly — QUIZ_WEEKS (Config.gs) says which two
 * weeks per phase actually have a quiz value.
 */

const NO_ET_SUBJECTS = ['Arabic Social Studies', 'Islamic Religion'];

// Both templates share the same 18-table layout now, so both variants
// use identical indices — kept as two keys so callers don't need to change.
const ACTION_PLAN_TABLES = {
  en: { weak1: 1, weak2: 2, means1: 3, evidence1: 4, cumWk5: 5, reviewWk5: 6,
        weeklyW1_5: 7, summary1: 8, graph1: 9, weak1AP2: 10, weak2AP2: 11,
        meansAP2: 12, evidenceAP2: 13, weeklyW6_10: 14, summary2: 15, graph2: 16,
        cumWk10: 17, summaryIsTable: true },
  ar: { weak1: 1, weak2: 2, means1: 3, evidence1: 4, cumWk5: 5, reviewWk5: 6,
        weeklyW1_5: 7, summary1: 8, graph1: 9, weak1AP2: 10, weak2AP2: 11,
        meansAP2: 12, evidenceAP2: 13, weeklyW6_10: 14, summary2: 15, graph2: 16,
        cumWk10: 17, summaryIsTable: true }
};

/** Determines which template variant a subject uses. */
function templateVariantForSubject_(subject) {
  return ARABIC_TEMPLATE_SUBJECTS.indexOf(subject) !== -1 ? 'ar' : 'en';
}

/** Converts a 1-based column number to its A1-notation letter(s), e.g. 27 -> "AA". */
function columnToLetter_(col) {
  let letter = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

// ---------- Header (identical position, both variants) ----------

function fillActionPlanHeader_(doc, studentName, cls, subject, teacherName) {
  const t = doc.getBody().getTables()[0];
  t.getRow(0).getCell(1).setText(studentName || '');
  t.getRow(1).getCell(1).setText(cls || '');
  t.getRow(2).getCell(1).setText(subject || '');
  t.getRow(3).getCell(1).setText(teacherName || '');
}

// ---------- Weakness / Steps (single weakness field only, per spec) ----------

function fillWeaknessTable_(doc, tableIndex, weakness, steps) {
  const t = doc.getBody().getTables()[tableIndex];
  t.getRow(0).getCell(1).setText(weakness || '');
  t.getRow(1).getCell(1).setText(steps || '');
}

// ---------- Means of Assessment / Evidence of Improvement (single blank cell) ----------

function fillSingleCellTable_(doc, tableIndex, text) {
  doc.getBody().getTables()[tableIndex].getRow(0).getCell(0).setText(text || '');
}

// ---------- Cumulative Result + Review Date ----------

function fillCumulativeAndReviewWeek5_(doc, tables, cumulativeResult, reviewDate) {
  const t = doc.getBody().getTables();
  t[tables.cumWk5].getRow(0).getCell(1).setText(cumulativeResult !== '' && cumulativeResult != null ? String(cumulativeResult) : '');
  t[tables.reviewWk5].getRow(0).getCell(1).setText(formatDateCell_(reviewDate));
}

function fillCumulativeAndReviewWeek10_(doc, tables, cumulativeResult, reviewDate) {
  const t = doc.getBody().getTables()[tables.cumWk10];
  t.getRow(0).getCell(1).setText(cumulativeResult !== '' && cumulativeResult != null ? String(cumulativeResult) : '');
  t.getRow(1).getCell(1).setText(formatDateCell_(reviewDate));
}

function formatDateCell_(d) {
  if (!d) return '';
  const date = (d instanceof Date) ? d : new Date(d);
  if (isNaN(date.getTime())) return String(d);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd/yyyy');
}

// ---------- Weekly scores table (ET / Quiz / Absences / Presences / Attendance %) ----------

/**
 * weeklyData: array of 5 {et, quiz, present} objects, present = 1 or 0.
 * quiz is only set on the 2 real quiz weeks for that subject — other
 * weeks leave it undefined, which renders as a blank Quiz cell here.
 * Attendance % is computed automatically, never typed by the teacher.
 * For NO_ET_SUBJECTS, the ET Score column is left as "N/A" for every week.
 */
function fillWeeklyScoresTable_(doc, tableIndex, weeklyData, subject) {
  const t = doc.getBody().getTables()[tableIndex];
  const noET = NO_ET_SUBJECTS.indexOf(subject) !== -1 || etNotApplicable_(weeklyData);

  for (let i = 0; i < 5; i++) {
    const row = t.getRow(i + 1); // row 0 is the header row
    const d = weeklyData[i] || {};
    const present = d.present === 1 || d.present === true;
    const absences = present ? 0 : 1;
    const presences = present ? 1 : 0;

    row.getCell(1).setText(noET ? 'N/A' : (d.et !== undefined && d.et !== '' ? String(d.et) : ''));
    row.getCell(2).setText(d.quiz !== undefined && d.quiz !== '' ? String(d.quiz) : '');
    row.getCell(3).setText(String(absences));
    row.getCell(4).setText(String(presences));
    row.getCell(5).setText(present ? '100%' : '0%');
  }
}

/** Overall attendance % across the 5 weeks in weeklyData (for the running Roster/Scores view, not the doc). */
function calcAttendancePercent_(weeklyData) {
  const total = weeklyData.filter(d => d && d.present !== undefined).length;
  if (!total) return '';
  const present = weeklyData.filter(d => d && (d.present === 1 || d.present === true)).length;
  return Math.round((present / total) * 100) + '%';
}

// ---------- Written Summary ----------

/** Both variants now use a real 1x1 table for this — fill directly into it. */
function fillWrittenSummary_(doc, tables, phase, summaryText) {
  if (tables.summaryIsTable) {
    const tableIndex = phase === 1 ? tables.summary1 : tables.summary2;
    doc.getBody().getTables()[tableIndex].getRow(0).getCell(0).setText(summaryText || '');
    return;
  }
  // Fallback for any older template still using bare paragraphs instead of a table.
  const occurrenceIndex = phase === 1 ? 0 : 1;
  const body = doc.getBody();
  const paragraphs = body.getParagraphs();
  let count = 0;
  for (let i = 0; i < paragraphs.length; i++) {
    if (paragraphs[i].getText().trim() === 'Written Summary') {
      if (count === occurrenceIndex) {
        const childIndex = body.getChildIndex(paragraphs[i]);
        body.insertParagraph(childIndex + 1, summaryText || '');
        return;
      }
      count++;
    }
  }
}

// ---------- Graph (combined ET + Quiz line chart) ----------

/** Builds ONE line chart with Exit Ticket and Quiz overlaid as two series
 * across the 5-week window. Quiz only has real values on its 2 actual
 * quiz weeks (biweekly) — the rest are null, and interpolateNulls draws a
 * single straight line between the two known points instead of gaps. */
function buildOverlaidLineChartBlob_(weekLabels, etValues, quizValues, noET) {
  const dataTable = Charts.newDataTable().addColumn(Charts.ColumnType.STRING, 'Week');
  if (!noET) dataTable.addColumn(Charts.ColumnType.NUMBER, 'Exit Ticket');
  dataTable.addColumn(Charts.ColumnType.NUMBER, 'Quiz');

  for (let i = 0; i < weekLabels.length; i++) {
    const row = [weekLabels[i]];
    if (!noET) row.push((etValues[i] === null || etValues[i] === undefined || etValues[i] === '') ? null : Number(etValues[i]));
    row.push((quizValues[i] === null || quizValues[i] === undefined || quizValues[i] === '') ? null : Number(quizValues[i]));
    dataTable.addRow(row);
  }

  return Charts.newLineChart()
    .setDataTable(dataTable.build())
    .setTitle('Exit Ticket & Quiz Scores')
    .setDimensions(420, 260)
    .setPointStyle(Charts.PointStyle.MEDIUM)
    .setLegendPosition(Charts.Position.BOTTOM)
    .setOption('interpolateNulls', true)
    .build()
    .getAs('image/png');
}

/** Both variants have a real 1x1 table for the graph — insert the chart image directly inside it. */
function insertGraphAtPlaceholder_(doc, tables, phase, weeklyData, subject, weekLabels) {
  const noET = NO_ET_SUBJECTS.indexOf(subject) !== -1 || etNotApplicable_(weeklyData);
  const etValues = weeklyData.map(d => (d && d.et !== undefined && d.et !== '') ? d.et : null);
  const quizValues = weeklyData.map(d => (d && d.quiz !== undefined && d.quiz !== '') ? d.quiz : null);
  const chartBlob = buildOverlaidLineChartBlob_(weekLabels, etValues, quizValues, noET);

  if (tables.summaryIsTable) {
    const tableIndex = phase === 1 ? tables.graph1 : tables.graph2;
    const cell = doc.getBody().getTables()[tableIndex].getRow(0).getCell(0);
    cell.clear();
    cell.appendImage(chartBlob);
    return;
  }

  // Fallback for any older template still using bare paragraphs instead of a table.
  const occurrenceIndex = phase === 1 ? 0 : 1;
  const body = doc.getBody();
  const paragraphs = body.getParagraphs();
  let count = 0;
  let targetIndex = -1;
  for (let i = 0; i < paragraphs.length; i++) {
    if (paragraphs[i].getText().trim() === 'Graph (insert chart image here)') {
      if (count === occurrenceIndex) { targetIndex = body.getChildIndex(paragraphs[i]); break; }
      count++;
    }
  }
  if (targetIndex === -1) return;
  body.insertImage(targetIndex + 1, chartBlob);
}

// ---------- Entry points called from the teacher workbook (via checkboxes) ----------

/**
 * Week 1 checkbox handler: creates the Action Plan doc (if it doesn't
 * exist yet) and fills the header + AP1 weakness/means/evidence/review-date.
 */
function generateOrFillActionPlan1FromRow_(ss, row) {
  const sh = ss.getSheetByName('Week 1');
  const c = WEEK1_COLS;
  assertRowComplete_(sh, row, 'Week 1');
  const existingLink = sh.getRange(row, c.ACTION_PLAN_LINK).getValue();
  const teacherName = extractTeacherNameFromWorkbook_(ss);
  const teacherRow = teacherRosterRowByName_(teacherName);
  if (!teacherRow) {
    throw new Error(`No Teacher Roster entry found matching "${teacherName}" — run "Refresh Teacher Roster" on the AD side, then try again.`);
  }
  const studentName = sh.getRange(row, c.NAME).getValue();
  const studentClass = sh.getRange(row, c.CLASS).getValue();
  const weakness1 = sh.getRange(row, c.WEAK1).getValue();
  const steps1 = sh.getRange(row, c.STEPS1).getValue();
  const weakness2 = sh.getRange(row, c.WEAK2).getValue();
  const steps2 = sh.getRange(row, c.STEPS2).getValue();
  const meansOfAssessment = sh.getRange(row, c.MEANS).getValue();
  const evidence = sh.getRange(row, c.EVIDENCE).getValue();
  const reviewDateWeek5 = sh.getRange(row, c.REVIEW_DATE).getValue();

  const settings = readSettingsMap_();
  const variant = templateVariantForSubject_(teacherRow.subject);
  const tables = ACTION_PLAN_TABLES[variant];
  let doc, url, fileId;

  if (existingLink) {
    fileId = extractFileIdFromUrl_(existingLink);
    doc = DocumentApp.openById(fileId);
    url = existingLink;
  } else {
    const templateId = variant === 'ar'
      ? (settings['Action Plan Arabic Template ID'] || settings['Action Plan English Template ID'])
      : settings['Action Plan English Template ID'];
    const termFolder = getTermFolder_(settings['Action Plan Folder ID'], settings);
    const fileName = `${studentName} - ${teacherRow.subject} - Action Plan - ${teacherName}`;
    const copy = DriveApp.getFileById(templateId).makeCopy(fileName, termFolder);
    applyLinkSharing_(copy, settings, teacherRow.email);
    fileId = copy.getId();
    doc = DocumentApp.openById(fileId);
    url = copy.getUrl();
    sh.getRange(row, c.ACTION_PLAN_LINK).setValue(url);
  }

  fillActionPlanHeader_(doc, studentName, studentClass, teacherRow.subject, teacherName);
  fillWeaknessTable_(doc, tables.weak1, weakness1, steps1);
  fillWeaknessTable_(doc, tables.weak2, weakness2, steps2);
  fillSingleCellTable_(doc, tables.means1, meansOfAssessment);
  fillSingleCellTable_(doc, tables.evidence1, evidence);
  fillCumulativeAndReviewWeek5_(doc, tables, '', reviewDateWeek5);
  doc.saveAndClose();

  writeActionPlanLinkToMasterRoster_(studentName, teacherName, url);
  return url;
}

/**
 * Week 5 checkbox handler: fills AP1's weekly scores/summary/graph
 * (the "Report from Action Plan 1" section), the Week 5 cumulative
 * result, starts AP2 (weakness/means/evidence), AND exports the
 * Phase 1 PDF snapshot into the Action Plan 1 subfolder.
 */
function fillActionPlan2FromWeek5Row_(ss, row) {
  const sh = ss.getSheetByName('Week 5');
  const c = WEEK5_COLS;
  assertRowComplete_(sh, row, 'Week 5');
  const studentId = sh.getRange(row, c.ID).getValue();
  const link = findActionPlanLiveDocLink_(ss, studentId);
  if (!link) {
    throw new Error('No Action Plan found for this student yet — generate it from Week 1 first.');
  }
  const fileId = extractFileIdFromUrl_(link);
  const doc = DocumentApp.openById(fileId);
  const teacherName = extractTeacherNameFromWorkbook_(ss);
  const teacherRow = teacherRosterRowByName_(teacherName);
  if (!teacherRow) {
    throw new Error(`No Teacher Roster entry found matching "${teacherName}" — run "Refresh Teacher Roster" on the AD side, then try again.`);
  }
  const subject = teacherRow.subject;
  const variant = templateVariantForSubject_(subject);
  const tables = ACTION_PLAN_TABLES[variant];
  const quizWeeks = quizWeeksForSubject_(subject);
  const studentName = sh.getRange(row, c.NAME).getValue();

  const cumulative1 = sh.getRange(row, c.CUMULATIVE_1).getValue();
  const etCols = [c.W1_ET, c.W2_ET, c.W3_ET, c.W4_ET, c.W5_ET];
  const presentCols = [c.W1_PRESENT, c.W2_PRESENT, c.W3_PRESENT, c.W4_PRESENT, c.W5_PRESENT];
  const weeklyData = [];
  for (let w = 0; w < 5; w++) {
    weeklyData.push({
      et: sh.getRange(row, etCols[w]).getValue(),
      present: sh.getRange(row, presentCols[w]).getValue() ? 1 : 0
    });
  }
  weeklyData[quizWeeks[0] - 1].quiz = sh.getRange(row, c.QZ1).getValue();
  weeklyData[quizWeeks[1] - 1].quiz = sh.getRange(row, c.QZ2).getValue();

  // Grades 1-2 have no Exit Ticket: force ET to N/A in the doc table and leave it out of the graph.
  const gradeRow = findRosterRowForStudent_(studentName, teacherName);
  if (gradeRow && isNoETGrade_(gradeRow[ROSTER_COL.GRADE - 1])) weeklyData.forEach(d => { d.et = 'N/A'; });

  const writtenSummary = sh.getRange(row, c.WRITTEN_SUMMARY).getValue();
  const ap2Weakness1 = sh.getRange(row, c.WEAK1_AP2).getValue();
  const ap2Steps1 = sh.getRange(row, c.STEPS1_AP2).getValue();
  const ap2Weakness2 = sh.getRange(row, c.WEAK2_AP2).getValue();
  const ap2Steps2 = sh.getRange(row, c.STEPS2_AP2).getValue();
  const means2 = sh.getRange(row, c.MEANS_AP2).getValue();
  const evidence2 = sh.getRange(row, c.EVIDENCE_AP2).getValue();

  fillCumulativeAndReviewWeek5_(doc, tables, cumulative1, new Date());
  fillWeeklyScoresTable_(doc, tables.weeklyW1_5, weeklyData, subject);
  fillWrittenSummary_(doc, tables, 1, writtenSummary);
  insertGraphAtPlaceholder_(doc, tables, 1, weeklyData, subject, ['W1', 'W2', 'W3', 'W4', 'W5']);
  fillWeaknessTable_(doc, tables.weak1AP2, ap2Weakness1, ap2Steps1);
  fillWeaknessTable_(doc, tables.weak2AP2, ap2Weakness2, ap2Steps2);
  fillSingleCellTable_(doc, tables.meansAP2, means2);
  fillSingleCellTable_(doc, tables.evidenceAP2, evidence2);
  doc.saveAndClose();

  sh.getRange(row, c.ACTION_PLAN_LINK).setValue(link);

  const settings = readSettingsMap_();
  const termFolder = getTermFolder_(settings['Action Plan Folder ID'], settings);
  const rosterRow = findRosterRowForStudent_(studentName, teacherName);
  const division = rosterRow ? (divisionFromGrade_(rosterRow[ROSTER_COL.GRADE - 1]) || 'MS') : 'MS';
  const divisionFolder = getDivisionSubfolder_(termFolder, division);
  const file = DriveApp.getFileById(fileId);
  const pdfBlob = file.getAs('application/pdf');
  const pdfFile = divisionFolder.createFile(pdfBlob).setName(file.getName() + ' - Action Plan (Week 5).pdf');
  applyLinkSharing_(pdfFile, settings, teacherRow.email);
  writeActionPlanLinkToMasterRoster_(sh.getRange(row, c.NAME).getValue(), teacherName, pdfFile.getUrl());

  return link;
}

/**
 * Week 10 checkbox handler: fills AP2's weekly scores/summary/graph
 * (the "Report from Action Plan 2" section) and the final Week 10
 * cumulative result + review date.
 */
function fillActionPlan10CloseoutFromRow_(ss, row) {
  const sh = ss.getSheetByName('Week 10');
  const c = WEEK10_COLS;
  assertRowComplete_(sh, row, 'Week 10');
  const teacherName = extractTeacherNameFromWorkbook_(ss);
  const teacherRow = teacherRosterRowByName_(teacherName);
  if (!teacherRow) {
    throw new Error(`No Teacher Roster entry found matching "${teacherName}" — run "Refresh Teacher Roster" on the AD side, then try again.`);
  }
  const subject = teacherRow.subject;
  const variant = templateVariantForSubject_(subject);
  const tables = ACTION_PLAN_TABLES[variant];
  const quizWeeks = quizWeeksForSubject_(subject);
  const studentId = sh.getRange(row, c.ID).getValue();
  const studentName = sh.getRange(row, c.NAME).getValue();

  const link = findActionPlanLiveDocLink_(ss, studentId);
  if (!link) {
    throw new Error('No Action Plan found for this student — it should have been created at Week 1.');
  }
  const fileId = extractFileIdFromUrl_(link);
  const doc = DocumentApp.openById(fileId);

  const cumulative2 = sh.getRange(row, c.CUMULATIVE_2).getValue();
  const etCols = [c.W6_ET, c.W7_ET, c.W8_ET, c.W9_ET, c.W10_ET];
  const presentCols = [c.W6_PRESENT, c.W7_PRESENT, c.W8_PRESENT, c.W9_PRESENT, c.W10_PRESENT];
  const weeklyData = [];
  for (let w = 0; w < 5; w++) {
    weeklyData.push({
      et: sh.getRange(row, etCols[w]).getValue(),
      present: sh.getRange(row, presentCols[w]).getValue() ? 1 : 0
    });
  }
  weeklyData[quizWeeks[2] - 6].quiz = sh.getRange(row, c.QZ3).getValue();
  weeklyData[quizWeeks[3] - 6].quiz = sh.getRange(row, c.QZ4).getValue();

  // Grades 1-2 have no Exit Ticket: force ET to N/A in the doc table and leave it out of the graph.
  const gradeRow = findRosterRowForStudent_(studentName, teacherName);
  if (gradeRow && isNoETGrade_(gradeRow[ROSTER_COL.GRADE - 1])) weeklyData.forEach(d => { d.et = 'N/A'; });

  const writtenSummary = sh.getRange(row, c.WRITTEN_SUMMARY).getValue();

  fillWeeklyScoresTable_(doc, tables.weeklyW6_10, weeklyData, subject);
  fillWrittenSummary_(doc, tables, 2, writtenSummary);
  insertGraphAtPlaceholder_(doc, tables, 2, weeklyData, subject, ['W6', 'W7', 'W8', 'W9', 'W10']);
  fillCumulativeAndReviewWeek10_(doc, tables, cumulative2, new Date());
  doc.saveAndClose();

  const settings = readSettingsMap_();
  const termFolder = getTermFolder_(settings['Action Plan Folder ID'], settings);
  const rosterRow = findRosterRowForStudent_(studentName, teacherName);
  const division = rosterRow ? (divisionFromGrade_(rosterRow[ROSTER_COL.GRADE - 1]) || 'MS') : 'MS';
  const divisionFolder = getDivisionSubfolder_(termFolder, division);
  const file = DriveApp.getFileById(fileId);
  const pdfBlob = file.getAs('application/pdf');
  const pdfFile = divisionFolder.createFile(pdfBlob).setName(file.getName() + ' - Action Plan (Final).pdf');
  applyLinkSharing_(pdfFile, settings, teacherRow.email);
  const pdfLink = pdfFile.getUrl();

  writeActionPlanLinkToMasterRoster_(studentName, teacherName, pdfLink);
  sh.getRange(row, c.ACTION_PLAN_PDF_LINK).setValue(pdfLink);

  return pdfLink;
}

/**
 * Finds the live Action Plan doc link within THIS teacher workbook —
 * checks Week 5 first (most likely to already be there), falls back to
 * Week 1. Never touches the AD's Roster, which only holds phase PDFs.
 */
function findActionPlanLiveDocLink_(ss, studentId) {
  const week5 = ss.getSheetByName('Week 5');
  if (week5 && week5.getLastRow() >= 4) {
    const linkCol = WEEK5_COLS.ACTION_PLAN_LINK;
    const data = week5.getRange(4, 1, week5.getLastRow() - 3, linkCol).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(studentId) && data[i][linkCol - 1]) return data[i][linkCol - 1];
    }
  }
  const week1 = ss.getSheetByName('Week 1');
  if (week1 && week1.getLastRow() >= 4) {
    const linkCol = WEEK1_COLS.ACTION_PLAN_LINK;
    const data = week1.getRange(4, 1, week1.getLastRow() - 3, linkCol).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(studentId) && data[i][linkCol - 1]) return data[i][linkCol - 1];
    }
  }
  return null;
}

/** Finds a Student Roster row (as raw array) matching a student+teacher. */
function findRosterRowForStudent_(studentName, teacherName) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ROSTER);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  const data = sh.getRange(2, 1, lastRow - 1, ROSTER_HEADERS.length).getValues();
  for (let i = 0; i < data.length; i++) {
    const name = data[i][ROSTER_COL.NAME - 1];
    const teacherField = data[i][ROSTER_COL.TEACHER - 1];
    if (name === studentName && String(teacherField).startsWith(teacherName)) return data[i];
  }
  return null;
}