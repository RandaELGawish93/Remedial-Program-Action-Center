/**
 * SETUP_SCORESPROGRESS.gs
 * Per-student, per-subject score + attendance tracking. Unlike the old
 * version, this is fully SCRIPT-POPULATED (not a live QUERY formula) —
 * attendance/absence/cumulative scores are pulled from each teacher's own
 * workbook, and a QUERY formula next to manually-typed values would
 * misalign every time the underlying row set changes. Run "Refresh Scores
 * Progress" any time you want the latest attendance/cumulative data pulled in.
 *
 * Final Score /65 is the one field YOU enter — it's the official
 * grade determination, not something derivable from teacher workbooks.
 * Refreshing preserves whatever you've already typed there.
 */

const SCORES_HEADERS = [
  'Student ID', 'Student Name', 'Grade', 'Class', 'Subject',
  'Remedial Teacher', 'Advisor', 'Starting Score',
  'Cumulative 1 (W5)', 'Cumulative 2 (W10)', 'Final Score /65 (%)'
];

function buildScoresProgress() {
  const sh = getOrCreateSheet_(SHEETS.SCORES);
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  sh.clear();

  sh.getRange('A1').setValue('Filter by Subject:').setFontWeight('bold');
  sh.getRange('B1').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(SpreadsheetApp.getActiveSpreadsheet()
        .getSheetByName(SHEETS.ROSTER).getRange('E2:E301'), true)
      .build()
  );
  sh.getRange('D1').setValue('Filter by Grade:').setFontWeight('bold');
  sh.getRange('E1').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(
      ['All', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12']
    ).build()
  ).setValue('All');
  sh.getRange('G1').setValue('Filter by Teacher:').setFontWeight('bold');

  const headerRow = 3;
  sh.getRange(headerRow, 1, 1, SCORES_HEADERS.length)
    .setValues([SCORES_HEADERS])
    .setFontWeight('bold').setBackground('#1c4587').setFontColor('#ffffff');
  sh.setFrozenRows(headerRow);

  const widths = [90, 170, 60, 70, 130, 150, 150, 100, 120, 120, 130];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  const formula =
    `=IFERROR(QUERY(_ScoresRaw!A2:K301,` +
    `"select * where A is not null" & ` +
    `IF($B$1<>"", " and E = '"&$B$1&"'", "") & ` +
    `IF(AND($E$1<>"", $E$1<>"All"), " and C = "&$E$1, "") & ` +
    `IF($G$1<>"", " and F = '"&$G$1&"'", ""), 0), "")`;
  sh.getRange(headerRow + 1, 1).setFormula(formula);

  const dataBlockRows = 300;
  sh.getRange(headerRow, 1, dataBlockRows + 1, SCORES_HEADERS.length).setBorder(true, true, true, true, true, true);
  [1, 3, 4, 8, 9, 10, 11].forEach(col => {
    sh.getRange(headerRow + 1, col, dataBlockRows, 1).setHorizontalAlignment('center');
  });
  sh.getRange(headerRow, 1, 1, SCORES_HEADERS.length).setHorizontalAlignment('center');
  sh.autoResizeColumns(1, SCORES_HEADERS.length);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Scores Progress built. Run "Refresh Scores Progress" to pull cumulative scores in from teacher workbooks.', 'Done', 7
  );
}
/**
 * Pulls cumulative scores from every teacher's own workbook, and writes
 * the full row set into the hidden _ScoresRaw sheet. Final Score /65 is
 * preserved across refreshes since it's the one field you type yourself.
 */
function refreshScoresProgress() {
  const rosterSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ROSTER);
  const lastRow = rosterSh.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('Student Roster is empty.'); return; }
  const rosterData = rosterSh.getRange(2, 1, lastRow - 1, ROSTER_HEADERS.length).getValues();

  const rawSh = getOrCreateSheet_('_ScoresRaw');
  const existingFinal = {};
  const rawLastRow = rawSh.getLastRow();
  if (rawLastRow > 1) {
    rawSh.getRange(2, 1, rawLastRow - 1, SCORES_HEADERS.length).getValues().forEach(r => {
      if (r[0]) existingFinal[r[0] + '|' + r[4]] = r[10];
    });
  }

  const teacherDataCache = {};
  function getTeacherData(teacherName) {
    if (teacherDataCache[teacherName] !== undefined) return teacherDataCache[teacherName];
    const teacherRow = teacherRosterRowByName_(teacherName);
    if (!teacherRow || !teacherRow.workbookLink) { teacherDataCache[teacherName] = null; return null; }
    let ss;
    try { ss = SpreadsheetApp.openById(extractFileIdFromUrl_(teacherRow.workbookLink)); }
    catch (e) { teacherDataCache[teacherName] = null; return null; }

    const week5Sh = ss.getSheetByName('Week 5');
    const week10Sh = ss.getSheetByName('Week 10');

    const cumulative1 = {};
    if (week5Sh && week5Sh.getLastRow() > 3) {
      week5Sh.getRange(4, 1, week5Sh.getLastRow() - 3, 5).getValues().forEach(r => {
        if (r[0]) cumulative1[r[0]] = r[4];
      });
    }
    const cumulative2 = {};
    if (week10Sh && week10Sh.getLastRow() > 3) {
      week10Sh.getRange(4, 1, week10Sh.getLastRow() - 3, 4).getValues().forEach(r => {
        if (r[0]) cumulative2[r[0]] = r[3];
      });
    }

    const result = { cumulative1, cumulative2 };
    teacherDataCache[teacherName] = result;
    return result;
  }

  const outRows = [];
  rosterData.forEach(row => {
    const id = row[ROSTER_COL.ID - 1], subject = row[ROSTER_COL.SUBJECT - 1];
    const teacherField = String(row[ROSTER_COL.TEACHER - 1] || '');
    const m = teacherField.match(/^(.*?)\s*\(/);
    const teacherName = m ? m[1].trim() : teacherField;
    const teacherData = getTeacherData(teacherName);

    const cum1 = teacherData ? teacherData.cumulative1[id] : '';
    const cum2 = teacherData ? teacherData.cumulative2[id] : '';
    const finalScore = existingFinal[id + '|' + subject] || '';

    outRows.push([
      id, row[ROSTER_COL.NAME - 1], row[ROSTER_COL.GRADE - 1], row[ROSTER_COL.CLASS - 1], subject,
      teacherName, row[ROSTER_COL.ADVISOR - 1], row[ROSTER_COL.STARTING_SCORE - 1],
      cum1 || '', cum2 || '', finalScore
    ]);
  });

  rawSh.clear();
  rawSh.getRange(1, 1, 1, SCORES_HEADERS.length).setValues([SCORES_HEADERS]);
  if (outRows.length) rawSh.getRange(2, 1, outRows.length, SCORES_HEADERS.length).setValues(outRows);
  rawSh.hideSheet();

  SpreadsheetApp.getActiveSpreadsheet().toast(`Scores Progress refreshed — ${outRows.length} row(s) pulled in.`, 'Done', 6);
}

