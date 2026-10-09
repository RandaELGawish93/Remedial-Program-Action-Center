/**
 * SETUP_STUDENTROSTER.gs
 * The master list of every student in Intervention or Prevention.
 * This tab holds names/IDs — it gets owner-only protection at the end
 * of buildStudentRoster(). Nobody else should ever have edit OR view
 * access to this specific sheet if the workbook is ever shared.
 */

const ROSTER_HEADERS = [
  'Student ID', 'Student Name', 'Grade', 'Class', 'Subject',
  'Program', 'Tier', 'Advisor', 'Remedial Teacher',
  'Starting Score', 'Status', 'Threshold Score (auto)',
  'Probation Letter Link', 'Action Plan Link', 'Progress Report'
];

function buildStudentRoster() {
  const sh = getOrCreateSheet_(SHEETS.ROSTER);
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  buildThresholdsHelper_(); // hidden lookup sheet used by the formula below

  sh.getRange(1, 1, 1, ROSTER_HEADERS.length)
    .setValues([ROSTER_HEADERS])
    .setFontWeight('bold')
    .setBackground('#1c4587')
    .setFontColor('#ffffff');
  sh.setFrozenRows(1);

  // Column widths
  const widths = [90, 170, 60, 70, 130, 100, 130, 150, 150, 100, 170, 140, 200, 160, 160];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  const startRow = 2;
  const numDataRows = 300; // pre-provision room to grow; extend later if needed

  // Dropdowns
  sh.getRange(startRow, 6, numDataRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(PROGRAM_OPTIONS).build()
  ); // Program
  sh.getRange(startRow, 7, numDataRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(TIER_OPTIONS).build()
  ); // Tier
  sh.getRange(startRow, 11, numDataRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(ROSTER_STATUS_OPTIONS).build()
  ); // Status

  // Threshold auto-formula: only applies to Prevention rows, looks up by Grade (col C)
  const formulas = [];
  for (let i = 0; i < numDataRows; i++) {
    const row = startRow + i;
    formulas.push([
      `=IF($F${row}="Prevention", IFERROR(VLOOKUP($C${row}, _Thresholds!$A$2:$B$13, 2, FALSE), "check grade"), "N/A — Intervention")`
    ]);
  }
  sh.getRange(startRow, 12, numDataRows, 1).setFormulas(formulas);

  // Table border, and center-align the identifier/numeric/dropdown columns
  // (ID, Grade, Class, Starting Score, Threshold, Program, Tier, Status) —
  // link columns stay left-aligned since they hold longer text.
  sh.getRange(1, 1, numDataRows + 1, ROSTER_HEADERS.length).setBorder(true, true, true, true, true, true);
  [1, 3, 4, 6, 7, 10, 11, 12].forEach(col => {
    sh.getRange(startRow, col, numDataRows, 1).setHorizontalAlignment('center');
  });
  sh.getRange(1, 1, 1, ROSTER_HEADERS.length).setHorizontalAlignment('center');
  sh.autoResizeColumns(1, ROSTER_HEADERS.length);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Student Roster built. Threshold Score fills in automatically for Prevention rows.', 'Done', 6
  );
}

/** Hidden helper sheet mapping Grade -> Prevention threshold score. */
function buildThresholdsHelper_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEETS.THRESHOLDS_HELPER);
  if (!sh) sh = ss.insertSheet(SHEETS.THRESHOLDS_HELPER);
  sh.clear();
  const rows = [['Grade', 'Threshold Score']];
  Object.keys(PREVENTION_THRESHOLDS).forEach(g => rows.push([Number(g), PREVENTION_THRESHOLDS[g]]));
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.hideSheet();
}

/**
 * Locks the Student Roster (and Scores Progress, which also carries
 * student-identifying data) down to the current owner only.
 * Run this AFTER you've finished entering/reviewing data.
 */
function protectStudentData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const me = Session.getEffectiveUser();
  [SHEETS.ROSTER, SHEETS.SCORES].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    // Remove any pre-existing protections on this sheet first
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
    const protection = sh.protect().setDescription(name + ' — owner only, contains student data');
    protection.removeEditors(protection.getEditors());
    if (protection.canDomainEdit()) protection.setDomainEdit(false);
    protection.addEditor(me);
  });
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Student Roster and Scores Progress are now locked to you only.', 'Protected', 6
  );
}
