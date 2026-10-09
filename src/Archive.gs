/**
 * ARCHIVE.gs
 * Run "Archive Trimester" from the menu at the END of a term, before
 * starting the next one. It:
 *   1. Makes a full copy of this workbook into an Archive folder
 *   2. Renames the copy "T<term> <label> — Archive — <today's date>"
 *   3. Clears data (not headers/formulas/dropdowns) from the LIVE tabs
 *      so you start the new term clean
 * Settings (contacts, folder/template IDs) are left untouched — those
 * don't change term to term.
 *
 * NOTE: this archives the MASTER workbook only. Teacher workbooks and
 * the Principals' workbooks are not copied by this step.
 */

const TABS_TO_RESET_ON_ARCHIVE = [
  SHEETS.ROSTER, SHEETS.SCORES, SHEETS.TEACHER_ROSTER,
  SHEETS.ATTENDANCE_TEACHERS, SHEETS.MOM_LOG, '_ScoresRaw', '_TeacherTrackStatus', '_ProgressReportDocs'
];

function archiveTrimester() {
  const ui = SpreadsheetApp.getUi();
  const settings = readSettingsMap_();
  const term = settings['Term # (1, 2, or 3)'] || '?';

  const confirm = ui.alert(
    'Archive Trimester ' + term + '?',
    'This will:\n' +
    '1. Save a full copy of this workbook to your Archive folder\n' +
    '2. Clear student data from the live Roster and Scores tabs\n\n' +
    'Settings will NOT be touched. This cannot be undone on the live tabs ' +
    '(but the archived copy will always have everything). Continue?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const file = DriveApp.getFileById(ss.getId());
  const dateLabel = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const archiveName = `T${term} Remedial Follow-Up — Archive — ${dateLabel}`;

  // Duplicate into the parent folder (or root if none) alongside the live file
  const parents = file.getParents();
  const parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  file.makeCopy(archiveName, parentFolder);

  // Reset live tabs — clear data rows below the header, keep formulas/dropdowns
  // where the builder put them (Roster's Threshold formula column, etc.)
  TABS_TO_RESET_ON_ARCHIVE.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    const headerRows = (name === SHEETS.SCORES) ? 3 : 1; // Scores has a filter row + header
    if (lastRow > headerRows) {
      sh.getRange(headerRows + 1, 1, lastRow - headerRows, lastCol).clearContent();
    }
  });

  ui.alert('Archived as "' + archiveName + '" in the same Drive folder. Live tabs are reset for the new term.');
}