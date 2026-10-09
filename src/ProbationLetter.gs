/**
 * PROBATIONLETTER.gs
 * The Probation Letter is never generated — you physically meet the
 * parent, get it signed, then upload the scanned/signed copy here.
 * Workflow: click the student's row anywhere in Student Roster, then
 * run "Upload Probation Letter" from the menu.
 */

function openProbationLetterUpload() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ROSTER);
  const activeSheet = SpreadsheetApp.getActiveSheet();
  if (activeSheet.getName() !== SHEETS.ROSTER) {
    ui.alert('Click a row in Student Roster first, then run this again.');
    return;
  }

  const row = SpreadsheetApp.getActiveRange().getRow();
  if (row < 2) {
    ui.alert('Select a student row (not the header) first.');
    return;
  }

  const name = sh.getRange(row, ROSTER_COL.NAME).getValue();
  const grade = sh.getRange(row, ROSTER_COL.GRADE).getValue();
  if (!name) {
    ui.alert('That row has no student name — select a valid student row.');
    return;
  }

  const template = HtmlService.createTemplateFromFile('UploadDialog');
  template.studentName = name;
  template.grade = grade;
  template.rowNum = row;
  const html = template.evaluate().setWidth(380).setHeight(220);
  ui.showModalDialog(html, 'Upload Probation Letter');
}

/**
 * Called from UploadDialog.html once a file is chosen. Saves it into the
 * Probation Letter Folder (from Settings), renamed to "Name - GradeX -
 * Probation Letter.<ext>", applies the standard link-sharing scope, and
 * writes the URL back into that row's Probation Letter Link column.
 */
function saveProbationLetter(base64Data, mimeType, originalFileName, rowNum) {
  const settings = readSettingsMap_();
  const folderId = settings['Probation Letter Folder ID'];
  if (!folderId) throw new Error('Set the Probation Letter Folder ID in Settings first.');

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ROSTER);
  const name = sh.getRange(rowNum, ROSTER_COL.NAME).getValue();
  const grade = sh.getRange(rowNum, ROSTER_COL.GRADE).getValue();
  const ext = (originalFileName.match(/\.[^.]+$/) || [''])[0];

  const bytes = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(bytes, mimeType, `${name} - Grade ${grade} - Probation Letter${ext}`);

  const folder = DriveApp.getFolderById(folderId);
  const file = folder.createFile(blob);
  applyLinkSharing_(file, settings);

  const url = file.getUrl();
  sh.getRange(rowNum, ROSTER_COL.PROBATION_LETTER).setValue(url);
  return url;
}