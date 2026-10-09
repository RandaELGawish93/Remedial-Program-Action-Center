/**
 * DIVISIONREPORT.gs
 * Generates one PDF per division (ES / MS / HS), broken out by subject
 * and grade level, with bar charts comparing same-subject-same-grade
 * students for a fair comparison. Only generates in Week 11 (after the
 * term's 10 weeks are done) — running it earlier just shows an alert.
 *
 * LOCATION: reports are filed into a "Division Reports" subfolder under
 * the Progress Report Folder's term folder.
 */

function generateDivisionReports() {
  const settings = readSettingsMap_();
  const termStart = TERM_START_DATES[String(settings['Term # (1, 2, or 3)'])];
  const currentWeek = termStart ? Math.floor((new Date() - termStart) / (7 * 86400000)) + 1 : 0;
  if (currentWeek !== 11) {
    SpreadsheetApp.getUi().alert(
      `Division Reports only generate in Week 11 — the current computed week is ${currentWeek}.`
    );
    return;
  }

  const rawSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('_ScoresRaw');
  if (!rawSh || rawSh.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Run "Refresh Scores Progress" first so there\'s data to report on.');
    return;
  }
  const data = rawSh.getRange(2, 1, rawSh.getLastRow() - 1, SCORES_HEADERS.length).getValues();

  const progressFolderId = settings['Progress Report Folder ID'];
  const termFolder = getTermFolder_(progressFolderId, settings);
  const reportsFolder = termFolder.getFoldersByName('Division Reports').hasNext()
    ? termFolder.getFoldersByName('Division Reports').next()
    : termFolder.createFolder('Division Reports');

  ['ES', 'MS', 'HS'].forEach(division => {
    const rows = data.filter(r => divisionFromGrade_(r[2]) === division); // r[2] = Grade column in _ScoresRaw
    if (!rows.length) return;
    const url = buildDivisionReportDoc_(division, rows, reportsFolder, settings);
    SpreadsheetApp.getActiveSpreadsheet().toast(`${division} Division Report: ${url}`, 'Generated', 5);
  });

  SpreadsheetApp.getUi().alert('Division Reports generated — check the Division Reports folder.');
}

function buildDivisionReportDoc_(division, rows, folder, settings) {
  const divisionLabel = division === 'ES' ? 'Elementary' : division === 'MS' ? 'Middle School' : 'High School';
  const doc = DocumentApp.create(`${divisionLabel} Division Report — Week 11`);
  const body = doc.getBody();
  body.appendParagraph(`${divisionLabel} Division Report`).setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph(`Week 11 — Term ${settings['Term # (1, 2, or 3)'] || ''}`)
    .setHeading(DocumentApp.ParagraphHeading.SUBTITLE);

  // Group by Subject, then by Grade within subject
  const bySubject = {};
  rows.forEach(r => {
    const subject = r[4], grade = r[2];
    if (!bySubject[subject]) bySubject[subject] = {};
    if (!bySubject[subject][grade]) bySubject[subject][grade] = [];
    bySubject[subject][grade].push(r);
  });

  Object.keys(bySubject).sort().forEach(subject => {
    body.appendParagraph(subject).setHeading(DocumentApp.ParagraphHeading.HEADING1);

    Object.keys(bySubject[subject]).sort((a, b) => Number(a) - Number(b)).forEach(grade => {
      const group = bySubject[subject][grade];
      body.appendParagraph(`Grade ${grade}`).setHeading(DocumentApp.ParagraphHeading.HEADING2);

      const names = group.map(r => r[1]);
      const finalScores = group.map(r => Number(r[10]) || 0);
      const chartBlob = buildGroupedBarChartBlob_(names, finalScores, `${subject} — Grade ${grade} Final Scores /65`);
      body.appendImage(chartBlob);

      const validScores = group.map(r => Number(r[10])).filter(v => !isNaN(v) && v > 0);
      const expectedOverall = validScores.length
        ? Math.round(validScores.reduce((a, b) => a + b, 0) / validScores.length) : '—';
      body.appendParagraph(`Expected Overall Grade for this group: ${expectedOverall} / 65`)
        .setBold(true);
    });
  });

  doc.saveAndClose();
  const file = DriveApp.getFileById(doc.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  const pdfBlob = file.getAs('application/pdf');
  const pdfFile = folder.createFile(pdfBlob).setName(`${divisionLabel} Division Report - Week 11.pdf`);
  applyLinkSharing_(pdfFile, settings);
  applyLinkSharing_(file, settings);

  return pdfFile.getUrl();
}

/** Builds a bar chart image comparing each student's score, for same-subject-same-grade fairness. */
function buildGroupedBarChartBlob_(names, values, title) {
  const dataTable = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, 'Student')
    .addColumn(Charts.ColumnType.NUMBER, 'Final Score /65');
  names.forEach((name, i) => dataTable.addRow([name, values[i]]));

  return Charts.newColumnChart()
    .setDataTable(dataTable.build())
    .setTitle(title)
    .setDimensions(500, 300)
    .build()
    .getAs('image/png');
}