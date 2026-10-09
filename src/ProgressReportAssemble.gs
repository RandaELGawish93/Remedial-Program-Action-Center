/**
 * PROGRESSREPORTASSEMBLE.gs
 * Run "Generate Progress Report (from Teacher Entries)" from the menu.
 * For every student:
 *   1. Works out which subjects they're actually in remedial for.
 *   2. For each of the template's 5 fixed subject columns (English, Math,
 *      Arabic, Science, Social Studies): if the student isn't in remedial
 *      for that subject, the whole column gets N/A. If they are, it looks
 *      up the assigned teacher, opens that teacher's workbook's Progress
 *      Report Entries tab, and pulls the 12 answers IF they are marked Confirmed.
 *   3. Once every required subject is confirmed, exports straight to PDF,
 *      files it, and writes the link back into Student Roster.
 *   4. Anything still missing a confirmation is left as-is and reported
 *      back to you in a summary — re-run any time to pick up new confirmations.
 *
 * IMPORTANT LIMITATION: the template only has 5 subject columns. Religion,
 * Arabic Social Studies, and any MS/HS-only subject (History, Chemistry,
 * Biology, etc.) have no column to go in — those are flagged in the
 * summary as "no column available", not silently dropped.
 */

function generateProgressReportsFromTeacherEntries() {
  const settings = readSettingsMap_();
  const folderId = settings['Progress Report Folder ID'];
  const templateId = settings['Progress Report Template ID'];
  if (!folderId || !templateId) {
    SpreadsheetApp.getUi().alert('Fill in the Progress Report Folder ID and Template ID in Settings first.');
    return;
  }
  const termFolder = getTermFolder_(folderId, settings);

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ROSTER);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  const data = sh.getRange(2, 1, lastRow - 1, ROSTER_HEADERS.length).getValues();

  // Group by student
  const byStudent = {};
  data.forEach((row, i) => {
    const id = row[ROSTER_COL.ID - 1];
    if (!id) return;
    if (!byStudent[id]) byStudent[id] = { rows: [], data: row };
    byStudent[id].rows.push(i + 2);
  });

  const adEmail = readContacts_()['Academic Dean'].email;
  let completed = 0, stillPending = 0, failed = 0;
  const unmappableSubjects = new Set();
  const pendingDetails = [];
  const failedDetails = [];

  Object.keys(byStudent).forEach(id => {
    const group = byStudent[id];
    const name = group.data[ROSTER_COL.NAME - 1];
    // Wrapped so ONE student's failure (bad template, Drive hiccup, etc.)
    // can never silently stop every student after them in this same run —
    // that's exactly the class of bug that caused Action Plan generation
    // to fail without explanation. Same fix, applied here too.
    try {
      processOneStudentProgressReport_(id, group, name, sh, termFolder, templateId, settings,
        adEmail, unmappableSubjects, pendingDetails,
        result => { if (result === 'completed') completed++; else stillPending++; });
    } catch (e) {
      failed++;
      failedDetails.push(`${name}: ${e.message}`);
    }
  });

  let summary = `${completed} Progress Report(s) completed and exported to PDF.\n${stillPending} still waiting on teacher confirmation.`;
  if (failed) summary += `\n${failed} FAILED with an error — see below.`;
  if (failedDetails.length) summary += '\n\n' + failedDetails.slice(0, 10).join('\n');
  if (pendingDetails.length) summary += '\n\n' + pendingDetails.slice(0, 15).join('\n');
  if (unmappableSubjects.size) {
    summary += '\n\nNo column exists in this template for:\n' + Array.from(unmappableSubjects).slice(0, 10).join('\n');
  }
  SpreadsheetApp.getUi().alert(summary);
}

/** One student's Progress Report assembly — pulled out so it can be wrapped in try/catch per-student. */
function processOneStudentProgressReport_(id, group, name, sh, termFolder, templateId, settings, adEmail, unmappableSubjects, pendingDetails, onResult) {
    const cls = group.data[ROSTER_COL.CLASS - 1];

    // Already exported? Skip — this function doesn't regenerate finished PDFs.
    // Roster's "Progress Report" column holds the final PDF only.
    const alreadyExported = group.rows.some(r => sh.getRange(r, ROSTER_COL.PROGRESS_REPORT).getValue());
    if (alreadyExported) return;

    // Which subjects is this student actually remediating in?
    const requiredSubjects = new Set();
    const subjectTeacher = {}; // subject -> teacher name
    group.rows.forEach(r => {
      const subj = sh.getRange(r, ROSTER_COL.SUBJECT).getValue();
      const teacherField = sh.getRange(r, ROSTER_COL.TEACHER).getValue();
      const m = String(teacherField).match(/^(.*?)\s*\(/);
      requiredSubjects.add(subj);
      subjectTeacher[subj] = m ? m[1].trim() : teacherField;
    });

    // Get or create the Doc for this student. The live doc link is tracked
    // in a hidden helper sheet — NOT in Roster's "Progress Report" column,
    // since that column is reserved for the final PDF only.
    let docUrl = getProgressReportDocUrl_(id);
    let fileId;
    if (!docUrl) {
      const division = divisionFromGrade_(group.data[ROSTER_COL.GRADE - 1]) || 'MS';
      const divisionFolder = getDivisionSubfolder_(termFolder, division);
      const fileName = `${name} - Grade ${group.data[ROSTER_COL.GRADE - 1]} - ${cls} - Progress Report`;
      const copy = DriveApp.getFileById(templateId).makeCopy(fileName, divisionFolder);
      applyLinkSharing_(copy, settings);
      fileId = copy.getId();
      docUrl = copy.getUrl();
      const doc = DocumentApp.openById(fileId);
      fillProgressReportHeader_(doc, name, cls, group.data[ROSTER_COL.GRADE - 1]);
      doc.saveAndClose();
      setProgressReportDocUrl_(id, docUrl);
    } else {
      fileId = extractFileIdFromUrl_(docUrl);
    }

    const doc = DocumentApp.openById(fileId);
    let allConfirmed = true;
    const missingFor = [];

    Object.keys(PR_SUBJECT_COLUMNS).forEach(templateSubject => {
      // Skip the duplicate 'Math'/'Mathematics' alias — only process once per column
      const colIndex = PR_SUBJECT_COLUMNS[templateSubject];
      if (templateSubject === 'Math') return; // 'Mathematics' handles column 2

      const isRequired = requiredSubjects.has(templateSubject) ||
        (templateSubject === 'Mathematics' && requiredSubjects.has('Math'));

      if (!isRequired) {
        fillProgressReportColumnNA_(doc, colIndex);
        return;
      }

      const teacherName = subjectTeacher[templateSubject] || subjectTeacher['Math'];
      const entry = getConfirmedProgressReportEntry_(teacherName, id);
      if (entry) {
        fillProgressReportColumn_(doc, colIndex, entry);
      } else {
        allConfirmed = false;
        missingFor.push(`${templateSubject} (${teacherName || 'unknown teacher'})`);
      }
    });

    // Flag any required subject that has no column in this template at all
    requiredSubjects.forEach(s => {
      if (!PR_SUBJECT_COLUMNS.hasOwnProperty(s)) unmappableSubjects.add(`${name} — ${s}`);
    });

    doc.saveAndClose();

    if (allConfirmed) {
      const pdfBlob = DriveApp.getFileById(fileId).getAs('application/pdf');
      const pdfFile = DriveApp.getFileById(fileId).getParents().next().createFile(pdfBlob)
        .setName(DriveApp.getFileById(fileId).getName() + '.pdf');
      applyLinkSharing_(pdfFile, settings);
      const pdfLink = pdfFile.getUrl();
      group.rows.forEach(r => sh.getRange(r, ROSTER_COL.PROGRESS_REPORT).setValue(pdfLink));

      const recipientEmails = new Set([adEmail]);
      Object.values(subjectTeacher).forEach(tn => {
        const email = teacherEmailByName_(tn);
        if (email) recipientEmails.add(email);
      });
      recipientEmails.forEach(email => {
        sendLoggedEmail_(email, `Progress Report ready — ${name}`,
          `The Progress Report for <b>${name}</b> is finalized: <a href="${pdfLink}">${pdfLink}</a>`,
          'Progress Report PDF', name);
      });

      // Also write the PDF link into each contributing teacher's OWN workbook
      // (Progress Report Entries tab), so it's clickable there too — not
      // just in your master Roster.
      new Set(Object.values(subjectTeacher)).forEach(tn => {
        writeProgressReportPdfLinkToTeacherWorkbook_(tn, id, pdfLink);
      });

      onResult('completed');
    } else {
      pendingDetails.push(`${name}: waiting on ${missingFor.join(', ')}`);
      onResult('pending');
    }
}

/** Writes the final Progress Report PDF link into a teacher's own Progress Report Entries row. */
function writeProgressReportPdfLinkToTeacherWorkbook_(teacherName, studentId, pdfLink) {
  const teacherRow = teacherRosterRowByName_(teacherName);
  if (!teacherRow || !teacherRow.workbookLink) return;
  const fileId = extractFileIdFromUrl_(teacherRow.workbookLink);
  let ss;
  try { ss = SpreadsheetApp.openById(fileId); } catch (e) { return; }
  const sh = ss.getSheetByName('Progress Report Entries');
  if (!sh) return;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  const ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  const pdfCol = 3 + PR_STATEMENTS.length + 1; // after ID, Name, 12 statements, Confirmed
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(studentId)) {
      sh.getRange(i + 2, pdfCol).setValue(pdfLink);
      return;
    }
  }
}

/** Reads a teacher's confirmed entry for one student from her own workbook. Returns null if not confirmed. */
function getConfirmedProgressReportEntry_(teacherName, studentId) {
  const teacherRow = teacherRosterRowByName_(teacherName);
  if (!teacherRow || !teacherRow.workbookLink) return null;
  const fileId = extractFileIdFromUrl_(teacherRow.workbookLink);
  let ss;
  try { ss = SpreadsheetApp.openById(fileId); } catch (e) { return null; }
  const sh = ss.getSheetByName('Progress Report Entries');
  if (!sh) return null;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  const numStatements = PR_STATEMENTS.length;
  const data = sh.getRange(2, 1, lastRow - 1, 2 + numStatements + 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(studentId)) {
      const done = data[i][2 + numStatements] === 'Done';
      if (!done) return null;
      return data[i].slice(2, 2 + numStatements); // the 12 response values
    }
  }
  return null;
}

/** Writes 12 confirmed responses into one subject column of the Progress Report doc. */
function fillProgressReportColumn_(doc, colIndex, responses) {
  const table = doc.getBody().getTables()[0];
  for (let i = 0; i < responses.length; i++) {
    table.getRow(i + 1).getCell(colIndex).setText(responses[i] || '');
  }
}

/** Sets an entire subject column to N/A — used when the student isn't in remedial for that subject. */
function fillProgressReportColumnNA_(doc, colIndex) {
  const table = doc.getBody().getTables()[0];
  for (let i = 1; i <= PR_STATEMENTS.length; i++) {
    table.getRow(i).getCell(colIndex).setText('N/A');
  }
}

/**
 * Hidden internal tracking for the Progress Report's LIVE doc link — kept
 * separate from Roster's visible "Progress Report" column, which is
 * reserved for the final PDF only, keeping the Roster to
 * just Probation Letter / Action Plan 1 / Action Plan 2 / Progress Report.
 */
function getProgressReportDocUrl_(studentId) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('_ProgressReportDocs');
  if (!sh) return null;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  const data = sh.getRange(2, 1, lastRow - 1, 2).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(studentId)) return data[i][1];
  }
  return null;
}

function setProgressReportDocUrl_(studentId, url) {
  const sh = getOrCreateSheet_('_ProgressReportDocs');
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 2).setValues([['Student ID', 'Doc URL']]);
  }
  sh.appendRow([studentId, url]);
  sh.hideSheet();
}