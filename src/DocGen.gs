/**
 * DOCGEN.gs
 * Action Plans and Progress Reports both read from Student Roster and
 * write links back into it. Nothing here ever re-generates a document
 * that already has a link in its column — that's what makes this
 * "generate once per term" instead of overwriting work in progress.
 *
 * FOLDER STRUCTURE
 * Each run checks Settings > "Term # (1, 2, or 3)" and creates/reuses a
 * subfolder named "Term <n> — <label>" inside the relevant Settings
 * folder ID, so every term's documents land in their own dated folder
 * instead of piling up together.
 */

const ROSTER_COL = { // 1-based column numbers, must match ROSTER_HEADERS order
  ID: 1, NAME: 2, GRADE: 3, CLASS: 4, SUBJECT: 5, PROGRAM: 6, TIER: 7,
  ADVISOR: 8, TEACHER: 9, STARTING_SCORE: 10, STATUS: 11, THRESHOLD: 12,
  PROBATION_LETTER: 13, ACTION_PLAN: 14, PROGRESS_REPORT: 15
};

/** Menu entry point: generate Action Plan 1 for every roster row missing one. */
function generateActionPlan1() { generateActionPlans_(1); }
/** Menu entry point: generate Action Plan 2 for every roster row missing one. */
function generateActionPlan2() { generateActionPlans_(2); }

/**
 * DEPRECATED as a content-generation path. Your Action Plan template has
 * no {{token}} placeholders — it's filled by matching table cell labels,
 * and the real weakness/steps content only exists in each teacher's own
 * Week 1/5 tab. Generation now happens correctly from the TEACHER's side
 * (see ActionPlanFill.gs + the "Teacher Tools" menu inside each teacher
 * workbook), which has access to that content and won't create a
 * duplicate file. This function is kept only so the menu item doesn't
 * error out if clicked; it explains where to actually generate instead.
 */
function generateActionPlans_(planNumber) {
  SpreadsheetApp.getUi().alert(
    'Action Plans are generated from each teacher\'s own workbook (Week 1 tab for AP1, ' +
    'Week 5 tab for the AP2 section of the same document) — not from here. Open the ' +
    'relevant teacher workbook and use "Teacher Tools" > Generate/Update Action Plan.'
  );
}

/**
 * Progress Report — ONE document per student (not per subject), since
 * several teachers may fill in different subject sections of the same
 * report. Run this once per term; it skips any student who already has
 * a Progress Report link on their roster rows.
 */
function generateProgressReports() {
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

  // Group rows by Student ID so each student gets exactly one document,
  // and we know every teacher across their subjects who needs edit access.
  const byStudent = {};
  data.forEach((row, i) => {
    const id = row[ROSTER_COL.ID - 1];
    if (!id) return;
    if (!byStudent[id]) byStudent[id] = { rows: [], data: row };
    byStudent[id].rows.push(i + 2);
  });

  let created = 0;
  Object.keys(byStudent).forEach(id => {
    const group = byStudent[id];
    const alreadyLinked = getProgressReportDocUrl_(id) || group.rows.some(r => sh.getRange(r, ROSTER_COL.PROGRESS_REPORT).getValue());
    if (alreadyLinked) return;

    const name = group.data[ROSTER_COL.NAME - 1];
    const grade = group.data[ROSTER_COL.GRADE - 1];
    const cls = group.data[ROSTER_COL.CLASS - 1];
    const fileName = `${name} - Grade ${grade} - ${cls} - Progress Report`;
    const division = divisionFromGrade_(grade) || 'MS';
    const divisionFolder = getDivisionSubfolder_(termFolder, division);

    const { url, fileId } = copyTemplateAndPrefill_(templateId, divisionFolder, fileName, {}, true);
    const doc = DocumentApp.openById(fileId);
    fillProgressReportHeader_(doc, name, cls, grade);
    doc.saveAndClose();

    // Share Editor access with every teacher across this student's subjects.
    // The Roster's Teacher column holds a name — look up their email from
    // Teacher Roster so sharing actually works.
    const teacherEmails = new Set();
    group.rows.forEach(r => {
      const teacherName = sh.getRange(r, ROSTER_COL.TEACHER).getValue();
      const email = teacherEmailByName_(teacherName);
      if (email) teacherEmails.add(email);
    });
    shareFileWithTeachers_(fileId, teacherEmails);

    // Live doc link is tracked internally — Roster's "Progress Report"
    // column stays reserved for the final PDF only.
    setProgressReportDocUrl_(id, url);
    created++;
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Progress Reports: ${created} document(s) created (one per student).`, 'Done', 6
  );
}

/**
 * Run this once teachers have filled in their sections. Converts each
 * Progress Report Doc to a PDF (same folder), stores the PDF link, and
 * emails both you and each involved teacher the final PDF link.
 */
function exportProgressReportsToPDF() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ROSTER);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  const data = sh.getRange(2, 1, lastRow - 1, ROSTER_HEADERS.length).getValues();
  const settings = readSettingsMap_();
  const adEmail = readContacts_()['Academic Dean'].email;

  const seen = new Set(); // avoid exporting the same student's doc twice
  let exported = 0;

  data.forEach((row, i) => {
    const rowNum = i + 2;
    const id = row[ROSTER_COL.ID - 1];
    const pdfUrl = row[ROSTER_COL.PROGRESS_REPORT - 1]; // final PDF, if already exported
    const docUrl = getProgressReportDocUrl_(id); // live doc, tracked internally
    if (!docUrl || pdfUrl || !id || seen.has(id)) return;
    seen.add(id);

    const fileId = extractFileIdFromUrl_(docUrl);
    const file = DriveApp.getFileById(fileId);
    const pdfBlob = file.getAs('application/pdf');
    const pdfFile = file.getParents().next().createFile(pdfBlob).setName(file.getName() + '.pdf');
    applyLinkSharing_(pdfFile, settings);
    const pdfLink = pdfFile.getUrl();

    // Write PDF link into every roster row for this student
    for (let j = 0; j < data.length; j++) {
      if (data[j][ROSTER_COL.ID - 1] === id) {
        sh.getRange(j + 2, ROSTER_COL.PROGRESS_REPORT).setValue(pdfLink);
      }
    }

    const name = row[ROSTER_COL.NAME - 1];
    sendLoggedEmail_(adEmail, `Progress Report ready — ${name}`,
      `The Progress Report for <b>${name}</b> has been exported to PDF: <a href="${pdfLink}">${pdfLink}</a>`,
      'Progress Report PDF', name);

    // Also email every teacher who contributed to this student's report —
    // not just the AD — since they should get the final PDF link too.
    const teacherEmails = new Set();
    data.forEach(r => {
      if (r[ROSTER_COL.ID - 1] !== id) return;
      const teacherField = String(r[ROSTER_COL.TEACHER - 1] || '');
      const m = teacherField.match(/\(([^)]+)\)\s*$/);
      if (m) teacherEmails.add(m[1].trim());
    });
    teacherEmails.forEach(email => {
      sendLoggedEmail_(email, `Progress Report ready — ${name}`,
        `The Progress Report for <b>${name}</b> is finalized: <a href="${pdfLink}">${pdfLink}</a>`,
        'Progress Report PDF', name);
    });

    exported++;
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(`${exported} Progress Report(s) exported to PDF.`, 'Done', 6);
}

// ---------- Shared helpers ----------

/** Gets (or creates) this term's subfolder inside the given parent folder ID. */
function getTermFolder_(parentFolderId, settings) {
  const term = settings['Term # (1, 2, or 3)'] || '?';
  const label = `Term ${term}`;
  const parent = DriveApp.getFolderById(parentFolderId);
  const existing = parent.getFoldersByName(label);
  return existing.hasNext() ? existing.next() : parent.createFolder(label);
}

/**
 * Gets (or creates) a division subfolder (Elementary / Middle School /
 * High School) inside a term folder. Used for Progress Reports, which you
 * want organized by division rather than sitting flat in the term folder.
 */
function getDivisionSubfolder_(termFolder, division) {
  const label = division === 'ES' ? 'Elementary' : division === 'MS' ? 'Middle School' : 'High School';
  const existing = termFolder.getFoldersByName(label);
  return existing.hasNext() ? existing.next() : termFolder.createFolder(label);
}

/** Fills a Doc's table cells by matching each row's label (column 1) against
 * a {label: value} map — used for templates with no {{token}} placeholders,
 * just labeled table rows (e.g. "Student Name" | <blank cell to fill>). */
function fillTemplateTableCells_(doc, labelValueMap) {
  const body = doc.getBody();
  const tables = body.getTables();
  tables.forEach(table => {
    for (let r = 0; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      if (row.getNumCells() < 2) continue;
      const label = row.getCell(0).getText().trim();
      if (labelValueMap.hasOwnProperty(label)) {
        row.getCell(1).setText(String(labelValueMap[label] || ''));
      }
    }
  });
}

/** Copies a Google Doc template and applies link sharing — no content fill;
 * use fillTemplateTableCells_() or replacements afterward as needed. */
function copyTemplateAndPrefill_(templateId, destFolder, fileName, replacements, returnFileId) {
  const copy = DriveApp.getFileById(templateId).makeCopy(fileName, destFolder);
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();
  Object.keys(replacements || {}).forEach(token => body.replaceText(token, String(replacements[token] || '')));
  doc.saveAndClose();

  const settings = readSettingsMap_();
  applyLinkSharing_(copy, settings);

  return returnFileId ? { url: copy.getUrl(), fileId: copy.getId() } : copy.getUrl();
}

/**
 * Fills the Progress Report's header. Your template has "Student Name: "
 * and "Class:" as plain paragraph text (no table, no {{token}}, and no
 * separate Grade field at all) — so this finds those paragraphs by their
 * starting text and appends the value. Grade is folded into the Class
 * line since the template has nowhere else for it.
 */
function fillProgressReportHeader_(doc, studentName, cls, grade) {
  const paragraphs = doc.getBody().getParagraphs();
  paragraphs.forEach(p => {
    const text = p.getText();
    if (text.trim().startsWith('Student Name:')) p.setText('Student Name: ' + studentName);
    else if (text.trim().startsWith('Class:')) p.setText(`Class: ${cls}     Grade: ${grade}`);
  });
}

/** Applies the sharing scope chosen in Settings ("Generated Document Link Sharing"). */
function applyLinkSharing_(file, settings) {
  const scope = settings['Generated Document Link Sharing'] || LINK_SHARING_OPTIONS[0];
  try {
    if (scope === LINK_SHARING_OPTIONS[1]) {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } else {
      file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
    }
  } catch (e) {
    // Link-sharing can be blocked by a Workspace admin policy even when
    // the file itself was created fine — don't let that take down the
    // whole generation. The file still exists and is owned by this
    // script's account; it just may not be link-shareable yet.
    file.setDescription('⚠️ Link sharing failed when this was generated: ' + e.message);
  }
}

/** Shares a file as Editor with a set of teacher emails. Silently skips non-email values. */
function shareFileWithTeachers_(fileId, teacherEmailsSet) {
  const file = DriveApp.getFileById(fileId);
  teacherEmailsSet.forEach(email => {
    if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      try { file.addEditor(email); } catch (e) { /* invalid/duplicate editor, ignore */ }
    }
  });
}

/** Pulls the Drive file ID out of a Docs/Sheets/Drive URL. */
function extractFileIdFromUrl_(url) {
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}