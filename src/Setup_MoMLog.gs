/**
 * SETUP_MOMLOG.gs
 * Generate MoM and Send MoM are two completely separate actions:
 *  - Generate: fills the template, exports a real PDF into the MoM
 *    Drive folder right then, writes that PDF's link into MoM Log.
 *    Sends nothing.
 *  - Send: a separate step, entirely about email — each teacher
 *    ready for that round gets their own individual email with only
 *    their own PDF attached, active Settings contacts in CC, always
 *    shown as a preview first.
 */

const MOM_ROUNDS = [
  { round: 1, attCol: 2, linkCol: 3, emailSentCol: 4, label: 'MoM #1 — Orientation' },
  { round: 2, attCol: 5, linkCol: 6, emailSentCol: 7, label: 'MoM #2' },
  { round: 3, attCol: 8, linkCol: 9, emailSentCol: 10, label: 'MoM #3' },
];

const MOM_ATTENDANCE_OPTIONS = ['Present', 'Absent'];

// MOM_CONTENT holds the real, authored wording that gets written into each
// teacher's MoM document (Purpose / Related to Action Plan / Discussion
// Points), keyed by round number. Round 1's text was recovered from the
// real MoM docs already generated and sent to teachers on 2026-09-19.
// Rounds 2 and 3 are intentionally left blank here — that wording has not
// been written yet, so it must be added (with the user's input) before
// those rounds can be generated or sent.
const MOM_CONTENT = {
  1: {
    purpose: 'Orientation — Workbook Navigation, Action Plan & Progress Report Preparation',
    relatedToActionPlan: 'N/A — Orientation session',
    discussionPoints: '• Remedial list and weekly schedule\n• Starting time and session logistics\n• Workbook navigation walkthrough\n• Action Plan 1 preparation (Weeks 1–5)\n• Progress Report preparation (due Week 5)\n• Overview of Cumulative Score 1 at Week 5',
    followUpWeekOffset: 5
  },
  2: {
    purpose: 'Mid-Term Review — Action Plan 1 Completion & Action Plan 2 Kickoff',
    relatedToActionPlan: 'Closing out Action Plan 1 (Weeks 1–5); opening Action Plan 2 (Weeks 6–10)',
    discussionPoints: '• Close out Action Plan 1: enter Cumulative Score 1 (Weeks 1–5) — ET grades, quiz grades, and attendance\n• Write the Action Plan 1 written summary\n• Start Action Plan 2: write Spotted Weakness #1 and #2, how each will be tackled, means of assessment, and evidence of improvement — then click Generate\n• Fill in the Progress Report and mark it "Filled"\n• Deadline: first Saturday after Week 5 is done, so reports can go to the Principal and be shared with parents',
    followUpWeekOffset: 10
  },
  3: {
    purpose: 'Final Review — Action Plan 2 Completion',
    relatedToActionPlan: 'Closing out Action Plan 2 (Weeks 6–10)',
    discussionPoints: '• Close out Action Plan 2: enter Cumulative Score 2 (Weeks 6–10) — ET grades and quiz grades\n• Write the Action Plan 2 written summary\n• Click Generate',
    followUpWeekOffset: null // final round — no further MoM follow-up
  }
};

/**
 * Builds the MoM Log tab. Preserves anything already entered (attendance,
 * links, sent status) if run again — matched by teacher name, not
 * position, same pattern as everything else in this project.
 */

function buildMoMLog() {
  const sh = getOrCreateSheet_(SHEETS.MOM_LOG);

  // Preserve existing Attendance + Link only — Email Sent is dropped
  // here since it's rebuilt as a live formula below, not a stored value.
  const existing = {};
  const oldLastRow = sh.getLastRow();
  if (oldLastRow > 1) {
    const oldLastCol = sh.getLastColumn();
    const oldValues = sh.getRange(2, 1, oldLastRow - 1, oldLastCol).getValues();
    const oldFormulas = sh.getRange(2, 1, oldLastRow - 1, oldLastCol).getFormulas();
    const keepCols = new Set([0]);
    MOM_ROUNDS.forEach(r => { keepCols.add(r.attCol - 1); keepCols.add(r.linkCol - 1); });
    oldValues.forEach((row, i) => {
      const name = row[0];
      if (!name) return;
      existing[name] = row.map((v, c) => (keepCols.has(c) ? (oldFormulas[i][c] || v) : ''));
    });
  }

  // Remove any protections left over from a previous build before
  // clearing — re-added fresh on the Email Sent columns at the end, so
  // re-running this never piles up duplicate protections.
  sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());
  sh.clear();

  // Defensive: sh.clear() removes content/formatting but NOT data
  // validation rules, so a stale rule from an earlier version of this
  // sheet (e.g. before the header restructure) can survive and show up
  // as a wrong/leftover dropdown on cells it no longer applies to.
  // Wipe validations clean here so only what this function sets below
  // (Attendance columns only) ever exists.
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();

  const headers = ['Teacher Name'];
  MOM_ROUNDS.forEach(r => headers.push(`Attendance MoM ${r.round}`, `MoM # ${r.round}`, `Email Sent MoM ${r.round}`));
  sh.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold').setBackground('#1c4587').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 170);
  for (let i = 2; i <= headers.length; i++) sh.setColumnWidth(i, 180);
  MOM_ROUNDS.forEach(r => sh.setColumnWidth(r.emailSentCol, 220));

  const roster = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.TEACHER_ROSTER);
  const lastRow = roster ? roster.getLastRow() : 1;
  const names = lastRow > 1 ? roster.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(String) : [];
  if (names.length) {
    sh.getRange(2, 1, names.length, 1).setValues(names.map(n => [n]));

    names.forEach((name, i) => {
      const prevRow = existing[name];
      if (!prevRow) return;
      const r = 2 + i;
      prevRow.forEach((cell, c) => {
        if (c === 0 || !cell) return;
        const cellRange = sh.getRange(r, c + 1);
        if (String(cell).startsWith('=')) cellRange.setFormula(cell);
        else cellRange.setValue(cell);
      });
    });

    MOM_ROUNDS.forEach(r => {
      sh.getRange(2, r.attCol, names.length, 1).setDataValidation(
        SpreadsheetApp.newDataValidation().requireValueInList(MOM_ATTENDANCE_OPTIONS).build()
      );
    });

    // Email Sent — live formula pulled from Email Log, not a stored value.
    MOM_ROUNDS.forEach(r => {
      const formulas = [];
      for (let i = 0; i < names.length; i++) formulas.push([momEmailStatusFormula_(2 + i, r.label)]);
      sh.getRange(2, r.emailSentCol, names.length, 1).setFormulas(formulas);
    });

    sh.getRange(1, 1, names.length + 1, headers.length).setBorder(true, true, true, true, true, true);
    MOM_ROUNDS.forEach(r => {
      sh.getRange(2, r.attCol, names.length, 1).setHorizontalAlignment('center');
      sh.getRange(2, r.emailSentCol, names.length, 1).setHorizontalAlignment('center');
    });
    sh.getRange(1, 1, 1, headers.length).setHorizontalAlignment('center');
    sh.autoResizeColumns(1, headers.length);

    // Read-only: Email Sent is computed from Email Log, so lock it to
    // the owner only — same pattern as protectStudentData().
    const me = Session.getEffectiveUser();
    MOM_ROUNDS.forEach(r => {
      const protection = sh.getRange(2, r.emailSentCol, names.length, 1)
        .protect().setDescription(`Email Sent MoM ${r.round} — read-only, pulled from Email Log`);
      protection.removeEditors(protection.getEditors());
      if (protection.canDomainEdit()) protection.setDomainEdit(false);
      protection.addEditor(me);
    });
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'MoM Log built. Mark attendance for a round, run "Generate MoM — Round 1/2/3" to create the PDFs, ' +
    'then "Send MoM — Round 1/2/3" separately whenever you\'re ready to actually email them. Email Sent ' +
    'updates itself from Email Log once you send.', 'Done', 8
  );
}

/**
 * Formula for one Email Sent cell: looks at Email Log for the most
 * recent row where To = this teacher's email (from Teacher Roster) and
 * Subject = this round's label, and returns "Sent", "Failed — <error>",
 * or "Not Sent" if nothing's there yet. Uses the LAST matching row (by
 * sheet row number) so a resend after a failure always shows the latest
 * outcome, not the first attempt.
 */
function momEmailStatusFormula_(row, roundLabel) {
  const label = roundLabel.replace(/"/g, '""');
  const match = `('Email Log'!$B$2:$B$2000=VLOOKUP($A${row},'Teacher Roster'!$A:$B,2,0))*('Email Log'!$C$2:$C$2000="${label}")`;
  const lastRow = `SUMPRODUCT(MAX(${match}*ROW('Email Log'!$B$2:$B$2000)))`;
  return `=IFERROR(IF(${lastRow}=0,"Not Sent",` +
    `IF(INDEX('Email Log'!$F:$F,${lastRow})="Failed","Failed — "&INDEX('Email Log'!$G:$G,${lastRow}),"Sent")` +
    `),"Not Sent")`;
}

function na_(v) {
  return (v === undefined || v === null || String(v).trim() === '') ? 'N/A' : v;
}

/** Fills the MoM doc per the exact template structure — verified directly
 * against the real MoM_AD_with_Teachers.docx (9 tables, 0-8). */
function fillMoMDoc_(doc, teacherName, teacherRow, adName, attendance, content, followUpDate, autoSummary, generalNote, specificNote, meetingInfo) {
  const info = meetingInfo || {};
  const tables = doc.getBody().getTables();

  // TABLE[0]: Date / Time / Location / Teacher Name / Subject / Classes
  tables[0].getRow(0).getCell(1).setText(na_(info.date));
  tables[0].getRow(1).getCell(1).setText(na_(info.time));
  tables[0].getRow(2).getCell(1).setText(na_(info.location));
  tables[0].getRow(3).getCell(1).setText(teacherName || '');
  tables[0].getRow(4).getCell(1).setText(teacherRow ? teacherRow.subject : '');
  tables[0].getRow(5).getCell(1).setText(teacherRow ? teacherRow.classes : '');

  // TABLE[1]: Attendees — Academic Dean / Teacher (+ status) / HoD
  tables[1].getRow(0).getCell(1).setText(adName || '');
  tables[1].getRow(1).getCell(1).setText(`${teacherName || ''} — Status: ${attendance}`);
  tables[1].getRow(2).getCell(1).setText(na_(info.hod));

  // TABLE[2]/[3]/[4]: Purpose / Related to Action Plan / Discussion Points
  tables[2].getRow(0).getCell(0).setText(na_(content.purpose));
  tables[3].getRow(0).getCell(0).setText(na_(content.relatedToActionPlan));

  let discussionText = content.discussionPoints || '';
  if (autoSummary) discussionText += '\n\nPer-student notes:\n' + autoSummary;
  if (generalNote) discussionText += '\n\nGeneral note to all teachers:\n' + generalNote;
  if (specificNote) discussionText += '\n\nNote specifically for you:\n' + specificNote;
  tables[4].getRow(0).getCell(0).setText(na_(discussionText));

  // TABLE[5]/[6]: Concerns Raised / Agreed Next Steps — left blank, filled live during the meeting

  // TABLE[7]: Follow-Up Date
  if (followUpDate) {
    tables[7].getRow(0).getCell(1).setText(Utilities.formatDate(followUpDate, Session.getScriptTimeZone(), 'MM/dd/yyyy'));
  } else {
    tables[7].getRow(0).getCell(1).setText('N/A');
  }

  // TABLE[8]: Sign-Off — Academic Dean + Teacher, name and email, replacing
  // the blank handwritten-signature placeholders.
  const adEmail = activeEmailFor_('Academic Dean') || '';
  const teacherEmail = teacherRow ? teacherRow.email : '';
  const signOffDate = na_(info.date);
  tables[8].getRow(1).getCell(0).setText(`${adName || ''}\n${adEmail}`);
  tables[8].getRow(1).getCell(1).setText(`${teacherName || ''}\n${teacherEmail}`);
  tables[8].getRow(2).getCell(0).setText(`Date: ${signOffDate}`);
  tables[8].getRow(2).getCell(1).setText(`Date: ${signOffDate}`);
}

/** MoM #1 — same generic content every time, so Date/Time/Location/HoD are
 * asked once as simple popups rather than needing the full compose screen. */
function generateMoMRound1() {
  const ui = SpreadsheetApp.getUi();

  const dateResp = ui.prompt('MoM #1 — Meeting Date', 'Enter the date of this orientation meeting. Leave blank for N/A.', ui.ButtonSet.OK_CANCEL);
  if (dateResp.getSelectedButton() !== ui.Button.OK) return;
  const timeResp = ui.prompt('MoM #1 — Meeting Time', 'Enter the time. Leave blank for N/A.', ui.ButtonSet.OK_CANCEL);
  if (timeResp.getSelectedButton() !== ui.Button.OK) return;
  const locationResp = ui.prompt('MoM #1 — Meeting Location', 'Enter the location or meeting method. Leave blank for N/A.', ui.ButtonSet.OK_CANCEL);
  if (locationResp.getSelectedButton() !== ui.Button.OK) return;
  const hodResp = ui.alert('MoM #1 — HoD Attendance', 'Is the Head of Department present at this meeting?', ui.ButtonSet.YES_NO);

  const meetingInfo = {
    date: dateResp.getResponseText().trim(),
    time: timeResp.getResponseText().trim(),
    location: locationResp.getResponseText().trim(),
    hod: hodResp === ui.Button.YES ? 'Present' : 'N/A'
  };

  generateMoMRound_(1, meetingInfo);
}

/** Generates the MoM PDF(s) for this round — exports straight to a real
 * PDF in Drive right now (not converted later at send time), writes that
 * PDF's link into MoM Log. Sends nothing — use "Send MoM" separately. */
function generateMoMRound_(roundNum, meetingInfo) {
  const settings = readSettingsMap_();
  const folderId = settings['MoM Folder ID'];
  const templateId = settings['MoM Template ID'];
  if (!folderId || !templateId) {
    SpreadsheetApp.getUi().alert('Fill in the MoM Folder ID and Template ID in Settings first.');
    return;
  }
  const termFolder = getTermFolder_(folderId, settings);
  const roundConfig = MOM_ROUNDS.find(r => r.round === roundNum);
  const content = MOM_CONTENT[roundNum];

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MOM_LOG);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const adName = readContacts_()['Academic Dean'].name;
  const termStart = TERM_START_DATES[String(settings['Term # (1, 2, or 3)'])];
  const followUpDate = (content.followUpWeekOffset && termStart)
    ? new Date(termStart.getTime() + content.followUpWeekOffset * 7 * 86400000) : null;

  let generated = 0, failed = 0;
  const failedNames = [];

  for (let row = 2; row <= lastRow; row++) {
    const teacherName = sh.getRange(row, 1).getValue();
    if (!teacherName) continue;
    const attendance = sh.getRange(row, roundConfig.attCol).getValue();
    if (!attendance) continue;
    if (sh.getRange(row, roundConfig.linkCol).getValue()) continue;

    try {
      const teacherRow = teacherRosterRowByName_(teacherName);
      const fileName = `MoM ${roundNum} - ${teacherName}`;
      const docCopy = DriveApp.getFileById(templateId).makeCopy(fileName, termFolder);
      const doc = DocumentApp.openById(docCopy.getId());
      fillMoMDoc_(doc, teacherName, teacherRow, adName, attendance, content, followUpDate, null, null, null, meetingInfo);
      doc.saveAndClose();

      // Export to a real PDF right now — this is the file "Send MoM" will
      // attach later. The live doc was only ever a means to fill it in.
      const pdfBlob = DriveApp.getFileById(docCopy.getId()).getAs('application/pdf').setName(fileName + '.pdf');
      const pdfFile = termFolder.createFile(pdfBlob);
      applyLinkSharing_(pdfFile, settings, teacherRow ? teacherRow.email : null);
      docCopy.setTrashed(true);

      const url = pdfFile.getUrl();
      sh.getRange(row, roundConfig.linkCol).setFormula(`=HYPERLINK("${url}", "Open MoM")`);
      generated++;
    } catch (e) {
      failed++;
      failedNames.push(`${teacherName}: ${e.message}`);
    }
  }

  let msg = `Round ${roundNum}: ${generated} MoM PDF(s) generated in Drive. Nothing was emailed — run "Send MoM — Round ${roundNum}" separately when ready.`;
  if (failed) msg += `\n${failed} FAILED — see below.\n\n` + failedNames.slice(0, 10).join('\n');
  SpreadsheetApp.getUi().alert(msg);
}

/** MoM #2/#3 — opens the compose screen so you can add per-teacher
 * observations (these go into each teacher's own document, not the
 * eventual shared email — which stays batched and generic). */
function openMoMComposeDialog_(roundNum) {
  const settings = readSettingsMap_();
  if (!settings['MoM Folder ID'] || !settings['MoM Template ID']) {
    SpreadsheetApp.getUi().alert('Fill in the MoM Folder ID and Template ID in Settings first.');
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const dateResp = ui.prompt(`MoM #${roundNum} — Meeting Date`, 'Enter the date of these meetings. Leave blank for N/A.', ui.ButtonSet.OK_CANCEL);
  if (dateResp.getSelectedButton() !== ui.Button.OK) return;
  const timeResp = ui.prompt(`MoM #${roundNum} — Meeting Time`, 'Enter the time. Leave blank for N/A.', ui.ButtonSet.OK_CANCEL);
  if (timeResp.getSelectedButton() !== ui.Button.OK) return;
  const locationResp = ui.prompt(`MoM #${roundNum} — Meeting Location`, 'Enter the location or meeting method. Leave blank for N/A.', ui.ButtonSet.OK_CANCEL);
  if (locationResp.getSelectedButton() !== ui.Button.OK) return;
  const hodResp = ui.alert(`MoM #${roundNum} — HoD Attendance`, 'Is the Head of Department present at these meetings?', ui.ButtonSet.YES_NO);

  const meetingInfo = {
    date: dateResp.getResponseText().trim(),
    time: timeResp.getResponseText().trim(),
    location: locationResp.getResponseText().trim(),
    hod: hodResp === ui.Button.YES ? 'Present' : 'N/A'
  };

  const template = HtmlService.createTemplateFromFile('MoMCompose');
  template.roundNum = roundNum;
  template.meetingInfoJson = JSON.stringify(meetingInfo);
  const html = template.evaluate().setWidth(900).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, `Generate MoM — Round ${roundNum}`);
}

/** Pulls each of a teacher's students' actual data for this round, as a
 * starting point for the AD's own observations — not a replacement for them. */
function buildTeacherObservationSummary_(teacherName, roundNum) {
  const teacherRow = teacherRosterRowByName_(teacherName);
  if (!teacherRow || !teacherRow.workbookLink) return 'No workbook data available.';
  let ss;
  try { ss = SpreadsheetApp.openById(extractFileIdFromUrl_(teacherRow.workbookLink)); }
  catch (e) { return 'Could not open her workbook.'; }

  const lines = [];
  if (roundNum === 2) {
    const sh = ss.getSheetByName('Week 5');
    if (!sh || sh.getLastRow() < 4) return 'No Week 5 data entered yet.';
    const c = WEEK5_COLS;
    const data = sh.getRange(4, 1, sh.getLastRow() - 3, c.STATUS).getValues();
    data.forEach(row => {
      const name = row[c.NAME - 1];
      if (!name) return;
      const cum = row[c.CUMULATIVE_1 - 1];
      const nextWeak = row[c.WEAK1_AP2 - 1];
      const summary = row[c.WRITTEN_SUMMARY - 1];
      const prStatus = row[c.PROGRESS_REPORT_STATUS - 1];
      lines.push(`• ${name}: Cumulative Score 1 = ${cum || 'not yet entered'}. ` +
        `${summary ? 'Summary: "' + summary + '". ' : ''}` +
        `${nextWeak ? 'Continuing focus into AP2: "' + nextWeak + '". ' : ''}` +
        `Progress Report: ${prStatus || 'Pending'}.`);
    });
  } else if (roundNum === 3) {
    const sh = ss.getSheetByName('Week 10');
    if (!sh || sh.getLastRow() < 4) return 'No Week 10 data entered yet.';
    const c = WEEK10_COLS;
    const data = sh.getRange(4, 1, sh.getLastRow() - 3, c.STATUS).getValues();
    data.forEach(row => {
      const name = row[c.NAME - 1];
      if (!name) return;
      const cum = row[c.CUMULATIVE_2 - 1];
      const summary = row[c.WRITTEN_SUMMARY - 1];
      lines.push(`• ${name}: Cumulative Score 2 = ${cum || 'not yet entered'}. ` +
        `${summary ? 'Final summary: "' + summary + '".' : ''}`);
    });
  }
  return lines.length ? lines.join('\n') : 'No student rows found yet.';
}

/** Called from the compose dialog on load — every teacher ready for this
 * round (attendance marked, not yet generated). */
function getMoMComposeData(roundNum) {
  const roundConfig = MOM_ROUNDS.find(r => r.round === roundNum);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MOM_LOG);
  const lastRow = sh.getLastRow();
  const teachers = [];
  if (lastRow >= 2) {
    for (let row = 2; row <= lastRow; row++) {
      const teacherName = sh.getRange(row, 1).getValue();
      if (!teacherName) continue;
      const attendance = sh.getRange(row, roundConfig.attCol).getValue();
      if (!attendance) continue;
      if (sh.getRange(row, roundConfig.linkCol).getValue()) continue;
      teachers.push({
        name: teacherName,
        attendance: attendance,
        summary: buildTeacherObservationSummary_(teacherName, roundNum)
      });
    }
  }

  return { teachers: teachers, roundLabel: roundConfig.label };
}

/** Generates the MoM PDF(s) for Round 2/3 with the compose dialog's
 * per-teacher observation notes baked into each document. Sends nothing. */
function submitMoMCompose(roundNum, generalNote, teacherNotes, meetingInfo) {
  const settings = readSettingsMap_();
  const folderId = settings['MoM Folder ID'];
  const templateId = settings['MoM Template ID'];
  const termFolder = getTermFolder_(folderId, settings);
  const roundConfig = MOM_ROUNDS.find(r => r.round === roundNum);
  const content = MOM_CONTENT[roundNum];

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MOM_LOG);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 'MoM Log is empty.';

  const adName = readContacts_()['Academic Dean'].name;
  const termStart = TERM_START_DATES[String(settings['Term # (1, 2, or 3)'])];
  const followUpDate = (content.followUpWeekOffset && termStart)
    ? new Date(termStart.getTime() + content.followUpWeekOffset * 7 * 86400000) : null;

  let generated = 0, failed = 0;
  const failedNames = [];

  for (let row = 2; row <= lastRow; row++) {
    const teacherName = sh.getRange(row, 1).getValue();
    if (!teacherName) continue;
    const attendance = sh.getRange(row, roundConfig.attCol).getValue();
    if (!attendance) continue;
    if (sh.getRange(row, roundConfig.linkCol).getValue()) continue;

    try {
      const teacherRow = teacherRosterRowByName_(teacherName);
      const autoSummary = buildTeacherObservationSummary_(teacherName, roundNum);
      const specificNote = teacherNotes[teacherName] || '';
      const fileName = `MoM ${roundNum} - ${teacherName}`;
      const docCopy = DriveApp.getFileById(templateId).makeCopy(fileName, termFolder);
      const doc = DocumentApp.openById(docCopy.getId());
      fillMoMDoc_(doc, teacherName, teacherRow, adName, attendance, content, followUpDate, autoSummary, generalNote, specificNote, meetingInfo);
      doc.saveAndClose();

      const pdfBlob = DriveApp.getFileById(docCopy.getId()).getAs('application/pdf').setName(fileName + '.pdf');
      const pdfFile = termFolder.createFile(pdfBlob);
      applyLinkSharing_(pdfFile, settings, teacherRow ? teacherRow.email : null);
      docCopy.setTrashed(true);

      const url = pdfFile.getUrl();
      sh.getRange(row, roundConfig.linkCol).setFormula(`=HYPERLINK("${url}", "Open MoM")`);
      generated++;
    } catch (e) {
      failed++;
      failedNames.push(`${teacherName}: ${e.message}`);
    }
  }

  let msg = `Round ${roundNum}: ${generated} MoM PDF(s) generated in Drive. Nothing was emailed — run "Send MoM — Round ${roundNum}" separately when ready.`;
  if (failed) msg += ` ${failed} FAILED — ` + failedNames.slice(0, 10).join('; ');
  return msg;
}

// ---------- Send MoM — entirely separate from generation ----------

function openSendMoMDialogRound1() { openSendMoMDialog_(1); }
function openSendMoMDialogRound2() { openSendMoMDialog_(2); }
function openSendMoMDialogRound3() { openSendMoMDialog_(3); }

function openSendMoMDialog_(roundNum) {
  const template = HtmlService.createTemplateFromFile('MoMSendPreview');
  template.roundNum = roundNum;
  const html = template.evaluate().setWidth(700).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, `Send MoM — Round ${roundNum}`);
}

/** Called from the send-preview dialog on load. Everyone in this round
 * who has a generated MoM PDF but hasn't been emailed yet — this whole
 * group goes out together, in ONE shared email. */
function getMoMSendPreviewData(roundNum) {
  const roundConfig = MOM_ROUNDS.find(r => r.round === roundNum);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MOM_LOG);
  const lastRow = sh.getLastRow();
  const teachers = [];
  if (lastRow >= 2) {
    for (let row = 2; row <= lastRow; row++) {
      const teacherName = sh.getRange(row, 1).getValue();
      if (!teacherName) continue;
      const linkFormula = sh.getRange(row, roundConfig.linkCol).getFormula();
      const alreadySent = sh.getRange(row, roundConfig.emailSentCol).getValue() === 'Sent';
      if (!linkFormula || alreadySent) continue;
      const teacherRow = teacherRosterRowByName_(teacherName);
      teachers.push({ name: teacherName, email: teacherRow ? teacherRow.email : '' });
    }
  }

  const contacts = readContacts_();
  const activeRoles = Object.keys(contacts)
    .filter(role => contacts[role].active && contacts[role].email)
    .map(role => ({ role: role, name: contacts[role].name, email: contacts[role].email }));

  return { teachers: teachers, activeRoles: activeRoles, roundLabel: roundConfig.label };
}

/** Sends each ready teacher their OWN individual email — just that
 * teacher in the To field, only their own MoM PDF attached, active
 * Settings contacts in CC, plus whatever additional note was typed
 * in the preview. One send per teacher, so a failure for one doesn't
 * affect the others. */
function submitMoMSend(roundNum, additionalNote, ccRoleEmails) {
  const roundConfig = MOM_ROUNDS.find(r => r.round === roundNum);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MOM_LOG);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 'MoM Log is empty.';

  const recipients = [];

  for (let row = 2; row <= lastRow; row++) {
    const teacherName = sh.getRange(row, 1).getValue();
    if (!teacherName) continue;
    const linkCell = sh.getRange(row, roundConfig.linkCol);
    const linkFormula = linkCell.getFormula();
    const alreadySent = sh.getRange(row, roundConfig.emailSentCol).getValue() === 'Sent';
    if (!linkFormula || alreadySent) continue;

    const teacherRow = teacherRosterRowByName_(teacherName);
    if (!teacherRow || !teacherRow.email) continue;

    try {
      const match = linkFormula.match(/HYPERLINK\("([^"]+)"/);
      const pdfUrl = match ? match[1] : null;
      if (!pdfUrl) throw new Error('no link found');
      const pdfFileId = extractFileIdFromUrl_(pdfUrl);
      const pdfBlob = DriveApp.getFileById(pdfFileId).getBlob();
      recipients.push({ row: row, name: teacherName, email: teacherRow.email, attachment: pdfBlob });
    } catch (e) {
      // skip this one silently, don't stop the whole batch
    }
  }

  if (!recipients.length) return 'No one is ready to be sent for this round.';

  const cc = (ccRoleEmails && ccRoleEmails.length) ? ccRoleEmails.join(',') : undefined;
  const roundLabel = roundConfig.label;
  const sentAt = new Date();
  let sentCount = 0, failedCount = 0;

  recipients.forEach(r => {
    const firstName = String(r.name).split(' ')[0];
    const body = emailWrapper_(roundLabel,
      `<p>Dear ${firstName},</p>
      <p>Please find attached your Minutes of Meeting for ${roundLabel}.</p>
      ${additionalNote ? `<p style="background:#f4f4f4;padding:10px;border-left:4px solid #1c4587;">${additionalNote}</p>` : ''}`);

    const options = { htmlBody: body, attachments: [r.attachment], name: 'Academic Dean\'s Office' };
    if (cc) options.cc = cc;

    let status = 'Sent', error = '';
    try {
      GmailApp.sendEmail(r.email, roundLabel, '', options);
      sentCount++;
    } catch (e) {
      status = 'Failed';
      error = e.message;
      failedCount++;
    }

    // One Email Log entry per recipient — this is purely the record, it
    // doesn't send anything a second time.
    logEmailOnly_(r.email, roundLabel, 'MoM Batch Send', '', status, error);
  });

  // Email Sent MoM column is a protected live formula pulled from Email Log,
  // so it updates itself once logEmailOnly_ writes the new Email Log rows above —
  // nothing to write here, and writing would fail against the protection anyway.

  const when = Utilities.formatDate(sentAt, Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm');
  if (failedCount === 0) {
    return `Sent to ${sentCount} teacher(s) individually at ${when}, each with only their own MoM PDF attached.`;
  }
  return `Sent to ${sentCount} teacher(s) individually at ${when}; ${failedCount} failed to send.`;
}
