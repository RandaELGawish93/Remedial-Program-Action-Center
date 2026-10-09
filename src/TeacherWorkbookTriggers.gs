/**
 * TEACHERWORKBOOKTRIGGERS.gs
 *
 * IMPORTANT ARCHITECTURE NOTE: Google caps a script at 20 installable
 * triggers TOTAL. This version uses exactly ONE time-driven trigger for
 * the whole system, regardless of teacher count — periodicTeacherWorkbookSync()
 * loops internally every ~12 seconds for up to ~50 seconds per firing,
 * so a ticked checkbox is usually picked up within 10-15 seconds instead
 * of waiting up to a full minute.
 *
 * Week 1 / Week 5 / Week 10 each have a "Generate?" checkbox column —
 * the teacher ticks it, and the next sync pass does the generation and
 * unticks it when done.
 */

function installPeriodicSync() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'periodicTeacherWorkbookSync') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('periodicTeacherWorkbookSync').timeBased().everyMinutes(1).create();
  SpreadsheetApp.getUi().alert('Periodic sync installed — runs every 1 minute across all teacher workbooks.');
}

function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  SpreadsheetApp.getUi().alert(`Deleted ${triggers.length} trigger(s). Now run "Install Periodic Sync" once.`);
}

function periodicTeacherWorkbookSync() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;

  try {
    const startTime = new Date().getTime();
    const MAX_LOOP_MS = 50 * 1000;
    const SLEEP_BETWEEN_PASSES_MS = 12 * 1000;

    while (true) {
      runOneSyncPass_();
      if (new Date().getTime() - startTime > MAX_LOOP_MS) break;
      Utilities.sleep(SLEEP_BETWEEN_PASSES_MS);
      if (new Date().getTime() - startTime > MAX_LOOP_MS) break;
    }
  } finally {
    lock.releaseLock();
  }
}

function runOneSyncPass_() {
  const trSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.TEACHER_ROSTER);
  const lastRow = trSh.getLastRow();
  if (lastRow < 2) return;
  const teacherRows = trSh.getRange(2, 1, lastRow - 1, TEACHER_ROSTER_HEADERS.length).getValues();

  teacherRows.forEach(row => {
    const name = row[0], workbookUrl = row[TEACHER_ROSTER_COL.WORKBOOK_LINK - 1];
    if (!name || !workbookUrl) return;
    const fileId = extractFileIdFromUrl_(workbookUrl);
    let ss;
    try { ss = SpreadsheetApp.openById(fileId); } catch (e) { return; }

    try {
      syncOneTeacherProgressReportStatus_(ss);
      processGenerateCheckboxes_(ss, name);
    } catch (e) {
      // Swallow and move to the next teacher — this one's sync failed
      // this cycle, but everyone else still gets processed.
    }
  });
}

function syncOneTeacherProgressReportStatus_(ss) {
  const prSh = ss.getSheetByName('Progress Report Entries');
  const week5 = ss.getSheetByName('Week 5');
  if (!prSh || !week5) return;

  const numStatements = PR_STATEMENTS.length;
  const statusCol = 3 + numStatements;
  const prLastRow = prSh.getLastRow();
  if (prLastRow < 2) return;
  const doneIds = new Set();
  prSh.getRange(2, 1, prLastRow - 1, statusCol).getValues().forEach(r => {
    if (r[statusCol - 1] === 'Done') doneIds.add(String(r[0]));
  });
  if (!doneIds.size) return;

  const week5LastRow = week5.getLastRow();
  if (week5LastRow < 4) return;
  const prStatusCol = WEEK5_COLS.PROGRESS_REPORT_STATUS;
  const week5Data = week5.getRange(4, 1, week5LastRow - 3, prStatusCol).getValues();
  week5Data.forEach((r, i) => {
    if (doneIds.has(String(r[0])) && r[prStatusCol - 1] !== 'Filled') {
      week5.getRange(4 + i, prStatusCol).setValue('Filled');
    }
  });
}

function processGenerateCheckboxes_(ss, teacherName) {
  processWeek1Checkboxes_(ss);
  processWeek5Checkboxes_(ss);
  processWeek10Checkboxes_(ss);
}

function processWeek1Checkboxes_(ss) {
  const sh = ss.getSheetByName('Week 1');
  if (!sh) return;
  const lastRow = sh.getLastRow();
  const checkboxCol = WEEK1_COLS.GENERATE;
  for (let row = 4; row <= lastRow; row++) {
    if (sh.getRange(row, checkboxCol).getValue() !== true) continue;
    const studentName = sh.getRange(row, 2).getValue();
    try {
      generateOrFillActionPlan1FromRow_(ss, row);
      sh.getRange(row, checkboxCol).setValue(false);
      sh.getRange(row, checkboxCol).clearNote();
      ss.toast(`Action Plan generated for ${studentName}.`, 'Done', 5);
    } catch (e) {
      sh.getRange(row, checkboxCol).setValue(false);
      sh.getRange(row, checkboxCol).setNote(`Generation failed: ${e.message}`);
      ss.toast(`${studentName}: ${e.message}`, '⚠️ Action Plan 1 failed', 8);
    }
  }
}

function processWeek5Checkboxes_(ss) {
  const sh = ss.getSheetByName('Week 5');
  if (!sh) return;
  const lastRow = sh.getLastRow();
  const checkboxCol = WEEK5_COLS.GENERATE;
  for (let row = 4; row <= lastRow; row++) {
    if (sh.getRange(row, checkboxCol).getValue() !== true) continue;
    const studentName = sh.getRange(row, 2).getValue();
    try {
      fillActionPlan2FromWeek5Row_(ss, row);
      sh.getRange(row, checkboxCol).setValue(false);
      sh.getRange(row, checkboxCol).clearNote();
      ss.toast(`Action Plan 1 finalized (Phase 1 PDF exported) for ${studentName}.`, 'Done', 5);
    } catch (e) {
      sh.getRange(row, checkboxCol).setValue(false);
      sh.getRange(row, checkboxCol).setNote(`Generation failed: ${e.message}`);
      ss.toast(`${studentName}: ${e.message}`, '⚠️ Action Plan 2 failed', 8);
    }
  }
}

function processWeek10Checkboxes_(ss) {
  const sh = ss.getSheetByName('Week 10');
  if (!sh) return;
  const lastRow = sh.getLastRow();
  const checkboxCol = WEEK10_COLS.GENERATE;
  for (let row = 4; row <= lastRow; row++) {
    if (sh.getRange(row, checkboxCol).getValue() !== true) continue;
    const studentName = sh.getRange(row, 2).getValue();
    try {
      fillActionPlan10CloseoutFromRow_(ss, row);
      sh.getRange(row, checkboxCol).setValue(false);
      sh.getRange(row, checkboxCol).clearNote();
      ss.toast(`Action Plan 2 closed out (Phase 2 PDF exported) for ${studentName}.`, 'Done', 5);
    } catch (e) {
      sh.getRange(row, checkboxCol).setValue(false);
      sh.getRange(row, checkboxCol).setNote(`Generation failed: ${e.message}`);
      ss.toast(`${studentName}: ${e.message}`, '⚠️ Action Plan 2 closeout failed', 8);
    }
  }
}

function extractTeacherNameFromWorkbook_(ss) {
  const profile = ss.getSheetByName('Profile');
  if (!profile) return null;
  return profile.getRange('B3').getValue();
}

function writeActionPlanLinkToMasterRoster_(studentName, teacherName, url) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ROSTER);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  const data = sh.getRange(2, 1, lastRow - 1, ROSTER_HEADERS.length).getValues();
  for (let i = 0; i < data.length; i++) {
    const name = data[i][ROSTER_COL.NAME - 1];
    const teacherField = data[i][ROSTER_COL.TEACHER - 1];
    if (name === studentName && String(teacherField).startsWith(teacherName)) {
      sh.getRange(i + 2, ROSTER_COL.ACTION_PLAN).setValue(url);
    }
  }
}

function syncTeacherProfiles() {
  const settings = readSettingsMap_();
  const termStart = TERM_START_DATES[String(settings['Term # (1, 2, or 3)'])];
  const currentWeek = termStart ? Math.max(1, Math.min(20, Math.floor((new Date() - termStart) / (7 * 86400000)) + 1)) : '—';

  const trSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.TEACHER_ROSTER);
  const lastRow = trSh.getLastRow();
  if (lastRow < 2) return;
  const teacherRows = trSh.getRange(2, 1, lastRow - 1, TEACHER_ROSTER_HEADERS.length).getValues();

  const momSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MOM_LOG);
  const attSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ATTENDANCE_TEACHERS);

  let synced = 0;
  teacherRows.forEach(row => {
    const name = row[0], workbookUrl = row[TEACHER_ROSTER_COL.WORKBOOK_LINK - 1];
    if (!name || !workbookUrl) return;

    const fileId = extractFileIdFromUrl_(workbookUrl);
    const ss = SpreadsheetApp.openById(fileId);
    const profile = ss.getSheetByName('Profile');
    if (!profile) return;

    profile.getRange('B4').setValue(currentWeek);

    if (momSh) {
      const momRow = findRowByFirstColumn_(momSh, name);
      if (momRow) {
        profile.getRange('B5').setValue(momSh.getRange(momRow, 3).getValue() ? 'Done' : 'Not yet');
        profile.getRange('B6').setValue(momSh.getRange(momRow, 5).getValue() ? 'Done' : 'Not yet');
        profile.getRange('B7').setValue(momSh.getRange(momRow, 7).getValue() ? 'Done' : 'Not yet');
      }
    }
    if (attSh) {
      const attLastRow = attSh.getLastRow();
      const matchingRows = [];
      if (attLastRow > 1) {
        attSh.getRange(2, 1, attLastRow - 1, 3).getValues().forEach((r, i) => {
          if (r[0] === name) matchingRows.push({ rowNum: i + 2, subject: r[1], division: r[2] });
        });
      }
      if (matchingRows.length && typeof currentWeek === 'number' && currentWeek >= 2) {
        const weekCol = 6 + (currentWeek - 2);
        const statuses = matchingRows.map(m => {
          const s = attSh.getRange(m.rowNum, weekCol).getValue();
          return matchingRows.length > 1 ? `${m.subject}/${m.division}: ${s || '—'}` : (s || '—');
        });
        profile.getRange('B8').setValue(statuses.join(' | '));
      } else if (matchingRows.length) {
        profile.getRange('B8').setValue('— (no remedial session yet)');
      }
    }
    profile.getRange('B9').setValue(row[5]);

    synced++;
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(`${synced} teacher profile(s) synced.`, 'Done', 5);
}

function findRowByFirstColumn_(sh, value) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  const values = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === value) return i + 2;
  }
  return null;
}






