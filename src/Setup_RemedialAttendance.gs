/**
 * SETUP_REMEDIALATTENDANCE.gs
 * Teacher punctuality tracker for the actual remedial teaching sessions
 * (Weeks 2–10 — Week 1 is orientation only, no remedial session yet).
 * Rows are grouped by Division, then Subject within each division, so
 * all Elementary sits together, all English together within it, etc. —
 * not a flat alphabetical list.
 */

const ATTENDANCE_STATUS_OPTIONS = ['On Time', 'Late', 'Absent (Replacement)'];
const REMEDIAL_WEEKS = 9; // Weeks 2 through 10

function buildRemedialAttendance() {
  const sh = getOrCreateSheet_(SHEETS.ATTENDANCE_TEACHERS);

  // Read and preserve every existing value BEFORE clearing anything —
  // Day, Time, all 9 weeks' status, and Notes, keyed by Teacher|Subject|Division
  // so a rebuild never loses what's already been entered.
  const existing = {};
  const oldLastRow = sh.getLastRow();
  if (oldLastRow > 1) {
    const oldLastCol = sh.getLastColumn();
    sh.getRange(2, 1, oldLastRow - 1, oldLastCol).getValues().forEach(r => {
      const key = `${r[0]}|${r[1]}|${r[2]}`;
      existing[key] = { day: r[3], time: r[4], weeks: r.slice(5, 5 + REMEDIAL_WEEKS), notes: r[5 + REMEDIAL_WEEKS] };
    });
  }

  // Must clear validations BEFORE sh.clear() and across the FULL sheet range —
  // clearing after (or only on the shrunk post-clear data range) misses rules
  // left over from a previous column layout.
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  sh.clear();

  const headers = ['Teacher Name', 'Subject', 'Division', 'Day', 'Time'];
  for (let w = 2; w <= 10; w++) headers.push(`Week ${w}`);
  headers.push('Notes (replacement details)');

  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#1c4587').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.setFrozenColumns(1);

  sh.setColumnWidth(1, 170);
  sh.setColumnWidth(2, 130);
  sh.setColumnWidth(3, 90);
  sh.setColumnWidth(4, 100);
  sh.setColumnWidth(5, 90);
  const weekColStart = 6;
  for (let w = 0; w < REMEDIAL_WEEKS; w++) sh.setColumnWidth(weekColStart + w, 85);
  sh.setColumnWidth(weekColStart + REMEDIAL_WEEKS, 280);

  // Pull directly from Student Roster — NOT Teacher Roster — because
  // Teacher Roster collapses a teacher's subjects into one comma-joined
  // string ("Math, Science"), losing which class belongs to which
  // subject. Student Roster still has the real Teacher+Subject+Class
  // pairing per row, so a teacher gets ONE ROW PER (Subject, Division)
  // combination she actually teaches — not just per division, and not
  // one row per teacher. A teacher teaching two subjects on two different
  // days, or one subject across two divisions, both come
  // out correct this way.
  const rosterSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ROSTER);
  const rosterLastRow = rosterSh ? rosterSh.getLastRow() : 1;
  if (rosterLastRow < 2) return;

  const seen = new Set(); // dedupe by teacher|subject|division
  const rows = [];
  rosterSh.getRange(2, 1, rosterLastRow - 1, ROSTER_HEADERS.length).getValues().forEach(r => {
    const teacherField = r[ROSTER_COL.TEACHER - 1];
    const subject = r[ROSTER_COL.SUBJECT - 1];
    const grade = r[ROSTER_COL.GRADE - 1];
    if (!teacherField) return;
    const m = String(teacherField).match(/^(.*?)\s*\(/);
    const teacherName = m ? m[1].trim() : String(teacherField).trim();
    const division = divisionFromGrade_(grade) || '—';
    const key = `${teacherName}|${subject}|${division}`;
    if (seen.has(key)) return;
    seen.add(key);
    rows.push({ name: teacherName, subject, division });
  });

  const divisionOrder = { ES: 0, MS: 1, HS: 2, '—': 3 };
  rows.sort((a, b) => {
    if (divisionOrder[a.division] !== divisionOrder[b.division]) return divisionOrder[a.division] - divisionOrder[b.division];
    if (a.subject !== b.subject) return String(a.subject).localeCompare(String(b.subject));
    return a.name.localeCompare(b.name);
  });

  if (rows.length) {
    const values = rows.map(r => {
      const key = `${r.name}|${r.subject}|${r.division}`;
      const prev = existing[key];
      const weeks = prev ? prev.weeks : Array(REMEDIAL_WEEKS).fill('');
      return [r.name, r.subject, r.division, prev ? prev.day : '', prev ? prev.time : '', ...weeks, prev ? prev.notes : ''];
    });
    sh.getRange(2, 1, values.length, headers.length).setValues(values);

    const statusRange = sh.getRange(2, weekColStart, rows.length, REMEDIAL_WEEKS);
    statusRange.setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(ATTENDANCE_STATUS_OPTIONS).build()
    );
    const lateRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Late').setBackground('#fff2cc').setRanges([statusRange]).build();
    const absentRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Absent (Replacement)').setBackground('#f4cccc').setRanges([statusRange]).build();
    sh.setConditionalFormatRules([lateRule, absentRule]);

    // Shade a visual break between division blocks
    let prevDivision = null;
    rows.forEach((r, i) => {
      if (prevDivision !== null && r.division !== prevDivision) {
        sh.getRange(i + 2, 1, 1, headers.length).setBorder(true, null, null, null, null, null, '#666666', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }
      prevDivision = r.division;
    });

    sh.getRange(1, 1, values.length + 1, headers.length).setBorder(true, true, true, true, true, true);
    sh.getRange(2, 3, values.length, 1).setHorizontalAlignment('center'); // Division
    sh.getRange(2, weekColStart, values.length, REMEDIAL_WEEKS).setHorizontalAlignment('center'); // Week status dropdowns
    sh.getRange(1, 1, 1, headers.length).setHorizontalAlignment('center');
    sh.autoResizeColumns(1, headers.length);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Remedial Attendance built — ${rows.length} teacher(s), grouped by Division then Subject, ` +
    'Weeks 2–10. Fill in Day/Time manually.', 'Done', 7
  );
}

/**
 * Overall attendance % for a teacher across ALL her rows — a teacher
 * spanning multiple divisions (e.g. both ES and MS) now has more than
 * one row here, so this aggregates across every row matching her name,
 * not just the first. On Time and Late both count as "attended"; only
 * Absent counts against her.
 */
function attendancePercentForTeacher_(teacherName) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ATTENDANCE_TEACHERS);
  if (!sh) return null;
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  const data = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const header = data[0];
  const weekCols = header.map((h, i) => (String(h).startsWith('Week') ? i : -1)).filter(i => i >= 0);

  const matchingRows = data.slice(1).filter(r => r[0] === teacherName);
  if (!matchingRows.length) return null;

  let attended = 0, total = 0;
  matchingRows.forEach(row => {
    weekCols.forEach(c => {
      const v = row[c];
      if (!v) return;
      total++;
      if (v !== 'Absent (Replacement)') attended++;
    });
  });
  return total ? Math.round((attended / total) * 100) : null;
}

/**
 * Returns the current remedial week's status for a teacher, per division
 * (a multi-division teacher gets one entry per division she teaches).
 */
function currentRemedialStatusForTeacher_(teacherName, currentWeek) {
  if (!currentWeek || currentWeek < 2) return [];
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ATTENDANCE_TEACHERS);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const data = sh.getRange(2, 1, lastRow - 1, 5 + REMEDIAL_WEEKS + 1).getValues();
  const weekCol = 5 + (currentWeek - 2); // 0-indexed: Week2 is at offset 0
  const results = [];
  data.forEach(row => {
    if (row[0] === teacherName) {
      results.push({ subject: row[1], division: row[2], status: row[weekCol], notes: row[5 + REMEDIAL_WEEKS] });
    }
  });
  return results;
}