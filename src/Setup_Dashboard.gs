/**
 * SETUP_DASHBOARD.gs
 * 7 cards. Cards 1-5 are live formulas reading Student Roster / Teacher
 * Roster directly. Cards 6-7 need data from every teacher's OWN
 * workbook (who hasn't marked attendance, which students are most
 * absent) — that can't be a live formula, so "Refresh Dashboard
 * Attendance Data" pulls it in, same pattern as Scores Progress.
 *
 * Card colors are just background fills on each section header — Sheets
 * can't do true floating "cards," so this is the closest equivalent:
 * distinct colored blocks stacked vertically.
 */

const CARD_COLORS = {
  students: '#1c4587', es: '#0b8043', ms: '#e69138', hs: '#a64d79',
  ad: '#674ea7', attendance: '#cc0000', absence: '#bf9000'
};

function buildDashboardShell() {
  const sh = getOrCreateSheet_(SHEETS.DASHBOARD);
  sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();

  sh.getRange('A1').setValue('ACADEMIC DEAN — REMEDIAL FOLLOW-UP DASHBOARD')
    .setFontWeight('bold').setFontSize(14);
  sh.getRange('A2').setValue('Cards 1–5 are live.')
    .setFontStyle('italic').setFontColor('#666666');  

  const R = `'${SHEETS.ROSTER}'`;
  let row = 4;

  // ============ CARD 1: STUDENTS ============
  row = cardHeader_(sh, row, 'CARD 1 — STUDENTS', CARD_COLORS.students);
  row = cardRow_(sh, row, 'Total students tracked (unique, has probation letter)',
    `=IF(COUNTIF(${R}!M2:M301,"<>")=0,0,COUNTA(UNIQUE(FILTER(${R}!A2:A301,${R}!M2:M301<>""))))`);
  row = cardRow_(sh, row, '  · Not Tiered', `=IF(COUNTIF(${R}!G2:G301,"Not Tiered (Elementary)")=0,0,COUNTA(UNIQUE(FILTER(${R}!A2:A301,${R}!G2:G301="Not Tiered (Elementary)"))))`);
  row = cardRow_(sh, row, '  · Tier 1', `=IF(COUNTIF(${R}!G2:G301,"Tier 1")=0,0,COUNTA(UNIQUE(FILTER(${R}!A2:A301,${R}!G2:G301="Tier 1"))))`);
  row = cardRow_(sh, row, '  · Tier 2', `=IF(COUNTIF(${R}!G2:G301,"Tier 2")=0,0,COUNTA(UNIQUE(FILTER(${R}!A2:A301,${R}!G2:G301="Tier 2"))))`);
  ['English', 'Science', 'Social Studies', 'Mathematics', 'Arabic', 'Islamic Religion', 'Arabic Social Studies'].forEach(subj => {
    row = cardRow_(sh, row, `  · ${subj} enrollments`, `=COUNTIF(${R}!E2:E301,"${subj}")`);
  });
  row += 1;

  // ============ CARDS 2-4: DIVISION BREAKDOWNS ============
  const divisions = [
    { label: 'CARD 2 — ELEMENTARY (G1-G5)', color: CARD_COLORS.es, lo: 1, hi: 5 },
    { label: 'CARD 3 — MIDDLE SCHOOL (G6-G8)', color: CARD_COLORS.ms, lo: 6, hi: 8 },
    { label: 'CARD 4 — HIGH SCHOOL (G9-G12)', color: CARD_COLORS.hs, lo: 9, hi: 12 }
  ];
  divisions.forEach(div => {
    row = cardHeader_(sh, row, div.label, div.color);
    row = cardRow_(sh, row, 'Total Teachers',
      `=IF(COUNTIFS(${R}!C2:C301,">="&${div.lo},${R}!C2:C301,"<="&${div.hi})=0,0,COUNTA(UNIQUE(FILTER(${R}!I2:I301,${R}!C2:C301>=${div.lo},${R}!C2:C301<=${div.hi}))))`);
    ['English', 'Science', 'Social Studies', 'Mathematics', 'Arabic', 'Arabic Social Studies', 'Islamic Religion'].forEach(subj => {
      row = cardRow_(sh, row, `  · ${subj}`,
        `=COUNTIFS(${R}!E2:E301,"${subj}",${R}!C2:C301,">="&${div.lo},${R}!C2:C301,"<="&${div.hi})`);
    });
    row += 1;
  });

  // ============ CARD 5: AD CARD ============
  row = cardHeader_(sh, row, 'CARD 5 — AD CARD', CARD_COLORS.ad);
  const totalStudents = `COUNTA(UNIQUE(${R}!A2:A301))`;
  row = cardRow_(sh, row, 'Probation letters — uploaded', `=IF(COUNTIF(${R}!M2:M301,"<>")=0,0,COUNTA(UNIQUE(FILTER(${R}!A2:A301,${R}!M2:M301<>""))))`);
  row = cardRow_(sh, row, 'Probation letters — total students', `=${totalStudents}`);
  row = cardRow_(sh, row, 'Probation letters — remaining',
    `=${totalStudents}-IF(COUNTIF(${R}!M2:M301,"<>")=0,0,COUNTA(UNIQUE(FILTER(${R}!A2:A301,${R}!M2:M301<>""))))`);
  row = cardRow_(sh, row, 'Action Plan 1 — generated', `=COUNTIFS(${R}!A2:A301,"<>",${R}!N2:N301,"<>")`);
  row = cardRow_(sh, row, 'Action Plan 1 — remaining (see list below)', `=COUNTIFS(${R}!A2:A301,"<>",${R}!N2:N301,"")`);
  row = cardRow_(sh, row, 'Action Plan 2 — closed out (PDF)', `=COUNTIFS(${R}!A2:A301,"<>",${R}!O2:O301,"<>")`);
  row = cardRow_(sh, row, 'Action Plan 2 — remaining (see list below)', `=COUNTIFS(${R}!A2:A301,"<>",${R}!O2:O301,"")`);
  row = cardRow_(sh, row, 'Progress Report — generated (PDF)', `=COUNTIFS(${R}!A2:A301,"<>",${R}!P2:P301,"<>")`);
  row = cardRow_(sh, row, 'Progress Report — remaining (see list below)', `=COUNTIFS(${R}!A2:A301,"<>",${R}!P2:P301,"")`);
  const teacherOnTrackRange = "'_TeacherTrackStatus'!B2:B200";
  row = cardRow_(sh, row, 'Teachers On Track', `=COUNTIF(${teacherOnTrackRange},"On Track")`);
  row = cardRow_(sh, row, 'Teachers Off Track (see list below)', `=COUNTIF(${teacherOnTrackRange},"Off Track")`);
  row += 1;

  // ---- Remaining-item lists for Card 5 ----
  row = cardHeader_(sh, row, 'Missing Action Plan 1 — list', '#999999');
  sh.getRange(row, 1).setFormula(
    `=IFERROR(QUERY(${R}!A2:P301,"select B,E,I where A is not null and N = ''",0),"None")`
  );
  row += 12;
  row = cardHeader_(sh, row, 'Missing Progress Report — list', '#999999');
  sh.getRange(row, 1).setFormula(`=IFERROR(QUERY(${R}!A2:P301,"select B,E,I where A is not null and P = ''",0),"None")`);
  row += 12;
  row = cardHeader_(sh, row, 'Off-Track Teachers — list', '#999999');
  sh.getRange(row, 1).setFormula(`=IFERROR(QUERY('_TeacherTrackStatus'!A2:B200,"select A where B = 'Off Track'",0),"None")`);
  row += 8;

  sh.setColumnWidth(1, 320);
  sh.setColumnWidth(2, 140);

  buildTeacherTrackStatusHelper_();
  sh.getRange(1, 2, sh.getLastRow(), 1).setHorizontalAlignment('center'); // center every numeric card value in column B
  sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).setBorder(true, true, true, true, true, true);
  sh.autoResizeColumns(1, sh.getLastColumn());

  SpreadsheetApp.getActiveSpreadsheet().toast('Dashboard built — 5 cards.', 'Done', 7);
}

function cardHeader_(sh, row, title, color) {
  sh.getRange(row, 1, 1, 4).merge().setValue(title)
    .setFontWeight('bold').setBackground(color).setFontColor('#ffffff').setFontSize(11);
  return row + 1;
}
function cardRow_(sh, row, label, formula) {
  sh.getRange(row, 1).setValue(label);
  sh.getRange(row, 2).setFormula(formula);
  return row + 1;
}

/** Hidden helper: On Track / Off Track per teacher, based on AP1/AP2/PR completeness across her students. */
function buildTeacherTrackStatusHelper_() {
  const sh = getOrCreateSheet_('_TeacherTrackStatus');
  sh.clear();
  sh.getRange(1, 1, 1, 2).setValues([['Teacher', 'Status']]);
  const roster = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.TEACHER_ROSTER);
  const lastRow = roster.getLastRow();
  if (lastRow < 2) { sh.hideSheet(); return; }
  const names = roster.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(String);
  const R = SHEETS.ROSTER;
  const rows = names.map((name, i) => {
    const r = i + 2;
    return [name, `=IF(COUNTIFS('${R}'!I2:I301,"*"&A${r}&"*",'${R}'!N2:N301,"")+` +
      `COUNTIFS('${R}'!I2:I301,"*"&A${r}&"*",'${R}'!O2:O301,"")+` +
      `COUNTIFS('${R}'!I2:I301,"*"&A${r}&"*",'${R}'!P2:P301,"")=0,"On Track","Off Track")`];
  });
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 1).setValues(rows.map(r => [r[0]]));
    rows.forEach((r, i) => sh.getRange(2 + i, 2).setFormula(r[1]));
  }
  sh.hideSheet();
}

function principalContactForDivision_(contacts, division) {
  const key = division === 'ES' ? 'Elementary School Principal' : division === 'MS' ? 'Middle School Principal' : 'High School Principal';
  const c = contacts[key];
  return (c && c.active && c.email) ? c : null;
}