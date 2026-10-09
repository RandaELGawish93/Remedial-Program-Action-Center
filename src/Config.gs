/**
 * CONFIG.gs
 * Central place for sheet names, settings keys, and static reference data.
 * Nothing in here talks to the spreadsheet directly — just constants.
 */

/**
 * THEME — shared palette used across every dashboard and workbook in this
 * project so all generated sheets look consistent.
 */
const THEME = {
  darkNavy: '#1a3c5e',   // main titles
  darkBlue: '#1c4587',   // secondary headers
  lightBlue: '#cfe2f3',  // section headers
  darkTeal: '#0b5394',   // KPIs / key highlights
  lightTeal: '#d0e0e3',  // KPI backgrounds
  white: '#ffffff',
  lightGray: '#f3f3f3',  // separators / inactive
  darkGray: '#434343',   // main text
  statusGreen: '#d9ead3',  // completed
  statusYellow: '#fff2cc', // in progress
  statusOrange: '#fce5cd', // at risk / attention
  statusRed: '#f4cccc',    // overdue / critical
  statusPurple: '#d9d2e9', // pending
  statusGray: '#f3f3f3'    // not started
};

const SHEETS = {
  DASHBOARD: 'Dashboard',
  SETTINGS: 'Settings',
  ROSTER: 'Student Roster',
  SCORES: 'Scores Progress',
  TEACHER_ROSTER: 'Teacher Roster',
  ATTENDANCE_TEACHERS: 'Remedial Attendance - Teachers',
  MOM_LOG: 'MoM Log',
  THRESHOLDS_HELPER: '_Thresholds' // hidden helper sheet, do not delete
};

// Row labels for the Contacts block in Settings (Section A)
const CONTACT_ROLES = [
  'Academic Dean',
  'High School Principal',
  'High School Office Assistant',
  'Middle School Principal',
  'Middle School Office Assistant',
  'Elementary School Principal',
  'Elementary School Office Assistant',
  'Campus Director',
  'DCI Elementary Director',
  'DCI Middle & High Director',
  'ADCI',
  'Student Success Center Manager'
];

// Row labels for the System Config block in Settings (Section B)
const SYSTEM_CONFIG_KEYS = [
  'Term # (1, 2, or 3)',
  'Term Start Date (first Sunday of the term)',
  'Number of Weeks in Term',
  '',
  'Action Plan Folder ID',
  'Action Plan English Template ID',
  'Action Plan Arabic Template ID (Arabic, Religion, Arabic Social Studies only)',
  'Progress Report Folder ID',
  'Progress Report Template ID',
  'Probation Letter Folder ID',
  'MoM Folder ID',
  'MoM Template ID',
  'Archive Folder ID',
  '',
  "Teacher's Workbook Folder ID", // where createTeacherWorkbooks() files each teacher's spreadsheet
  'Principals Workbook Folder ID', // where the 3 separate division workbooks get created
  'Teacher Guide File ID (Drive)',
  '',
  'Generated Document Link Sharing',            // dropdown, see LINK_SHARING_OPTIONS
];

const LINK_SHARING_OPTIONS = [
  'Anyone in your organization with the link (recommended)',
  'Anyone with the link (fully open)'
];

// Threshold table for Prevention-program students, by grade.
// Grade 1-5 per the AD's rule; grade 6-12 (MS/HS) flat at 75.
const PREVENTION_THRESHOLDS = {
  1: 90, 2: 88, 3: 85, 4: 80, 5: 75,
  6: 75, 7: 75, 8: 75, 9: 75, 10: 75, 11: 75, 12: 75
};

const PROGRAM_OPTIONS = ['Intervention', 'Prevention'];
const TIER_OPTIONS = ['Not Tiered (Elementary)', 'Tier 1', 'Tier 2'];
const ROSTER_STATUS_OPTIONS = ['On Probation', 'Recommended (Prevention)'];

// Subjects that use the Arabic Action Plan template instead of the English one
const ARABIC_TEMPLATE_SUBJECTS = ['Arabic', 'Islamic Religion', 'Arabic Social Studies'];

// SAMPLE dates — set the first teaching day of each term from your school calendar.
// Update this map at the start of each school year when the new calendar comes out.
const TERM_START_DATES = {
  '1': new Date(2026, 8, 13),  // Sunday, 13 September 2026 — Term 1 begins
  '2': new Date(2026, 11, 7),  // Monday, 7 December 2026 — Term 2 begins
  '3': new Date(2027, 2, 29)   // Monday, 29 March 2027 — Term 3 begins
};

/** Utility: get a sheet, creating it if it doesn't exist yet. */
function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

/** Utility: read the Settings tab's System Config block into a {label: value} map. */
function readSettingsMap_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SETTINGS);
  if (!sh) throw new Error('Settings tab not found — run "Build Settings Tab" first.');
  const data = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  const map = {};
  data.forEach(row => {
    const label = String(row[0]).trim();
    if (label && !CONTACT_ROLES.includes(label)) map[label] = row[1];
  });
  return map;
}

/**
 * Utility: read the Contacts block into {role: {name, email, active}}.
 * Use this — never read Settings directly — whenever an email function
 * needs a contact, so toggling "Active" to No always takes effect.
 */
function readContacts_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SETTINGS);
  if (!sh) throw new Error('Settings tab not found — run "Build Settings Tab" first.');
  const data = sh.getRange(1, 1, sh.getLastRow(), 4).getValues();
  const map = {};
  data.forEach(row => {
    const role = String(row[0]).trim();
    if (CONTACT_ROLES.includes(role)) {
      map[role] = { name: row[1], email: row[2], active: String(row[3]).trim().toLowerCase() === 'yes' };
    }
  });
  return map;
}

/**
 * Returns the contact's email only if they're marked Active, else null.
 * Every email-sending function should check this before adding someone
 * to a recipient list (e.g. CC'ing a principal or the ADCI).
 */
function activeEmailFor_(role) {
  const contact = readContacts_()[role];
  return (contact && contact.active && contact.email) ? contact.email : null;
}

/**
 * Comma-separated CC list of every Contact role currently marked Active
 * in Settings (blank/No-active roles are skipped), for bulk emails that
 * should keep leadership in the loop — e.g. sendTeacherWorkbookLinksAndGuide().
 * Was referenced but never defined anywhere in this project; added 2026-09-22.
 */
function activeContactsCcString_() {
  const contacts = readContacts_();
  const emails = Object.keys(contacts)
    .map(role => contacts[role])
    .filter(c => c && c.active && c.email)
    .map(c => c.email);
  return [...new Set(emails)].join(',');
}

// Week 1 / Week 5 / Week 10 tab column layouts (1-based). Centralized so
// every function that reads/writes/protects these tabs agrees on positions.
const WEEK1_COLS = {
  ID: 1, NAME: 2, CLASS: 3, WEAK1: 4, STEPS1: 5, WEAK2: 6, STEPS2: 7,
  MEANS: 8, EVIDENCE: 9, REVIEW_DATE: 10, ACTION_PLAN_LINK: 11, GENERATE: 12, STATUS: 13
};

const WEEK5_COLS = {
  ID: 1, NAME: 2, CLASS: 3, STARTING_SCORE: 4, CUMULATIVE_1: 5,
  W1_ET: 6, W1_PRESENT: 7, W2_ET: 8, W2_PRESENT: 9, W3_ET: 10, W3_PRESENT: 11,
  W4_ET: 12, W4_PRESENT: 13, W5_ET: 14, W5_PRESENT: 15,
  QZ1: 16, QZ2: 17,
  WRITTEN_SUMMARY: 18, WEAK1_AP2: 19, STEPS1_AP2: 20, WEAK2_AP2: 21, STEPS2_AP2: 22,
  MEANS_AP2: 23, EVIDENCE_AP2: 24,
  PROGRESS_REPORT_STATUS: 25, ACTION_PLAN_LINK: 26, GENERATE: 27, STATUS: 28
};

const WEEK10_COLS = {
  ID: 1, NAME: 2, CLASS: 3, CUMULATIVE_2: 4,
  W6_ET: 5, W6_PRESENT: 6, W7_ET: 7, W7_PRESENT: 8, W8_ET: 9, W8_PRESENT: 10,
  W9_ET: 11, W9_PRESENT: 12, W10_ET: 13, W10_PRESENT: 14,
  QZ3: 15, QZ4: 16,
  WRITTEN_SUMMARY: 17, ACTION_PLAN_PDF_LINK: 18,
  GENERATE: 19, STATUS: 20
};

// Biweekly quiz schedule — which two weeks per phase actually have a
// quiz, split by whether the subject has an Exit Ticket at all.
const QUIZ_WEEKS = {
  noET: [2, 4, 6, 8],   // Arabic Social Studies, Islamic Religion
  hasET: [3, 5, 7, 9]   // English, Math, Science, Arabic, Social Studies, etc.
};
function quizWeeksForSubject_(subject) {
  return NO_ET_SUBJECTS.indexOf(subject) !== -1 ? QUIZ_WEEKS.noET : QUIZ_WEEKS.hasET;
}