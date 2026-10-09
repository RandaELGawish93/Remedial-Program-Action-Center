/**
 * SETUP_SETTINGS.gs
 * Builds the Settings tab. Re-running this will NOT wipe values you've
 * already typed in column B — it only rebuilds labels/formatting/dropdowns.
 */

function buildSettingsTab() {
  const sh = getOrCreateSheet_(SHEETS.SETTINGS);
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();

  // Preserve any existing values before rebuilding labels/formatting
  const existingContacts = {}; // role -> {name, email, active}
  const existingConfig = {};   // key -> value
  const lastRow = sh.getLastRow();
  if (lastRow > 0) {
    const data = sh.getRange(1, 1, lastRow, 4).getValues();
    data.forEach(row => {
      const label = String(row[0]).trim();
      if (!label) return;
      if (CONTACT_ROLES.includes(label)) {
        existingContacts[label] = { name: row[1], email: row[2], active: row[3] };
      } else {
        existingConfig[label] = row[1];
      }
    });
  }

  sh.clear();
  sh.setColumnWidth(1, 260);
  sh.setColumnWidth(2, 170);
  sh.setColumnWidth(3, 220);
  sh.setColumnWidth(4, 80);

  let r = 1;

  // --- Section A: Contacts ---
  sh.getRange(r, 1, 1, 4).merge().setValue('CONTACTS')
    .setFontWeight('bold').setBackground('#1c4587').setFontColor('#ffffff');
  r++;
  sh.getRange(r, 1).setValue('Role').setFontWeight('bold');
  sh.getRange(r, 2).setValue('Name').setFontWeight('bold');
  sh.getRange(r, 3).setValue('Email').setFontWeight('bold');
  sh.getRange(r, 4).setValue('Active').setFontWeight('bold');
  const contactHeaderRow = r;
  r++;

  const firstContactRow = r;
  CONTACT_ROLES.forEach(role => {
    const prior = existingContacts[role] || {};
    sh.getRange(r, 1).setValue(role);
    sh.getRange(r, 2).setValue(prior.name || '');
    sh.getRange(r, 3).setValue(prior.email || '');
    sh.getRange(r, 4).setValue(prior.active !== undefined && prior.active !== '' ? prior.active : 'Yes');
    r++;
  });
  const lastContactRow = r - 1;

  // Active dropdown (Yes/No) for every contact row
  sh.getRange(firstContactRow, 4, lastContactRow - firstContactRow + 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No']).build()
  );
  // Grey out inactive rows visually so it's obvious at a glance
  const activeRange = sh.getRange(firstContactRow, 1, lastContactRow - firstContactRow + 1, 4);
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$D${firstContactRow}="No"`)
    .setBackground('#f4f4f4').setFontColor('#999999')
    .setRanges([activeRange]).build();
  const rules = sh.getConditionalFormatRules();
  rules.push(rule);
  sh.setConditionalFormatRules(rules);

  r++; // spacer

  // --- Section B: System Config ---
  sh.getRange(r, 1, 1, 4).merge().setValue('SYSTEM CONFIGURATION')
    .setFontWeight('bold').setBackground('#1c4587').setFontColor('#ffffff');
  r++;

  SYSTEM_CONFIG_KEYS.forEach(key => {
    if (key === '') { r++; return; } // spacer row
    sh.getRange(r, 1).setValue(key);
    sh.getRange(r, 2).setValue(existingConfig[key] || '');
    r++;
  });

  // Data validation: Term # dropdown
  const termRow = findRowByLabel_(sh, 'Term # (1, 2, or 3)');
  if (termRow) {
    sh.getRange(termRow, 2).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['1', '2', '3']).build()
    );
  }

  // Term Start Date — auto-filled from TERM_START_DATES based on Term #,
  // format as date so TODAY()-based week formulas work
  const startDateRow = findRowByLabel_(sh, 'Term Start Date (first Sunday of the term)');
  if (startDateRow) {
    sh.getRange(startDateRow, 2).setNumberFormat('yyyy-mm-dd');
    const currentTerm = String(sh.getRange(termRow, 2).getValue() || '1');
    if (TERM_START_DATES[currentTerm]) {
      sh.getRange(startDateRow, 2).setValue(TERM_START_DATES[currentTerm]);
    }
  }

  // Data validation: link sharing dropdown
  const sharingRow = findRowByLabel_(sh, 'Generated Document Link Sharing');
  if (sharingRow) {
    sh.getRange(sharingRow, 2).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(LINK_SHARING_OPTIONS).build()
    );
    if (!sh.getRange(sharingRow, 2).getValue()) {
      sh.getRange(sharingRow, 2).setValue(LINK_SHARING_OPTIONS[0]); // default: organization-only
    }
  }

  sh.setFrozenRows(2);
  SpreadsheetApp.getActiveSpreadsheet().toast('Settings tab built. Fill in column B.', 'Done', 5);
}

/** Finds the row number of a label in column A of a given sheet. */
function findRowByLabel_(sh, label) {
  const values = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === label) return i + 1;
  }
  return null;
}