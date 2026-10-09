/**
 * CODE.gs
 * Custom menus, split by tab/workflow so it's clear at a glance which
 * part of the system each function belongs to. Run each "Build" item
 * once to create/reset that tab. Safe to re-run; each builder clears
 * and rebuilds only its own tab.
 */

function onOpen() {
  // All menus are defined in AUDIT_HOD.gs (buildAllMenus_) so they live in one place.
  buildAllMenus_(SpreadsheetApp.getUi());
}

/**
 * Simple trigger — fires on any manual edit in the workbook. Only acts
 * when the edit is the Term # cell in Settings, so changing terms mid-year
 * (or at the start of a new one) updates Term Start Date automatically
 * without needing to re-run "Build Settings Tab".
 */
function onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== SHEETS.SETTINGS) return;

  const termRow = findRowByLabel_(sh, 'Term # (1, 2, or 3)');
  const startDateRow = findRowByLabel_(sh, 'Term Start Date (first Sunday of the term)');
  if (!termRow || !startDateRow) return;
  if (e.range.getRow() !== termRow || e.range.getColumn() !== 2) return;

  const term = String(e.value || '').trim();
  if (TERM_START_DATES[term]) {
    sh.getRange(startDateRow, 2).setValue(TERM_START_DATES[term]);
  }
}
