/* eslint-disable no-unused-vars */
/* global SpreadsheetApp */

const SHEET_UNIT_OVERVIEW = 'Unit Overview';

/**
 * Create or refresh a 'Unit Overview' sheet using a single QUERY pivot over the
 * Grades table so the sheet does not depend on separate student named ranges.
 *
 * The QUERY aggregates average mastery (Grades column F) by student name/email and pivots
 * on the Unit (Grades column C) to produce one column per unit.
 */
function createUnitOverview() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_UNIT_OVERVIEW);
  if (!sh) sh = ss.insertSheet(SHEET_UNIT_OVERVIEW);
  else sh.clear();

  // Build a QUERY that uses a 3-column inline range: Name, Unit, Mastery
  // and pivots unit into columns with AVG of mastery values. We filter out rows without name.
  // Result: first column is student name, following columns are per-unit averages.
  const q = `=QUERY({Grades!A2:A,Grades!C2:C,Grades!F2:F},"select Col1, avg(Col3) where Col1 is not null group by Col1 pivot Col2",0)`;
  sh.getRange('A1').setFormula(q);
  // Wait for the formula to take effect, then apply basic formatting
  SpreadsheetApp.flush();
  try {
    const dataRange = sh.getDataRange();
    const lastCol = Math.max(1, dataRange.getLastColumn());
    const lastRow = Math.max(1, dataRange.getLastRow());
    sh.getRange(1, 1, 1, lastCol).setFontWeight('bold').setWrap(true);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 180);
    if (lastCol >= 2) sh.setColumnWidth(2, 240);
    for (let c = 3; c <= lastCol; c++) sh.setColumnWidth(c, 100);
    // Apply numeric formatting to the pivot numeric area (rows 2..lastRow, cols 2..lastCol)
    if (lastRow > 1 && lastCol > 1) {
      sh.getRange(2, 2, lastRow - 1, lastCol - 1).setNumberFormat('0.00');
    }
  } catch (e) {
    // best-effort formatting; ignore if QUERY hasn't expanded yet
  }
}

