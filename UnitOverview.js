/* UnitOverview.js Last Update 2025-08-24 22:25:58 <f4e98bd879d04bf55291b096d17bd816b79eb8a2cffb308acf6c729df07e0c9e>
/* eslint-disable no-unused-vars */
/* global SpreadsheetApp, STYLE */

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
  // Label Col1 as 'Name' so the header shows a readable title instead of Col1
  const q = `=QUERY({Grades!A2:A,Grades!C2:C,Grades!F2:F},"select Col1, avg(Col3) where Col1 is not null group by Col1 pivot Col2 label Col1 'Name'",0)`;
  sh.getRange('A1').setFormula(q);
  // Wait for the formula to take effect, then apply basic formatting
  SpreadsheetApp.flush();
  try {
    const dataRange = sh.getDataRange();
    const lastCol = Math.max(1, dataRange.getLastColumn());
    const lastRow = Math.max(1, dataRange.getLastRow());
    // Header styling: apply theme header bg/text if available
    const headerBg = (typeof STYLE !== 'undefined' && STYLE && STYLE.COLORS && STYLE.COLORS.UI && STYLE.COLORS.UI.HEADER_BG) || '#f0f3f5';
    const headerText = (typeof STYLE !== 'undefined' && STYLE && STYLE.COLORS && STYLE.COLORS.UI && STYLE.COLORS.UI.HEADER_TEXT) || '#000000';
    sh.getRange(1, 1, 1, lastCol).setFontWeight('bold').setWrap(true).setBackground(headerBg).setFontColor(headerText);
    sh.setFrozenRows(1);
    // Freeze first column so teachers can scroll horizontally
    sh.setFrozenColumns(1);
    sh.setColumnWidth(1, 180);
    if (lastCol >= 2) sh.setColumnWidth(2, 240);
    for (let c = 3; c <= lastCol; c++) sh.setColumnWidth(c, 100);
    // Apply numeric formatting to the pivot numeric area (rows 2..lastRow, cols 2..lastCol)
    if (lastRow > 1 && lastCol > 1) {
      sh.getRange(2, 2, lastRow - 1, lastCol - 1).setNumberFormat('0.00');
    }
    // Apply subtle row striping (alternate rows) using theme colors if STYLE is available
    try {
      const ui = (typeof STYLE !== 'undefined' && STYLE && STYLE.COLORS && STYLE.COLORS.UI) || {};
      const neutralBg = ui.NEUTRAL_BG || '#ffffff';
      const neutralAlt = ui.NEUTRAL_BG_ALT || '#f7f7f7';
      // Apply base background then overwrite even rows with alt
      sh.getRange(2, 1, Math.max(1, lastRow - 1), lastCol).setBackground(neutralBg);
      // Build conditional format rule for even rows
      const dataRangeA1 = sh.getRange(2, 1, Math.max(1, lastRow - 1), lastCol).getA1Notation();
      const rules = sh.getConditionalFormatRules().filter(r => !r.getRanges().some(rg => rg.getA1Notation() === dataRangeA1));
      const stripe = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=ISEVEN(ROW())')
        .setBackground(neutralAlt)
        .setRanges([sh.getRange(2, 1, Math.max(1, lastRow - 1), lastCol)])
        .build();
      rules.push(stripe);
      sh.setConditionalFormatRules(rules);
    } catch (e) {
      // ignore stripe errors
    }
  } catch (e) {
    // best-effort formatting; ignore if QUERY hasn't expanded yet
  }
}

