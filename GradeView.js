/* GradeView.js Last Update 2025-09-04 16:27 <1292da3664dbb5515062830b2578cabf9fa0b0bea3700d1c6d853716ce91938d>
/* eslint-disable no-unused-vars */
/* global SpreadsheetApp, STYLE,
RANGE_SKILL_UNITS,
          RANGE_STUDENT_NAMES, RANGE_STUDENT_EMAILS,
          RANGE_LEVEL_NAMES, RANGE_LEVEL_STREAK, RANGE_LEVEL_SCORES,
          RANGE_NONE_CORRECT_SCORE, RANGE_SOME_CORRECT_SCORE */

const SHEET_GRADE_VIEW = 'Grade View';
// Toggle to enable/disable applying header highlight to the first row of each Unit
const HIGHLIGHT_FIRST_LINE_OF_UNIT = false;

/**
 * Build or rebuild a read-friendly Grade View sheet.
 * - Student dropdown (by name)
 * - Explanations for each level and fallback rules
 * - Sorted table of Unit, Skill #, Description, Mastery Grade, and per-level symbol strings
 */
function setupGradeViewSheet(studentName) {
  const ss = SpreadsheetApp.getActive();
  const sheetName = studentName ? safeSheetName_(studentName) : SHEET_GRADE_VIEW;
  let sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  sh.clear();
  // Global fonts
  try {
    sh.getRange(1, 1, Math.max(1, sh.getMaxRows()), Math.max(1, sh.getMaxColumns()))
      .setFontFamily(STYLE.FONT_FAMILY)
      .setFontSize(Number(STYLE.FONT_SIZE));
  } catch (e) { /* STYLE optional */ }

  // Read settings arrays
  const levelNames = ss.getRangeByName(RANGE_LEVEL_NAMES).getValues().flat().filter(String);
  const levelStreak = ss.getRangeByName(RANGE_LEVEL_STREAK).getValues().flat().slice(0, levelNames.length);
  const levelScores = ss.getRangeByName(RANGE_LEVEL_SCORES).getValues().flat().slice(0, levelNames.length);

  // Calculate dynamic column positions based on levels
  const tableColCount = 4 + levelNames.length; // Base: Unit(1), Skill#(2), Desc(3), Grade(4) + attempts(5+)
  const summaryStartCol = tableColCount + 2; // Start summary after table with 1 spacer column

  // Layout constants matching Grades sheet
  const baseHeadersCount = 6; // Name, Email, Unit, Skill #, Skill Description, Mastery Grade
  const firstUtilCol = baseHeadersCount + 1; // util starts after base
  const masteryCol = baseHeadersCount; // 6
  const symbolCols = levelNames.map((_, i) => firstUtilCol + i * 3 + 2); // per-level symbols column in Grades

  // Title
  sh.getRange('A1:G1').merge();
  const brandPrimary = (STYLE && STYLE.COLORS && STYLE.COLORS.BRAND_PRIMARY) || '#0033a0';
  const headerBg = (STYLE && STYLE.COLORS && STYLE.COLORS.UI && STYLE.COLORS.UI.HEADER_BG) || '#f0f3f5';
  sh.getRange('A1').setValue('Standards Based Grades').setFontSize(Number(STYLE.FONT_SIZE_XLARGE || 18)).setFontWeight('bold').setHorizontalAlignment('left').setFontColor(brandPrimary);

  // Student selector (merge A2:B2 for label, C2:G2 for dropdown or fixed name)
  sh.getRange('A2:B2').merge();
  sh.getRange('A2').setValue('Student').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange('C2:G2').merge();
  const selCell = sh.getRange('C2');
  const namesRange = ss.getRangeByName(RANGE_STUDENT_NAMES);
  const inputBg = (STYLE && STYLE.COLORS && STYLE.COLORS.UI && STYLE.COLORS.UI.INPUT_HIGHLIGHT) || '#fffbe6';
  if (studentName) {
    selCell.setValue(studentName).setFontSize(12).setHorizontalAlignment('left').setBackground(inputBg);
  } else {
    if (namesRange) {
      const rule = SpreadsheetApp.newDataValidation().requireValueInRange(namesRange, true).setAllowInvalid(false).build();
      selCell.setDataValidation(rule);
    }
    selCell.setFontSize(12).setHorizontalAlignment('left').setBackground(inputBg);
  }
  // The dropdown spans C2:G2; widths set later below

  // Hidden helper: selected email from name (place outside visible table area)
  // =IF(C2="","",INDEX(StudentEmails, MATCH(C2, StudentNames, 0)))
  sh.getRange('Z2').setFormula(`=IF($C$2="","",INDEX(${RANGE_STUDENT_EMAILS}, MATCH($C$2, ${RANGE_STUDENT_NAMES}, 0)))`);
  sh.hideColumns(26);

  // Spacer before table
  let row = 4;

  // Table headers
  const tableHeaderRow = row;
  const tableHeaders = ['Unit', 'Skill #', 'Skill Description', 'Grade', ...levelNames.map(n => `${n} Attempts`)];
  sh.getRange(tableHeaderRow, 1, 1, tableHeaders.length).setValues([tableHeaders]).setFontWeight('bold').setBackground(headerBg);

  // Subheader row: put descriptions where they go
  const subHeaderRow = tableHeaderRow + 1;
  // Mastery Grade column: conditional fallback description only when relevant to selected student
  const masteryDescFormula = `=IF($Z$2="","",
    TRIM(
      IF(COUNTIF(FILTER(INDEX(Grades!A2:ZZ,,${masteryCol}), Grades!B2:B=$Z$2), ${RANGE_NONE_CORRECT_SCORE})>0,
         "If no correct attempts, score is "&${RANGE_NONE_CORRECT_SCORE}&". ", "") &
      IF(COUNTIF(FILTER(INDEX(Grades!A2:ZZ,,${masteryCol}), Grades!B2:B=$Z$2), ${RANGE_SOME_CORRECT_SCORE})>0,
         "If at least one correct, score is "&${RANGE_SOME_CORRECT_SCORE}&".", "")
    )
  )`;
  sh.getRange(subHeaderRow, 4).setFormula(masteryDescFormula).setWrap(true).setFontStyle('italic');

  // Per-level attempt column descriptors
  levelNames.forEach((_, i) => {
    const col = 5 + i;
    const descFormula = `="To earn a score of "&INDEX(${RANGE_LEVEL_SCORES},${i + 1})&", you must show proficiency "&INDEX(${RANGE_LEVEL_STREAK},${i + 1})&" times in a row."`;
    sh.getRange(subHeaderRow, col).setFormula(descFormula).setWrap(true).setFontStyle('italic');
  });

  // Table data formula (array): sorted by Unit (3) then Skill # (4), columns pulled via CHOOSECOLS
  // We filter by selected email in Z2 to disambiguate duplicate names
  const chooseCols = [3, 4, 5, masteryCol, ...symbolCols];
  const formula = `=IF($Z$2="","",
    CHOOSECOLS(
      SORT(FILTER(Grades!A2:ZZ, Grades!B2:B=$Z$2, Grades!F2:F<>""), 3, TRUE, 4, TRUE),
      ${chooseCols.join(',')}
    )
  )`;
  sh.getRange(subHeaderRow + 1, 1).setFormula(formula);

  // Helper column for unit-index mapping (hidden):
  // We'll place them immediately after the last column of the generated table so they stay adjacent
  // - helperCol1: Derive the unit number for this unit

  let helperCol1A = null;
  try {
    const helperStartRow = subHeaderRow + 1;
    // compute where the table ends so helper columns are placed just after it
    const lastTableCol = tableHeaders.length; // number of columns in the presented table
    // ensure helper columns don't overlap the Unit Summary. Start at least at summaryStartCol + 3.
    const minHelperStart = summaryStartCol + 3;
    const helperCol1 = Math.max(lastTableCol + 1, minHelperStart); // first helper column (right after table or after summary)

    // ensure sheet has enough columns
    if (sh.getMaxColumns() < helperCol1) {
      sh.insertColumnsAfter(sh.getMaxColumns(), helperCol1 - sh.getMaxColumns());
    }

    // small helper to convert column number to A1 letter(s)
    const colToA1 = function (c) {
      let s = '';
      while (c > 0) {
        const m = (c - 1) % 26;
        s = String.fromCharCode(65 + m) + s;
        c = Math.floor((c - 1) / 26);
      }
      return s;
    };
    // UNIQUE/MATCH spill at helperCol1 starting at helperStartRow
    sh.getRange(helperStartRow, helperCol1).setFormula(
      `=ARRAYFORMULA(MATCH($A${helperStartRow}:$A,UNIQUE(${RANGE_SKILL_UNITS}),0))`
    );
    // store the A1 letter for CF formulas
    helperCol1A = colToA1(helperCol1);
    sh.hideColumns(helperCol1, 1);
  } catch (e) {
    if (console && console.warn) console.warn('Grade View helper columns warn', e);
  }

  // Pretty formatting
  sh.setFrozenRows(tableHeaderRow); // freeze header
  // Column widths
  sh.setColumnWidth(1, 90);  // Unit
  sh.setColumnWidth(2, 70);  // Skill #
  sh.setColumnWidth(3, 360); // Description
  sh.setColumnWidth(4, 110); // Mastery Grade
  for (let i = 0; i < levelNames.length; i++) {
    sh.setColumnWidth(5 + i, 180);
  }
  // We'll apply custom color striping and gradient rules instead of generic banding.

  // Alignments for readability
  const dataRowsGV = Math.max(1, sh.getMaxRows() - subHeaderRow);
  sh.getRange(subHeaderRow + 1, 2, dataRowsGV, 1).setHorizontalAlignment('center'); // Skill #
  sh.getRange(subHeaderRow + 1, 4, dataRowsGV, 1 + levelNames.length).setHorizontalAlignment('center'); // Mastery + attempts
  sh.getRange(subHeaderRow + 1, 3, dataRowsGV, 1).setWrap(true); // Description

  // Emphasize Mastery Grade: larger font for header and data
  sh.getRange(tableHeaderRow, 4).setFontSize(14).setFontWeight('bold');
  sh.getRange(subHeaderRow + 1, 4, dataRowsGV, 1).setFontSize(14).setFontWeight('bold');

  // Set requested column widths for A..G
  sh.setColumnWidth(1, 41);
  sh.setColumnWidth(2, 47);
  sh.setColumnWidth(3, 185);
  sh.setColumnWidth(4, 88);
  sh.setColumnWidth(5, 180);
  sh.setColumnWidth(6, 180);
  sh.setColumnWidth(7, 180);

  // Optional: Unit summary to the right (dynamic position based on levels)
  sh.getRange(3, summaryStartCol).setValue('Unit Summary').setFontWeight('bold');
  sh.getRange(4, summaryStartCol, 1, 3).setValues([['Unit', 'Average Grade', 'Skills']]).setFontWeight('bold').setBackground(headerBg);
  // Unit summary: average only numeric grades (ignore non-numeric like "-")
  sh.getRange(5, summaryStartCol).setFormula(`=IF($Z$2="","",
    QUERY(A5:D,      
      "select A, avg(D), count(D) group by A order by A",
      0
    )
  )`);
  sh.setColumnWidth(summaryStartCol, 120);     // Dynamic: Unit
  sh.setColumnWidth(summaryStartCol + 1, 120); // Dynamic: Average Grade
  sh.setColumnWidth(summaryStartCol + 2, 90);  // Dynamic: Skills
  // Format averages in the dynamic Average Grade column to two decimals
  try {
    const avgColRange = sh.getRange(5, summaryStartCol + 1, Math.max(1, sh.getMaxRows() - 4), 1);
    avgColRange.setNumberFormat('0.00');
  } catch (e) { /* formatting best-effort */ }

  // Color application: neutral stripes, per-level attempts stripes, mastery gradient
  try {
    const neutralBg = (STYLE && STYLE.COLORS && STYLE.COLORS.UI && STYLE.COLORS.UI.NEUTRAL_BG) || '#f7f7f7';
    const neutralBgAlt = (STYLE && STYLE.COLORS && STYLE.COLORS.UI && STYLE.COLORS.UI.NEUTRAL_BG_ALT) || '#f0f0f0';
    const neutralText = (STYLE && STYLE.COLORS && STYLE.COLORS.UI && STYLE.COLORS.UI.NEUTRAL_TEXT) || '#333333';
    const dataStart = subHeaderRow + 1;
    const dataCount = Math.max(1, sh.getMaxRows() - subHeaderRow);
    // Base neutral background for Unit, Skill #, Description (A..C)
    sh.getRange(dataStart, 1, dataCount, 3).setBackground(neutralBg).setFontColor(neutralText);

    // Build conditional formatting rules, filtering out duplicates for target ranges
    let rules = sh.getConditionalFormatRules();
    const targetA1s = [];

    // Neutral stripe on even data rows for A..C
    const neutralRange = sh.getRange(dataStart, 1, dataCount, 3);
    targetA1s.push(neutralRange.getA1Notation());
    const neutralStripe = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=ISEVEN(ROW())')
      .setBackground(neutralBgAlt)
      .setRanges([neutralRange])
      .build();

    // Build ordered rules container (we don't use header-row highlighting here)
    const orderedRules = [];

    // Per-level attempts columns start at column 5
    const attemptsStartCol = 5;
    const attemptRules = [];
    for (let i = 0; i < levelNames.length; i++) {
      const col = attemptsStartCol + i;
      const levelIdx = i + 1; // levels are 1-based in STYLE
      const levelDef = (STYLE && STYLE.COLORS && STYLE.COLORS.LEVELS && STYLE.COLORS.LEVELS[levelIdx]) || {};
      const baseBg = levelDef.BG || neutralBg;
      const baseText = levelDef.TEXT || '#000000';
      const altBg = levelDef.BG_ALT || baseBg;
      // Base background for data area
      sh.getRange(dataStart, col, dataCount, 1).setBackground(baseBg).setFontColor(baseText);
      // Stripe rule on even rows
      const r = sh.getRange(dataStart, col, dataCount, 1);
      targetA1s.push(r.getA1Notation());
      // (no per-level header rules here)
      const stripe = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=ISEVEN(ROW())')
        .setBackground(altBg)
        .setRanges([r])
        .build();
      attemptRules.push(stripe);
      // then push stripe so header takes precedence
      orderedRules.push(stripe);
    }

    // Unit-based text color rules (cycle via helper column AB (28)).
    try {
      const unitColorRanges = [];
      // A..C (Unit, Skill #, Description)
      const unitCoreRange = sh.getRange(dataStart, 1, dataCount, 3);
      unitColorRanges.push(unitCoreRange);
      targetA1s.push(unitCoreRange.getA1Notation());
      // Add all attempt columns to the unit-color target ranges
      for (let i = 0; i < levelNames.length; i++) {
        const r = sh.getRange(dataStart, attemptsStartCol + i, dataCount, 1);
        unitColorRanges.push(r);
        targetA1s.push(r.getA1Notation());
      }
      // Formula anchored to top data row. Use MOD to cycle colors for many units.
      // Build N conditional-format rules from STYLE.COLORS.UI.UNIT_TEXT_COLORS and cycle via helper index.
      const unitColors = (STYLE && STYLE.COLORS && STYLE.COLORS.UI && STYLE.COLORS.UI.UNIT_TEXT_COLORS) || ['#3a3a3a', '#0d47a1'];
      const nColors = Math.max(1, unitColors.length);
      // Use dynamically computed helper column letter if available; otherwise fall back to AB.
      const helperIdxCol = helperCol1A ? `$${helperCol1A}` : '$AB';
      const unitRules = [];
      // Separate ranges: restrict unit-based formatting to core A:C only so attempt columns keep their own styling
      const unitTextRanges = [unitCoreRange];
      const unitCoreRangeOnly = unitCoreRange;
      for (let k = 0; k < nColors; k++) {
        // odd (non-even) rows rule: set font color across all target ranges
        const fOdd = `=AND(${helperIdxCol}${dataStart}<>"", MOD(${helperIdxCol}${dataStart}-1, ${nColors})=${k}, NOT(ISEVEN(ROW())))`;
        const ruleFontOdd = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(fOdd)
          .setFontColor(unitColors[k])
          .setRanges(unitTextRanges)
          .build();
        // even rows rule: set font color across all target ranges
        const fEven = `=AND(${helperIdxCol}${dataStart}<>"", MOD(${helperIdxCol}${dataStart}-1, ${nColors})=${k}, ISEVEN(ROW()))`;
        const ruleFontEven = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(fEven)
          .setFontColor(unitColors[k])
          .setRanges(unitTextRanges)
          .build();
        // even rows bg rule applied only to core unit columns (A:C) so per-level attempt backgrounds stay intact
        /* const ruleBgEven = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(fEven)
          .setBackground(neutralBgAlt)
          .setRanges([unitCoreRangeOnly])
          .build(); */
        // push bg rule first so it wins for background on core columns, then font rules
        unitRules.push(
          /* ruleBgEven, */
          ruleFontEven, ruleFontOdd);
      }
      // Insert unit rules early so they take precedence over generic stripes (preserve order)
      orderedRules.unshift(...unitRules);
    } catch (e) { if (console && console.warn) console.warn('Unit color rules warn', e); }

    // Mastery Grade gradient on column 4
    const gradeCol = 4;
    const gradeRange = sh.getRange(dataStart, gradeCol, dataCount, 1);
    targetA1s.push(gradeRange.getA1Notation());
    const minColor = STYLE.COLORS.GRADE_SCALE.MIN;
    const midColor = STYLE.COLORS.GRADE_SCALE.MID;
    const maxColor = STYLE.COLORS.GRADE_SCALE.MAX;
    const textOnScale = STYLE.COLORS.GRADE_SCALE.TEXT;
    let maxScore = 1;
    try {
      const lvl = ss.getRangeByName(RANGE_LEVEL_SCORES);
      if (lvl) {
        const nums = lvl.getValues().flat().map(v => Number(v)).filter(v => !isNaN(v));
        if (nums.length) maxScore = Math.max.apply(null, nums);
      }
    } catch (e) { /* default */ }
    let minScore = 0;
    try {
      const none = ss.getRangeByName(RANGE_NONE_CORRECT_SCORE);
      if (none) {
        const n = Number(none.getValue());
        if (!isNaN(n)) minScore = n;
      }
    } catch (e) { /* default */ }
    let midScore = null;
    try {
      const some = ss.getRangeByName(RANGE_SOME_CORRECT_SCORE);
      if (some) {
        const m = Number(some.getValue());
        if (!isNaN(m)) midScore = m;
      }
    } catch (e) { /* optional */ }
    let gradientBuilder = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpointWithValue(minColor, SpreadsheetApp.InterpolationType.NUMBER, String(minScore));
    if (midScore !== null && midColor) {
      gradientBuilder = gradientBuilder.setGradientMidpointWithValue(midColor, SpreadsheetApp.InterpolationType.NUMBER, String(midScore));
    }
    const gradeGradient = gradientBuilder
      .setGradientMaxpointWithValue(maxColor, SpreadsheetApp.InterpolationType.NUMBER, String(maxScore))
      .setRanges([gradeRange])
      .build();
    const gradeText = SpreadsheetApp.newConditionalFormatRule()
      .whenCellNotEmpty()
      .setFontColor(textOnScale)
      .setRanges([gradeRange])
      .build();

    // De-duplicate rules for our target ranges
    rules = rules.filter(r => !r.getRanges().some(rg => targetA1s.includes(rg.getA1Notation())));
    // Push header rules (ordered) before stripe rules so they take priority
    rules.push(...orderedRules);
    // Neutral stripe for A..C should come after header rule for A..C
    rules.push(neutralStripe);
    // Any remaining attempt stripe rules that weren't already pushed (attemptRules) can be pushed too
    // (orderedRules already contains per-column stripes), but include any extras for safety
    rules.push(...attemptRules.filter(ar => !orderedRules.includes(ar)));
    rules.push(gradeGradient, gradeText);
    sh.setConditionalFormatRules(rules);
  } catch (e) {
    if (console && console.warn) console.warn('Grade View color rules warn', e);
  }
}

// Make a safe sheet name from an arbitrary string (strip forbidden chars, trim length)
function safeSheetName_(name) {
  const cleaned = String(name).replace(/[\\/*?:[\]]/g, ' ').trim();
  return cleaned.substring(0, 99) || 'Student';
}
