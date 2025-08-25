/* eslint-disable no-unused-vars */
/* exported setupGradesSheet, populateGrades, reformatGradesSheet */
/* global SpreadsheetApp, STYLE,
          RANGE_SYMBOL_CHARS, RANGE_SYMBOL_MASTERY, RANGE_SYMBOL_SYMBOL,
          RANGE_LEVEL_STREAK, RANGE_LEVEL_SCORES, RANGE_NONE_CORRECT_SCORE, RANGE_SOME_CORRECT_SCORE,
          RANGE_LEVEL_SHORTCODES, RANGE_LEVEL_NAMES, RANGE_LEVEL_DEFAULTATTEMPTS,
          RANGE_STUDENT_NAMES, RANGE_STUDENT_EMAILS,
          RANGE_SKILL_UNITS, RANGE_SKILL_NUMBERS, RANGE_SKILL_DESCRIPTORS */
const SHEET_GRADES = 'Grades';

/**
 * Public entry: orchestrates sheet creation by delegating to two parts
 * (1) headers and (2) formulas. Behavior preserved, just clearer.
 */
function setupGradesSheet() {
  const ss = SpreadsheetApp.getActive();
  const settings = readLevelSettings_(ss);
  const sh = ensureGradesSheet_(ss);

  const ctx = setupGradesHeaders_(sh, settings); // (1) headers
  setupGradesFormulas_(sh, settings, ctx);       // (2) formulas

  // Basic sheet niceties (frozen header + autoresize + fonts)
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, ctx.headers.length);
  try {
    sh.getRange(1, 1, Math.max(1, sh.getMaxRows()), Math.max(1, sh.getMaxColumns()))
      .setFontFamily(STYLE.FONT_FAMILY)
      .setFontSize(Number(STYLE.FONT_SIZE));
  } catch (e) { /* STYLE may not be defined in some contexts; ignore */ }

  // Light formatting and validation
  applyGradesFormatting_(sh, settings, ctx);
}

/**
 * Return the first row index (>=2) where columns A..E are all blank. Returns null if none found.
 */
function findFirstEmptyDataRow_(sh) {
  const startRow = 2;
  const lastRow = Math.max(sh.getLastRow(), startRow);
  if (lastRow < startRow) return startRow;
  const numRows = lastRow - startRow + 1;
  // Read columns A..E for existing rows
  const vals = sh.getRange(startRow, 1, numRows, 5).getValues();
  for (let i = 0; i < vals.length; i++) {
    const row = vals[i];
    const allBlank = row.every(c => c === null || c === undefined || String(c).trim() === '');
    if (allBlank) return startRow + i;
  }
  return null;
}

/* -------------------- (1) HEADERS -------------------- */
function setupGradesHeaders_(sh, settings) {
  // Base columns that are typed-in by the teacher
  // Split Skill into Unit | Skill # | Skill Description
  const baseHeaders = ['Name', 'Email', 'Unit', 'Skill #', 'Skill Description', 'Mastery Grade'];

  // Utility columns per level: Streak + String + Mastery (display-only)
  const utilHeaders = settings.codes.flatMap((_, i) => [
    `${settings.names[i]} Streak`,
    `${settings.names[i]} String`,
    `${settings.names[i]}`,
  ]);

  // Attempt columns per level: make names unique per level (e.g., B1..Bn)
  const attemptHeaders = settings.codes.flatMap((code, i) =>
    Array.from({ length: Number(settings.defaultAttempts[i] || 0) }, (_, k) => `${code}${k + 1}`)
  );

  const headers = [...baseHeaders, ...utilHeaders, ...attemptHeaders];
  sh.clear();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Compute column indices we need later
  const firstUtilCol = baseHeaders.length + 1;
  const firstAttemptCol = baseHeaders.length + utilHeaders.length + 1;
  const lastCol = headers.length;
  const masteryCol = baseHeaders.length; // 'Mastery Grade' is last in base headers

  return {
    headers,
    baseHeaders,
    firstUtilCol,
    firstAttemptCol,
    lastCol,
    masteryCol,
  };
}

/* -------------------- (2) FORMULAS -------------------- */
/**
 * Populate the computed formulas in the Grades sheet.
 *
 * Why the “look at the header” pattern?
 * - Teachers may add more attempt columns at any time (e.g., insert 5 more columns for level "B").
 * - We label attempt headers with a short level prefix (like B1, B2, …) and then, per data row,
 *   we dynamically FILTER only the attempt cells whose header begins with that prefix.
 * - To make this robust with Google Sheets’ new Table views (sorting, Group by), we avoid hard-coded
 *   references to row 1 and instead derive the header vector relative to the current row using
 *   ROW()-based OFFSET/INDEX. We also scan out to ZZ so newly added attempt columns are automatically
 *   discovered by the formulas.
 *
 * How do we compute “streaks” from symbol entries?
 * - Each attempt cell contains a symbol (✓, ✗, etc.). Settings define a Symbols table mapping
 *   each symbol to a mastery bit (1 for proficient/correct, 0 otherwise) and a display string.
 * - For each level on a row we:
 *   1) Look up each attempt symbol’s mastery bit via XLOOKUP and TEXTJOIN them into a bitstring
 *      like "011101" (String column).
 *   2) Compute the longest consecutive run of 1s via MAX(LEN(SPLIT(bitstring, "0"))). Splitting on 0s
 *      yields only the 1-runs; taking their lengths and MAX gives the streak length.
 *
 * Mastery Grade logic (summary):
 * - If a row has no attempts, return "-".
 * - Otherwise, evaluate levels from highest to lowest; the highest level whose required streak
 *   threshold is met determines the grade.
 * - If no correct attempts exist anywhere on the row, emit the configured "none correct" score.
 *   Otherwise emit the configured "some correct" score.
 */
function setupGradesFormulas_(sh, settings, ctx) {
  const { firstUtilCol, firstAttemptCol, lastCol } = ctx;

  // Shared A1 ranges for header row (1) and first data row (2)
  // Use ROW()-relative OFFSET/INDEX so formulas tolerate sorting and Table views.
  const startA1 = columnA1(firstAttemptCol);
  // Header vector for the current row: first row above this row, across all attempt columns
  const headerGeneric = `OFFSET(INDEX(${startA1}:${startA1},ROW()),(ROW()-1)*-1,0,1,COLUMNS(${startA1}:ZZ))`;
  // Current row attempt values across all attempt columns (expands to new columns automatically)
  const rowValsGeneric = `OFFSET(INDEX(${startA1}:${startA1},ROW()),0,0,1,COLUMNS(${startA1}:ZZ))`;

  // Per-level String (mastery bits) and Streak formulas into row 2
  settings.codes.forEach((code, i) => {
    const streakCol = firstUtilCol + i * 3;
    const stringCol = streakCol + 1;
    const symbolsCol = streakCol + 2;

    // Map symbol chars in attempt cells to mastery bits using the Symbols table.
    // We FILTER the current row of attempt values by headers matching this level’s prefix (e.g., "^B").
    // Then XLOOKUP the symbols to their mastery bits and TEXTJOIN into a bitstring for streak analysis.
    const stringFormula =
      `=LET(hdr, ${headerGeneric}, rowvals, ${rowValsGeneric},` +
      `TEXTJOIN("",TRUE,ARRAYFORMULA(` +
      `XLOOKUP(FILTER(rowvals, REGEXMATCH(hdr, "^"&"${code}")), ${RANGE_SYMBOL_CHARS}, ${RANGE_SYMBOL_MASTERY}, 0)` +
      `)))`;
    sh.getRange(2, stringCol).setFormula(stringFormula);

    // Longest run of 1s in the per-level string. We split on 0 (treating 0 as a divider),
    // measure each run’s length, and take the maximum. Empty string -> no streak.
    const stringCellA1 = `${columnA1(stringCol)}2`;
    const streakFormula = `=IF(${stringCellA1}="","",MAX(ARRAYFORMULA(LEN(SPLIT(${stringCellA1},"0",FALSE,FALSE)))))`;
    sh.getRange(2, streakCol).setFormula(streakFormula);

    // Symbols: join the display symbols (e.g., ✓ ✗) corresponding to attempts for this level
    // to provide a compact visual summary in the utility area.
    const symbolsFormula =
      `=LET(hdr, ${headerGeneric}, rowvals, ${rowValsGeneric},` +
      `TEXTJOIN("",TRUE,ARRAYFORMULA(` +
      `XLOOKUP(FILTER(rowvals, REGEXMATCH(hdr, "^"&"${code}")), ${RANGE_SYMBOL_CHARS}, ${RANGE_SYMBOL_SYMBOL}, "-")` +
      `)))`;
    sh.getRange(2, symbolsCol).setFormula(symbolsFormula);
  });

  // Mastery Grade formula (highest level whose streak threshold is met wins)
  // noneCorrectCheck: detect if there are zero "1" bits across all per-level String columns for this row.
  const noneCorrectCheck = `ISERROR(SEARCH("1", TEXTJOIN("", TRUE, {${settings.codes.map((_, i) => {
    const strCol = columnA1(firstUtilCol + i * 3 + 1);
    return `INDEX(${strCol}:${strCol},ROW())`;
  }).join(',')}} )))`;
  const parts = settings.codes
    .map((_, i) => ({
      // Compare this row’s streak for the level against that level’s required streak threshold.
      cond: `INDEX(${columnA1(firstUtilCol + i * 3)}:${columnA1(firstUtilCol + i * 3)},ROW())>=INDEX(${RANGE_LEVEL_STREAK},${i + 1})`,
      val: `INDEX(${RANGE_LEVEL_SCORES},${i + 1})`,
    }))
    .reverse(); // evaluate highest first

  const ifs =
    `=IFS(` +
    // No attempts on this row -> show "-" to indicate ungraded.
    `COUNTA(${rowValsGeneric})=0,"-",` +
    parts.map(p => `${p.cond},${p.val}`).join(',') + (parts.length ? ',' : '') +
    // If no "1" anywhere, emit the configured "none correct" score; otherwise fall back to "some correct".
    `${noneCorrectCheck},${RANGE_NONE_CORRECT_SCORE},` +
    `TRUE,${RANGE_SOME_CORRECT_SCORE}` +
    `)`;

  // Mastery Grade uses dynamic column index from ctx
  sh.getRange(2, ctx.masteryCol).setFormula(ifs);
}

/* -------------------- helpers -------------------- */
function readLevelSettings_(ss) {
  const codes = ss.getRangeByName(RANGE_LEVEL_SHORTCODES).getValues().flat().filter(String);
  const names = ss.getRangeByName(RANGE_LEVEL_NAMES).getValues().flat().slice(0, codes.length);
  const defaultAttempts = ss.getRangeByName(RANGE_LEVEL_DEFAULTATTEMPTS).getValues().flat().slice(0, codes.length);
  // We read these to maintain parity with prior logic; formulas reference named ranges directly
  // const streaks = ss.getRangeByName(RANGE_LEVEL_STREAK).getValues().flat().slice(0, codes.length);
  // const scores  = ss.getRangeByName(RANGE_LEVEL_SCORES).getValues().flat().slice(0, codes.length);
  return { codes, names, defaultAttempts };
}

function ensureGradesSheet_(ss) {
  return ss.getSheetByName(SHEET_GRADES) || ss.insertSheet(SHEET_GRADES);
}

function columnA1(n) {
  let s = '';
  while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
  return s;
}

/**
 * Hide utility columns, validate attempt inputs, and visually flag computed cells.
 */
function applyGradesFormatting_(sh, settings, ctx) {
  const ss = sh.getParent();

  // Hide the computing columns (Streak and String) but leave Mastery visible
  if (settings.codes.length > 0) {
    settings.codes.forEach((_, i) => {
      const streakCol = ctx.firstUtilCol + i * 3; // Streak
      sh.hideColumns(streakCol, 2); // hide Streak and String
    });
  }

  // Mark computed columns (Mastery Grade + util) with a subtle fill and readable text (with light striping)
  const neutralBg = (STYLE && STYLE.COLORS && STYLE.COLORS.UI && STYLE.COLORS.UI.NEUTRAL_BG) || '#f7f7f7';
  const neutralBgAlt = (STYLE && STYLE.COLORS && STYLE.COLORS.UI && STYLE.COLORS.UI.NEUTRAL_BG_ALT) || '#f0f0f0';
  const neutralText = (STYLE && STYLE.COLORS && STYLE.COLORS.UI && STYLE.COLORS.UI.NEUTRAL_TEXT) || '#333333';
  const totalRows = Math.max(2, sh.getMaxRows());
  // Mastery column base background (we'll overlay gradient CF on data rows)
  sh.getRange(1, 4, totalRows, 1).setBackground(neutralBg).setFontColor(neutralText);
  // Apply background per level for both hidden Streak/String and visible Mastery
  settings.codes.forEach((_, i) => {
    const streakCol = ctx.firstUtilCol + i * 3;
    const symbolsCol = streakCol + 2;
    const rows = totalRows;
    sh.getRange(1, streakCol, rows, 2).setBackground(neutralBg);
    // Visible per-level REFLECTION column (Symbols) uses regular level BG.
    // Map by header name to level index to avoid mismatch if order changes.
    const headerName = sh.getRange(1, symbolsCol).getValue();
    const levelIdx = settings.names.findIndex(n => n && headerName && String(headerName).startsWith(n));
    const level = levelIdx >= 0 ? (levelIdx + 1) : (i + 1);
    const levelDef = (STYLE && STYLE.COLORS && STYLE.COLORS.LEVELS[level]) || {};
    const levelBg = levelDef.BG || neutralBg;
    const levelBgAlt = levelDef.BG_ALT || levelBg;
    const levelText = levelDef.TEXT || '#000000';
    // Base background for entire column (header + data)
    sh.getRange(1, symbolsCol, rows, 1).setBackground(levelBg).setFontColor(levelText);
    // Conditional format stripe on even-numbered rows (data rows only)
    try {
      const dataRange = sh.getRange(2, symbolsCol, Math.max(1, sh.getMaxRows() - 1), 1);
      const rules = sh.getConditionalFormatRules();
      const dataA1 = dataRange.getA1Notation();
      const filtered = rules.filter(r => !r.getRanges().some(rg => rg.getA1Notation() === dataA1));
      const stripeRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=ISEVEN(ROW())')
        .setBackground(levelBgAlt)
        .setRanges([dataRange])
        .build();
      filtered.push(stripeRule);
      sh.setConditionalFormatRules(filtered);
    } catch (e) { /* CF not applied; static base color remains */ }
  });

  // Fixed base widths (A..F): Name, Email, Unit, Skill #, Description, Mastery Grade
  const baseWidths = [115, 232, 40, 46, 172, 60];
  baseWidths.forEach((w, i) => sh.setColumnWidth(1 + i, w));

  // Set a reasonable width for visible per-level Symbols (display) columns
  settings.codes.forEach((_, i) => {
    const symbolsCol = ctx.firstUtilCol + i * 3 + 2;
    sh.setColumnWidth(symbolsCol, 82);
  });

  // Attempt columns: restrict input, set as text, narrow width (~34px), center align, and highlight per level
  if (ctx.firstAttemptCol <= ctx.lastCol) {
    const attemptsWidth = ctx.lastCol - ctx.firstAttemptCol + 1;
    const symbolsRange = ss.getRangeByName(RANGE_SYMBOL_CHARS);
    if (symbolsRange) {
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(symbolsRange, true)
        .setAllowInvalid(true) // warn, don’t block
        .setHelpText('Type a symbol or choose from the list. See the Symbols sheet for allowed entries (e.g., 1, 1o, H, PC, X, 0, N).')
        .build();
      sh.getRange(2, ctx.firstAttemptCol, sh.getMaxRows() - 1, attemptsWidth).setDataValidation(rule);
    }
    sh.getRange(1, ctx.firstAttemptCol, sh.getMaxRows(), attemptsWidth).setNumberFormat('@');

    let col = ctx.firstAttemptCol;
    settings.defaultAttempts.forEach((cntRaw, levelIdx) => {
      const cnt = Number(cntRaw || 0);
      if (cnt > 0) {
        // Determine level index from attempt headers (e.g., B1...). Use the first column's header to map.
        const headerVal = sh.getRange(1, col).getValue();
        const code = (headerVal && String(headerVal).trim()) ? String(headerVal).trim().charAt(0) : '';
        const levelFromCode = settings.codes.findIndex(c => String(c).toUpperCase() === code.toUpperCase());
        const level = levelFromCode >= 0 ? (levelFromCode + 1) : (levelIdx + 1);
        const levelDef = (STYLE && STYLE.COLORS && STYLE.COLORS.LEVELS[level]) || {};
        const levelBgBright = levelDef.BG_BRIGHT || '#fff7d6';
        const levelBgBrightAlt = levelDef.BG_BRIGHT_ALT || levelBgBright;
        const levelTextBright = levelDef.TEXT_BRIGHT || '#000000';
        // Narrow width and center align
        sh.setColumnWidths(col, cnt, 34);
        sh.getRange(1, col, sh.getMaxRows(), cnt).setHorizontalAlignment('center');
        // Highlight the whole attempt area for this level (ACTION area)
        sh.getRange(1, col, totalRows, cnt).setBackground(levelBgBright).setFontColor(levelTextBright);
        // Conditional format stripe on even-numbered rows
        try {
          const dataRange = sh.getRange(2, col, Math.max(1, sh.getMaxRows() - 1), cnt);
          const rules = sh.getConditionalFormatRules();
          const dataA1 = dataRange.getA1Notation();
          const filtered = rules.filter(r => !r.getRanges().some(rg => rg.getA1Notation() === dataA1));
          const stripeRule = SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied('=ISEVEN(ROW())')
            .setBackground(levelBgBrightAlt)
            .setRanges([dataRange])
            .build();
          filtered.push(stripeRule);
          sh.setConditionalFormatRules(filtered);
        } catch (e) { /* CF not applied; static base color remains */ }
        col += cnt;
      }
    });
  }

  // Header bold for readability and header background
  const headerCols = ctx.headers && ctx.headers.length ? ctx.headers.length : ctx.lastCol;
  const headerBg = (STYLE && STYLE.COLORS && STYLE.COLORS.UI && STYLE.COLORS.UI.HEADER_BG) || '#f0f3f5';
  const headerText = (STYLE && STYLE.COLORS && STYLE.COLORS.UI && STYLE.COLORS.UI.HEADER_TEXT) || '#000000';
  sh.getRange(1, 1, 1, headerCols).setFontWeight('bold').setBackground(headerBg).setFontColor(headerText);

  // Conditional formatting for Mastery Grade column using GRADE_SCALE
  try {
    const minColor = STYLE.COLORS.GRADE_SCALE.MIN; // background low color
    const midColor = STYLE.COLORS.GRADE_SCALE.MID; // midpoint color
    const maxColor = STYLE.COLORS.GRADE_SCALE.MAX; // background high color
    const textOnScale = STYLE.COLORS.GRADE_SCALE.TEXT; // foreground text color
    const gradeCol = ctx.masteryCol;
    const dataRange = sh.getRange(2, gradeCol, Math.max(1, sh.getMaxRows() - 1), 1);
    const rules = sh.getConditionalFormatRules();
    // Remove old rules targeting Mastery Grade to avoid duplicates
    const dataA1 = dataRange.getA1Notation();
    const filtered = rules.filter(r => !r.getRanges().some(rg => rg.getA1Notation() === dataA1));
    // Determine thresholds: min = NoneCorrectScore, max from LevelScores
    let maxScore = 1;
    try {
      const lvl = sh.getParent().getRangeByName(RANGE_LEVEL_SCORES);
      if (lvl) {
        const nums = lvl.getValues().flat().map(v => Number(v)).filter(v => !isNaN(v));
        if (nums.length) maxScore = Math.max.apply(null, nums);
      }
    } catch (e) { /* default to 1 */ }
    let minScore = 0;
    try {
      const none = sh.getParent().getRangeByName(RANGE_NONE_CORRECT_SCORE);
      if (none) {
        const n = Number(none.getValue());
        if (!isNaN(n)) minScore = n;
      }
    } catch (e) { /* default to 0 */ }
    let midScore = null;
    try {
      const some = sh.getParent().getRangeByName(RANGE_SOME_CORRECT_SCORE);
      if (some) {
        const m = Number(some.getValue());
        if (!isNaN(m)) midScore = m;
      }
    } catch (e) { /* optional */ }
    // Background gradient scale (built-in) anchored at minScore..maxScore
    let gradientBuilder = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpointWithValue(minColor, SpreadsheetApp.InterpolationType.NUMBER, String(minScore));
    if (midScore !== null && midColor) {
      gradientBuilder = gradientBuilder.setGradientMidpointWithValue(midColor, SpreadsheetApp.InterpolationType.NUMBER, String(midScore));
    }
    const gradient = gradientBuilder
      .setGradientMaxpointWithValue(maxColor, SpreadsheetApp.InterpolationType.NUMBER, String(maxScore))
      .setRanges([dataRange])
      .build();
    // Foreground text on the gradient
    const textColorRule = SpreadsheetApp.newConditionalFormatRule()
      .whenCellNotEmpty()
      .setFontColor(textOnScale)
      .setRanges([dataRange])
      .build();
    filtered.push(gradient, textColorRule);
    // Before applying mastery-gradient rules, add unit-based text color rules for Unit/Skill#/Description
    try {
      const dataStart = 2;
      const dataCount = Math.max(1, sh.getMaxRows() - 1);
      // Place a hidden helper column that maps the Unit text to an index via UNIQUE(RANGE_SKILL_UNITS)
      const minHelperStart = 12;
      const fallbackLast = (ctx && ctx.lastCol) ? ctx.lastCol : sh.getLastColumn();
      const helperCol = Math.max((ctx && ctx.lastCol) ? ctx.lastCol + 1 : fallbackLast + 1, minHelperStart);
      if (sh.getMaxColumns() < helperCol) {
        sh.insertColumnsAfter(sh.getMaxColumns(), helperCol - sh.getMaxColumns());
      }
      // Write an ARRAYFORMULA that MATCHes the Unit column (C) against the unique skill units list.
      // This produces a 1-based index per data row that we can reference from conditional formatting.
      sh.getRange(dataStart, helperCol).setFormula(`=ARRAYFORMULA(MATCH($C${dataStart}:$C, UNIQUE(${RANGE_SKILL_UNITS}), 0))`);
      const helperColA = columnA1(helperCol);
      // Hide the helper column
      try { sh.hideColumns(helperCol, 1); } catch (e) { /* ignore */ }

      // Build unit color conditional-format rules for columns C:E (Unit, Skill #, Skill Description)
      const unitRange = sh.getRange(dataStart, 3, dataCount, 3);
      const rulesBefore = filtered; // start from filtered set we've been building
      const unitColors = (STYLE && STYLE.COLORS && STYLE.COLORS.UI && STYLE.COLORS.UI.UNIT_TEXT_COLORS) || ['#3a3a3a'];
      const nColors = Math.max(1, unitColors.length);
      const helperIdxCol = helperColA ? `$${helperColA}` : '$AB';
      // Remove any existing rules that target the same unitRange to avoid duplicates
      const targetA1 = unitRange.getA1Notation();
      let baseRules = rulesBefore.filter(r => !r.getRanges().some(rg => rg.getA1Notation() === targetA1));

      const unitRules = [];
      for (let k = 0; k < nColors; k++) {
        const fEven = `=AND(${helperIdxCol}${dataStart}<>"", MOD(${helperIdxCol}${dataStart}-1, ${nColors})=${k}, ISEVEN(ROW()))`;
        const fOdd = `=AND(${helperIdxCol}${dataStart}<>"", MOD(${helperIdxCol}${dataStart}-1, ${nColors})=${k}, NOT(ISEVEN(ROW())))`;
        const ruleEven = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(fEven)
          .setFontColor(unitColors[k])
          .setRanges([unitRange])
          .build();
        const ruleOdd = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(fOdd)
          .setFontColor(unitColors[k])
          .setRanges([unitRange])
          .build();
        unitRules.push(ruleEven, ruleOdd);
      }
      // Prepend unit rules so they take precedence over generic stripes/backgrounds
      baseRules.unshift(...unitRules);
      sh.setConditionalFormatRules(baseRules);
    } catch (e) {
      if (console && console.warn) console.warn('Grades unit color rules warn', e);
      // Fallback: apply previously computed filtered rules
      sh.setConditionalFormatRules(filtered);
    }
  } catch (e) { /* STYLE or method may be unavailable; skip gracefully */ }
}

/**
 * Reapply formatting only (no content or headers changes).
 * Useful after manual edits or palette tweaks.
 */
function reformatGradesSheet() {
  const ss = SpreadsheetApp.getActive();
  const sh = ensureGradesSheet_(ss);
  const settings = readLevelSettings_(ss);
  const layout = computeGradesLayoutFromSettings_(settings);
  const headersLen = Math.max(sh.getLastColumn(), layout.lastCol);
  const ctx = {
    headers: new Array(headersLen).fill(''),
    firstUtilCol: layout.firstUtilCol,
    firstAttemptCol: layout.firstAttemptCol,
    lastCol: Math.max(layout.lastCol, sh.getLastColumn()),
    masteryCol: layout.masteryCol,
  };
  // 1) Re-apply formatting, colors, and data validation
  applyGradesFormatting_(sh, settings, ctx);
  // 2) Re-apply all computed formulas across existing data rows
  fillComputedFormulas_(sh, settings, layout);
}

/**
 * Normalize user edits in attempt cells to reduce validation friction.
 * - Trim spaces, uppercase letters (except keep '1' and digits),
 * - Map '0' to 'N' (explicit none), allow 'pc' -> 'PC', 'xo'/'xc'/'xs' -> uppercased,
 * - Leave anything else as-is (validation is warn-only and rules will still compute).
 */
function onEdit(e) {
  try {
    if (!e || !e.range || !e.value) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== SHEET_GRADES) return;
    // Determine attempt columns from headers (look for XY1 style like B1, I1, etc.)
    const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const firstAttemptCol = header.findIndex(h => /^([A-Za-z])1$/.test(String(h || ''))) + 1;
    if (firstAttemptCol <= 0) return;
    const lastCol = sh.getLastColumn();
    const r = e.range;
    const inAttemptArea = r.getRow() >= 2 && r.getColumn() >= firstAttemptCol && r.getColumn() <= lastCol;
    if (!inAttemptArea) return;
    const raw = String(e.value).trim();
    if (raw === '') return;
    let norm = raw;
    // quick aliases
    if (norm === '0') norm = 'N';
    // Preserve leading '1' then optional letter
    const m = norm.match(/^1([a-zA-Z])?$/);
    if (m) {
      norm = '1' + (m[1] ? m[1].toLowerCase() : '');
    } else {
      // Uppercase general tokens like pc, h, g, x, xo, xc, xs
      norm = norm.toUpperCase();
    }
    // Replace only if changed
    if (norm !== raw) {
      r.setValue(norm);
    }
  } catch (err) {
    // Best-effort; ignore errors to avoid blocking user typing
  }
}



/* -------------------- population -------------------- */
/**
 * Populate Grades with one row per Student x Skill and fill formulas.
 * Idempotent: won’t duplicate existing (Name, Email, Skill) rows.
 */
function populateGrades() {
  const ss = SpreadsheetApp.getActive();
  const settings = readLevelSettings_(ss);
  const sh = ensureGradesSheet_(ss);

  // Derive layout from settings (must match setupGradesHeaders_)
  const layout = computeGradesLayoutFromSettings_(settings);

  // Load students (skip blanks)
  const names = ss.getRangeByName(RANGE_STUDENT_NAMES).getValues().flat();
  const emails = ss.getRangeByName(RANGE_STUDENT_EMAILS).getValues().flat();
  const students = [];
  for (let i = 0; i < Math.max(names.length, emails.length); i++) {
    const name = (names[i] || '').toString().trim();
    const email = (emails[i] || '').toString().trim();
    if (name) students.push({ name, email });
  }

  // Load skills (skip rows with no unit and no descriptor)
  const units = ss.getRangeByName(RANGE_SKILL_UNITS).getValues().flat();
  const numbers = ss.getRangeByName(RANGE_SKILL_NUMBERS).getValues().flat();
  const descs = ss.getRangeByName(RANGE_SKILL_DESCRIPTORS).getValues().flat();
  const skills = [];
  const maxSkills = Math.max(units.length, numbers.length, descs.length);
  for (let i = 0; i < maxSkills; i++) {
    const unit = (units[i] || '').toString().trim();
    const num = numbers[i] != null && numbers[i] !== '' ? numbers[i] : '';
    const desc = (descs[i] || '').toString().trim();
    if (unit || desc) {
      skills.push({ unit, num, desc });
    }
  }

  if (students.length === 0 || skills.length === 0) return; // nothing to do

  // Existing rows map: key = email|unit|num|desc, value -> row index
  const existing = new Map();
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    const rng = sh.getRange(2, 1, lastRow - 1, 5).getValues();
    rng.forEach((r, idx) => {
      const email = (r[1] || '').toString().trim();
      const unit = (r[2] || '').toString().trim();
      const num = (r[3] != null && r[3] !== '') ? String(r[3]) : '';
      const desc = (r[4] || '').toString().trim();
      const key = `${email}|${unit}|${num}|${desc}`;
      if (email) existing.set(key, 2 + idx);
    });
  }

  // Build missing rows
  const newRows = [];
  students.forEach(s => {
    skills.forEach(sk => {
      const key = `${s.email}|${sk.unit}|${sk.num !== '' ? String(sk.num) : ''}|${sk.desc}`;
      if (!existing.has(key)) {
        newRows.push([s.name, s.email, sk.unit, sk.num, sk.desc]);
        existing.set(key, lastRow + newRows.length); // provisional index
      }
    });
  });

  // Append missing rows in one batch
  if (newRows.length > 0) {
    // Find the first fully-empty data row among typed columns (A..E) so we don't append at the
    // very bottom when ARRAYFORMULA or helper columns are present. This ignores auto-filled
    // helper columns and targets rows where all of A..E are blank.
    const insertRow = findFirstEmptyDataRow_(sh) || (sh.getLastRow() + 1);
    sh.getRange(insertRow, 1, newRows.length, 5).setValues(newRows);
  }

  // Fill computed formulas down for all data rows
  fillComputedFormulas_(sh, settings, layout);
}

/** Compute layout (column indices and last column) from settings without touching the sheet. */
function computeGradesLayoutFromSettings_(settings) {
  const baseHeadersCount = 6; // Name, Email, Unit, Skill #, Skill Description, Mastery Grade
  const utilHeadersCount = settings.codes.length * 3; // Streak, String, Mastery per level
  const attemptCount = settings.defaultAttempts.reduce((a, b) => a + Number(b || 0), 0);
  const firstUtilCol = baseHeadersCount + 1;
  const firstAttemptCol = baseHeadersCount + utilHeadersCount + 1;
  const lastCol = baseHeadersCount + utilHeadersCount + attemptCount;
  const masteryCol = baseHeadersCount; // 6
  return { firstUtilCol, firstAttemptCol, lastCol, masteryCol };
}

/** Fill formulas for columns: Mastery Grade, Streak/String/Mastery for each level, across all rows.
 * This mirrors the logic of setupGradesFormulas_ but writes row-agnostic formulas for an entire range.
 * Notes:
 * - We use the same ROW()-relative header/row derivations so formulas stay correct under sorting/Table views.
 * - Attempt scans extend to ZZ so newly inserted attempt columns are included automatically.
 */
function fillComputedFormulas_(sh, settings, layout) {
  const { firstUtilCol, firstAttemptCol, lastCol } = layout;
  const startRow = 2;
  const endRow = sh.getLastRow();
  const rowCount = Math.max(0, endRow - startRow + 1);
  if (rowCount <= 0) return;

  // Precompute header attempt A1 range (row 1) once
  const startA1 = columnA1(firstAttemptCol);
  const headerGeneric = `OFFSET(INDEX(${startA1}:${startA1},ROW()),(ROW()-1)*-1,0,1,COLUMNS(${startA1}:ZZ))`;
  const rowValsGeneric = `OFFSET(INDEX(${startA1}:${startA1},ROW()),0,0,1,COLUMNS(${startA1}:ZZ))`;

  // Mastery Grade formulas for all rows
  const masteryCol = layout.masteryCol;
  const masteryFormulas = new Array(rowCount);
  // noneCorrectCheck: same concept as above, but expressed generically for a filled range.
  const noneCorrectCheckGeneric = `ISERROR(SEARCH("1", TEXTJOIN("", TRUE, {${settings.codes.map((_, i) => {
    const strCol = columnA1(firstUtilCol + i * 3 + 1);
    return `INDEX(${strCol}:${strCol},ROW())`;
  }).join(',')}} )))`;
  const partsGeneric = settings.codes.map((_, i) => {
    const streakCol = columnA1(firstUtilCol + i * 3);
    return `INDEX(${streakCol}:${streakCol},ROW())>=INDEX(${RANGE_LEVEL_STREAK},${i + 1}),INDEX(${RANGE_LEVEL_SCORES},${i + 1})`;
  }).reverse().join(',');
  const masteryFormulaGeneric = `=IFS(COUNTA(${rowValsGeneric})=0,"-",${partsGeneric}${partsGeneric ? ',' : ''}${noneCorrectCheckGeneric},${RANGE_NONE_CORRECT_SCORE},TRUE,${RANGE_SOME_CORRECT_SCORE})`;
  for (let r = 0; r < rowCount; r++) {
    masteryFormulas[r] = [masteryFormulaGeneric];
  }
  sh.getRange(startRow, masteryCol, rowCount, 1).setFormulas(masteryFormulas);

  // Per-level Streak/String/Mastery display formulas
  settings.codes.forEach((code, i) => {
    const streakCol = firstUtilCol + i * 3;
    const stringCol = streakCol + 1;
    const symbolsCol = streakCol + 2;

    const streakArr = new Array(rowCount);
    const stringArr = new Array(rowCount);
    const symbolsArr = new Array(rowCount);
    const stringColLetter = columnA1(stringCol);
    const stringCellGeneric = `INDEX(${stringColLetter}:${stringColLetter},ROW())`;
    // Generic formulas used for each row of the filled ranges:
    // - stringFormula: collect mastery bits for this level across attempts on the row
    // - streakFormula: longest run of 1s
    // - symbolsFormula: pretty symbol string for display
    const stringFormulaGeneric = `=LET(hdr, ${headerGeneric}, rowvals, ${rowValsGeneric},TEXTJOIN("",TRUE,ARRAYFORMULA(XLOOKUP(FILTER(rowvals, REGEXMATCH(hdr, "^"&"${code}")), ${RANGE_SYMBOL_CHARS}, ${RANGE_SYMBOL_MASTERY}, 0))))`;
    const streakFormulaGeneric = `=IF(${stringCellGeneric}="","",MAX(ARRAYFORMULA(LEN(SPLIT(${stringCellGeneric},"0",FALSE,FALSE)))))`;
    const symbolsFormulaGeneric = `=LET(hdr, ${headerGeneric}, rowvals, ${rowValsGeneric},TEXTJOIN("",TRUE,ARRAYFORMULA(XLOOKUP(FILTER(rowvals, REGEXMATCH(hdr, "^"&"${code}")), ${RANGE_SYMBOL_CHARS}, ${RANGE_SYMBOL_SYMBOL}, "-"))))`;
    for (let r = 0; r < rowCount; r++) {
      stringArr[r] = [stringFormulaGeneric];
      streakArr[r] = [streakFormulaGeneric];
      symbolsArr[r] = [symbolsFormulaGeneric];
    }
    sh.getRange(startRow, stringCol, rowCount, 1).setFormulas(stringArr);
    sh.getRange(startRow, streakCol, rowCount, 1).setFormulas(streakArr);
    sh.getRange(startRow, symbolsCol, rowCount, 1).setFormulas(symbolsArr);
  });
}