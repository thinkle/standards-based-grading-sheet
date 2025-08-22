/* global SpreadsheetApp, RANGE_SYMBOL_CHARS, RANGE_SYMBOL_MASTERY, RANGE_SYMBOL_SYMBOL, RANGE_LEVEL_STREAK, RANGE_LEVEL_SCORES, RANGE_NONE_CORRECT_SCORE, RANGE_SOME_CORRECT_SCORE, RANGE_LEVEL_SHORTCODES, RANGE_LEVEL_NAMES, RANGE_LEVEL_DEFAULTATTEMPTS */
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

  // Basic sheet niceties (frozen header + autoresize)
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, ctx.headers.length);

  // Light formatting and validation
  applyGradesFormatting_(sh, settings, ctx);
}

/* -------------------- (1) HEADERS -------------------- */
function setupGradesHeaders_(sh, settings) {
  // Base columns that are typed-in by the teacher
  const baseHeaders = ['Name', 'Email', 'Skill', 'Mastery Grade'];

  // Utility columns per level: Streak + String + Symbols
  const utilHeaders = settings.codes.flatMap((_, i) => [
    `${settings.names[i]} Streak`,
    `${settings.names[i]} String`,
    `${settings.names[i]} Symbols`,
  ]);

  // Attempt columns per level: repeat short code per defaultAttempts[i]
  const attemptHeaders = settings.codes.flatMap((code, i) =>
    Array.from({ length: Number(settings.defaultAttempts[i] || 0) }, () => code)
  );

  const headers = [...baseHeaders, ...utilHeaders, ...attemptHeaders];
  sh.clear();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Compute column indices we need later
  const firstUtilCol = baseHeaders.length + 1;
  const firstAttemptCol = baseHeaders.length + utilHeaders.length + 1;
  const lastCol = headers.length;

  return {
    headers,
    baseHeaders,
    firstUtilCol,
    firstAttemptCol,
    lastCol,
  };
}

/* -------------------- (2) FORMULAS -------------------- */
function setupGradesFormulas_(sh, settings, ctx) {
  const { firstUtilCol, firstAttemptCol, lastCol } = ctx;

  // Shared A1 ranges for header row (1) and first data row (2)
  const attemptHeaderA1 = `${columnA1(firstAttemptCol)}1:${columnA1(lastCol)}1`;
  const attemptRowA1 = `${columnA1(firstAttemptCol)}2:${columnA1(lastCol)}2`;

  // Per-level String (mastery bits) and Streak formulas into row 2
  settings.codes.forEach((code, i) => {
    const streakCol = firstUtilCol + i * 3;
    const stringCol = streakCol + 1;
    const symbolsCol = streakCol + 2;

    // Map symbol chars in attempt cells to mastery bits using Symbols table
    const stringFormula =
      `=TEXTJOIN("",TRUE,ARRAYFORMULA(` +
      `XLOOKUP(FILTER(${attemptRowA1}, REGEXMATCH(${attemptHeaderA1}, "^"&"${code}")), ${RANGE_SYMBOL_CHARS}, ${RANGE_SYMBOL_MASTERY}, "-")` +
      `))`;
    sh.getRange(2, stringCol).setFormula(stringFormula);

    // Longest run of 1s in the per-level string
    const stringCellA1 = `${columnA1(stringCol)}2`;
    const streakFormula = `=IF(${stringCellA1}="","",MAX(ARRAYFORMULA(LEN(SPLIT(${stringCellA1},"0",FALSE,FALSE)))))`;
    sh.getRange(2, streakCol).setFormula(streakFormula);

    // Symbols: join pretty symbols corresponding to attempts for this level
    const symbolsFormula =
      `=TEXTJOIN("",TRUE,ARRAYFORMULA(` +
      `XLOOKUP(FILTER(${attemptRowA1}, REGEXMATCH(${attemptHeaderA1}, "^"&"${code}")), ${RANGE_SYMBOL_CHARS}, ${RANGE_SYMBOL_SYMBOL}, "-")` +
      `))`;
    sh.getRange(2, symbolsCol).setFormula(symbolsFormula);
  });

  // Mastery Grade formula (highest level whose streak threshold is met wins)
  const noneCorrectCheck = `ISERROR(SEARCH("1", TEXTJOIN("", TRUE, {${settings.codes.map((_, i) => columnA1(firstUtilCol + i * 3 + 1) + '2').join(',')}} )))`;
  const parts = settings.codes
    .map((_, i) => ({
      cond: `${columnA1(firstUtilCol + i * 3)}2>=INDEX(${RANGE_LEVEL_STREAK},${i + 1})`,
      val: `INDEX(${RANGE_LEVEL_SCORES},${i + 1})`,
    }))
    .reverse(); // evaluate highest first

  const ifs =
    `=IFS(` +
    `COUNTA(${attemptRowA1})=0,"-",` +
    parts.map(p => `${p.cond},${p.val}`).join(',') + (parts.length ? ',' : '') +
    `${noneCorrectCheck},${RANGE_NONE_CORRECT_SCORE},` +
    `TRUE,${RANGE_SOME_CORRECT_SCORE}` +
    `)`;

  // Mastery Grade is column 4
  sh.getRange(2, 4).setFormula(ifs);
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

  // Hide the computing columns (String and Streak)
  const utilCols = settings.codes.length * 3;
  if (utilCols > 0) sh.hideColumns(ctx.firstUtilCol, utilCols);

  // Mark computed columns (Mastery Grade + util) with a subtle fill and protection (warning only)
  const computedBg = '#f5f5f5';
  const gradeRange = sh.getRange(1, 4, Math.max(2, sh.getMaxRows() - 0), 1).setBackground(computedBg);
  if (utilCols > 0) {
    const utilRange = sh.getRange(1, ctx.firstUtilCol, Math.max(2, sh.getMaxRows() - 0), utilCols).setBackground(computedBg);
    try {
      const pUtil = utilRange.protect();
      pUtil.setWarningOnly(true);
    } catch (e) {
      console && console.warn && console.warn('Protection (util) warning', e);
    }
  }
  try {
    const pGrade = gradeRange.protect();
    pGrade.setWarningOnly(true);
  } catch (e) {
    console && console.warn && console.warn('Protection (grade) warning', e);
  }

  // Attempt columns: restrict input to entries in SymbolChars named range and keep as text
  if (ctx.firstAttemptCol <= ctx.lastCol) {
    const attemptsWidth = ctx.lastCol - ctx.firstAttemptCol + 1;
    const symbolsRange = ss.getRangeByName(RANGE_SYMBOL_CHARS);
    if (symbolsRange) {
      const rule = SpreadsheetApp.newDataValidation().requireValueInRange(symbolsRange, true).setAllowInvalid(false).build();
      sh.getRange(2, ctx.firstAttemptCol, sh.getMaxRows() - 1, attemptsWidth).setDataValidation(rule);
    }
    sh.getRange(1, ctx.firstAttemptCol, sh.getMaxRows(), attemptsWidth).setNumberFormat('@STRING@');
  }

  // Header bold for readability
  sh.getRange(1, 1, 1, ctx.headers.length).setFontWeight('bold');
}