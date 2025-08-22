/* global SpreadsheetApp,
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

  // Basic sheet niceties (frozen header + autoresize)
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, ctx.headers.length);

  // Light formatting and validation
  applyGradesFormatting_(sh, settings, ctx);
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
      `XLOOKUP(FILTER(${attemptRowA1}, REGEXMATCH(${attemptHeaderA1}, "^"&"${code}")), ${RANGE_SYMBOL_CHARS}, ${RANGE_SYMBOL_MASTERY}, 0)` +
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

  // Mark computed columns (Mastery Grade + util) with a subtle fill
  sh.getRange(1, 4, Math.max(2, sh.getMaxRows() - 0), 1).setBackground('#f5f5f5');
  // Apply background per level for both hidden Streak/String and visible Mastery
  settings.codes.forEach((_, i) => {
    const streakCol = ctx.firstUtilCol + i * 3;
    const symbolsCol = streakCol + 2;
    const rows = Math.max(2, sh.getMaxRows() - 0);
    sh.getRange(1, streakCol, rows, 2).setBackground('#f5f5f5');
    sh.getRange(1, symbolsCol, rows, 1).setBackground('#f5f5f5');
  });

  // Attempt columns: restrict input to entries in SymbolChars named range and keep as text
  if (ctx.firstAttemptCol <= ctx.lastCol) {
    const attemptsWidth = ctx.lastCol - ctx.firstAttemptCol + 1;
    const symbolsRange = ss.getRangeByName(RANGE_SYMBOL_CHARS);
    if (symbolsRange) {
      const rule = SpreadsheetApp.newDataValidation().requireValueInRange(symbolsRange, true).setAllowInvalid(false).build();
      sh.getRange(2, ctx.firstAttemptCol, sh.getMaxRows() - 1, attemptsWidth).setDataValidation(rule);
    }
    sh.getRange(1, ctx.firstAttemptCol, sh.getMaxRows(), attemptsWidth).setNumberFormat('@STRING@');
    // Make dropdown columns compact and centered
    sh.setColumnWidths(ctx.firstAttemptCol, attemptsWidth, 48);
    sh.getRange(1, ctx.firstAttemptCol, sh.getMaxRows(), attemptsWidth).setHorizontalAlignment('center');
  }

  // Header bold for readability
  sh.getRange(1, 1, 1, ctx.headers.length).setFontWeight('bold');
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
    sh.getRange(sh.getLastRow() + 1, 1, newRows.length, 5).setValues(newRows);
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

/** Fill formulas for columns: Mastery Grade, Streak/String/Mastery for each level, across all rows. */
function fillComputedFormulas_(sh, settings, layout) {
  const { firstUtilCol, firstAttemptCol, lastCol } = layout;
  const startRow = 2;
  const endRow = sh.getLastRow();
  const rowCount = Math.max(0, endRow - startRow + 1);
  if (rowCount <= 0) return;

  // Precompute header attempt A1 range (row 1) once
  const attemptHeaderA1 = `${columnA1(firstAttemptCol)}1:${columnA1(lastCol)}1`;

  // Mastery Grade formulas for all rows
  const masteryCol = layout.masteryCol;
  const masteryFormulas = new Array(rowCount);
  for (let r = 0; r < rowCount; r++) {
    const row = startRow + r;
    const noneCorrectCheck = `ISERROR(SEARCH("1", TEXTJOIN("", TRUE, {${settings.codes.map((_, i) => columnA1(firstUtilCol + i * 3 + 1) + row).join(',')}} )))`;
    const parts = settings.codes.map((_, i) => ({
      cond: `${columnA1(firstUtilCol + i * 3)}${row}>=INDEX(${RANGE_LEVEL_STREAK},${i + 1})`,
      val: `INDEX(${RANGE_LEVEL_SCORES},${i + 1})`,
    })).reverse();
    masteryFormulas[r] = [
      `=IFS(` +
      `COUNTA(${columnA1(firstAttemptCol)}${row}:${columnA1(lastCol)}${row})=0,"-",` +
      parts.map(p => `${p.cond},${p.val}`).join(',') + (parts.length ? ',' : '') +
      `${noneCorrectCheck},${RANGE_NONE_CORRECT_SCORE},` +
      `TRUE,${RANGE_SOME_CORRECT_SCORE}` +
      `)`
    ];
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
    for (let r = 0; r < rowCount; r++) {
      const row = startRow + r;
      const attemptRowA1 = `${columnA1(firstAttemptCol)}${row}:${columnA1(lastCol)}${row}`;
      const stringCellA1 = `${columnA1(stringCol)}${row}`;
      stringArr[r] = [
        `=TEXTJOIN("",TRUE,ARRAYFORMULA(` +
        `XLOOKUP(FILTER(${attemptRowA1}, REGEXMATCH(${attemptHeaderA1}, "^"&"${code}")), ${RANGE_SYMBOL_CHARS}, ${RANGE_SYMBOL_MASTERY}, 0)` +
        `))`
      ];
      streakArr[r] = [
        `=IF(${stringCellA1}="","",MAX(ARRAYFORMULA(LEN(SPLIT(${stringCellA1},"0",FALSE,FALSE)))))`
      ];
      symbolsArr[r] = [
        `=TEXTJOIN("",TRUE,ARRAYFORMULA(` +
        `XLOOKUP(FILTER(${attemptRowA1}, REGEXMATCH(${attemptHeaderA1}, "^"&"${code}")), ${RANGE_SYMBOL_CHARS}, ${RANGE_SYMBOL_SYMBOL}, "-")` +
        `))`
      ];
    }
    sh.getRange(startRow, stringCol, rowCount, 1).setFormulas(stringArr);
    sh.getRange(startRow, streakCol, rowCount, 1).setFormulas(streakArr);
    sh.getRange(startRow, symbolsCol, rowCount, 1).setFormulas(symbolsArr);
  });
}