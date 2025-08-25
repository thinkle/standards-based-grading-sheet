/* eslint-disable no-unused-vars */
/* exported nukeAllSheetsExceptInstructions, populateDemoStudents, populateDemoSkills, populateDemoGrades, runDemoSetup */
/* exported repairTextFormats */
/* global SpreadsheetApp, Browser, Menu, setupNamedRanges, setupStudents, setupSkills, setupGradeSheet, setupGradesSheet, populateGrades, setupGradeViewSheet,
          RANGE_STUDENT_NAMES, RANGE_STUDENT_EMAILS, RANGE_SKILL_UNITS, RANGE_SKILL_NUMBERS, RANGE_SKILL_DESCRIPTORS,
          RANGE_LEVEL_NAMES, RANGE_LEVEL_SHORTCODES, RANGE_LEVEL_DEFAULTATTEMPTS, RANGE_LEVEL_STREAK, RANGE_LEVEL_SCORES,
          RANGE_NONE_CORRECT_SCORE, RANGE_SOME_CORRECT_SCORE,
          RANGE_SYMBOL_CHARS, RANGE_SYMBOL_MASTERY, RANGE_SYMBOL_SYMBOL
*/

/**
 * Danger zone: delete all sheets except Instructions (if present). Double-confirm with text prompt.
 */
function nukeAllSheetsExceptInstructions() {
  const ui = SpreadsheetApp.getUi();
  const proceed = ui.alert('NUKE WARNING', 'This will DELETE all sheets except "Instructions". This cannot be undone. Continue?', ui.ButtonSet.YES_NO);
  if (proceed !== ui.Button.YES) return;
  const resp = ui.prompt('Final confirmation', 'Type exactly: Yes, nuke my data', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK || resp.getResponseText().trim() !== 'Yes, nuke my data') return;

  const ss = SpreadsheetApp.getActive();
  const sheets = ss.getSheets();
  sheets.forEach(sh => {
    const name = sh.getName();
    if (name !== 'Instructions') {
      try { ss.deleteSheet(sh); } catch (e) { /* ignore */ }
    }
  });
}

/**
 * Populate demo students into Students sheet.
 */
function populateDemoStudents() {
  const ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName('Students')) setupStudents();
  const sh = ss.getSheetByName('Students');
  const demo = [
    ['Alice Johnson', 'alice@example.com'],
    ['Ben Rivera', 'ben@example.com'],
    ['Chloe Kim', 'chloe@example.com'],
    ['Diego Patel', 'diego@example.com'],
    ['Emi Sato', 'emi@example.com']
  ];
  sh.getRange(2, 1, demo.length, 2).setValues(demo);
}

/**
 * Populate demo Algebra 1 skills (Common Core flavored, simplified sample).
 */
function populateDemoSkills() {
  const ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName('Skills')) setupSkills();
  const sh = ss.getSheetByName('Skills');
  const demo = [
    ['Unit 1: Expressions & Equations', '1.1', 'Interpret expressions'],
    ['Unit 1: Expressions & Equations', '1.2', 'Write expressions for word problems'],
    ['Unit 1: Expressions & Equations', '1.3', 'Evaluate expressions'],
    ['Unit 2: Linear Equations', '2.1', 'Solve one-variable linear equations'],
    ['Unit 2: Linear Equations', '2.2', 'Solve two-step linear equations'],
    ['Unit 2: Linear Equations', '2.3', 'Solve multi-step linear equations'],
    ['Unit 3: Linear Functions', '3.1', 'Identify slope and intercept'],
    ['Unit 3: Linear Functions', '3.2', 'Graph linear functions'],
    ['Unit 3: Linear Functions', '3.3', 'Write linear equations from contexts']
  ];
  sh.getRange(2, 1, demo.length, 3).setValues(demo);
}

/**
 * Populate demo grades across students and skills with a variety of cases.
 */
function populateDemoGrades() {
  const ss = SpreadsheetApp.getActive();
  // Ensure base structures
  if (!ss.getSheetByName('Students')) setupStudents();
  if (!ss.getSheetByName('Skills')) setupSkills();
  if (!ss.getSheetByName('Grades')) setupGradesSheet();

  // Build Student x Skill grid in Grades
  if (typeof populateGrades === 'function') populateGrades();

  // Read settings for attempts
  const codes = ss.getRangeByName(RANGE_LEVEL_SHORTCODES).getValues().flat().filter(String);
  const defaultAttempts = ss.getRangeByName(RANGE_LEVEL_DEFAULTATTEMPTS).getValues().flat().slice(0, codes.length).map(n => Number(n || 0));

  const gradesSh = ss.getSheetByName('Grades');
  const lastRow = gradesSh.getLastRow();
  if (lastRow < 2) return;

  // Build symbol pools from Symbols sheet so we exercise everything teachers might enter
  const symbolChars = ss.getRangeByName(RANGE_SYMBOL_CHARS).getValues().flat().map(s => String(s || '').trim());
  const symbolMastery = ss.getRangeByName(RANGE_SYMBOL_MASTERY).getValues().flat().slice(0, symbolChars.length).map(n => Number(n || 0));
  const masterySymbols = [];
  const failSymbols = [];
  for (let i = 0; i < symbolChars.length; i++) {
    const ch = symbolChars[i];
    if (!ch) continue;
    if (symbolMastery[i] === 1) masterySymbols.push(ch); else failSymbols.push(ch);
  }
  if (masterySymbols.length === 0) masterySymbols.push('1');
  if (failSymbols.length === 0) failSymbols.push('X');

  // Determine attempt columns region by header scan
  const headerRow = gradesSh.getRange(1, 1, 1, gradesSh.getLastColumn()).getValues()[0];
  const firstAttemptCol = headerRow.findIndex(h => /^([A-Za-z])1$/.test(String(h || ''))) + 1; // like B1, I1, M1
  if (firstAttemptCol <= 0) return;
  const attemptWidth = gradesSh.getLastColumn() - firstAttemptCol + 1;

  // Offsets for each level's attempts across the row
  const levelOffsets = [];
  let off = 0;
  for (let li = 0; li < codes.length; li++) {
    const len = defaultAttempts[li] || 0;
    levelOffsets.push({ code: String(codes[li]), start: off, len });
    off += len;
  }

  // Per-student mastery probabilities by level code (noise added later per task)
  const probsByStudent = {
    'Alice Johnson': { B: 0.85, I: 0.75, M: 0.65 },
    'Ben Rivera': { B: 0.55, I: 0.45, M: 0.30 },
    'Chloe Kim': { B: 0.15, I: 0.10, M: 0.05 },
    'Diego Patel': { B: 0.70, I: 0.55, M: 0.40 },
    'Emi Sato': { B: 0.60, I: 0.60, M: 0.55 },
    __default__: { B: 0.55, I: 0.45, M: 0.35, __fallback: 0.40 }
  };

  // Cycle through symbol pools so all variants show up
  let masteryIdx = 0, failIdx = 0;
  function pickSymbol(isMastery) {
    if (isMastery) { const s = masterySymbols[masteryIdx % masterySymbols.length]; masteryIdx++; return s; }
    const s = failSymbols[failIdx % failSymbols.length]; failIdx++; return s;
  }

  // Fill attempt cells per row with randomized outcomes; track which symbols we used
  const rowCount = lastRow - 1;
  const allRowValues = [];
  const usedSymbols = Object.create(null);
  function markUsed(sym) { if (sym) usedSymbols[sym] = (usedSymbols[sym] || 0) + 1; }

  // helper: random integer inclusive
  function randInt(min, max) { return Math.floor(Math.random() * (max - min + 1)) + min; }

  for (let r = 2; r <= lastRow; r++) {
    const studentName = String(gradesSh.getRange(r, 1).getValue() || '');
    const base = probsByStudent[studentName] || probsByStudent.__default__;
    const rowVals = new Array(attemptWidth).fill('');

    var allowNextLevel = true; // start with first level permitted
    for (let i = 0; i < levelOffsets.length; i++) {
      const lvl = levelOffsets[i];
      if (!lvl.len) continue;
      // Gate: only proceed if previous level had some success; 10% chance to still proceed
      if (!allowNextLevel && Math.random() > 0.10) {
        // Leave this level blank
        continue;
      }

      // Jitter probability per skill/task so the sheet looks more “real”
      var baseVal = (base.hasOwnProperty(lvl.code) && base[lvl.code] != null) ? base[lvl.code]
        : (base.hasOwnProperty('__fallback') && base.__fallback != null) ? base.__fallback : 0.4;
      const baseP = Number(baseVal);
      const jitter = (Math.random() - 0.5) * 0.20; // ±10%
      const p = Math.max(0.02, Math.min(0.98, baseP + jitter));

      // Only fill a subset of attempts (typically 2–4), bounded by available attempts
      var minFill = Math.min(2, lvl.len);
      var maxFill = Math.min(4, lvl.len);
      var fillCount = (maxFill >= minFill) ? randInt(minFill, maxFill) : lvl.len;

      var successCount = 0;
      for (let k = 0; k < lvl.len; k++) {
        if (k >= fillCount) { rowVals[lvl.start + k] = ''; continue; }
        // Small chance to skip an attempt entirely
        if (Math.random() < 0.05) { rowVals[lvl.start + k] = ''; continue; }
        const isMastery = Math.random() < p;
        const sym = pickSymbol(isMastery);
        rowVals[lvl.start + k] = sym;
        if (isMastery) successCount++;
        markUsed(sym);
      }

      // Must have at least one success to unlock the next level
      allowNextLevel = successCount > 0;
    }
    allRowValues.push(rowVals);
  }

  // Write entire attempt area in one go for speed (only if there are attempt columns)
  if (attemptWidth > 0 && rowCount > 0) {
    gradesSh.getRange(2, firstAttemptCol, rowCount, attemptWidth).setValues(allRowValues);
  }

  // Coverage pass: ensure every symbol from the Symbols sheet appears at least once
  const ensureSymbols = masterySymbols.concat(failSymbols).filter(Boolean);
  if (ensureSymbols.length && attemptWidth > 0 && rowCount > 0) {
    let rr = 2, cc = firstAttemptCol;
    for (const sym of ensureSymbols) {
      if (!usedSymbols[sym]) {
        gradesSh.getRange(rr, cc).setValue(sym);
        markUsed(sym);
        cc++;
        if (cc > firstAttemptCol + attemptWidth - 1) { cc = firstAttemptCol; rr++; if (rr > lastRow) rr = 2; }
      }
    }
  }

  // Make sure formats are correct (text) so digits don’t coerce
  try { repairTextFormats(); } catch (e) { if (console && console.warn) console.warn('repairTextFormats() failed after populateDemoGrades', e); }
}

/**
 * One-click demo setup: named ranges, student/skills, grade sheet, populate grid, then demo attempts.
 */
function runDemoSetup() {
  const ss = SpreadsheetApp.getActive();
  setupNamedRanges();
  setupStudents();
  setupSkills();
  setupGradesSheet();
  // Populate grid (rows)
  if (typeof populateGrades === 'function') populateGrades();
  // Demo values
  populateDemoStudents();
  populateDemoSkills();
  // Rebuild grid now that demo data exists
  if (typeof populateGrades === 'function') populateGrades();
  // Fill sample attempts
  populateDemoGrades();
  // Build Grade View tab
  setupGradeViewSheet();
  // Ensure key entry columns are plain text
  try { repairTextFormats(); } catch (e) { /* best-effort */ }
  ss.toast('Demo setup complete.', 'Standards-Based Grading', 5);
}

/**
 * Advanced: Re-apply plain text formats to key entry columns to avoid numeric coercion.
 */
function repairTextFormats() {
  const ss = SpreadsheetApp.getActive();
  // Symbols Character and Display columns
  const sym = ss.getSheetByName('Symbols');
  if (sym) {
    try { sym.getRange('A:A').setNumberFormat('@'); } catch (e) { if (console && console.warn) console.warn('Symbols A text format warn', e); }
    try { sym.getRange('C:C').setNumberFormat('@'); } catch (e) { if (console && console.warn) console.warn('Symbols C text format warn', e); }
  }
  // Grades attempt columns
  const grades = ss.getSheetByName('Grades');
  if (grades) {
    const header = grades.getRange(1, 1, 1, grades.getLastColumn()).getValues()[0];
    const firstAttemptCol = header.findIndex(h => /^([A-Za-z])1$/.test(String(h || ''))) + 1;
    if (firstAttemptCol > 0) {
      const width = grades.getLastColumn() - firstAttemptCol + 1;
      try { grades.getRange(1, firstAttemptCol, grades.getMaxRows(), width).setNumberFormat('@'); } catch (e) { if (console && console.warn) console.warn('Grades attempts text format warn', e); }
    }
  }
  ss.toast('Reapplied plain-text formats to Symbols and attempt columns.', 'Standards-Based Grading', 3);
}
