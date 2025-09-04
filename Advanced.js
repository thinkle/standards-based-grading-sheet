/* Advanced.js Last Update 2025-08-27 13:36:54 <603e14dfd250495677e7609c21f246bae9f01744bc71fdd7551e0c700f488af6>
/* eslint-disable no-unused-vars */
/* exported nukeAllSheetsExceptInstructions, populateDemoStudents, populateDemoSkills, populateDemoGrades, runDemoSetup */
/* exported repairTextFormats */
/* global SpreadsheetApp, Browser, Menu, setupNamedRanges, setupStudents, setupSkills, setupGradeSheet, setupGradesSheet, populateGrades, setupGradeViewSheet,
          RANGE_STUDENT_NAMES, RANGE_STUDENT_EMAILS, RANGE_SKILL_UNITS, RANGE_SKILL_NUMBERS, RANGE_SKILL_DESCRIPTORS,
          RANGE_LEVEL_NAMES, RANGE_LEVEL_SHORTCODES, RANGE_LEVEL_DEFAULTATTEMPTS, RANGE_LEVEL_STREAK, RANGE_LEVEL_SCORES,
          RANGE_NONE_CORRECT_SCORE, RANGE_SOME_CORRECT_SCORE,
          RANGE_SYMBOL_CHARS, RANGE_SYMBOL_MASTERY, RANGE_SYMBOL_SYMBOL,
          createUnitOverview
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
    ['Emi Sato', 'emi@example.com'],
    ['Frankie Lee', 'frankie@example.com'],
    ['Grace Chen', 'grace@example.com'],
    ['Maria da Silva', 'maria@example.com'],
    ['Noah Brown', 'noah@example.com'],
    ['Olivia Martinez', 'olivia@example.com'],
    ['Luzdivina Rodriguez', 'luzdivina@example.com'],
    ['Mia Wong', 'mia@example.com'],
    ['Nguyen Van An', 'nguyen@example.com'],
    ['Khang Nguyen', 'khang@example.com'],
    ['Hana Lee', 'hana@example.com']
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
    ['Unit 3: Linear Functions', '3.3', 'Write linear equations from contexts'],
    ['Unit 4: Quadratic Functions', '4.1', 'Identify quadratic functions'],
    ['Unit 4: Quadratic Functions', '4.2', 'Graph quadratic functions'],
    ['Unit 4: Quadratic Functions', '4.3', 'Solve quadratic equations'],
    ['Unit 5: Factoring', '5.1', 'Identify factors'],
    ['Unit 5: Factoring', '5.2', 'Factor trinomials'],
    ['Unit 5: Factoring', '5.3', 'Factor by grouping'],
    ['Unit 6: Rational Expressions', '6.1', 'Simplify rational expressions'],
    ['Unit 6: Rational Expressions', '6.2', 'Multiply rational expressions'],
    ['Unit 6: Rational Expressions', '6.3', 'Divide rational expressions']
  ];
  sh.getRange(2, 1, demo.length, 3).setValues(demo);
}

/**
 * Populate demo grades across students and skills with a variety of cases.
 */
function populateDemoGrades() {
  const absoluteStart = new Date();
  const ss = SpreadsheetApp.getActive();
  // Ensure base structures

  if (!ss.getSheetByName('Students')) setupStudents();
  if (!ss.getSheetByName('Skills')) setupSkills();
  if (!ss.getSheetByName('Grades')) setupGradesSheet();
  console.log(`Set up any necessary missing sheets, took ${new Date() - absoluteStart}ms`);

  // Build Student x Skill grid in Grades
  //if (typeof populateGrades === 'function') populateGrades();

  // Read settings for attempts
  let start = new Date();
  const codes = ss.getRangeByName(RANGE_LEVEL_SHORTCODES).getValues().flat().filter(String);
  const defaultAttempts = ss.getRangeByName(RANGE_LEVEL_DEFAULTATTEMPTS).getValues().flat().slice(0, codes.length).map(n => Number(n || 0));
  console.log(`Read level settings, took ${new Date() - start}ms`);
  const gradesSheet = ss.getSheetByName('Grades');
  const lastRow = gradesSheet.getLastRow();
  if (lastRow < 2) return;

  // Build symbol pools from Symbols sheet so we exercise everything teachers might enter
  const symbolChars = ss.getRangeByName(RANGE_SYMBOL_CHARS).getValues().flat().map(s => String(s || '').trim());
  const symbolMastery = ss.getRangeByName(RANGE_SYMBOL_MASTERY).getValues().flat().slice(0, symbolChars.length).map(n => Number(n || 0));
  console.log(`Done reading symbols, now reading has taken ${new Date() - start}ms`);
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
  const headerRow = gradesSheet.getRange(1, 1, 1, gradesSheet.getLastColumn()).getValues()[0];
  const firstAttemptCol = headerRow.findIndex(h => /^([A-Za-z])1$/.test(String(h || ''))) + 1; // like B1, I1, M1
  if (firstAttemptCol <= 0) return;
  const attemptWidth = gradesSheet.getLastColumn() - firstAttemptCol + 1;

  // Offsets for each level's attempts across the row
  const levelOffsets = [];
  let off = 0;
  for (let li = 0; li < codes.length; li++) {
    const len = defaultAttempts[li] || 0;
    levelOffsets.push({ code: String(codes[li]), start: off, len });
    off += len;
  }
  console.log('Finished reading header data, etc total time so far is ', new Date() - absoluteStart);
  let generatingGradeStart = new Date();
  // Per-student mastery probabilities by level code (noise added later per task)
  const probsByStudent = {
    'Alice Johnson': { B: 0.85, I: 0.75, M: 0.65 },
    'Ben Rivera': { B: 0.55, I: 0.45, M: 0.30 },
    'Chloe Kim': { B: 0.15, I: 0.10, M: 0.05 },
    'Diego Patel': { B: 0.70, I: 0.55, M: 0.40 },
    'Emi Sato': { B: 0.60, I: 0.60, M: 0.55 },
    'Franki Lee': { B: .95, I: 0.85, M: 0.75 },
    'Grace Chen': { B: 0.80, I: 0.80, M: 0.20 },
    'Hana Lee': { B: 0.90, I: 0.70, M: 0.60 },
    'Maria da Silva': { B: 0.7, I: 0.85, M: 0.9 },
    'Luzdivina Rodriguez': { B: 1, I: 0.9, M: 0.7 },
    'Mia Wong': { B: 0.6, I: 0.5, M: 0.4 },
    'Nguyen Van An': { B: 0.8, I: 0.7, M: 0.6 },
    'Khang Nguyen': { B: 0.9, I: 0.8, M: 0.7 },
    'Olivia Martinez': { B: 0.75, I: 0.85, M: 0.95 },
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

  // Batch-read columns A..E (Name, Email, Unit, Skill #, Descriptor) so we can skip empty rows
  const rowsData = gradesSheet.getRange(2, 1, rowCount, 5).getValues();
  // Progress instrumentation
  const runStart = Date.now();
  let lastLog = runStart;
  const progressInterval = Math.max(100, Math.floor(rowCount / 10)); // log at least every ~10% or every 100 rows

  if (console && console.log) console.log(`Filling demo grades: ${rowCount} rows to process...`);

  for (let r = 2; r <= lastRow; r++) {
    const row = rowsData[r - 2] || [];
    const studentName = String(row[0] || '').trim();
    const unitVal = String(row[2] || '').trim();
    // Skip rows without a student name or without a skill unit (these are not valid grade rows)

    const base = probsByStudent[studentName] || probsByStudent.__default__;
    const rowVals = new Array(attemptWidth).fill('');
    if (!studentName || !unitVal) {
      allRowValues.push(rowVals);
      continue;
    }

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

    // Periodic progress log
    if ((r - 1) % progressInterval === 0) {
      const now = Date.now();
      const elapsed = now - runStart;
      const sinceLast = now - lastLog;
      lastLog = now;
      if (console && console.log) console.log(`populateDemoGrades: processed ${r - 1}/${rowCount} rows (elapsed ${Math.round(elapsed / 1000)}s, +${Math.round(sinceLast / 1000)}s)`);
    }
  }
  console.log('Done generating grades - generating time is ', new Date() - generatingGradeStart);
  start = new Date();
  // Write entire attempt area in one go for speed (only if there are attempt columns)
  if (attemptWidth > 0 && rowCount > 0) {
    if (console && console.log) console.log(`Writing ${rowCount}×${attemptWidth} attempt cells to sheet...`);
    const writeStart = Date.now();
    try {
      gradesSheet.getRange(2, firstAttemptCol, rowCount, attemptWidth).setValues(allRowValues);
      if (console && console.log) console.log(`Write complete (${Math.round((Date.now() - writeStart) / 1000)}s).`);
    } catch (e) {
      if (console && console.log) console.log(`Error writing attempts: ${e && e.message}`);
      throw e;
    }
  }
  console.log('Pushed grades to sheet in ', new Date() - start);
  start = new Date();
  // Replace the "Coverage" pass with structured test cases
  function populateTestCases() {
    const ss = SpreadsheetApp.getActive();
    const gradesTestSheet = ss.getSheetByName('Grades');
    if (!gradesTestSheet) return;

    const testCases = [
      { description: 'Symbol Coverage', attempts: ['1', '1o', '1s', 'X', 'Xo', 'Xs', 'P', 'G', 'H', 'N'] },
      {
        description: 'Empty Set', attempts: [
          '', '', '', '', '',
          '', '', '', '', '',
          '', '', '', '', '',
        ]
      },
      { description: 'Should get 0s', attempts: ['x', 'x', 'x', 'x', 'x'] },
      { description: 'Should get 0s', attempts: ['x', 'P', 'G', 'H', 'N'] },
      { description: 'Should get 0s', attempts: ['x', 'x', 'x', 'x', 'x', 'x', 'x', 'x', 'x', 'x'] },
      { description: 'Should get 0s', attempts: ['x', 'x', 'x', 'x', 'x', 'x', 'x', 'x', 'x', 'x', 'x', 'x', 'x'] },
      { description: 'Should get 1s', attempts: ['x', '1', 'x', '1o', 'x'] },
      { description: 'Should get 1s', attempts: ['1s', 'x', '1', 'x', '1o', 'x', 'x', '1', 'x', '1o', 'x'] },
      { description: 'Should get 2s', attempts: ['x', 'x', '1', '1', 'x'] },
      {
        description: 'Should get 2s', attempts: ['1o', '1s', '1', '1', 'x',
          'x', '1', 'x', '1', 'x',
        ]
      },
      {
        description: 'Should get 2s', attempts: ['x', 'x', 'x', '1', '1o',
          'x', '1', 'x', '1', 'x',
          'N', '1', 'N', '1', 'N',
        ]
      },
      {
        description: 'Should get 3s', attempts: [
          '1', '1', 'x', '1', 'x',
          '1', '1', 'x', '1', 'x',
        ]
      }, {
        description: 'Should get 3s', attempts: [
          'x', '1', 'x', '1', 'x',
          'x', '1', 'x', '1', '1s',
          'N', '1', 'N', '1', 'N',
        ]
      },
      {
        description: 'Should get 4s', attempts: [
          '1', '1', '1', '1', '1',
          '1', '1', '1', '1', '1',
          '1', '1', '1', '1', '1'
        ]
      },
      {
        description: 'Should get 4s', attempts: [
          '', '', '', '', '',
          '1', 'x', 'G', '1s', 'x',
          'H', 'G', 'P', '1o', '1s'
        ]
      },
      {
        description: 'Should get 4s', attempts: [
          'N', 'N', 'N', 'H', 'G',
          'N', 'x', 'G', '1s', '1',
          'H', 'G', 'P', '1o', '1s'
        ]
      },
    ];

    const firstAttemptCol = gradesTestSheet.getRange(1, 1, 1, gradesTestSheet.getLastColumn())
      .getValues()[0]
      .findIndex(h => /^([A-Za-z])1$/.test(String(h || ''))) + 1;
    if (firstAttemptCol <= 0) return;

    const attemptWidth = gradesTestSheet.getLastColumn() - firstAttemptCol + 1;
    const rowCount = testCases.length;

    // Populate test cases in the first n rows
    testCases.forEach((testCase, index) => {
      const row = index + 2; // Start from row 2
      const rowValues = new Array(attemptWidth).fill('');
      testCase.attempts.forEach((value, colIndex) => {
        if (colIndex < attemptWidth) {
          rowValues[colIndex] = value;
        }
      });
      gradesTestSheet.getRange(row, firstAttemptCol, 1, attemptWidth).setValues([rowValues]);
    });

    console.log(`Populated ${rowCount} test cases.`);
  }

  // Call the new function in place of the "Coverage" pass
  const gradesDemoSheet = ss.getSheetByName('Grades');
  if (!gradesDemoSheet) return;

  // Replace the "Coverage" pass with test cases
  populateTestCases();

  // Ensure key entry columns are plain text
  try { repairTextFormats(); } catch (e) { if (console && console.warn) console.warn('repairTextFormats() failed after populateDemoGrades', e); }
}

/**
 * One-click demo setup: named ranges, student/skills, grade sheet, populate grid, then demo attempts.
 */

function timer(stamp) {
  const now = Date.now();
  const elapsed = now - stamp;
  return { now, elapsed };
}

function toastAndLog(ss, message) {
  if (ss && typeof ss.toast === 'function') ss.toast(message, 'Standards-Based Grading', 3);
  if (console && console.log) console.log(message);
}

function runDemoSetup() {
  const ss = SpreadsheetApp.getActive();
  let startTime = Date.now();
  setupNamedRanges();
  let { now, elapsed } = timer(startTime);
  toastAndLog(ss, `Set up settings (${elapsed}ms)`);
  setupStudents();
  setupSkills();
  ({ now, elapsed } = timer(now));
  toastAndLog(ss, `Set up student and skills sheets (${elapsed}ms)`);
  setupGradesSheet();
  ({ now, elapsed } = timer(now));
  toastAndLog(ss, `Set up grades sheet (${elapsed}ms)`);
  createUnitOverview();
  ({ now, elapsed } = timer(now));
  toastAndLog(ss, `Set up unit overview (${elapsed}ms)`);
  // Demo values
  populateDemoStudents();
  populateDemoSkills();
  ({ now, elapsed } = timer(now));
  toastAndLog(ss, `Populated demo students and skills (${elapsed}ms)`);
  // Populate grid (rows)
  if (typeof populateGrades === 'function') populateGrades();
  ({ now, elapsed } = timer(now));
  toastAndLog(ss, `Set up grade grid for students+assignments (${elapsed}ms)`);
  // Fill sample attempts
  populateDemoGrades();
  ({ now, elapsed } = timer(now));
  toastAndLog(ss, `Filled in demo grades (${elapsed}ms)`);
  // Build Grade View tab
  setupGradeViewSheet();
  ({ now, elapsed } = timer(now));
  toastAndLog(ss, `Set up demo sheet (${elapsed}ms)`);
  // Ensure key entry columns are plain text
  try { repairTextFormats(); } catch (e) { /* best-effort */ }
  let totalElapsed = Date.now() - startTime;
  toastAndLog(ss, `Demo setup complete. (${totalElapsed}ms)`, 'Standards-Based Grading', 5);
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
