/* Menu.js Last Update 2025-09-13 11:35 <9916259cb9425f7d1fa7c010699c5cdd4681405455cf343af037c496bc44a28f>
/* eslint-disable no-unused-vars */
/* exported onOpen, doInitialSetup, setupGradeSheet, addStudentsAndSkills, reformatGradesOnly,
  addSkillsToAspenAssignmentsMenu, addUnitAveragesToAspenAssignmentsMenu, createMissingAspenAssignmentsMenu */
/* global SpreadsheetApp, setupNamedRanges, setupStudents, setupSkills, setupGradesSheet, populateGrades, reformatGradesOnly,
writePostSetupInstructions, addAllSkillsToAspenAssignments, addUnitAveragesToAspenAssignments, createMissingAssignmentsFromSheet, reapplyAspenAssignmentsValidation */

/**
 * Adds a custom menu to the spreadsheet.
 */

const MENU = {
  TITLE: 'Standards-Based Grading',
  SETUP: 'Initial setup',
  REGENERATE_INSTRUCTIONS: 'Generate Instructions Sheet',
  SETUP_GRADE_SHEET: 'Setup Grade Sheet (Editable)',
  REFORMAT_GRADES: 'Reformat Grades (no contents will change)',
  ADD_STUDENTS_AND_SKILLS: 'Add Students & Skills (Rerun when adding new students/skills)',
  SETUP_GRADE_VIEW: 'Setup Grade View (Read Only)',
  GENERATE_STUDENT_VIEWS: 'Generate student views',
  SHARE_STUDENT_VIEWS: 'Share student views',
  GENERATE_UNIT_OVERVIEW: 'Create Unit Overview (Averages)',
  // Aspen Integration
  ASPEN_SETUP: 'Setup Aspen Integration...',
  ASPEN_TEST: 'Test Aspen Connection',
  ASPEN_STATUS: 'Show Aspen Status',
  ASPEN_CREATE_ASSIGNMENT: 'Create Assignment...',
  ASPEN_SYNC_GRADES: 'Sync Grades...',
  ASPEN_ADD_SKILLS_TO_ASSIGNMENTS: 'Add Skills to Assignments Tab',
  ASPEN_ADD_UNIT_AVERAGES: 'Add Unit Averages to Assignments Tab',
  ASPEN_CREATE_MISSING_ASSIGNMENTS: 'Create Missing Assignments (from due dates)'
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const main = ui.createMenu(MENU.TITLE)
    .addItem(MENU.REGENERATE_INSTRUCTIONS, 'writePostSetupInstructions') // dev only
    .addItem(MENU.SETUP, 'doInitialSetup')
    .addItem(MENU.SETUP_GRADE_SHEET, 'setupGradeSheet')
    .addSeparator()
    .addItem(MENU.ADD_STUDENTS_AND_SKILLS, 'addStudentsAndSkills')
    .addSeparator()
    .addItem(MENU.REFORMAT_GRADES, 'reformatGradesSheet')
    .addSeparator()
    .addItem(MENU.SETUP_GRADE_VIEW, 'setupGradeViewSheet')
    .addItem(MENU.GENERATE_UNIT_OVERVIEW, 'createUnitOverview')
    .addItem(MENU.GENERATE_STUDENT_VIEWS, 'generateStudentViews')
    .addItem(MENU.SHARE_STUDENT_VIEWS, 'shareStudentViews')
    .addSeparator()
    .addSubMenu(ui.createMenu('Aspen Integration')
      .addItem(MENU.ASPEN_SETUP, 'setupAspenIntegration')
      .addSeparator()
      .addItem(MENU.ASPEN_TEST, 'testAspenConnection')
      .addItem(MENU.ASPEN_STATUS, 'showAspenStatus')
      .addSeparator()
      .addItem(MENU.ASPEN_ADD_SKILLS_TO_ASSIGNMENTS, 'addSkillsToAspenAssignmentsMenu')
      .addItem(MENU.ASPEN_ADD_UNIT_AVERAGES, 'addUnitAveragesToAspenAssignmentsMenu')
      .addItem(MENU.ASPEN_CREATE_MISSING_ASSIGNMENTS, 'createMissingAspenAssignmentsMenu')
      .addItem('Reapply Assignments Validation', 'reapplyAspenAssignmentsValidation')
      .addSeparator()
      .addItem(MENU.ASPEN_CREATE_ASSIGNMENT, 'createAspenAssignmentUI')
      .addItem(MENU.ASPEN_SYNC_GRADES, 'syncGradesUI'))
    ;



  // Advanced submenu
  const adv = ui.createMenu('Advanced/Testing')
    .addItem('Nuke! (delete all sheets except Instructions)', 'nukeAllSheetsExceptInstructions')
    .addSeparator()
    .addItem('Populate Demo Students', 'populateDemoStudents')
    .addItem('Populate Demo Skills', 'populateDemoSkills')
    .addItem('Populate Demo Grades', 'populateDemoGrades')
    .addItem('Repair Text Formats', 'repairTextFormats')
    .addItem('Reinsert Formulas', 'reinsertGradesFormulas')
    .addSeparator()
    .addItem('Set Up Demo (all-in-one)', 'runDemoSetup');


  main.addSubMenu(adv).addToUi();
}

/**
 * Initial one-time setup: named ranges + source sheets.
 */
function doInitialSetup() {
  const ss = SpreadsheetApp.getActive();
  // Settings and symbols/levels
  setupNamedRanges();
  // Source sheets
  setupStudents();
  setupSkills();
  writePostSetupInstructions();
  ss.toast('Initial setup complete.', 'Standards-Based Grading', 3);
}

/**
 * Build or rebuild the Grades sheet from current settings.
 */
function setupGradeSheet() {
  const ss = SpreadsheetApp.getActive();
  if (typeof setupGradesSheet === 'function') setupGradesSheet();
  ss.toast('Grades sheet (headers, formulas, formatting) set up.', 'Standards-Based Grading', 3);
}

/**
 * Add any missing Student x Skill rows to Grades and fill formulas.
 */
function addStudentsAndSkills() {
  const ss = SpreadsheetApp.getActive();
  if (typeof populateGrades === 'function') populateGrades();
  ss.toast('Added any missing Student × Skill rows and filled formulas.', 'Standards-Based Grading', 3);
}

// ---------------- Aspen menu helpers ----------------

/** Add all unique (Unit, Skill) pairs from Grades into the Aspen Assignments tab. */
function addSkillsToAspenAssignmentsMenu() {
  try {
    if (typeof addAllSkillsToAspenAssignments === 'function') addAllSkillsToAspenAssignments();
  } catch (e) {
    if (typeof console !== 'undefined' && console.error) console.error('Menu:addSkillsToAspenAssignmentsMenu error', e);
    SpreadsheetApp.getUi().alert('Error', `Failed to add skills to assignments:\n\n${e}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/** Add one "Unit Average" item per unit into the Aspen Assignments tab. */
function addUnitAveragesToAspenAssignmentsMenu() {
  try {
    if (typeof addUnitAveragesToAspenAssignments === 'function') addUnitAveragesToAspenAssignments();
  } catch (e) {
    if (typeof console !== 'undefined' && console.error) console.error('Menu:addUnitAveragesToAspenAssignmentsMenu error', e);
    SpreadsheetApp.getUi().alert('Error', `Failed to add unit averages:\n\n${e}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Create missing Aspen assignments for rows with due dates but no Aspen ID.
 * Uses the first configured class from the "Aspen Config" sheet if multiple exist, prompts when none.
 */
function createMissingAspenAssignmentsMenu() {
  try {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActive();
    const cfg = ss.getSheetByName('Aspen Config');
    if (!cfg || cfg.getLastRow() < 2) {
      ui.alert('Aspen', 'No configured class found. Run "Setup Aspen Integration" first.', ui.ButtonSet.OK);
      return;
    }
    const vals = cfg.getRange(2, 1, cfg.getLastRow() - 1, Math.max(2, cfg.getLastColumn())).getValues();
    const classes = vals.map(r => ({ id: String(r[0] || '').trim(), name: String(r[1] || '').trim() })).filter(c => c.id);
    if (classes.length === 0) {
      ui.alert('Aspen', 'No configured class found. Run "Setup Aspen Integration" first.', ui.ButtonSet.OK);
      return;
    }
    let classId = classes[0].id;
    if (classes.length > 1) {
      // Build a simple selection list
      const list = classes.map((c, i) => `${i + 1}. ${c.name || c.id}  [${c.id}]`).join('\n');
      const res = ui.prompt('Select Class', `Multiple classes configured. Enter the number to target:\n\n${list}`, ui.ButtonSet.OK_CANCEL);
      if (res.getSelectedButton() !== ui.Button.OK) return;
      const n = parseInt(res.getResponseText().trim(), 10);
      if (!isNaN(n) && n >= 1 && n <= classes.length) {
        classId = classes[n - 1].id;
      }
    }
    const result = createMissingAssignmentsFromSheet(classId);
    ui.alert('Aspen', `Assignments created: ${result.created}\nSkipped: ${result.skipped}\nErrors: ${result.errors.length}`, ui.ButtonSet.OK);
  } catch (e) {
    if (typeof console !== 'undefined' && console.error) console.error('Menu:createMissingAspenAssignmentsMenu error', e);
    SpreadsheetApp.getUi().alert('Error', `Failed to create missing assignments:\n\n${e}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
