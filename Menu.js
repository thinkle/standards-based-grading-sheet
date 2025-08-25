/* eslint-disable no-unused-vars */
/* exported onOpen, doInitialSetup, setupGradeSheet, addStudentsAndSkills, reformatGradesOnly */
/* global SpreadsheetApp, setupNamedRanges, setupStudents, setupSkills, setupGradesSheet, populateGrades, reformatGradesOnly,
writePostSetupInstructions */

/**
 * Adds a custom menu to the spreadsheet.
 */

const MENU = {
  TITLE: 'Standards-Based Grading',
  SETUP: 'Initial setup',
  REGENERATE_INSTRUCTIONS: 'Regenerate Instructions',
  SETUP_GRADE_SHEET: 'Setup Grade Sheet',
  REFORMAT_GRADES: 'Reformat Grades (no content)',
  ADD_STUDENTS_AND_SKILLS: 'Add Students & Skills',
  SETUP_GRADE_VIEW: 'Setup Grade View',
  GENERATE_STUDENT_VIEWS: 'Generate student views',
  SHARE_STUDENT_VIEWS: 'Share student views',
  GENERATE_UNIT_OVERVIEW: 'Create Unit Overview (Averages)'
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const main = ui.createMenu(MENU.TITLE)
    .addItem(MENU.SETUP, 'doInitialSetup')
    .addItem(MENU.REGENERATE_INSTRUCTIONS, 'writePostSetupInstructions') // dev only
    .addSeparator()
    .addItem(MENU.SETUP_GRADE_SHEET, 'setupGradeSheet')
    .addItem(MENU.REFORMAT_GRADES, 'reformatGradesSheet')
    .addItem(MENU.ADD_STUDENTS_AND_SKILLS, 'addStudentsAndSkills')
    .addItem(MENU.SETUP_GRADE_VIEW, 'setupGradeViewSheet')
    .addItem(MENU.GENERATE_UNIT_OVERVIEW, 'createUnitOverview')
    .addSeparator()
    .addItem(MENU.GENERATE_STUDENT_VIEWS, 'generateStudentViews')
    .addItem(MENU.SHARE_STUDENT_VIEWS, 'shareStudentViews');


  // Advanced submenu
  const adv = ui.createMenu('Advanced')
    .addItem('Nuke! (delete all sheets except Instructions)', 'nukeAllSheetsExceptInstructions')
    .addSeparator()
    .addItem('Populate Demo Students', 'populateDemoStudents')
    .addItem('Populate Demo Skills', 'populateDemoSkills')
    .addItem('Populate Demo Grades', 'populateDemoGrades')
    .addItem('Repair Text Formats', 'repairTextFormats')
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
