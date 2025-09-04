/* Menu.js Last Update 2025-09-04 16:27 <b9f04010ce719bafa1b7c3eef780c20554bd08f66203c1b165ac97414c5e224e>
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
  REGENERATE_INSTRUCTIONS: 'Generate Instructions Sheet',
  SETUP_GRADE_SHEET: 'Setup Grade Sheet (Editable)',
  REFORMAT_GRADES: 'Reformat Grades (no contents will change)',
  ADD_STUDENTS_AND_SKILLS: 'Add Students & Skills (Rerun when adding new students/skills)',
  SETUP_GRADE_VIEW: 'Setup Grade View (Read Only)',
  GENERATE_STUDENT_VIEWS: 'Generate student views',
  SHARE_STUDENT_VIEWS: 'Share student views',
  GENERATE_UNIT_OVERVIEW: 'Create Unit Overview (Averages)'
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
    .addItem(MENU.SHARE_STUDENT_VIEWS, 'shareStudentViews');


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
