/* eslint-disable no-unused-vars */
/* exported onOpen, doInitialSetup, setupGradeSheet, addStudentsAndSkills */
/* global SpreadsheetApp, setupNamedRanges, setupStudents, setupSkills, setupGradesSheet, populateGrades */

/**
 * Adds a custom menu to the spreadsheet.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Standards-Based Grading')
    .addItem('Initial setup', 'doInitialSetup')
    .addSeparator()
    .addItem('Setup Grade Sheet', 'setupGradeSheet')
    .addItem('Add Students & Skills', 'addStudentsAndSkills')
    .addItem('Setup Grade View', 'setupGradeViewSheet')
    .addToUi();
}

/**
 * Initial one-time setup: named ranges + source sheets.
 */
function doInitialSetup() {
  const ss = SpreadsheetApp.getActive();
  // Settings and symbols/levels
  if (typeof setupNamedRanges === 'function') setupNamedRanges();
  // Source sheets
  if (typeof setupStudents === 'function') setupStudents();
  if (typeof setupSkills === 'function') setupSkills();
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
