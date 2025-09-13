/* AspenGrades.js Last Update 2025-09-13 12:08 <6dc47d71f94e8bfeac41c4812bfa1a10d38d244bbeeba2d90ad4f3227e510151>
// filepath: /Users/thinkle/BackedUpProjects/gas/standards-based-grading-sheet/AspenGrades.js

/* Sheet and manager code for Aspen grades */
/* global SpreadsheetApp, getAspenStudents, getAspenAssignments, postGrade */

// Aspen Grades Sheet
const ASPEN_GRADES_HEADERS = {
  studentEmail: 'Student Email',
  unit: 'Unit',
  skill: 'Skill',
  score: 'Score',
  comment: 'Comment',
  dateSynced: 'Date Synced',
  studentId: 'Student ID',
  assignmentId: 'Assignment ID',
};

const ASPEN_GRADES_COLS = [
  ASPEN_GRADES_HEADERS.studentEmail,
  ASPEN_GRADES_HEADERS.unit,
  ASPEN_GRADES_HEADERS.skill,
  ASPEN_GRADES_HEADERS.score,
  ASPEN_GRADES_HEADERS.comment,
  ASPEN_GRADES_HEADERS.dateSynced,
  ASPEN_GRADES_HEADERS.studentId,
  ASPEN_GRADES_HEADERS.assignmentId
];

/* Helper Functions */
/**
 * Gets column index for a header (0-based)
 * @param {Array} cols - Column array
 * @param {string} header - Header to find
 * @returns {number} Column index
 */
function getColumnIndex(cols, header) {
  return cols.indexOf(header);
}

/**
 * Ensures Unit/Skill lookup formulas exist for a given row.
 * Looks up Unit/Skill from 'Aspen Assignments' based on this row's Assignment ID.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Grades sheet
 * @param {number} row - 1-based row index
 */
function ensureLookupFormulasForRow(sheet, row) {
  const unitCol = getColumnIndex(ASPEN_GRADES_COLS, ASPEN_GRADES_HEADERS.unit) + 1; // 1-based
  const skillCol = getColumnIndex(ASPEN_GRADES_COLS, ASPEN_GRADES_HEADERS.skill) + 1;
  const studentEmailCol = getColumnIndex(ASPEN_GRADES_COLS, ASPEN_GRADES_HEADERS.studentEmail) + 1;
  const assignmentIdCol = getColumnIndex(ASPEN_GRADES_COLS, ASPEN_GRADES_HEADERS.assignmentId) + 1;
  const studentIdCol = getColumnIndex(ASPEN_GRADES_COLS, ASPEN_GRADES_HEADERS.studentId) + 1;

  const assignmentCellA1 = sheet.getRange(row, assignmentIdCol).getA1Notation();
  const studentIdCellA1 = sheet.getRange(row, studentIdCol).getA1Notation();

  // Use FILTER to pull Unit/Skill by matching Aspen ID (column C) in 'Aspen Assignments'
  const unitFormula = `=IF(${assignmentCellA1}="","",IFERROR(INDEX(FILTER('Aspen Assignments'!$A:$A, 'Aspen Assignments'!$C:$C = ${assignmentCellA1}),1),""))`;
  const skillFormula = `=IF(${assignmentCellA1}="","",IFERROR(INDEX(FILTER('Aspen Assignments'!$B:$B, 'Aspen Assignments'!$C:$C = ${assignmentCellA1}),1),""))`;

  // Use FILTER to pull Student Email by matching Student ID (column A) in 'Aspen Students'
  const emailFormula = `=IF(${studentIdCellA1}="","",IFERROR(INDEX(FILTER('Aspen Students'!$C:$C, 'Aspen Students'!$A:$A = ${studentIdCellA1}),1),""))`;

  const unitCell = sheet.getRange(row, unitCol);
  const skillCell = sheet.getRange(row, skillCol);
  const emailCell = sheet.getRange(row, studentEmailCol);

  if (!unitCell.getFormula()) unitCell.setFormula(unitFormula);
  if (!skillCell.getFormula()) skillCell.setFormula(skillFormula);
  if (!emailCell.getFormula()) emailCell.setFormula(emailFormula);
}

/* Aspen Grade Sheet Functions */

/**
 * Gets or creates the Aspen Grades sheet
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getAspenGradesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Aspen Grades');

  if (!sheet) {
    sheet = ss.insertSheet('Aspen Grades');
    // Set up headers using constants
    sheet.getRange(1, 1, 1, ASPEN_GRADES_COLS.length).setValues([ASPEN_GRADES_COLS]);
    sheet.getRange(1, 1, 1, ASPEN_GRADES_COLS.length).setFontWeight('bold');
  }

  return sheet;
}

/**
 * Records a synced grade (derived columns are looked up by formulas)
 * @param {string} studentId - Aspen student ID
 * @param {string} assignmentId - Aspen assignment ID
 * @param {number} score - Grade score
 * @param {string} [comment] - Optional grade comment
 */
function recordSyncedGrade(studentId, assignmentId, score, comment = '') {
  const sheet = getAspenGradesSheet();
  const dateSynced = new Date();

  // Check if grade already exists
  const data = sheet.getDataRange().getValues();
  const studentIdCol = getColumnIndex(ASPEN_GRADES_COLS, ASPEN_GRADES_HEADERS.studentId);
  const assignmentIdCol = getColumnIndex(ASPEN_GRADES_COLS, ASPEN_GRADES_HEADERS.assignmentId);
  const scoreCol = getColumnIndex(ASPEN_GRADES_COLS, ASPEN_GRADES_HEADERS.score);
  const commentCol = getColumnIndex(ASPEN_GRADES_COLS, ASPEN_GRADES_HEADERS.comment);
  const dateSyncedCol = getColumnIndex(ASPEN_GRADES_COLS, ASPEN_GRADES_HEADERS.dateSynced);
  const studentEmailCol = getColumnIndex(ASPEN_GRADES_COLS, ASPEN_GRADES_HEADERS.studentEmail);

  for (let i = 1; i < data.length; i++) {
    if (data[i][studentIdCol] === studentId && data[i][assignmentIdCol] === assignmentId) {
      // Update existing grade: update score, comment, and date; keep formulas intact
      console.time && console.time('gradeSheetUpdate');
      sheet.getRange(i + 1, scoreCol + 1).setValue(score);
      sheet.getRange(i + 1, commentCol + 1).setValue(comment);
      sheet.getRange(i + 1, dateSyncedCol + 1).setValue(dateSynced);
      console.timeEnd && console.timeEnd('gradeSheetUpdate');
      // Ensure lookup formulas are present for this row (in case of older rows)
      ensureLookupFormulasForRow(sheet, i + 1);
      return;
    }
  }

  // Add new grade record
  console.time && console.time('gradeSheetAppend');
  sheet.appendRow([
    '',                  // Student Email (lookup via formula)
    '',                  // Unit (formula added below)
    '',                  // Skill (formula added below)
    score,               // Score
    comment,             // Comment
    dateSynced,          // Date Synced
    studentId,           // Student ID
    assignmentId         // Assignment ID
  ]);
  console.timeEnd && console.timeEnd('gradeSheetAppend');

  // Add lookup formulas for the new row
  const newRow = sheet.getLastRow();
  ensureLookupFormulasForRow(sheet, newRow);
}

/**
 * Gets synced grades for a student and assignment
 * @param {string} studentId - Aspen student ID
 * @param {string} assignmentId - Aspen assignment ID
 * @returns {Object|null} Grade record or null if not found
 */
function getSyncedGrade(studentId, assignmentId) {
  const sheet = getAspenGradesSheet();
  const data = sheet.getDataRange().getValues();
  const studentIdCol = getColumnIndex(ASPEN_GRADES_COLS, ASPEN_GRADES_HEADERS.studentId);
  const assignmentIdCol = getColumnIndex(ASPEN_GRADES_COLS, ASPEN_GRADES_HEADERS.assignmentId);
  const scoreCol = getColumnIndex(ASPEN_GRADES_COLS, ASPEN_GRADES_HEADERS.score);
  const commentCol = getColumnIndex(ASPEN_GRADES_COLS, ASPEN_GRADES_HEADERS.comment);
  const dateSyncedCol = getColumnIndex(ASPEN_GRADES_COLS, ASPEN_GRADES_HEADERS.dateSynced);

  for (let i = 1; i < data.length; i++) {
    if (data[i][studentIdCol] === studentId && data[i][assignmentIdCol] === assignmentId) {
      return {
        studentId: data[i][studentIdCol],
        assignmentId: data[i][assignmentIdCol],
        score: data[i][scoreCol],
        comment: data[i][commentCol],
        dateSynced: data[i][dateSyncedCol]
      };
    }
  }

  return null;
}

/* Grade Sync Manager Class */

/**
 * Grade Sync Manager - handles bulk grade sync operations efficiently
 */
class GradeSyncManager {
  constructor(classId) {
    this.classId = classId;
    this.students = null;
    this.assignments = null;
    this.syncedGrades = null;
    this.studentsLookup = null;
    this.assignmentsLookup = null;
    this.syncedGradesLookup = null;
    this.pendingWrites = [];
  }

  /**
   * Loads all data needed for sync operations
   */
  loadData() {
    console.log('Loading sync data...');

    // Load students and create lookup by email
    this.students = getAspenStudents(this.classId);
    this.studentsLookup = {};
    this.students.forEach(student => {
      this.studentsLookup[student.email] = student;
    });

    // Load assignments and create lookup by ID
    this.assignments = getAspenAssignments();
    this.assignmentsLookup = {};
    this.assignments.forEach(assignment => {
      this.assignmentsLookup[assignment.assignmentId] = assignment;
    });

    // Load synced grades and create lookup by student+assignment
    const sheet = getAspenGradesSheet();
    const data = sheet.getDataRange().getValues();
    this.syncedGrades = [];
    this.syncedGradesLookup = {};

    const studentIdCol = getColumnIndex(ASPEN_GRADES_COLS, ASPEN_GRADES_HEADERS.studentId);
    const assignmentIdCol = getColumnIndex(ASPEN_GRADES_COLS, ASPEN_GRADES_HEADERS.assignmentId);
    const scoreCol = getColumnIndex(ASPEN_GRADES_COLS, ASPEN_GRADES_HEADERS.score);
    const commentCol = getColumnIndex(ASPEN_GRADES_COLS, ASPEN_GRADES_HEADERS.comment);
    const dateSyncedCol = getColumnIndex(ASPEN_GRADES_COLS, ASPEN_GRADES_HEADERS.dateSynced);

    for (let i = 1; i < data.length; i++) {
      const grade = {
        studentId: data[i][studentIdCol],
        assignmentId: data[i][assignmentIdCol],
        score: data[i][scoreCol],
        comment: data[i][commentCol],
        dateSynced: data[i][dateSyncedCol]
      };
      this.syncedGrades.push(grade);
      const key = `${grade.studentId}_${grade.assignmentId}`;
      this.syncedGradesLookup[key] = grade;
    }

    console.log(`Loaded ${this.students.length} students, ${this.assignments.length} assignments, ${this.syncedGrades.length} synced grades`);
  }

  /**
   * Checks if a grade needs syncing (fast lookup operation)
   * @private
   * @param {string} studentEmail - Student email
   * @param {string} assignmentId - Assignment ID
   * @param {number} score - Current score
   * @param {string} [comment] - Optional comment
   * @returns {boolean} True if sync is needed
   */
  _needsSync(studentEmail, assignmentId, score, comment) {
    if (!this.studentsLookup || !this.assignmentsLookup) {
      throw new Error('Data not loaded. Call loadData() first.');
    }

    const student = this.studentsLookup[studentEmail];
    if (!student) {
      return false; // Student not in Aspen class
    }

    const assignment = this.assignmentsLookup[assignmentId];
    if (!assignment) {
      return false; // Assignment not found
    }

    const key = `${student.studentId}_${assignmentId}`;
    const syncedGrade = this.syncedGradesLookup[key];

    if (!syncedGrade) {
      return true; // Never synced before
    }

    const normalizedNewComment = (comment == null) ? '' : String(comment);
    const normalizedOldComment = (syncedGrade.comment == null) ? '' : String(syncedGrade.comment);
    return syncedGrade.score !== score || normalizedOldComment !== normalizedNewComment; // Score or comment has changed
  }

  /**
   * Syncs a grade if needed (combines check + sync)
   * @param {string} studentEmail - Student email
   * @param {string} assignmentId - Assignment ID  
   * @param {number} score - Current score
   * @param {string} [comment] - Optional comment
   * @returns {Object} Sync result
   */
  maybeSync(studentEmail, assignmentId, score, comment) {
    if (!this._needsSync(studentEmail, assignmentId, score, comment)) {
      return {
        success: true,
        synced: false,
        message: 'No sync needed - score unchanged'
      };
    }

    return this._syncGrade(studentEmail, assignmentId, score, comment);
  }

  /**
   * @private
   * Syncs a grade to Aspen (assumes data is already loaded)
   * @param {string} studentEmail - Student email
   * @param {string} assignmentId - Assignment ID
  * @param {number} score - Current score
  * @param {string} [comment] - Optional comment
   * @returns {Object} Sync result
   */
  _syncGrade(studentEmail, assignmentId, score, comment) {
    try {
      const student = this.studentsLookup[studentEmail];
      if (!student) {
        throw new Error(`Student with email ${studentEmail} not found in Aspen class`);
      }

      const assignment = this.assignmentsLookup[assignmentId];
      if (!assignment) {
        throw new Error(`Assignment ${assignmentId} not found`);
      }

      // Generate a unique result ID
      const resultId = `${assignmentId}_${student.studentId}_${Date.now()}`;

      // Post grade to Aspen
      const result = postGrade(resultId, assignment.assignmentData, student.studentData, score, comment || '');

      // Update our in-memory lookup immediately
      const key = `${student.studentId}_${assignmentId}`;
      const syncedGrade = {
        studentId: student.studentId,
        assignmentId: assignmentId,
        score: score,
        comment: comment || '',
        dateSynced: new Date()
      };
      this.syncedGradesLookup[key] = syncedGrade;

      // Record the synced grade to sheet (write immediately for safety)
      recordSyncedGrade(student.studentId, assignmentId, score, comment || '');

      return {
        success: true,
        synced: true,
        message: `Grade synced successfully for ${student.name}`,
        result: result
      };

    } catch (error) {
      console.error('Error syncing grade:', error);
      return {
        success: false,
        synced: false,
        message: error.toString()
      };
    }
  }

  /**
   * Gets sync statistics
   * @returns {Object} Statistics about loaded data
   */
  getStats() {
    return {
      studentsCount: this.students ? this.students.length : 0,
      assignmentsCount: this.assignments ? this.assignments.length : 0,
      syncedGradesCount: this.syncedGrades ? this.syncedGrades.length : 0,
      classId: this.classId
    };
  }

  /**
   * Checks if a student exists in the Aspen class
   * @param {string} studentEmail - Student email
   * @returns {boolean} True if student exists
   */
  hasStudent(studentEmail) {
    return this.studentsLookup && !!this.studentsLookup[studentEmail];
  }

  /**
   * Checks if an assignment exists
   * @param {string} assignmentId - Assignment ID
   * @returns {boolean} True if assignment exists
   */
  hasAssignment(assignmentId) {
    return this.assignmentsLookup && !!this.assignmentsLookup[assignmentId];
  }
}

/**
 * Creates a new Grade Sync Manager for efficient batch operations
 * @param {string} classId - Aspen class ID
 * @returns {GradeSyncManager} Initialized sync manager
 */
function createGradeSyncManager(classId) {
  const manager = new GradeSyncManager(classId);
  manager.loadData();
  return manager;
}