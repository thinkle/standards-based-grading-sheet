/* AspenGrades.js Last Update 2025-09-13 08:36 <d0c7c6821168ce61173cbd56b1817e42baf5239eecd2b097351d0fbdc59fa19f>
// filepath: /Users/thinkle/BackedUpProjects/gas/standards-based-grading-sheet/AspenGrades.js

/* Sheet and manager code for Aspen grades */

// Aspen Grades Sheet
const ASPEN_GRADES_HEADERS = {
  studentId: 'Student ID',
  assignmentId: 'Assignment ID',
  score: 'Score',
  dateSynced: 'Date Synced'
};

const ASPEN_GRADES_COLS = [
  ASPEN_GRADES_HEADERS.studentId,
  ASPEN_GRADES_HEADERS.assignmentId,
  ASPEN_GRADES_HEADERS.score,
  ASPEN_GRADES_HEADERS.dateSynced
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
 * Records a synced grade
 * @param {string} studentId - Aspen student ID
 * @param {string} assignmentId - Aspen assignment ID
 * @param {number} score - Grade score
 */
function recordSyncedGrade(studentId, assignmentId, score) {
  const sheet = getAspenGradesSheet();
  const dateSynced = new Date();

  // Check if grade already exists
  const data = sheet.getDataRange().getValues();
  const studentIdCol = getColumnIndex(ASPEN_GRADES_COLS, ASPEN_GRADES_HEADERS.studentId);
  const assignmentIdCol = getColumnIndex(ASPEN_GRADES_COLS, ASPEN_GRADES_HEADERS.assignmentId);

  for (let i = 1; i < data.length; i++) {
    if (data[i][studentIdCol] === studentId && data[i][assignmentIdCol] === assignmentId) {
      // Update existing grade
      sheet.getRange(i + 1, 1, 1, ASPEN_GRADES_COLS.length).setValues([[studentId, assignmentId, score, dateSynced]]);
      return;
    }
  }

  // Add new grade record
  sheet.appendRow([studentId, assignmentId, score, dateSynced]);
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
  const dateSyncedCol = getColumnIndex(ASPEN_GRADES_COLS, ASPEN_GRADES_HEADERS.dateSynced);

  for (let i = 1; i < data.length; i++) {
    if (data[i][studentIdCol] === studentId && data[i][assignmentIdCol] === assignmentId) {
      return {
        studentId: data[i][studentIdCol],
        assignmentId: data[i][assignmentIdCol],
        score: data[i][scoreCol],
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
    const dateSyncedCol = getColumnIndex(ASPEN_GRADES_COLS, ASPEN_GRADES_HEADERS.dateSynced);

    for (let i = 1; i < data.length; i++) {
      const grade = {
        studentId: data[i][studentIdCol],
        assignmentId: data[i][assignmentIdCol],
        score: data[i][scoreCol],
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
   * @param {string} studentEmail - Student email
   * @param {string} assignmentId - Assignment ID
   * @param {number} score - Current score
   * @returns {boolean} True if sync is needed
   */
  needsSync(studentEmail, assignmentId, score) {
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

    return syncedGrade.score !== score; // Score has changed
  }

  /**
   * Syncs a grade if needed (combines check + sync)
   * @param {string} studentEmail - Student email
   * @param {string} assignmentId - Assignment ID  
   * @param {number} score - Current score
   * @returns {Object} Sync result
   */
  maybeSync(studentEmail, assignmentId, score) {
    if (!this.needsSync(studentEmail, assignmentId, score)) {
      return {
        success: true,
        synced: false,
        message: 'No sync needed - score unchanged'
      };
    }

    return this.syncGrade(studentEmail, assignmentId, score);
  }

  /**
   * Syncs a grade to Aspen (assumes data is already loaded)
   * @param {string} studentEmail - Student email
   * @param {string} assignmentId - Assignment ID
   * @param {number} score - Current score
   * @returns {Object} Sync result
   */
  syncGrade(studentEmail, assignmentId, score) {
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
      const result = postGrade(resultId, assignment.assignmentData, student.studentData, score, '');

      // Update our in-memory lookup immediately
      const key = `${student.studentId}_${assignmentId}`;
      const syncedGrade = {
        studentId: student.studentId,
        assignmentId: assignmentId,
        score: score,
        dateSynced: new Date()
      };
      this.syncedGradesLookup[key] = syncedGrade;

      // Record the synced grade to sheet (write immediately for safety)
      recordSyncedGrade(student.studentId, assignmentId, score);

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