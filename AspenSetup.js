/* AspenSetup.js Last Update 2025-09-13 10:18 <f99f9ee6b6fd07bf8e497c29cd339b7e1e659598e938833cebb300cab9caa298>
/* GradeSync.js Last Update 2025-09-12 16:25 <760306f37429c4f9f913127370e645115e6a2823b1730625957463c82f6202fe>
// filepath: /Users/thinkle/BackedUpProjects/gas/standards-based-grading-sheet/GradeSync.js

/* Main Grade Sync Orchestrator */
// This module coordinates between our spreadsheet data and Aspen grade syncing
// It uses the other Aspen modules: AspenIdGen, AspenAssignments, AspenGrades

/* Sheet Column Definitions for Core Sheets */
// Aspen Config Sheet
const ASPEN_CONFIG_HEADERS = {
  classId: 'Class ID',
  className: 'Class Name',
  categoriesJson: 'Categories JSON',
  dateCreated: 'Date Created'
};

const ASPEN_CONFIG_COLS = [
  ASPEN_CONFIG_HEADERS.classId,
  ASPEN_CONFIG_HEADERS.className,
  ASPEN_CONFIG_HEADERS.categoriesJson,
  ASPEN_CONFIG_HEADERS.dateCreated
];

// Aspen Students Sheet
const ASPEN_STUDENTS_HEADERS = {
  studentId: 'Student ID',
  name: 'Name',
  email: 'Email',
  classId: 'Class ID',
  studentJson: 'Student JSON'
};

const ASPEN_STUDENTS_COLS = [
  ASPEN_STUDENTS_HEADERS.studentId,
  ASPEN_STUDENTS_HEADERS.name,
  ASPEN_STUDENTS_HEADERS.email,
  ASPEN_STUDENTS_HEADERS.classId,
  ASPEN_STUDENTS_HEADERS.studentJson
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

/* Core Aspen Integration Functions */

/**
 * Gets or creates the Aspen Config sheet
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getAspenConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Aspen Config');

  if (!sheet) {
    sheet = ss.insertSheet('Aspen Config');
    sheet.getRange(1, 1, 1, ASPEN_CONFIG_COLS.length).setValues([ASPEN_CONFIG_COLS]);
    sheet.getRange(1, 1, 1, ASPEN_CONFIG_COLS.length).setFontWeight('bold');
  }

  return sheet;
}

/**
 * Stores Aspen class configuration
 * @param {string} classId - Aspen class ID
 * @param {string} className - Aspen class name
 * @param {Array} categories - Array of category objects
 */
function storeAspenClassConfig(classId, className, categories) {
  const sheet = getAspenConfigSheet();
  const categoriesJson = JSON.stringify(categories);
  const dateCreated = new Date();

  const data = sheet.getDataRange().getValues();
  const classIdCol = getColumnIndex(ASPEN_CONFIG_COLS, ASPEN_CONFIG_HEADERS.classId);

  for (let i = 1; i < data.length; i++) {
    if (data[i][classIdCol] === classId) {
      sheet.getRange(i + 1, 1, 1, ASPEN_CONFIG_COLS.length).setValues([[classId, className, categoriesJson, dateCreated]]);
      return;
    }
  }

  sheet.appendRow([classId, className, categoriesJson, dateCreated]);
}

/**
 * Gets Aspen class configuration
 * @param {string} classId - Aspen class ID
 * @returns {Object|null} Configuration object or null if not found
 */
function getAspenClassConfig(classId) {
  const sheet = getAspenConfigSheet();
  const data = sheet.getDataRange().getValues();
  const classIdCol = getColumnIndex(ASPEN_CONFIG_COLS, ASPEN_CONFIG_HEADERS.classId);
  const classNameCol = getColumnIndex(ASPEN_CONFIG_COLS, ASPEN_CONFIG_HEADERS.className);
  const categoriesCol = getColumnIndex(ASPEN_CONFIG_COLS, ASPEN_CONFIG_HEADERS.categoriesJson);
  const dateCreatedCol = getColumnIndex(ASPEN_CONFIG_COLS, ASPEN_CONFIG_HEADERS.dateCreated);

  for (let i = 1; i < data.length; i++) {
    if (data[i][classIdCol] === classId) {
      return {
        classId: data[i][classIdCol],
        className: data[i][classNameCol],
        categories: JSON.parse(data[i][categoriesCol]),
        dateCreated: data[i][dateCreatedCol]
      };
    }
  }

  return null;
}

/**
 * Gets or creates the Aspen Students sheet
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getAspenStudentsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Aspen Students');

  if (!sheet) {
    sheet = ss.insertSheet('Aspen Students');
    sheet.getRange(1, 1, 1, ASPEN_STUDENTS_COLS.length).setValues([ASPEN_STUDENTS_COLS]);
    sheet.getRange(1, 1, 1, ASPEN_STUDENTS_COLS.length).setFontWeight('bold');
  }

  return sheet;
}

/**
 * Stores Aspen students for a class
 * @param {string} classId - Aspen class ID
 * @param {Array} students - Array of student objects
 */
function storeAspenStudents(classId, students) {
  const sheet = getAspenStudentsSheet();

  const data = sheet.getDataRange().getValues();
  const classIdCol = getColumnIndex(ASPEN_STUDENTS_COLS, ASPEN_STUDENTS_HEADERS.classId);
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][classIdCol] === classId) {
      sheet.deleteRow(i + 1);
    }
  }

  students.forEach(student => {
    const studentJson = JSON.stringify(student);
    sheet.appendRow([
      student.sourcedId,
      student.givenName + ' ' + student.familyName,
      student.email,
      classId,
      studentJson
    ]);
  });
}

/**
 * Gets Aspen students for a class
 * @param {string} classId - Aspen class ID
 * @returns {Array} Array of student objects
 */
function getAspenStudents(classId) {
  const sheet = getAspenStudentsSheet();
  const data = sheet.getDataRange().getValues();
  const students = [];
  const studentIdCol = getColumnIndex(ASPEN_STUDENTS_COLS, ASPEN_STUDENTS_HEADERS.studentId);
  const nameCol = getColumnIndex(ASPEN_STUDENTS_COLS, ASPEN_STUDENTS_HEADERS.name);
  const emailCol = getColumnIndex(ASPEN_STUDENTS_COLS, ASPEN_STUDENTS_HEADERS.email);
  const classIdCol = getColumnIndex(ASPEN_STUDENTS_COLS, ASPEN_STUDENTS_HEADERS.classId);
  const studentJsonCol = getColumnIndex(ASPEN_STUDENTS_COLS, ASPEN_STUDENTS_HEADERS.studentJson);

  for (let i = 1; i < data.length; i++) {
    if (data[i][classIdCol] === classId) {
      students.push({
        studentId: data[i][studentIdCol],
        name: data[i][nameCol],
        email: data[i][emailCol],
        classId: data[i][classIdCol],
        studentData: JSON.parse(data[i][studentJsonCol])
      });
    }
  }

  return students;
}

/**
 * Initializes Aspen integration for a class
 * @param {string} classId - Aspen class ID
 */
function initializeAspenIntegration(classId) {
  try {
    const teacher = fetchTeacherByEmail(Session.getActiveUser().getEmail());
    const courses = fetchAspenCourses(teacher);

    const selectedCourse = courses.find(course => course.sourcedId === classId);
    if (!selectedCourse) {
      throw new Error('Course not found or access denied');
    }

    const categories = fetchCategories(selectedCourse);
    storeAspenClassConfig(classId, selectedCourse.title, categories);

    const students = fetchStudents(selectedCourse);
    storeAspenStudents(classId, students);

    console.log(`Initialized Aspen integration for class: ${selectedCourse.title}`);
    return {
      success: true,
      message: `Successfully initialized integration for ${selectedCourse.title}`,
      classId: classId,
      categoriesCount: categories.length,
      studentsCount: students.length
    };

  } catch (error) {
    console.error('Error initializing Aspen integration:', error);
    return {
      success: false,
      message: error.toString()
    };
  }
}

/**
 * Gets available Aspen courses for the current user
 * @returns {Array} Array of available courses
 */
function getAvailableAspenCourses() {
  try {
    const teacher = fetchTeacherByEmail(Session.getActiveUser().getEmail());
    return fetchAspenCourses(teacher);
  } catch (error) {
    console.error('Error fetching Aspen courses:', error);
    return [];
  }
}
