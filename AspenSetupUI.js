/* AspenSetupUI.js Last Update 2025-09-12 16:39 <9b19f8798ea58a272a7b19ad4a41f34a47e65e900d3320f8f5b68db2586cc917>
// filepath: /Users/thinkle/BackedUpProjects/gas/standards-based-grading-sheet/AspenSetupUI.js

/* Simple UI functions for Aspen integration setup */

/**
 * Main function to set up Aspen integration - called from menu or manually
 * This is a one-time setup process
 */
function setupAspenIntegration() {
  try {
    // Step 1: Get available courses
    console.log('Fetching available Aspen courses...');
    const courses = getAvailableAspenCourses();

    if (!courses || courses.length === 0) {
      SpreadsheetApp.getUi().alert(
        'No Courses Found',
        'No Aspen courses found for your account. Please check your Aspen access and try again.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    // Step 2: Show course selection dialog
    const selectedCourse = showCourseSelectionDialog(courses);
    if (!selectedCourse) {
      return; // User cancelled
    }

    // Step 3: Confirm selection
    const confirmed = confirmCourseSelection(selectedCourse);
    if (!confirmed) {
      return; // User cancelled
    }

    // Step 4: Initialize integration
    console.log(`Initializing integration for course: ${selectedCourse.title}`);
    const result = initializeAspenIntegration(selectedCourse.sourcedId);

    if (result.success) {
      SpreadsheetApp.getUi().alert(
        'Setup Complete!',
        `Successfully set up Aspen integration!\n\n` +
        `Class: ${selectedCourse.title}\n` +
        `Categories: ${result.categoriesCount}\n` +
        `Students: ${result.studentsCount}\n\n` +
        `You can now create assignments and sync grades.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      SpreadsheetApp.getUi().alert(
        'Setup Failed',
        `Failed to initialize Aspen integration:\n\n${result.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }

  } catch (error) {
    console.error('Error in setupAspenIntegration:', error);
    SpreadsheetApp.getUi().alert(
      'Error',
      `An error occurred during setup:\n\n${error.toString()}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Shows a dialog for the user to select their course
 * @param {Array} courses - Array of course objects
 * @returns {Object|null} Selected course or null if cancelled
 */
function showCourseSelectionDialog(courses) {
  const ui = SpreadsheetApp.getUi();

  // Build the course list message
  let courseList = 'Available Aspen Courses:\n\n';
  courses.forEach((course, index) => {
    courseList += `${index + 1}. ${course.title}\n`;
    if (course.courseCode) {
      courseList += `   Code: ${course.courseCode}\n`;
    }
    courseList += '\n';
  });

  courseList += `\nEnter the number (1-${courses.length}) of the course you want to sync with:`;

  // Show input dialog
  const response = ui.prompt(
    'Select Aspen Course',
    courseList,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return null; // User cancelled
  }

  const userInput = response.getResponseText().trim();
  const courseNumber = parseInt(userInput);

  // Validate input
  if (isNaN(courseNumber) || courseNumber < 1 || courseNumber > courses.length) {
    ui.alert(
      'Invalid Selection',
      `Please enter a number between 1 and ${courses.length}.`,
      ui.ButtonSet.OK
    );
    return showCourseSelectionDialog(courses); // Try again
  }

  return courses[courseNumber - 1];
}

/**
 * Shows a confirmation dialog for the selected course
 * @param {Object} course - Selected course object
 * @returns {boolean} True if confirmed, false if cancelled
 */
function confirmCourseSelection(course) {
  const ui = SpreadsheetApp.getUi();

  const confirmMessage =
    `You selected:\n\n` +
    `Course: ${course.title}\n` +
    `${course.courseCode ? `Code: ${course.courseCode}\n` : ''}` +
    `${course.schoolYear ? `Year: ${course.schoolYear}\n` : ''}` +
    `\nThis will set up grade syncing for this course. ` +
    `The system will download student lists and grading categories.\n\n` +
    `Continue with setup?`;

  const response = ui.alert(
    'Confirm Course Selection',
    confirmMessage,
    ui.ButtonSet.YES_NO
  );

  return response === ui.Button.YES;
}

/**
 * Quick test function to check if Aspen connection is working
 */
function testAspenConnection() {
  try {
    console.log('Testing Aspen connection...');
    const courses = getAvailableAspenCourses();

    if (courses && courses.length > 0) {
      SpreadsheetApp.getUi().alert(
        'Connection Test',
        `✅ Aspen connection successful!\n\nFound ${courses.length} courses:\n\n` +
        courses.slice(0, 3).map(c => `• ${c.title}`).join('\n') +
        (courses.length > 3 ? `\n... and ${courses.length - 3} more` : ''),
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      SpreadsheetApp.getUi().alert(
        'Connection Test',
        '❌ No courses found. Please check your Aspen access.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (error) {
    console.error('Error testing Aspen connection:', error);
    SpreadsheetApp.getUi().alert(
      'Connection Test Failed',
      `❌ Error connecting to Aspen:\n\n${error.toString()}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Shows current integration status
 */
function showAspenStatus() {
  try {
    // Try to find existing config
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Aspen Config');

    if (!sheet || sheet.getDataRange().getNumRows() <= 1) {
      SpreadsheetApp.getUi().alert(
        'Integration Status',
        '❌ No Aspen integration configured.\n\nRun "Setup Aspen Integration" to get started.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    // Get config data
    const data = sheet.getDataRange().getValues();
    let statusMessage = '✅ Aspen Integration Active\n\n';

    for (let i = 1; i < data.length; i++) {
      const classId = data[i][0];
      const className = data[i][1];
      const dateCreated = data[i][3];

      statusMessage += `Class: ${className}\n`;
      statusMessage += `ID: ${classId}\n`;
      statusMessage += `Setup: ${dateCreated}\n\n`;
    }

    // Check for assignments
    const assignmentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Aspen Assignments');
    const assignmentCount = assignmentsSheet ? assignmentsSheet.getDataRange().getNumRows() - 1 : 0;

    // Check for synced grades
    const gradesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Aspen Grades');
    const gradeCount = gradesSheet ? gradesSheet.getDataRange().getNumRows() - 1 : 0;

    statusMessage += `Assignments created: ${Math.max(0, assignmentCount)}\n`;
    statusMessage += `Grades synced: ${Math.max(0, gradeCount)}`;

    SpreadsheetApp.getUi().alert(
      'Integration Status',
      statusMessage,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (error) {
    console.error('Error checking Aspen status:', error);
    SpreadsheetApp.getUi().alert(
      'Status Check Failed',
      `Error checking integration status:\n\n${error.toString()}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// Placeholder functions for advanced features
function createAspenAssignmentUI() {
  SpreadsheetApp.getUi().alert('Coming Soon', 'Assignment creation UI will be implemented next!', SpreadsheetApp.getUi().ButtonSet.OK);
}

function syncGradesUI() {
  SpreadsheetApp.getUi().alert('Coming Soon', 'Grade sync UI will be implemented next!', SpreadsheetApp.getUi().ButtonSet.OK);
}