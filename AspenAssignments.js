/* AspenAssignments.js Last Update 2025-09-13 09:34 <61003ae3444b9277089caea78b8621eafac7486e8d86332cb3d9f2584ac75f16>
// filepath: /Users/thinkle/BackedUpProjects/gas/standards-based-grading-sheet/AspenAssignments.js

/* global SpreadsheetApp, createAssignmentId, createAssignmentTitle, getAspenClassConfig, createLineItem */

/* Sheet and manager code for Aspen assignments */
/**
 * @typedef {Object} AssignmentSpec
 * @property {string} unit - Unit name/identifier
 * @property {string} skill - Skill/standard name
 * @property {string} categoryTitle - Category title from Aspen
 * @property {Date} dueDate - Due date for the assignment
 * @property {number} [minValue=0] - Minimum score value
 * @property {number} [maxValue=4] - Maximum score value
 */
// Aspen Assignments Sheet
const ASPEN_ASSIGNMENTS_HEADERS = {
  unit: 'Unit',
  skill: 'Skill',
  assignmentId: 'Aspen ID',
  title: 'Title',
  category: 'Category',
  dueDate: 'Due Date',
  minValue: 'Min Value',
  maxValue: 'Max Value',
  dateCreated: 'Date Created',
  assignmentJson: 'Assignment JSON'
};

const ASPEN_ASSIGNMENTS_COLS = [
  ASPEN_ASSIGNMENTS_HEADERS.unit,
  ASPEN_ASSIGNMENTS_HEADERS.skill,
  ASPEN_ASSIGNMENTS_HEADERS.assignmentId,
  ASPEN_ASSIGNMENTS_HEADERS.title,
  ASPEN_ASSIGNMENTS_HEADERS.category,
  ASPEN_ASSIGNMENTS_HEADERS.dueDate,
  ASPEN_ASSIGNMENTS_HEADERS.minValue,
  ASPEN_ASSIGNMENTS_HEADERS.maxValue,
  ASPEN_ASSIGNMENTS_HEADERS.dateCreated,
  ASPEN_ASSIGNMENTS_HEADERS.assignmentJson
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

/* Aspen Assignment Sheet Functions */

/**
 * Gets or creates the Aspen Assignments sheet
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getAspenAssignmentsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Aspen Assignments');

  if (!sheet) {
    sheet = ss.insertSheet('Aspen Assignments');
    // Set up headers using constants
    sheet.getRange(1, 1, 1, ASPEN_ASSIGNMENTS_COLS.length).setValues([ASPEN_ASSIGNMENTS_COLS]);
    sheet.getRange(1, 1, 1, ASPEN_ASSIGNMENTS_COLS.length).setFontWeight('bold');
  }

  return sheet;
}

/**
 * Stores an Aspen assignment
 * @param {Object} assignment - Assignment object from Aspen API
 * @param {AssignmentSpec} assignmentSpec - Assignment specification
 */
function storeAspenAssignment(assignment, assignmentSpec) {
  const sheet = getAspenAssignmentsSheet();
  const assignmentJson = JSON.stringify(assignment);
  const dateCreated = new Date();

  sheet.appendRow([
    assignmentSpec.unit,
    assignmentSpec.skill,
    assignment.sourcedId,
    assignment.title,
    assignmentSpec.categoryTitle,
    assignment.dueDate || '',
    assignment.resultValueMin || 0,
    assignment.resultValueMax || 4,
    dateCreated,
    assignmentJson
  ]);
}

/**
 * Gets stored Aspen assignments
 * @returns {Array} Array of assignment objects
 */
function getAspenAssignments() {
  const sheet = getAspenAssignmentsSheet();
  const data = sheet.getDataRange().getValues();
  const assignments = [];
  const unitCol = getColumnIndex(ASPEN_ASSIGNMENTS_COLS, ASPEN_ASSIGNMENTS_HEADERS.unit);
  const skillCol = getColumnIndex(ASPEN_ASSIGNMENTS_COLS, ASPEN_ASSIGNMENTS_HEADERS.skill);
  const assignmentIdCol = getColumnIndex(ASPEN_ASSIGNMENTS_COLS, ASPEN_ASSIGNMENTS_HEADERS.assignmentId);
  const titleCol = getColumnIndex(ASPEN_ASSIGNMENTS_COLS, ASPEN_ASSIGNMENTS_HEADERS.title);
  const categoryCol = getColumnIndex(ASPEN_ASSIGNMENTS_COLS, ASPEN_ASSIGNMENTS_HEADERS.category);
  const dueDateCol = getColumnIndex(ASPEN_ASSIGNMENTS_COLS, ASPEN_ASSIGNMENTS_HEADERS.dueDate);
  const minValueCol = getColumnIndex(ASPEN_ASSIGNMENTS_COLS, ASPEN_ASSIGNMENTS_HEADERS.minValue);
  const maxValueCol = getColumnIndex(ASPEN_ASSIGNMENTS_COLS, ASPEN_ASSIGNMENTS_HEADERS.maxValue);
  const dateCreatedCol = getColumnIndex(ASPEN_ASSIGNMENTS_COLS, ASPEN_ASSIGNMENTS_HEADERS.dateCreated);
  const assignmentJsonCol = getColumnIndex(ASPEN_ASSIGNMENTS_COLS, ASPEN_ASSIGNMENTS_HEADERS.assignmentJson);

  for (let i = 1; i < data.length; i++) {
    assignments.push({
      unit: data[i][unitCol],
      skill: data[i][skillCol],
      assignmentId: data[i][assignmentIdCol],
      title: data[i][titleCol],
      category: data[i][categoryCol],
      dueDate: data[i][dueDateCol],
      minValue: data[i][minValueCol],
      maxValue: data[i][maxValueCol],
      dateCreated: data[i][dateCreatedCol],
      assignmentData: JSON.parse(data[i][assignmentJsonCol])
    });
  }

  return assignments;
}

/* Assignment Manager - Creates Aspen assignments from spreadsheet skills/standards */

/**
 * Assignment Manager - handles creating and managing Aspen assignments
 * from the spreadsheet's skills/standards system
 */
class AspenAssignmentManager {
  constructor(classId) {
    this.classId = classId;
    this.aspenConfig = null;
    this.existingAssignments = null;
    this.assignmentLookup = null;
  }

  /**
   * Loads configuration and existing assignments
   */
  loadData() {
    console.log('Loading assignment manager data...');

    // Load Aspen class configuration (includes categories)
    this.aspenConfig = getAspenClassConfig(this.classId);
    if (!this.aspenConfig) {
      throw new Error(`Aspen configuration not found for class ${this.classId}. Run initializeAspenIntegration first.`);
    }

    // Load existing assignments and create lookup
    this.existingAssignments = getAspenAssignments();
    this.assignmentLookup = {};
    this.existingAssignments.forEach(assignment => {
      this.assignmentLookup[assignment.assignmentId] = assignment;
    });

    console.log(`Loaded ${this.existingAssignments.length} existing assignments, ${this.aspenConfig.categories.length} categories`);
  }

  /**
   * Gets available categories for assignments
   * @returns {Array} Array of category objects
   */
  getCategories() {
    if (!this.aspenConfig) {
      throw new Error('Data not loaded. Call loadData() first.');
    }
    return this.aspenConfig.categories;
  }

  /**
   * Creates a unique assignment ID from unit and skill info
   * @param {string} unit - Unit name/identifier
   * @param {string} skill - Skill/standard identifier  
   * @returns {string} Unique assignment ID
   */
  createAssignmentId(unit, skill) {
    return createAssignmentId(this.classId, unit, skill);
  }

  /**
   * Creates an assignment title from unit and skill info
   * @param {string} unit - Unit name
   * @param {string} skill - Skill/standard name
   * @returns {string} Assignment title
   */


  /**
   * Checks if an assignment already exists
   * @param {string} unit - Unit identifier
   * @param {string} skill - Skill identifier
   * @returns {boolean} True if assignment exists
   */
  assignmentExists(unit, skill) {
    if (!this.assignmentLookup) {
      throw new Error('Data not loaded. Call loadData() first.');
    }
    const assignmentId = this.createAssignmentId(unit, skill);
    return !!this.assignmentLookup[assignmentId];
  }

  /**
   * Gets an existing assignment
   * @param {string} unit - Unit identifier
   * @param {string} skill - Skill identifier
   * @returns {Object|null} Assignment object or null if not found
   */
  getAssignment(unit, skill) {
    const assignmentId = this.createAssignmentId(unit, skill);
    return this.assignmentLookup[assignmentId] || null;
  }

  /**
   * Creates a new assignment in Aspen
   * @param {AssignmentSpec} assignmentSpec - Assignment specification   
   * @returns {Object} Creation result
   */
  createAssignment(assignmentSpec) {
    try {
      const { unit, skill, categoryTitle, dueDate, minValue = 0, maxValue = 4 } = assignmentSpec;

      // Check if assignment already exists
      if (this.assignmentExists(unit, skill)) {
        return {
          success: false,
          message: `Assignment for ${unit} - ${skill} already exists`,
          assignmentId: this.createAssignmentId(unit, skill)
        };
      }

      // Find the category
      const category = this.aspenConfig.categories.find(cat => cat.title === categoryTitle);
      if (!category) {
        throw new Error(`Category '${categoryTitle}' not found. Available categories: ${this.aspenConfig.categories.map(c => c.title).join(', ')}`);
      }

      // Create assignment data structure for Aspen API
      const assignmentId = this.createAssignmentId(unit, skill);
      const title = createAssignmentTitle(unit, skill);

      const lineItemData = {
        sourcedId: assignmentId,
        title: title,
        description: `Standards-based assessment for ${unit} - ${skill}`,
        assignDate: new Date().toISOString().split('T')[0], // Today
        dueDate: dueDate.toISOString().split('T')[0],
        class: {
          sourcedId: this.classId,
        },
        category: {
          sourcedId: category.sourcedId,
        },
        gradingPeriod: category.gradingPeriod || null,
        resultValueMin: minValue,
        resultValueMax: maxValue,
        metadata: {
          unit: unit,
          skill: skill,
          createdBy: 'standards-based-grading-sheet'
        }
      };

      console.log('Creating assignment with assignmentId:', assignmentId);
      console.log('Line item data:', JSON.stringify(lineItemData, null, 2));

      // Create the assignment in Aspen
      const result = createLineItem(assignmentId, lineItemData);

      // Store the assignment locally
      storeAspenAssignment(result, assignmentSpec);

      // Update our lookup
      const assignmentRecord = {
        assignmentId: assignmentId,
        title: title,
        category: categoryTitle,
        dueDate: dueDate,
        minValue: minValue,
        maxValue: maxValue,
        dateCreated: new Date(),
        assignmentData: result
      };
      this.assignmentLookup[assignmentId] = assignmentRecord;

      return {
        success: true,
        message: `Successfully created assignment: ${title}`,
        assignmentId: assignmentId,
        assignment: result
      };

    } catch (error) {
      console.error('Error creating assignment:', error);
      return {
        success: false,
        message: error.toString()
      };
    }
  }

  /**
   * Creates assignments for multiple skills at once
   * @param {Array} skillsList - Array of skill objects
   * @param {string} categoryTitle - Category for all assignments
   * @param {Date} dueDate - Due date for all assignments
   * @param {number} [minValue=0] - Minimum score
   * @param {number} [maxValue=4] - Maximum score
   * @returns {Object} Batch creation result
   */
  createBatchAssignments(skillsList, categoryTitle, dueDate, minValue = 0, maxValue = 4) {
    const results = {
      success: true,
      created: [],
      skipped: [],
      errors: []
    };

    for (const skill of skillsList) {
      const result = this.createAssignment({
        unit: skill.unit,
        skill: skill.name,
        categoryTitle: categoryTitle,
        dueDate: dueDate,
        minValue: minValue,
        maxValue: maxValue
      });

      if (result.success) {
        results.created.push({
          unit: skill.unit,
          skill: skill.name,
          assignmentId: result.assignmentId
        });
      } else if (result.message.includes('already exists')) {
        results.skipped.push({
          unit: skill.unit,
          skill: skill.name,
          assignmentId: result.assignmentId
        });
      } else {
        results.errors.push({
          unit: skill.unit,
          skill: skill.name,
          error: result.message
        });
        results.success = false;
      }
    }

    return {
      ...results,
      message: `Created ${results.created.length}, skipped ${results.skipped.length}, errors ${results.errors.length}`
    };
  }

  /**
   * Gets the assignment ID for a given unit/skill combination
   * This is the key function for linking spreadsheet data to Aspen assignments
   * @param {string} unit - Unit identifier
   * @param {string} skill - Skill identifier
   * @returns {string|null} Assignment ID or null if not found
   */
  getAssignmentIdForSkill(unit, skill) {
    const assignmentId = this.createAssignmentId(unit, skill);
    return this.assignmentExists(unit, skill) ? assignmentId : null;
  }

  /**
   * Gets stats about assignments
   * @returns {Object} Assignment statistics
   */
  getStats() {
    return {
      totalAssignments: this.existingAssignments ? this.existingAssignments.length : 0,
      categoriesCount: this.aspenConfig ? this.aspenConfig.categories.length : 0,
      classId: this.classId
    };
  }
}

/**
 * Creates a new Assignment Manager for a class
 * @param {string} classId - Aspen class ID
 * @returns {AspenAssignmentManager} Initialized assignment manager
 */
function createAspenAssignmentManager(classId) {
  const manager = new AspenAssignmentManager(classId);
  manager.loadData();
  return manager;
}