/* AspenAssignments.js Last Update 2025-09-13 11:35 <344b4c845fdc1f0b86bd68b7f878ea1c7c5d29b10d1be00b7476c4cb8e826d93>
// filepath: /Users/thinkle/BackedUpProjects/gas/standards-based-grading-sheet/AspenAssignments.js

/* global SpreadsheetApp, createAssignmentId, createAssignmentTitle, getAspenClassConfig, createLineItem, STYLE */

/* Sheet and manager code for Aspen assignments */
/**
 * @typedef {Object} AssignmentSpec
 * @property {string} unit - Unit name/identifier
 * @property {string} skill - Skill/standard name
 * @property {string} categoryTitle - Category title from Aspen
 * @property {Date} dueDate - Due date for the assignment
 * @property {Date} [assignDate] - Assign date for the assignment (optional, defaults to today)
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
  assignDate: 'Assign Date',
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
  ASPEN_ASSIGNMENTS_HEADERS.assignDate,
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

/** Convert 1-based column number to A1 letter(s). */
function columnA1_(n) {
  let s = '';
  while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
  return s;
}

/**
 * Read Skills sheet into a lookup map: { unit: { skillNum: descriptor } }
 * Best-effort: if Skills sheet missing or empty, returns {}.
 */
function readSkillsMap_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Skills');
  const map = {};
  if (!sh) return map;
  const data = sh.getDataRange().getValues();
  if (!data || data.length < 2) return map;
  const header = data[0];
  const unitIdx = header.indexOf('Unit');
  const numIdx = header.indexOf('Skill #');
  const descIdx = header.indexOf('Descriptor');
  if (unitIdx === -1 || numIdx === -1 || descIdx === -1) return map;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const unit = String(row[unitIdx] || '').trim();
    const num = String(row[numIdx] || '').trim();
    const desc = String(row[descIdx] || '').trim();
    if (!unit || !num) continue;
    if (!map[unit]) map[unit] = {};
    if (!map[unit][num]) map[unit][num] = desc;
  }
  return map;
}

/** Apply conditional formatting to the Aspen Assignments sheet. */
function applyAspenAssignmentsConditionalFormatting() {
  const sh = getAspenAssignmentsSheet();
  const headers = ASPEN_ASSIGNMENTS_COLS;
  const idCol = getColumnIndex(headers, ASPEN_ASSIGNMENTS_HEADERS.assignmentId) + 1; // 1-based
  const dueCol = getColumnIndex(headers, ASPEN_ASSIGNMENTS_HEADERS.dueDate) + 1; // 1-based
  const catCol = getColumnIndex(headers, ASPEN_ASSIGNMENTS_HEADERS.category) + 1; // 1-based
  const totalRows = Math.max(1, sh.getMaxRows() - 1);
  const totalCols = headers.length;
  if (totalRows <= 0) return;
  const dataRange = sh.getRange(2, 1, totalRows, totalCols);

  const rules = sh.getConditionalFormatRules();
  const dataA1 = dataRange.getA1Notation();
  // Remove any prior rules targeting this full data range to avoid duplicates
  const filtered = rules.filter(r => !r.getRanges().some(rg => rg.getA1Notation() === dataA1));

  // Colors from STYLE tokens with safe fallbacks
  const neutralBg = (typeof STYLE !== 'undefined' && STYLE.COLORS && STYLE.COLORS.UI && STYLE.COLORS.UI.NEUTRAL_BG) || '#f7f7f7';
  const neutralText = (typeof STYLE !== 'undefined' && STYLE.COLORS && STYLE.COLORS.UI && STYLE.COLORS.UI.NEUTRAL_TEXT) || '#333333';
  const inputBg = (typeof STYLE !== 'undefined' && STYLE.COLORS && STYLE.COLORS.UI && STYLE.COLORS.UI.INPUT_HIGHLIGHT) || '#fffdf0';

  const idA = columnA1_(idCol);
  const dueA = columnA1_(dueCol);
  const catA = columnA1_(catCol);
  const rowStart = 2;

  // 1) Existing (Aspen ID present) => subtle solid background, normal text
  const hasIdFormula = `=LEN($${idA}${rowStart})>0`;
  const ruleHasId = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(hasIdFormula)
    .setBackground(neutralBg)
    .setFontColor(neutralText)
    .setRanges([dataRange])
    .build();

  // 2) Ready to roll (no ID yet, has due date) => highlight background
  const readyFormula = `=AND($${idA}${rowStart}="", $${dueA}${rowStart}<>"", $${catA}${rowStart}<>"")`;
  const ruleReady = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(readyFormula)
    .setBackground(inputBg)
    .setFontColor('#000000')
    .setRanges([dataRange])
    .build();

  // 3) Not real yet (no ID, no due date) => grey and italic
  const notRealFormula = `=AND($${idA}${rowStart}="", $${dueA}${rowStart}="")`;
  const ruleNotReal = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(notRealFormula)
    .setFontColor(neutralText)
    .setItalic(true)
    .setRanges([dataRange])
    .build();

  // Apply in priority order: Has ID > Ready > Not real
  filtered.push(ruleNotReal, ruleReady, ruleHasId);
  sh.setConditionalFormatRules(filtered);
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
    try { applyAspenAssignmentsConditionalFormatting(); } catch (e) { /* optional styling */ }
    try { applyAspenAssignmentsCategoryValidation(sheet); } catch (e) { /* optional */ }
  }

  return sheet;
}

/**
 * Apply data validation for Category column using titles from Aspen Config (Categories JSON)
 */
function applyAspenAssignmentsCategoryValidation(sheet) {
  const ss = sheet.getParent();
  const cfg = ss.getSheetByName('Aspen Config');
  if (!cfg || cfg.getLastRow() < 2) return;
  const headers = ASPEN_ASSIGNMENTS_COLS;
  const catCol = getColumnIndex(headers, ASPEN_ASSIGNMENTS_HEADERS.category) + 1; // 1-based
  const dataRows = Math.max(1, sheet.getMaxRows() - 1);

  // Parse category titles from JSON in column C of Aspen Config
  const data = cfg.getDataRange().getValues();
  const categoriesJsonIdx = 2; // zero-based: column C
  let titles = [];
  for (let i = 1; i < data.length; i++) {
    const raw = data[i][categoriesJsonIdx];
    if (!raw) continue;
    try {
      const arr = (typeof raw === 'string') ? JSON.parse(raw) : raw;
      if (Array.isArray(arr)) {
        arr.forEach(c => { if (c && c.title) titles.push(String(c.title)); });
      }
    } catch (e) {
      if (typeof console !== 'undefined' && console.warn) console.warn('Category JSON parse failed for Aspen Config row', i + 1, e);
    }
  }
  titles = Array.from(new Set(titles)).filter(Boolean);
  if (titles.length === 0) return;

  // Ensure we have a helper sheet/range for validation list (or inline list if short)
  if (titles.join(',').length <= 50000) {
    // Inline list via data validation
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(titles, true)
      .setAllowInvalid(false)
      .setHelpText('Choose an Aspen category title (from configured class categories).')
      .build();
    sheet.getRange(2, catCol, dataRows, 1).setDataValidation(rule);
  } else {
    // Fallback: write to a hidden helper sheet range
    const helperName = 'Aspen Category Titles (helper)';
    let helper = ss.getSheetByName(helperName);
    if (!helper) helper = ss.insertSheet(helperName);
    helper.clear();
    helper.getRange(1, 1, titles.length, 1).setValues(titles.map(t => [t]));
    try { helper.hideSheet(); } catch (e) { }
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(helper.getRange(1, 1, titles.length, 1), true)
      .setAllowInvalid(false)
      .setHelpText('Choose an Aspen category title (from configured class categories).')
      .build();
    sheet.getRange(2, catCol, dataRows, 1).setDataValidation(rule);
  }
}

/**
 * Stores an Aspen assignment
 * @param {Object} assignment - Assignment object from Aspen API
 * @param {AssignmentSpec} assignmentSpec - Assignment specification
 */
function storeAspenAssignment(assignment, assignmentSpec) {
  const sheet = getAspenAssignmentsSheet();
  const headers = ASPEN_ASSIGNMENTS_COLS;
  const col = (name) => getColumnIndex(headers, name);
  const unitC = col(ASPEN_ASSIGNMENTS_HEADERS.unit) + 1; // 1-based
  const skillC = col(ASPEN_ASSIGNMENTS_HEADERS.skill) + 1;
  const idC = col(ASPEN_ASSIGNMENTS_HEADERS.assignmentId) + 1;
  const titleC = col(ASPEN_ASSIGNMENTS_HEADERS.title) + 1;
  const catC = col(ASPEN_ASSIGNMENTS_HEADERS.category) + 1;
  const assignDateC = col(ASPEN_ASSIGNMENTS_HEADERS.assignDate) + 1;
  const dueDateC = col(ASPEN_ASSIGNMENTS_HEADERS.dueDate) + 1;
  const minC = col(ASPEN_ASSIGNMENTS_HEADERS.minValue) + 1;
  const maxC = col(ASPEN_ASSIGNMENTS_HEADERS.maxValue) + 1;
  const createdC = col(ASPEN_ASSIGNMENTS_HEADERS.dateCreated) + 1;
  const jsonC = col(ASPEN_ASSIGNMENTS_HEADERS.assignmentJson) + 1;

  const assignmentJson = JSON.stringify(assignment);
  const dateCreated = new Date();
  const assignDate = assignmentSpec.assignDate ? assignmentSpec.assignDate : new Date();

  // Prepare row payload
  const rowValues = [];
  rowValues[unitC - 1] = assignmentSpec.unit;
  rowValues[skillC - 1] = assignmentSpec.skill;
  rowValues[idC - 1] = assignment.sourcedId;
  rowValues[titleC - 1] = assignment.title;
  rowValues[catC - 1] = assignmentSpec.categoryTitle;
  rowValues[assignDateC - 1] = assignDate;
  rowValues[dueDateC - 1] = assignmentSpec.dueDate || '';
  rowValues[minC - 1] = assignment.resultValueMin || 0;
  rowValues[maxC - 1] = assignment.resultValueMax || 4;
  rowValues[createdC - 1] = dateCreated;
  rowValues[jsonC - 1] = assignmentJson;

  // Find existing row by Aspen ID or (Unit+Skill)
  const lastRow = sheet.getLastRow();
  let targetRow = null;
  if (lastRow >= 2) {
    const rng = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    for (let i = 0; i < rng.length; i++) {
      const r = rng[i];
      const rowIndex = 2 + i;
      const curId = (r[idC - 1] || '').toString().trim();
      const curUnit = (r[unitC - 1] || '').toString().trim();
      const curSkill = (r[skillC - 1] || '').toString().trim();
      if (curId && curId === assignment.sourcedId) { targetRow = rowIndex; break; }
      if (!targetRow && curUnit === assignmentSpec.unit && curSkill === assignmentSpec.skill) { targetRow = rowIndex; }
    }
  }

  if (targetRow) {
    // Update existing row in place (only our known columns)
    const rowRange = sheet.getRange(targetRow, 1, 1, headers.length);
    const existing = rowRange.getValues()[0];
    // Merge: only overwrite our known fields, keep any extra columns untouched
    existing[unitC - 1] = rowValues[unitC - 1];
    existing[skillC - 1] = rowValues[skillC - 1];
    existing[idC - 1] = rowValues[idC - 1];
    existing[titleC - 1] = rowValues[titleC - 1];
    existing[catC - 1] = rowValues[catC - 1];
    existing[assignDateC - 1] = rowValues[assignDateC - 1];
    existing[dueDateC - 1] = rowValues[dueDateC - 1];
    existing[minC - 1] = rowValues[minC - 1];
    existing[maxC - 1] = rowValues[maxC - 1];
    existing[createdC - 1] = rowValues[createdC - 1];
    existing[jsonC - 1] = rowValues[jsonC - 1];
    rowRange.setValues([existing]);
  } else {
    // Append a new row (align with headers length)
    const row = new Array(headers.length).fill('');
    row[unitC - 1] = rowValues[unitC - 1];
    row[skillC - 1] = rowValues[skillC - 1];
    row[idC - 1] = rowValues[idC - 1];
    row[titleC - 1] = rowValues[titleC - 1];
    row[catC - 1] = rowValues[catC - 1];
    row[assignDateC - 1] = rowValues[assignDateC - 1];
    row[dueDateC - 1] = rowValues[dueDateC - 1];
    row[minC - 1] = rowValues[minC - 1];
    row[maxC - 1] = rowValues[maxC - 1];
    row[createdC - 1] = rowValues[createdC - 1];
    row[jsonC - 1] = rowValues[jsonC - 1];
    sheet.appendRow(row);
  }

  try { applyAspenAssignmentsConditionalFormatting(); } catch (e) { /* optional styling */ }
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
  const assignDateCol = getColumnIndex(ASPEN_ASSIGNMENTS_COLS, ASPEN_ASSIGNMENTS_HEADERS.assignDate);
  const dueDateCol = getColumnIndex(ASPEN_ASSIGNMENTS_COLS, ASPEN_ASSIGNMENTS_HEADERS.dueDate);
  const minValueCol = getColumnIndex(ASPEN_ASSIGNMENTS_COLS, ASPEN_ASSIGNMENTS_HEADERS.minValue);
  const maxValueCol = getColumnIndex(ASPEN_ASSIGNMENTS_COLS, ASPEN_ASSIGNMENTS_HEADERS.maxValue);
  const dateCreatedCol = getColumnIndex(ASPEN_ASSIGNMENTS_COLS, ASPEN_ASSIGNMENTS_HEADERS.dateCreated);
  const assignmentJsonCol = getColumnIndex(ASPEN_ASSIGNMENTS_COLS, ASPEN_ASSIGNMENTS_HEADERS.assignmentJson);

  for (let i = 1; i < data.length; i++) {
    let rawJson = data[i][assignmentJsonCol];
    let parsedJson = null;
    if (rawJson != null && rawJson !== '') {
      try {
        // Some environments might store objects; preserve if already parsed
        parsedJson = (typeof rawJson === 'string') ? JSON.parse(rawJson) : rawJson;
      } catch (e) {
        if (typeof console !== 'undefined' && console.warn) {
          console.warn('AspenAssignments: Invalid JSON at row', i + 1, e);
        }
        parsedJson = null; // degrade gracefully
      }
    }
    assignments.push({
      unit: data[i][unitCol],
      skill: data[i][skillCol],
      assignmentId: data[i][assignmentIdCol],
      title: data[i][titleCol],
      category: data[i][categoryCol],
      assignDate: data[i][assignDateCol],
      dueDate: data[i][dueDateCol],
      minValue: data[i][minValueCol],
      maxValue: data[i][maxValueCol],
      dateCreated: data[i][dateCreatedCol],
      assignmentData: parsedJson
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
    this.skillsMap = null;
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

    // Load skills map for descriptor lookups
    try {
      this.skillsMap = readSkillsMap_();
      const unitCount = Object.keys(this.skillsMap).length;
      console.log(`Loaded skills map for ${unitCount} unit(s)`);
    } catch (e) {
      this.skillsMap = {};
      if (typeof console !== 'undefined' && console.warn) console.warn('Failed to load skills map', e);
    }
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
      const { unit, skill, categoryTitle, dueDate, assignDate, minValue = 0, maxValue = 4 } = assignmentSpec;

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
      // Lookup descriptor using Skills sheet (treat `skill` as Skill #)
      let descriptor = '';
      if (this.skillsMap && this.skillsMap[unit] && this.skillsMap[unit][skill]) {
        descriptor = this.skillsMap[unit][skill];
      } else {
        // Fallback: if not found, don't block creation
        descriptor = '';
      }
      const title = createAssignmentTitle(unit, skill, descriptor);

      const lineItemData = {
        sourcedId: assignmentId,
        title: title,
        description: descriptor ? `${descriptor} (${unit} - ${skill})` : `Standards-based assessment for ${unit} - ${skill}`,
        assignDate: (assignDate ? assignDate : new Date()).toISOString().split('T')[0],
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

  /**
   * Scan the Aspen Assignments sheet and create assignments for rows missing Aspen ID
   * but with required fields present. Only creates when dueDate is present.
   * Idempotent: rows with Aspen ID are skipped; duplicates prevented by sourcedId.
   * @returns {{created: number, skipped: number, errors: Array<{row:number,error:string}>, updatedRows:number[]}}
   */
  createMissingAssignmentsFromSheet() {
    const sh = getAspenAssignmentsSheet();
    const data = sh.getDataRange().getValues();
    if (data.length <= 1) return { created: 0, skipped: 0, errors: [], updatedRows: [] };

    const headers = ASPEN_ASSIGNMENTS_COLS;
    const col = (name) => getColumnIndex(headers, name);
    const unitC = col(ASPEN_ASSIGNMENTS_HEADERS.unit);
    const skillC = col(ASPEN_ASSIGNMENTS_HEADERS.skill);
    const idC = col(ASPEN_ASSIGNMENTS_HEADERS.assignmentId);
    const titleC = col(ASPEN_ASSIGNMENTS_HEADERS.title);
    const catC = col(ASPEN_ASSIGNMENTS_HEADERS.category);
    const assignDateC = col(ASPEN_ASSIGNMENTS_HEADERS.assignDate);
    const dueDateC = col(ASPEN_ASSIGNMENTS_HEADERS.dueDate);
    const minC = col(ASPEN_ASSIGNMENTS_HEADERS.minValue);
    const maxC = col(ASPEN_ASSIGNMENTS_HEADERS.maxValue);
    const createdC = col(ASPEN_ASSIGNMENTS_HEADERS.dateCreated);
    const jsonC = col(ASPEN_ASSIGNMENTS_HEADERS.assignmentJson);

    const result = { created: 0, skipped: 0, errors: [], updatedRows: [] };

    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      const aspenId = (row[idC] || '').toString().trim();
      if (aspenId) { result.skipped++; continue; }

      const unit = (row[unitC] || '').toString().trim();
      const skill = (row[skillC] || '').toString().trim();
      const categoryTitle = (row[catC] || '').toString().trim();
      const dueDate = row[dueDateC];
      const assignDate = row[assignDateC];
      const minValue = (row[minC] !== '' && row[minC] != null) ? Number(row[minC]) : 0;
      const maxValue = (row[maxC] !== '' && row[maxC] != null) ? Number(row[maxC]) : 4;

      // Required: unit, skill, categoryTitle, dueDate
      if (!unit || !skill || !categoryTitle || !dueDate) { result.skipped++; continue; }
      if (!(dueDate instanceof Date)) { result.skipped++; continue; }

      const spec = { unit, skill, categoryTitle, dueDate, assignDate: (assignDate instanceof Date) ? assignDate : undefined, minValue, maxValue };
      const created = this.createAssignment(spec);
      if (created && created.success) {
        // Upsert is handled inside storeAspenAssignment (called by createAssignment).
        // Avoid duplicate row creation by not writing here again.
        result.created++;
      } else {
        const errMsg = (created && created.message) ? created.message : 'Unknown error';
        if (typeof console !== 'undefined' && console.error) {
          console.error('AspenAssignments: Failed to create assignment at row', r + 1, errMsg);
        }
        result.errors.push({ row: r + 1, error: errMsg });
      }
    }

    try { applyAspenAssignmentsConditionalFormatting(); } catch (e) { /* optional styling */ }
    try { applyAspenAssignmentsCategoryValidation(sh); } catch (e) { /* optional */ }
    return result;
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

/**
 * Top-level helper to create any missing assignments from the sheet for a class.
 * Reads rows with due dates but no Aspen ID and creates them via API.
 * Shows a toast with a short summary.
 * @param {string} classId
 */
function createMissingAssignmentsFromSheet(classId) {
  const ss = SpreadsheetApp.getActive();
  const mgr = createAspenAssignmentManager(classId);
  const res = mgr.createMissingAssignmentsFromSheet();
  const msg = `Assignments: created ${res.created}, skipped ${res.skipped}, errors ${res.errors.length}`;
  try { ss.toast(msg, 'Aspen Assignments', 5); } catch (e) { /* ignore */ }
  return res;
}

/** Quick utility: reapply formatting + category validation on the Aspen Assignments sheet */
function reapplyAspenAssignmentsValidation() {
  const ss = SpreadsheetApp.getActive();
  const sh = getAspenAssignmentsSheet();
  try { applyAspenAssignmentsConditionalFormatting(); } catch (e) { if (console && console.warn) console.warn('conditional formatting apply failed', e); }
  try { applyAspenAssignmentsCategoryValidation(sh); } catch (e) { if (console && console.warn) console.warn('category validation apply failed', e); }
  try { ss.toast('Reapplied assignments formatting and category validation.', 'Aspen Assignments', 5); } catch (e) { /* ignore */ }
}