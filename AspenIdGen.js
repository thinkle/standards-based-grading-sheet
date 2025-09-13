/* AspenIdGen.js Last Update 2025-09-12 18:19 <7275a2cb52b34b3f8642c34a2881662ffcdb58aed12b13998fd1566b4811eece>
// filepath: /Users/thinkle/BackedUpProjects/gas/standards-based-grading-sheet/AspenIdGen.js

/* A module for generating IDs for Aspen Assignments from our skill #/name combos */

/**
   * Creates a hash from a string for collision detection
   * @param {string} str - String to hash
   * @returns {string} Simple hash
   */
function simpleHash(str) {
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    const char = str.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash; // Convert to 32-bit integer
  }
  return Math.abs(hash).toString(36);
}

/**
 * Sanitizes a string for use in IDs - preserves more info to avoid conflicts
 * @param {string} str - String to sanitize
 * @param {number} [maxLength=50] - Maximum length for the sanitized string
 * @returns {string} Sanitized string with preserved structure
 */
function sanitizeForId(str, maxLength = 50) {
  let sanitized = str.toString()
    .replace(/\s+/g, '_')           // Replace spaces with underscores
    .replace(/\./g, 'DOT')          // Replace dots with 'DOT' 
    .replace(/:/g, 'COLON')         // Replace colons with 'COLON'
    .replace(/-/g, 'DASH')          // Replace dashes with 'DASH'
    .replace(/\//g, 'SLASH')        // Replace slashes with 'SLASH'
    .replace(/[^a-zA-Z0-9_]/g, 'X') // Replace other chars with 'X'
    .replace(/_+/g, '_')            // Collapse multiple underscores
    .replace(/^_|_$/g, '');         // Remove leading/trailing underscores

  // Truncate if too long, but preserve some hash for uniqueness
  if (sanitized.length > maxLength) {
    const hash = simpleHash(str).substring(0, 6);
    sanitized = sanitized.substring(0, maxLength - 7) + '_' + hash;
  }

  return sanitized;
}

/**
 * Creates a unique assignment ID from unit and skill info with length constraints
 * @param {string} classId - Class identifier
 * @param {string} unit - Unit name/identifier
 * @param {string} skill - Skill/standard identifier  
 * @param {number} [maxLength=200] - Maximum total ID length
 * @returns {string} Unique assignment ID
 */
function createAssignmentId(classId, unit, skill, maxLength = 200) {
  // Clean the inputs with appropriate length limits
  const cleanUnit = sanitizeForId(unit, 40);    // Max 40 chars for unit
  const cleanSkill = sanitizeForId(skill, 60);  // Max 60 chars for skill

  // Create a hash of the original strings for uniqueness
  const originalText = `${unit}|||${skill}`;
  const hash = simpleHash(originalText);

  // Build the ID with components
  let assignmentId = `${classId}_${cleanUnit}_${cleanSkill}_H${hash}`;

  // If still too long, use a more aggressive approach
  if (assignmentId.length > maxLength) {
    // Truncate components more aggressively and rely more on hash
    const shortUnit = sanitizeForId(unit, 20);
    const shortSkill = sanitizeForId(skill, 30);
    const longHash = simpleHash(originalText).substring(0, 12); // Longer hash for uniqueness

    assignmentId = `${classId}_${shortUnit}_${shortSkill}_H${longHash}`;

    // Final safety check
    if (assignmentId.length > maxLength) {
      // Last resort: just use class ID and a hash
      const fullHash = simpleHash(`${classId}_${originalText}`);
      assignmentId = `${classId}_AUTO_${fullHash}`;
    }
  }

  return assignmentId;
}

/**
 * Validates an assignment ID for OneRoster compliance
 * @param {string} id - ID to validate
 * @returns {Object} Validation result
 */
function validateAssignmentId(id) {
  const issues = [];

  if (id.length > 255) {
    issues.push(`Too long: ${id.length} chars (recommended max: 255)`);
  }

  if (id.length > 500) {
    issues.push(`CRITICAL: ${id.length} chars (likely database limit: 500)`);
  }

  if (!/^[a-zA-Z0-9_\-.]+$/.test(id)) {
    issues.push('Contains non-URL-safe characters');
  }

  if (id.startsWith('_') || id.endsWith('_')) {
    issues.push('Starts or ends with underscore');
  }

  return {
    valid: issues.length === 0,
    issues: issues,
    length: id.length,
    id: id
  };
}

/**
 * Test ID generation for potential conflicts
 * @param {string} classId - Class identifier
 * @param {Array} testCases - Array of {unit, skill} objects to test
 * @returns {Object} Test results
 */
function testIdGeneration(classId, testCases) {
  const results = {
    ids: {},
    conflicts: [],
    total: testCases.length
  };

  for (const testCase of testCases) {
    const id = createAssignmentId(classId, testCase.unit, testCase.skill);
    const key = `${testCase.unit} - ${testCase.skill}`;

    if (results.ids[id]) {
      results.conflicts.push({
        id: id,
        case1: results.ids[id],
        case2: key
      });
    } else {
      results.ids[id] = key;
    }
  }

  return results;
}

/**
 * Test function for ID generation conflicts
 * You can run this to test various unit/skill combinations
 */
function testAspenIdGeneration() {
  console.log('Testing Aspen ID generation for conflicts...');

  // Test cases that could cause conflicts
  const testCases = [
    { unit: "Unit A", skill: "1.11" },
    { unit: "Unit A", skill: "11.1" },
    { unit: "Unit A", skill: "1:11" },
    { unit: "Unit A", skill: "1-11" },
    { unit: "Unit A", skill: "1/11" },
    { unit: "Unit 1", skill: "A.11" },
    { unit: "Unit 1", skill: "A11" },
    { unit: "Algebra I", skill: "3.2.1" },
    { unit: "Algebra I", skill: "32.1" },
    { unit: "Algebra I", skill: "3.21" },
    { unit: "Geometry", skill: "Circle Properties" },
    { unit: "Geometry", skill: "Circle-Properties" },
    { unit: "Geometry", skill: "Circle/Properties" },
    { unit: "Unit: Fractions", skill: "Add & Subtract" },
    { unit: "Unit Fractions", skill: "Add & Subtract" },
    { unit: "Pre-Calc", skill: "Trig Functions" },
    { unit: "PreCalc", skill: "Trig Functions" },
  ];

  // Test for conflicts
  const results = testIdGeneration('TEST_CLASS', testCases);

  console.log(`Tested ${results.total} cases`);
  console.log(`Found ${results.conflicts.length} conflicts`);

  if (results.conflicts.length > 0) {
    console.log('CONFLICTS FOUND:');
    results.conflicts.forEach(conflict => {
      console.log(`  ID: ${conflict.id}`);
      console.log(`    Case 1: ${conflict.case1}`);
      console.log(`    Case 2: ${conflict.case2}`);
      console.log('');
    });
  } else {
    console.log('✅ No conflicts found!');
  }

  // Show some example IDs with validation
  console.log('\nExample generated IDs with validation:');
  testCases.slice(0, 5).forEach(testCase => {
    const id = createAssignmentId('TEST_CLASS', testCase.unit, testCase.skill);
    const validation = validateAssignmentId(id);
    console.log(`  "${testCase.unit}" + "${testCase.skill}"`);
    console.log(`    → ${id} (${validation.length} chars) ${validation.valid ? '✅' : '❌'}`);
    if (!validation.valid) {
      console.log(`    Issues: ${validation.issues.join(', ')}`);
    }
  });

  return results;
}

function createAssignmentTitle(unit, skill) {
  // Aspen auto-generates an ID from the first 10 characters of the title which we *cannot change*
  // Therefore, we need our skill and unit names to be reasonably short (at least at the start of the title).
  if (unit.length + skill.length + 3 < 10) {
    return `${unit} - ${skill}`;
  } else {
    // get short unit and skill
    unit = unit.replace(/unit/i, '').trim();
    skill = skill.replace(/standard|skill/i, '').trim();
    let shortUnit = unit;
    if (skill.length < 7) {
      shortUnit = unit.substring(0, 10 - skill.length - 1);
      return `${shortUnit}-${skill} (${unit})`;
    } else {
      // in this case, we need a short hash for the skill and unit to ensure it is unique
      const hash = simpleHash(`${unit}${skill}`).substring(0, 8);
      return `[${hash}] ${unit} - ${skill}`;
    }
  }
}

function testAssignmentTitles() {
  for (let [unit, skill] of [
    ['A', '1.11'],

    ['Algebra', '1:11'],
    ['Geometry', 'Circle Properties'],
    ['Geometry', 'Circle Properties and Attributes'],
    ['Geometry', 'Circle Properties and Attributes and Things'],
    ['Unit A', '1-11'],
    ['Unit A', '1/11'],
    ['Unit 1', 'A.11'],
    ['Geometry', '2.a.c'],
    ['Geometry', '2.a.d'],
    ['Geometry', '3.1.a'],
    ['Geometry', '2.a.c Circle Properties'],
    ['Geometry', '2.a.d Circle Properties and Attributes'],
    ['Geometry', '3.1.a Circle Properties and Attributes and Things'],

  ]) {
    const title = createAssignmentTitle(unit, skill);
    console.log(`"${unit}" + "${skill}" → ${title}`);
  }
}