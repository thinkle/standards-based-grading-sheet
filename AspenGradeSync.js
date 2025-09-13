/* AspenGradeSync.js Last Update 2025-09-13 10:30:25 <188771d7c5443cdf0c08b736b76ffbb5de11a2e0193f29532c046dcfa57d1d6a> */

/* AspenGradeSync.js Last Update 2025-09-13 10:17 <e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855>
 * Implements: Add all skill items to Aspen Assignments Tab (WIP)
 */

/**
 * Adds all skill/unit items from the Grades sheet to the Aspen Assignments tab.
 * Foregrounds the Assignments tab and shows a toast prompting the user to fill in dueDate.
 * WIP: Only scaffolding, does not yet add menu item or full logic.
 */
function addAllSkillsToAspenAssignments() {
  const ss = SpreadsheetApp.getActive();
  const gradesSheet = ss.getSheetByName('Grades');
  if (!gradesSheet) throw new Error('Grades sheet not found.');

  // Read all rows from Grades sheet (skip header)
  const data = gradesSheet.getDataRange().getValues();
  if (data.length < 2) return; // no data
  const header = data[0];
  const unitCol = header.indexOf('Unit');
  const skillNumCol = header.indexOf('Skill #');
  if (unitCol === -1 || skillNumCol === -1) throw new Error('Unit or Skill # column not found.');

  // Build set of unique (unit, skill) pairs
  const uniqueSkills = new Set();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const unit = String(row[unitCol] || '').trim();
    const skillNum = String(row[skillNumCol] || '').trim();
    if (unit && skillNum) uniqueSkills.add(`${unit}|||${skillNum}`);
  }

  // Use AspenAssignments.js helper to get/create the sheet
  const assignmentsSheet = getAspenAssignmentsSheet();
  const assignmentsData = assignmentsSheet.getDataRange().getValues();
  const assignmentsHeader = assignmentsData[0];
  const aUnitCol = assignmentsHeader.indexOf('Unit');
  const aSkillCol = assignmentsHeader.indexOf('Skill');
  if (aUnitCol === -1 || aSkillCol === -1) throw new Error('Unit or Skill column not found in Aspen Assignments.');

  // Build Skills sheet maps to normalize existing rows that may still hold descriptors
  const skillsSheet = ss.getSheetByName('Skills');
  let byUnitDesc = {};
  if (skillsSheet) {
    const sData = skillsSheet.getDataRange().getValues();
    if (sData.length >= 2) {
      const sHead = sData[0];
      const uI = sHead.indexOf('Unit');
      const nI = sHead.indexOf('Skill #');
      const dI = sHead.indexOf('Descriptor');
      if (uI !== -1 && nI !== -1 && dI !== -1) {
        for (let i = 1; i < sData.length; i++) {
          const u = String(sData[i][uI] || '').trim();
          const n = String(sData[i][nI] || '').trim();
          const d = String(sData[i][dI] || '').trim();
          if (u && n && d) byUnitDesc[`${u}|||${d}`] = n;
        }
      }
    }
  }

  // Build set of existing (unit, skillNum) pairs in assignments
  const existingAssignments = new Set();
  for (let i = 1; i < assignmentsData.length; i++) {
    const row = assignmentsData[i];
    const unit = String(row[aUnitCol] || '').trim();
    let skillVal = String(row[aSkillCol] || '').trim();
    if (!unit || !skillVal) continue;
    // Normalize: if this looks like a descriptor we can map, convert to its Skill #
    const mappedNum = byUnitDesc[`${unit}|||${skillVal}`];
    const normalized = mappedNum || skillVal;
    existingAssignments.add(`${unit}|||${normalized}`);
  }

  // Add missing (unit, skill) pairs to Aspen Assignments
  let addedCount = 0;
  uniqueSkills.forEach(pair => {
    if (!existingAssignments.has(pair)) {
      const [unit, skillNum] = pair.split('|||');
      // Add a new row with unit and Skill # as the Skill column
      const newRow = new Array(assignmentsHeader.length).fill('');
      newRow[aUnitCol] = unit;
      newRow[aSkillCol] = skillNum;
      assignmentsSheet.appendRow(newRow);
      addedCount++;
    }
  });

  // Foreground the Assignments tab and show a toast
  ss.setActiveSheet(assignmentsSheet);
  ss.toast(
    addedCount > 0
      ? `Added ${addedCount} new skill(s) to Aspen Assignments (using Skill #). Please fill in dueDate to create assignments.`
      : 'No new skills to add. All skills already present.',
    'Aspen Assignments',
    5
  );
}

/**
 * Adds one "Unit Average" item per unique Unit from the Grades sheet into the Aspen Assignments tab.
 * If a (Unit, Skill) pair where Skill=="Unit Average" already exists, it won't be duplicated.
 */
function addUnitAveragesToAspenAssignments() {
  const ss = SpreadsheetApp.getActive();
  const gradesSheet = ss.getSheetByName('Grades');
  if (!gradesSheet) throw new Error('Grades sheet not found.');

  const data = gradesSheet.getDataRange().getValues();
  if (data.length < 2) return; // no data
  const header = data[0];
  const unitCol = header.indexOf('Unit');
  if (unitCol === -1) throw new Error('Unit column not found.');

  // Collect unique units (skip blanks)
  const uniqueUnits = new Set();
  for (let i = 1; i < data.length; i++) {
    const unit = String(data[i][unitCol] || '').trim();
    if (unit) uniqueUnits.add(unit);
  }

  // Target Assignments sheet and build lookup of existing (unit, skill)
  const assignmentsSheet = getAspenAssignmentsSheet();
  const assignmentsData = assignmentsSheet.getDataRange().getValues();
  if (assignmentsData.length === 0) return;
  const aHeader = assignmentsData[0];
  const aUnitCol = aHeader.indexOf('Unit');
  const aSkillCol = aHeader.indexOf('Skill');
  if (aUnitCol === -1 || aSkillCol === -1) throw new Error('Unit or Skill column not found in Aspen Assignments.');

  const existing = new Set();
  for (let i = 1; i < assignmentsData.length; i++) {
    const row = assignmentsData[i];
    const unit = String(row[aUnitCol] || '').trim();
    const skill = String(row[aSkillCol] || '').trim();
    if (unit && skill) existing.add(`${unit}|||${skill}`);
  }

  // Add a row per unit for skill = "Unit Average" if missing
  let added = 0;
  const MAGIC_SKILL = 'Unit Average';
  uniqueUnits.forEach(unit => {
    const key = `${unit}|||${MAGIC_SKILL}`;
    if (!existing.has(key)) {
      const newRow = new Array(aHeader.length).fill('');
      newRow[aUnitCol] = unit;
      newRow[aSkillCol] = MAGIC_SKILL;
      assignmentsSheet.appendRow(newRow);
      added++;
    }
  });

  ss.setActiveSheet(assignmentsSheet);
  ss.toast(
    added > 0
      ? `Added ${added} Unit Average item(s). Fill in dueDate to create assignments.`
      : 'No new unit averages to add. All present.',
    'Aspen Assignments',
    5
  );
}
