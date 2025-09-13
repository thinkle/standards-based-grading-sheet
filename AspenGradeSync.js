/* AspenGradeSync.js Last Update 2025-09-13 11:43 <12d38998e82624cc75c21940469a3cd5acd8c4a53e813d6db20c81038845b0a2>

/* AspenGradeSync.js Last Update 2025-09-13 12:05 <sync-grades-impl>
 * Implements: Add all skill items to Aspen Assignments Tab, Unit Average helper, and Grade Sync (skill & unit-average modes)
 */
/* global SpreadsheetApp, createGradeSyncManager, getAspenAssignments */

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

// ---------------- Grade Sync helpers ----------------

/**
 * Read the Grades sheet once and build fast lookups.
 * Returns { rows, byEmailUnitSkill, byEmailUnit, header, indices, symbolCols }.
 */
function readGradesSnapshot_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Grades');
  if (!sh) throw new Error('Grades sheet not found.');

  const data = sh.getDataRange().getValues();
  if (data.length < 2) {
    return { rows: [], byEmailUnitSkill: new Map(), byEmailUnit: new Map(), header: [], indices: {}, symbolCols: [] };
  }
  const header = data[0];
  const idx = {
    name: header.indexOf('Name'),
    email: header.indexOf('Email'),
    unit: header.indexOf('Unit'),
    skillNum: header.indexOf('Skill #'),
    desc: header.indexOf('Skill Description'),
    mastery: header.indexOf('Mastery Grade')
  };
  ['email','unit','skillNum','mastery'].forEach(k => { if (idx[k] === -1) throw new Error(`Grades header missing: ${k}`); });

  // Detect per-level Symbols columns by pattern: [.. Streak], [.. String], [LevelName]
  const symbolCols = [];
  let col = idx.mastery + 1; // start after Mastery Grade
  while (col + 2 < header.length) {
    const h0 = String(header[col] || '');
    const h1 = String(header[col + 1] || '');
    const h2 = String(header[col + 2] || '');
    if (h0.endsWith(' Streak') && h1.endsWith(' String') && h2 && !/^\s*$/.test(h2)) {
      symbolCols.push(col + 2);
      col += 3;
    } else {
      break;
    }
  }

  const rows = [];
  const byEmailUnitSkill = new Map(); // key email|unit|skill#
  const byEmailUnit = new Map();      // key email|unit -> array of row objects
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const email = String(r[idx.email] || '').trim();
    const unit = String(r[idx.unit] || '').trim();
    const skillNum = r[idx.skillNum] != null && r[idx.skillNum] !== '' ? String(r[idx.skillNum]) : '';
    const desc = idx.desc !== -1 ? String(r[idx.desc] || '').trim() : '';
    const sVal = r[idx.mastery];
    const score = (typeof sVal === 'number') ? sVal : (sVal != null && sVal !== '' && !isNaN(Number(sVal)) ? Number(sVal) : null);
    // Build per-level symbols comment block
    const levelParts = [];
    symbolCols.forEach(ci => {
      const label = String(header[ci] || '').trim(); // Level name
      const val = String(r[ci] || '').trim();
      if (label && val) levelParts.push(`${label}: ${val}`);
    });
    const levelComment = levelParts.join(' | ');

    if (!email) continue; // require student
    const row = { email, unit, skillNum, desc, score, levelComment };
    rows.push(row);
    if (unit && skillNum) {
      byEmailUnitSkill.set(`${email}|${unit}|${skillNum}`, row);
    }
    if (unit) {
      const keyEU = `${email}|${unit}`;
      if (!byEmailUnit.has(keyEU)) byEmailUnit.set(keyEU, []);
      byEmailUnit.get(keyEU).push(row);
    }
  }

  return { rows, byEmailUnitSkill, byEmailUnit, header, indices: idx, symbolCols };
}

/**
 * Sync grades in Skill mode: one assignment per (Unit, Skill#) per student.
 * Reads Grades once and posts only changed grades.
 */
function syncGradesSkillMode(classId) {
  const mgr = createGradeSyncManager(classId);
  const snapshot = readGradesSnapshot_();
  const assignments = getAspenAssignments();

  // Build lookup for assignments by (unit, skill#)
  const byUnitSkill = new Map();
  assignments.forEach(a => {
    const unit = String(a.unit || '').trim();
    const skill = String(a.skill || '').trim();
    const id = String(a.assignmentId || '').trim();
    if (unit && skill && id && skill !== 'Unit Average') byUnitSkill.set(`${unit}|${skill}`, id);
  });

  let attempted = 0, synced = 0, skipped = 0, errors = 0;
  snapshot.rows.forEach(row => {
    if (!row.unit || !row.skillNum) return;
    const assignmentId = byUnitSkill.get(`${row.unit}|${row.skillNum}`);
    if (!assignmentId) return;
    if (row.score == null) return; // nothing to sync
    const comment = row.levelComment || '';
    attempted++;
    const res = mgr.maybeSync(row.email, assignmentId, row.score, comment);
    if (res && res.success) {
      if (res.synced) synced++; else skipped++;
    } else {
      errors++;
    }
  });

  return { mode: 'skill', attempted, synced, skipped, errors };
}

/**
 * Sync grades in Unit Average mode: one assignment "Unit Average" per Unit per student.
 * Calculates per-student unit averages and posts only changed grades.
 */
function syncGradesUnitAverageMode(classId) {
  const mgr = createGradeSyncManager(classId);
  const snapshot = readGradesSnapshot_();
  const assignments = getAspenAssignments();

  // Map Unit -> assignmentId for Unit Average entries
  const unitAvgMap = new Map();
  assignments.forEach(a => {
    const unit = String(a.unit || '').trim();
    const skill = String(a.skill || '').trim();
    const id = String(a.assignmentId || '').trim();
    if (unit && id && skill === 'Unit Average') unitAvgMap.set(unit, id);
  });

  let attempted = 0, synced = 0, skipped = 0, errors = 0;

  // For each (email, unit) compute average and comment
  snapshot.byEmailUnit.forEach((rows, key) => {
    const [email, unit] = key.split('|');
    const assignmentId = unitAvgMap.get(unit);
    if (!assignmentId) return; // no Unit Average assignment yet
    // Collect numeric scores and build breakdown
    const scored = rows.filter(r => r.score != null);
    if (scored.length === 0) return;
    const avg = scored.reduce((s, r) => s + Number(r.score), 0) / scored.length;
    const breakdown = scored.map(r => `${r.skillNum || '?'}: ${r.desc || ''} - ${r.score}`).join('\n');
    attempted++;
    const res = mgr.maybeSync(email, assignmentId, avg, breakdown);
    if (res && res.success) {
      if (res.synced) synced++; else skipped++;
    } else {
      errors++;
    }
  });

  return { mode: 'unit-average', attempted, synced, skipped, errors };
}
