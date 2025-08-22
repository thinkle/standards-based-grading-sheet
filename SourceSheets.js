/* ---------- SHEET NAMES ---------- */
const SHEET_STUDENTS = 'Students';
const SHEET_SKILLS   = 'Skills';

/* ---------- NAMED RANGES (Students) ---------- */
const RANGE_STUDENTS        = 'Students';        // full table
const RANGE_STUDENT_NAMES   = 'StudentNames';
const RANGE_STUDENT_EMAILS  = 'StudentEmails';

/* ---------- NAMED RANGES (Skills) ---------- */
const RANGE_SKILLS             = 'Skills';       // full table
const RANGE_SKILL_UNITS        = 'SkillUnits';
const RANGE_SKILL_NUMBERS      = 'SkillNumbers';
const RANGE_SKILL_DESCRIPTORS  = 'SkillDescriptors';

/* ---------- SETUP ---------- */
function setupStudents() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_STUDENTS) || ss.insertSheet(SHEET_STUDENTS);

  // headers
  const headers = ['Name','Email'];
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  sh.setFrozenRows(1);

  // formatting
  sh.getRange('A:A').setNumberFormat('@STRING@'); // names as text
  sh.getRange('B:B').setNumberFormat('@STRING@'); // keep emails as text
  sh.autoResizeColumns(1, 2);

  // named ranges (open-ended)
  upsertNamedRange_(ss, RANGE_STUDENT_NAMES,  sh.getRange('A2:A'));
  upsertNamedRange_(ss, RANGE_STUDENT_EMAILS, sh.getRange('B2:B'));
  upsertNamedRange_(ss, RANGE_STUDENTS,       sh.getRange('A2:B'));
}

function setupSkills() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_SKILLS) || ss.insertSheet(SHEET_SKILLS);

  // headers
  const headers = ['Unit','Skill #','Descriptor'];
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  sh.setFrozenRows(1);

  // formatting
  sh.getRange('A:A').setNumberFormat('@STRING@'); // Unit labels
  sh.getRange('B:B').setNumberFormat('0');        // numeric skill id
  sh.getRange('C:C').setNumberFormat('@STRING@'); // description
  sh.autoResizeColumns(1, 3);

  // named ranges (open-ended)
  upsertNamedRange_(ss, RANGE_SKILL_UNITS,       sh.getRange('A2:A'));
  upsertNamedRange_(ss, RANGE_SKILL_NUMBERS,     sh.getRange('B2:B'));
  upsertNamedRange_(ss, RANGE_SKILL_DESCRIPTORS, sh.getRange('C2:C'));
  upsertNamedRange_(ss, RANGE_SKILLS,            sh.getRange('A2:C'));
}

/* helper (reuse your existing one) */
function upsertNamedRange_(ss, name, range) {
  ss.getNamedRanges().filter(n => n.getName() === name).forEach(n => n.remove());
  ss.setNamedRange(name, range);
}