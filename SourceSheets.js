/* SourceSheets.js Last Update 2025-09-04 16:27 <189059e6ce49f04dd3ec379e3f0d06f43b3872a39760eedb4ca0df3b0ba78108>
/* eslint-disable no-unused-vars */
/* global SpreadsheetApp, STYLE */
/* ---------- SHEET NAMES ---------- */
const SHEET_STUDENTS = 'Students';
const SHEET_SKILLS = 'Skills';

/* ---------- NAMED RANGES (Students) ---------- */
const RANGE_STUDENTS = 'Students';        // full table
const RANGE_STUDENT_NAMES = 'StudentNames';
const RANGE_STUDENT_EMAILS = 'StudentEmails';

/* ---------- NAMED RANGES (Skills) ---------- */
const RANGE_SKILLS = 'Skills';       // full table
const RANGE_SKILL_UNITS = 'SkillUnits';
const RANGE_SKILL_NUMBERS = 'SkillNumbers';
const RANGE_SKILL_DESCRIPTORS = 'SkillDescriptors';

/* ---------- SETUP ---------- */
function setupStudents() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_STUDENTS) || ss.insertSheet(SHEET_STUDENTS);

  // headers
  const headers = ['Name', 'Email'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);

  // formatting
  sh.getRange('A:A').setNumberFormat('@STRING@'); // names as text
  sh.getRange('B:B').setNumberFormat('@STRING@'); // keep emails as text
  sh.autoResizeColumns(1, 2);

  // named ranges (open-ended)
  upsertNamedRange_(ss, RANGE_STUDENT_NAMES, sh.getRange('A2:A'));
  upsertNamedRange_(ss, RANGE_STUDENT_EMAILS, sh.getRange('B2:B'));
  upsertNamedRange_(ss, RANGE_STUDENTS, sh.getRange('A2:B'));
}

function setupSkills() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_SKILLS) || ss.insertSheet(SHEET_SKILLS);

  // headers
  const headers = ['Unit', 'Skill #', 'Descriptor'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);

  // formatting
  sh.getRange('A:A').setNumberFormat('@STRING@'); // Unit labels
  sh.getRange('B:B').setNumberFormat('@STRING@'); // skill id as text (preserve 1.2.3 etc.)
  sh.getRange('C:C').setNumberFormat('@STRING@'); // description
  sh.autoResizeColumns(1, 3);

  // Unit title text colors: cycle through configured UI colors based on unique Unit order
  try {
    const unitColors = (STYLE && STYLE.COLORS && STYLE.COLORS.UI && STYLE.COLORS.UI.UNIT_TEXT_COLORS) || ['#3a3a3a', '#0d47a1'];
    const nColors = Math.max(1, unitColors.length);
    const unitDataRange = sh.getRange(2, 1, Math.max(1, sh.getMaxRows() - 1), 1); // A2:A
    const unitA1 = unitDataRange.getA1Notation();
    // Remove any existing CF rules that target this exact range to avoid duplicates
    let rules = sh.getConditionalFormatRules().filter(r => !r.getRanges().some(rg => rg.getA1Notation() === unitA1));
    for (let k = 0; k < nColors; k++) {
      const f = `=IFERROR(MOD(MATCH($A2, UNIQUE($A$2:$A), 0)-1, ${nColors})=${k}, FALSE)`;
      const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(f)
        .setFontColor(unitColors[k])
        .setRanges([unitDataRange])
        .build();
      rules.push(rule);
    }
    sh.setConditionalFormatRules(rules);
  } catch (e) { /* best-effort: don't fail sheet setup */ }

  // named ranges (open-ended)
  upsertNamedRange_(ss, RANGE_SKILL_UNITS, sh.getRange('A2:A'));
  upsertNamedRange_(ss, RANGE_SKILL_NUMBERS, sh.getRange('B2:B'));
  upsertNamedRange_(ss, RANGE_SKILL_DESCRIPTORS, sh.getRange('C2:C'));
  upsertNamedRange_(ss, RANGE_SKILLS, sh.getRange('A2:C'));
}

/* helper (reuse your existing one) */
function upsertNamedRange_(ss, name, range) {
  ss.getNamedRanges().filter(n => n.getName() === name).forEach(n => n.remove());
  ss.setNamedRange(name, range);
}

// Set up conditional formatting on units based on...
// e.g. MOD(MATCH(A2,UNIQUE(A$2:A)), nUnitColors)