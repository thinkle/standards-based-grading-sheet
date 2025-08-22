/* eslint-disable no-unused-vars */
/* global SpreadsheetApp,
          RANGE_STUDENT_NAMES, RANGE_STUDENT_EMAILS,
          RANGE_LEVEL_NAMES, RANGE_LEVEL_STREAK, RANGE_LEVEL_SCORES,
          RANGE_NONE_CORRECT_SCORE, RANGE_SOME_CORRECT_SCORE */

const SHEET_GRADE_VIEW = 'Grade View';

/**
 * Build or rebuild a read-friendly Grade View sheet.
 * - Student dropdown (by name)
 * - Explanations for each level and fallback rules
 * - Sorted table of Unit, Skill #, Description, Mastery Grade, and per-level symbol strings
 */
function setupGradeViewSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_GRADE_VIEW) || ss.insertSheet(SHEET_GRADE_VIEW);
  sh.clear();

  // Read settings arrays
  const levelNames = ss.getRangeByName(RANGE_LEVEL_NAMES).getValues().flat().filter(String);
  const levelStreak = ss.getRangeByName(RANGE_LEVEL_STREAK).getValues().flat().slice(0, levelNames.length);
  const levelScores = ss.getRangeByName(RANGE_LEVEL_SCORES).getValues().flat().slice(0, levelNames.length);

  // Layout constants matching Grades sheet
  const baseHeadersCount = 6; // Name, Email, Unit, Skill #, Skill Description, Mastery Grade
  const firstUtilCol = baseHeadersCount + 1; // util starts after base
  const masteryCol = baseHeadersCount; // 6
  const symbolCols = levelNames.map((_, i) => firstUtilCol + i * 3 + 2); // per-level symbols column in Grades

  // Title
  sh.getRange('A1').setValue('Grade View').setFontSize(18).setFontWeight('bold');

  // Student selector
  sh.getRange('A2').setValue('Student').setFontWeight('bold');
  const selCell = sh.getRange('B2');
  const namesRange = ss.getRangeByName(RANGE_STUDENT_NAMES);
  if (namesRange) {
    const rule = SpreadsheetApp.newDataValidation().requireValueInRange(namesRange, true).setAllowInvalid(false).build();
    selCell.setDataValidation(rule);
  }
  selCell.setFontSize(12).setHorizontalAlignment('left').setBackground('#fffbe6');
  sh.setColumnWidth(2, 240); // B: student dropdown width

  // Hidden helper: selected email from name
  // =IF(B2="","",INDEX(StudentEmails, MATCH(B2, StudentNames, 0)))
  sh.getRange('C2').setFormula(`=IF($B$2="","",INDEX(${RANGE_STUDENT_EMAILS}, MATCH($B$2, ${RANGE_STUDENT_NAMES}, 0)))`);
  sh.hideColumns(3);

  // Explanations header and lines
  let row = 4;
  sh.getRange(row, 1).setValue('What mastery means').setFontWeight('bold');
  row++;
  levelNames.forEach((name, i) => {
    const line = `${name}: To master this level and score a ${levelScores[i]}, you must earn credit ${levelStreak[i]} times in a row.`;
    sh.getRange(row++, 1).setValue(line);
  });
  // Fallback rules line using named ranges so it stays dynamic
  sh.getRange(row++, 1).setFormula(`="If no correct attempts, score is "&${RANGE_NONE_CORRECT_SCORE}&". If at least one correct, score is "&${RANGE_SOME_CORRECT_SCORE}&"."`);

  row += 1; // spacer before table

  // Table headers
  const tableHeaderRow = row;
  const tableHeaders = ['Unit', 'Skill #', 'Skill Description', 'Mastery Grade', ...levelNames.map(n => `${n} Attempts`)];
  sh.getRange(tableHeaderRow, 1, 1, tableHeaders.length).setValues([tableHeaders]).setFontWeight('bold').setBackground('#f0f3f5');

  // Table data formula (array): sorted by Unit (3) then Skill # (4), columns pulled via CHOOSECOLS
  // We filter by selected email in C2 to disambiguate duplicate names
  const chooseCols = [3, 4, 5, masteryCol, ...symbolCols];
  const formula = `=IF($C$2="","",
    CHOOSECOLS(
      SORT(FILTER(Grades!A2:ZZ, Grades!B2:B=$C$2), 3, TRUE, 4, TRUE),
      ${chooseCols.join(',')}
    )
  )`;
  sh.getRange(tableHeaderRow + 1, 1).setFormula(formula);

  // Pretty formatting
  sh.setFrozenRows(tableHeaderRow); // freeze through header area
  // Column widths
  sh.setColumnWidth(1, 90);  // Unit
  sh.setColumnWidth(2, 70);  // Skill #
  sh.setColumnWidth(3, 360); // Description
  sh.setColumnWidth(4, 110); // Mastery Grade
  for (let i = 0; i < levelNames.length; i++) {
    sh.setColumnWidth(5 + i, 180);
  }
  // Banding for the table region (large height to cover dynamic rows)
  try {
    sh.getRange(tableHeaderRow, 1, 2000, tableHeaders.length).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  } catch (e) {
    if (console && console.warn) console.warn('Banding apply warning', e);
  }

  // Alignments for readability
  const dataRows = Math.max(1, sh.getMaxRows() - tableHeaderRow);
  sh.getRange(tableHeaderRow + 1, 2, dataRows, 1).setHorizontalAlignment('center'); // Skill #
  sh.getRange(tableHeaderRow + 1, 4, dataRows, 1 + levelNames.length).setHorizontalAlignment('center'); // Mastery + attempts
  sh.getRange(tableHeaderRow + 1, 3, dataRows, 1).setWrap(true); // Description
}
