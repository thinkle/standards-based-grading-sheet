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
function setupGradeViewSheet(studentName) {
  const ss = SpreadsheetApp.getActive();
  const sheetName = studentName ? safeSheetName_(studentName) : SHEET_GRADE_VIEW;
  let sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
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
  sh.getRange('A1:G1').merge();
  sh.getRange('A1').setValue('Standards Based Grades').setFontSize(18).setFontWeight('bold').setHorizontalAlignment('left');

  // Student selector (merge A2:B2 for label, C2:G2 for dropdown or fixed name)
  sh.getRange('A2:B2').merge();
  sh.getRange('A2').setValue('Student').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange('C2:G2').merge();
  const selCell = sh.getRange('C2');
  const namesRange = ss.getRangeByName(RANGE_STUDENT_NAMES);
  if (studentName) {
    selCell.setValue(studentName).setFontSize(12).setHorizontalAlignment('left').setBackground('#fffbe6');
  } else {
    if (namesRange) {
      const rule = SpreadsheetApp.newDataValidation().requireValueInRange(namesRange, true).setAllowInvalid(false).build();
      selCell.setDataValidation(rule);
    }
    selCell.setFontSize(12).setHorizontalAlignment('left').setBackground('#fffbe6');
  }
  // The dropdown spans C2:G2; widths set later below

  // Hidden helper: selected email from name (place outside visible table area)
  // =IF(C2="","",INDEX(StudentEmails, MATCH(C2, StudentNames, 0)))
  sh.getRange('Z2').setFormula(`=IF($C$2="","",INDEX(${RANGE_STUDENT_EMAILS}, MATCH($C$2, ${RANGE_STUDENT_NAMES}, 0)))`);
  sh.hideColumns(26);

  // Spacer before table
  let row = 4;

  // Table headers
  const tableHeaderRow = row;
  const tableHeaders = ['Unit', 'Skill #', 'Skill Description', 'Grade', ...levelNames.map(n => `${n} Attempts`)];
  sh.getRange(tableHeaderRow, 1, 1, tableHeaders.length).setValues([tableHeaders]).setFontWeight('bold').setBackground('#f0f3f5');

  // Subheader row: put descriptions where they go
  const subHeaderRow = tableHeaderRow + 1;
  // Mastery Grade column: conditional fallback description only when relevant to selected student
  const masteryDescFormula = `=IF($Z$2="","",
    TRIM(
      IF(COUNTIF(FILTER(INDEX(Grades!A2:ZZ,,${masteryCol}), Grades!B2:B=$Z$2), ${RANGE_NONE_CORRECT_SCORE})>0,
         "If no correct attempts, score is "&${RANGE_NONE_CORRECT_SCORE}&". ", "") &
      IF(COUNTIF(FILTER(INDEX(Grades!A2:ZZ,,${masteryCol}), Grades!B2:B=$Z$2), ${RANGE_SOME_CORRECT_SCORE})>0,
         "If at least one correct, score is "&${RANGE_SOME_CORRECT_SCORE}&".", "")
    )
  )`;
  sh.getRange(subHeaderRow, 4).setFormula(masteryDescFormula).setWrap(true).setFontStyle('italic');

  // Per-level attempt column descriptors
  levelNames.forEach((_, i) => {
    const col = 5 + i;
    const descFormula = `="To earn a score of "&INDEX(${RANGE_LEVEL_SCORES},${i + 1})&", you must show proficiency "&INDEX(${RANGE_LEVEL_STREAK},${i + 1})&" times in a row."`;
    sh.getRange(subHeaderRow, col).setFormula(descFormula).setWrap(true).setFontStyle('italic');
  });

  // Table data formula (array): sorted by Unit (3) then Skill # (4), columns pulled via CHOOSECOLS
  // We filter by selected email in Z2 to disambiguate duplicate names
  const chooseCols = [3, 4, 5, masteryCol, ...symbolCols];
  const formula = `=IF($Z$2="","",
    CHOOSECOLS(
      SORT(FILTER(Grades!A2:ZZ, Grades!B2:B=$Z$2), 3, TRUE, 4, TRUE),
      ${chooseCols.join(',')}
    )
  )`;
  sh.getRange(subHeaderRow + 1, 1).setFormula(formula);

  // Pretty formatting
  sh.setFrozenRows(subHeaderRow); // freeze header + subheader
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
  const dataRowsGV = Math.max(1, sh.getMaxRows() - subHeaderRow);
  sh.getRange(subHeaderRow + 1, 2, dataRowsGV, 1).setHorizontalAlignment('center'); // Skill #
  sh.getRange(subHeaderRow + 1, 4, dataRowsGV, 1 + levelNames.length).setHorizontalAlignment('center'); // Mastery + attempts
  sh.getRange(subHeaderRow + 1, 3, dataRowsGV, 1).setWrap(true); // Description

  // Emphasize Mastery Grade: larger font for header and data
  sh.getRange(tableHeaderRow, 4).setFontSize(14).setFontWeight('bold');
  sh.getRange(subHeaderRow + 1, 4, dataRowsGV, 1).setFontSize(14).setFontWeight('bold');

  // Set requested column widths for A..G
  sh.setColumnWidth(1, 41);
  sh.setColumnWidth(2, 47);
  sh.setColumnWidth(3, 185);
  sh.setColumnWidth(4, 88);
  sh.setColumnWidth(5, 180);
  sh.setColumnWidth(6, 180);
  sh.setColumnWidth(7, 180);

  // Optional: Unit summary to the right (I:K)
  sh.getRange('I4').setValue('Unit Summary').setFontWeight('bold');
  sh.getRange('I5:K5').setValues([['Unit', 'Average Grade', 'Skills']]).setFontWeight('bold').setBackground('#f0f3f5');
  // Unit summary: average only numeric grades (ignore non-numeric like "-")
  sh.getRange('I6').setFormula(`=IF($Z$2="","",
    QUERY(
      FILTER(
        { INDEX(FILTER(Grades!A2:ZZ, Grades!B2:B=$Z$2),,3),
          IFERROR(VALUE(INDEX(FILTER(Grades!A2:ZZ, Grades!B2:B=$Z$2),,6)))
        },
        ISNUMBER(IFERROR(VALUE(INDEX(FILTER(Grades!A2:ZZ, Grades!B2:B=$Z$2),,6))))
      ),
      "select Col1, avg(Col2), count(Col2) group by Col1 order by Col1",
      0
    )
  )`);
  sh.setColumnWidth(9, 120);  // I
  sh.setColumnWidth(10, 120); // J
  sh.setColumnWidth(11, 90);  // K

  // Legacy alignment block retained for compatibility; main alignment handled above
}

// Make a safe sheet name from an arbitrary string (strip forbidden chars, trim length)
function safeSheetName_(name) {
  const cleaned = String(name).replace(/[\\/*?:[\]]/g, ' ').trim();
  return cleaned.substring(0, 99) || 'Student';
}
