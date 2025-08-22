const SHEET_GRADES = 'Grades';

function setupGradesSheet() {
  const ss = SpreadsheetApp.getActive();

  // Read level settings once
  const codes   = ss.getRangeByName(RANGE_LEVEL_SHORTCODES).getValues().flat().filter(String);
  const names   = ss.getRangeByName(RANGE_LEVEL_NAMES).getValues().flat().slice(0, codes.length);
  const streaks = ss.getRangeByName(RANGE_LEVEL_STREAK).getValues().flat().slice(0, codes.length);
  const scores  = ss.getRangeByName(RANGE_LEVEL_SCORES).getValues().flat().slice(0, codes.length);

  const fallbackScores = ss.getRangeByName(RANGE_FALLBACK_SCORES).getValues().flat().filter(v => v !== '');
  const scoreNone = RANGE_NONE_CORRECT_SCORE;
  const scoreSome = RANGE_SOME_CORRECT_SCORE;

  let sh = ss.getSheetByName(SHEET_GRADES) || ss.insertSheet(SHEET_GRADES);
  sh.clear();

  // Build headers
  const baseHeaders = ['Name','Email','Skill','Mastery Grade'];
  const utilHeaders = codes.flatMap((c, i) => ([`${names[i]} Streak`, `${names[i]} String`]));
  const attemptHeaders = codes.flatMap((c, i) => Array.from({length:  Number(ss.getRangeByName(RANGE_LEVEL_DEFAULTATTEMPTS).getValues()[i][0] || 0)}, () => c));
  const headers = [...baseHeaders, ...utilHeaders, ...attemptHeaders];
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);

  // Coordinates
  const firstUtilCol = baseHeaders.length + 1;
  const firstAttemptCol = baseHeaders.length + utilHeaders.length + 1;

  // Build shared attempt ranges for row 2 (header and row)
  const lastCol = headers.length;
  const attemptHeaderA1 = `${columnA1(firstAttemptCol)}1:${columnA1(lastCol)}1`;
  const attemptRowA1    = `${columnA1(firstAttemptCol)}2:${columnA1(lastCol)}2`;

  // Put per-level STRING and STREAK formulas in row 2
  codes.forEach((code, i) => {
    const streakCol = firstUtilCol + i*2;
    const stringCol = streakCol + 1;

    // String: map symbols->mastery bits for columns whose header starts with ^code
    const stringFormula =
      `=TEXTJOIN("",TRUE,ARRAYFORMULA(` +
      `XLOOKUP(FILTER(${attemptRowA1}, REGEXMATCH(${attemptHeaderA1}, "^"&"${code}")), ${RANGE_SYMBOL_CHARS}, ${RANGE_SYMBOL_MASTERY}, "-")` +
      `))`;

    sh.getRange(2, stringCol).setFormula(stringFormula);

    // Streak: longest run of 1s in that string
    const stringCellA1 = `${columnA1(stringCol)}2`;
    const streakFormula = `=IF(${stringCellA1}="","",MAX(ARRAYFORMULA(LEN(SPLIT(${stringCellA1},"0",FALSE,FALSE)))))`;
    sh.getRange(2, streakCol).setFormula(streakFormula);
  });

  // Mastery Grade formula (generated IFS using settings)
  // Check: any attempts entered?
  const attemptsRowConcat = `TEXTJOIN("",TRUE,${attemptRowA1})`;
  const noneCorrectCheck  = `ISERROR(SEARCH("1", TEXTJOIN("", TRUE, {${codes.map((_,i)=>columnA1(firstUtilCol+i*2+1)+'2').join(',')}} )))`;

  // Build ordered IFS parts from highest level to lowest (so highest mastered wins).
  const parts = codes.map((c,i)=>({
    cond: `${columnA1(firstUtilCol + i*2)}2>=INDEX(${RANGE_LEVEL_STREAK},${i+1})`,
    val:  `INDEX(${RANGE_LEVEL_SCORES},${i+1})`
  })).reverse();

  const ifs =
    `=IFS(` +
    // nothing entered?
    `COUNTA(${attemptRowA1})=0,"-",` +
    // per-level thresholds (highest first)
    parts.map(p => `${p.cond},${p.val}`).join(',') + (parts.length ? ',' : '') +
    // none correct?
    `${noneCorrectCheck},${scoreNone},` +
    // some correct
    `TRUE,${scoreSome}` +
    `)`;

  sh.getRange(2, 4).setFormula(ifs); // Mastery Grade is col 4
}

/* --- tiny util --- */
function columnA1(n) {
  let s = '';
  while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
  return s;
}