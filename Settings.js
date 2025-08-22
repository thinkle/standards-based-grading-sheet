/** 
 * Standards Based Grading Sheet
 * 
 * The Goal of this sheet:
 * -> Set up standards based grading where teachers grade students for mastery. 
 * - The idea is for each standard, students have various chances to "attempt" the problem at various levels.
 *   Students need to accumulate a "streak" of success before they demonstrate "mastery" 
 * In our main spreadsheet, we then have little rows like this...
 * 
 * Skill | Score | Basic Streak | Interm Streak | Adv Streak | Basic String | Interm String | Adv String | B | B | B | B | I | I | I | I | A | A | A |A |A |A
 * Adding | 2    | 0011         |  00101011     | 001010
 */

/* ---------- CONSTANTS ---------- */
// Sheet names
const SHEET_SYMBOLS       = 'Symbols';
const SHEET_LEVELSETTINGS = 'LevelSettings';

// Named ranges (Symbols)
const RANGE_SYMBOL_CHARS   = 'SymbolChars';
const RANGE_SYMBOL_MASTERY = 'SymbolMastery';
const RANGE_SYMBOL_SYMBOL  = 'SymbolSymbol';
const RANGE_SYMBOLS        = 'Symbols';

// Named ranges (Level Settings)
const RANGE_LEVEL_NAMES           = 'LevelNames';
const RANGE_LEVEL_SHORTCODES      = 'LevelShortCodes';
const RANGE_LEVEL_DEFAULTATTEMPTS = 'LevelDefaultAttempts';
const RANGE_LEVEL_STREAK          = 'LevelStreakForMastery';
const RANGE_LEVELSETTINGS         = 'LevelSettings';
const RANGE_LEVEL_SCORES        = 'LevelScores';
const RANGE_FALLBACK_LABELS     = 'FallbackLabels';
const RANGE_FALLBACK_SCORES     = 'FallbackScores';
// NEW named ranges for individual fallback scores
const RANGE_NONE_CORRECT_SCORE = 'NoneCorrectScore';
const RANGE_SOME_CORRECT_SCORE = 'SomeCorrectScore';

/* ---------- SETUP FUNCTIONS ---------- */
function setupNamedRanges () {
  setupSymbols();
  setupLevelSettings();
  setupSettings();
}




function setupLevelSettings () {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_LEVELSETTINGS) || ss.insertSheet(SHEET_LEVELSETTINGS);

  // headers
  const headers = ['Level Name','Short Code','Default Attempts','Streak for Mastery','Score','','Fallback','Score'];
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  sh.setFrozenRows(1);

  // seed rows if empty
  if (sh.getLastRow() < 2) {
    sh.getRange(2,1,3,5).setValues([
      ['Basic','B',5,2,2],
      ['Intermediate','I',5,2,3],
      ['Mastery','M',5,2,4],
    ]);
    sh.getRange(2,7,2,2).setValues([
      ['None correct',0],
      ['Some correct',1],
    ]);
  }

  sh.getRange('A:B').setNumberFormat('@STRING@');
  sh.getRange('C:E').setNumberFormat('0');
  sh.getRange('G:G').setNumberFormat('@STRING@');
  sh.getRange('H:H').setNumberFormat('0');
  sh.autoResizeColumns(1, 8);

  // level ranges
  upsertNamedRange_(ss, RANGE_LEVEL_NAMES,           sh.getRange('A2:A'));
  upsertNamedRange_(ss, RANGE_LEVEL_SHORTCODES,      sh.getRange('B2:B'));
  upsertNamedRange_(ss, RANGE_LEVEL_DEFAULTATTEMPTS, sh.getRange('C2:C'));
  upsertNamedRange_(ss, RANGE_LEVEL_STREAK,          sh.getRange('D2:D'));
  upsertNamedRange_(ss, RANGE_LEVEL_SCORES,          sh.getRange('E2:E'));
  upsertNamedRange_(ss, RANGE_LEVELSETTINGS,         sh.getRange('A2:E'));

  // fallback ranges
  upsertNamedRange_(ss, RANGE_FALLBACK_LABELS,  sh.getRange('G2:G'));
  upsertNamedRange_(ss, RANGE_FALLBACK_SCORES,  sh.getRange('H2:H'));
  // individual named ranges
  upsertNamedRange_(ss, RANGE_NONE_CORRECT_SCORE, sh.getRange('H2'));
  upsertNamedRange_(ss, RANGE_SOME_CORRECT_SCORE, sh.getRange('H3'));
}

function setupSymbols () {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_SYMBOLS) || ss.insertSheet(SHEET_SYMBOLS);

  // headers
  const desired = ['Character','Mastery','Symbol'];
  const hdr = sh.getRange(1, 1, 1, 3).getValues()[0];
  for (let i = 0; i < 3; i++) if (!hdr[i]) hdr[i] = desired[i];
  sh.getRange(1, 1, 1, 3).setValues([hdr]);
  sh.setFrozenRows(1);

  // seed rows if empty
  if (sh.getLastRow() < 2) {
    sh.getRange(2, 1, 2, 3).setValues([
      ['✓', 1, '✅'],
      ['X', 0, '❌'],
    ]);
  }

  // named ranges (open-ended)
  upsertNamedRange_(ss, RANGE_SYMBOL_CHARS,   sh.getRange('A2:A'));
  upsertNamedRange_(ss, RANGE_SYMBOL_MASTERY, sh.getRange('B2:B'));
  upsertNamedRange_(ss, RANGE_SYMBOL_SYMBOL,  sh.getRange('C2:C'));
  upsertNamedRange_(ss, RANGE_SYMBOLS,        sh.getRange('A2:C'));
}

function setupSettings () {
  console.log('fix me');
}

/* ---------- HELPER ---------- */
function upsertNamedRange_(ss, name, range) {
  ss.getNamedRanges().filter(n => n.getName() === name).forEach(n => n.remove());
  ss.setNamedRange(name, range);
}