/* DevUtils.js Last Update 2025-09-04 16:27 <74ff530efabf6338a9955d5406d14101b89af70bb11c078b1980091bc33e3935>
/* eslint-disable no-unused-vars */
/* exported checkWidths, printWidthsCode, getSheetForDev */
/* global SpreadsheetApp */

/**
 * Dev helper: get a sheet by name (or active if falsy) with a friendly error.
 */
function getSheetForDev(name) {
  const ss = SpreadsheetApp.getActive();
  if (!name) return ss.getActiveSheet();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Sheet not found: ${name}`);
  return sh;
}

/**
 * Dev helper: log and return column widths for a sheet.
 * Usage: checkWidths('Grade View') or checkWidths() for the active sheet.
 */
function checkWidths(name) {
  const sh = getSheetForDev(name);
  const last = sh.getLastColumn();
  const widths = [];
  for (let c = 1; c <= last; c++) widths.push({ col: c, a1: _colToA1(c), width: sh.getColumnWidth(c) });
  const title = `Column widths for "${sh.getName()}" (1..${last})`;
  const lines = widths.map(w => `${w.a1} (${w.col}): ${w.width}px`).join('\n');
  if (console && console.log) console.log(`${title}\n${lines}`);
  return widths;
}

/**
 * Dev helper: print code to set widths, useful after manual tweaks.
 * Usage: printWidthsCode('Grade View')
 */
function printWidthsCode(name) {
  const sh = getSheetForDev(name);
  const widths = checkWidths(name);
  const sheetVar = 'sh';
  const header = `// Paste inside a function where ${sheetVar} is a Sheet for "${sh.getName()}"`;
  const code = widths.map(w => `${sheetVar}.setColumnWidth(${w.col}, ${w.width});`).join('\n');
  const out = `${header}\n${code}`;
  if (console && console.log) console.log(out);
  return out;
}

// --- local util ---
function _colToA1(n) {
  let s = '';
  while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
  return s;
}
