/* eslint-disable no-unused-vars */
/* exported generateStudentViews, shareStudentViews */
/* global SpreadsheetApp, DriveApp, Session, STYLE, setupGradeViewSheet, RANGE_STUDENT_NAMES, RANGE_STUDENT_EMAILS, safeSheetName_ */

/**
 * Step 1: Generate per-student tabs AND external view docs (no sharing yet).
 * - Creates/refreshes a hidden tab per student via setupGradeViewSheet(name)
 * - Creates/reuses an external spreadsheet per student that IMPORTRANGEs that tab
 * - Writes a teacher-facing index with external View URLs (no internal tab links)
 */
function generateStudentViews() {
  const ss = SpreadsheetApp.getActive();
  const names = ss.getRangeByName(RANGE_STUDENT_NAMES).getValues().flat().filter(String);
  const emails = ss.getRangeByName(RANGE_STUDENT_EMAILS).getValues().flat();
  const pairs = names.map((n, i) => ({ name: n, email: emails[i] || '' })).filter(p => p.name);

  // Ensure organized folder for per-student files
  const parentId = ss.getId();
  const parentName = ss.getName();
  const parentUrl = ss.getUrl();
  const parentFile = DriveApp.getFileById(parentId);
  const parentFolder = parentFile.getParents().hasNext() ? parentFile.getParents().next() : DriveApp.getRootFolder();
  const childFolderName = `${parentName} — Student Views`;
  const childFolder = ensureOrCreateFolder_(parentFolder, childFolderName);

  const results = [];
  pairs.forEach(p => {
    const name = p.name;
    // Create/refresh student tab and hide it
    setupGradeViewSheet(name);
    const tabName = typeof safeSheetName_ === 'function' ? safeSheetName_(name) : name;
    const tab = ss.getSheetByName(tabName);
    if (tab && !tab.isSheetHidden()) tab.hideSheet();

    // Create/reuse external child file inside organized folder
    const childName = `${parentName} — ${name}`;
    let childFile = findChildFile_(childName, childFolder);
    if (!childFile) {
      const child = SpreadsheetApp.create(childName);
      childFile = DriveApp.getFileById(child.getId());
      try { childFile.moveTo(childFolder); } catch (e) { if (console && console.warn) console.warn('Move child file warning', e); }
      const childSS = SpreadsheetApp.openById(child.getId());
      // Copy the fully formatted per-student tab to child, then clear contents and set IMPORTRANGE
      const viewSh = ensureChildViewFromSource_(childSS, tab, parentUrl, tabName);
    } else {
      // Keep any existing file organized in the folder
      try { childFile.moveTo(childFolder); } catch (e) { if (console && console.warn) console.warn('Move existing child to folder warning', e); }
      // Rebuild the child 'View' sheet from the source each time to guarantee formatting
      try {
        const childSS = SpreadsheetApp.openById(childFile.getId());
        ensureChildViewFromSource_(childSS, tab, parentUrl, tabName);
      } catch (e) {
        if (console && console.warn) console.warn('Rebuild child view warn', e);
        throw e;
      }
    }
    results.push({ name, email: p.email, url: childFile.getUrl() });
  });

  // Build/update index with teacher-facing warning and external links only
  const idxName = 'Student Views';
  let idx = ss.getSheetByName(idxName) || ss.insertSheet(idxName);
  idx.clear();
  // Default fonts
  try {
    idx.getRange(1, 1, Math.max(1, idx.getMaxRows()), Math.max(1, idx.getMaxColumns()))
      .setFontFamily(STYLE.FONT_FAMILY)
      .setFontSize(Number(STYLE.FONT_SIZE));
  } catch (e) { /* STYLE optional */ }
  idx.getRange('A1').setValue('List of student view links: give students/families READ or COMMENT ACCESS only. Do NOT grant edit access or they can potentially see the grades of peers. Open a view once to authorize IMPORTRANGE. You can share from the menu.');
  const warnBg = (typeof STYLE !== 'undefined' && STYLE.COLORS && STYLE.COLORS.UI && STYLE.COLORS.UI.WARNING_BG) || '#fff3cd';
  idx.getRange('A1:C1').merge().setFontWeight('bold').setBackground(warnBg).setWrap(true);
  idx.getRange(2, 1, 1, 3).setValues([['Name', 'Email', 'View URL']]).setFontWeight('bold');
  if (results.length) idx.getRange(3, 1, results.length, 3).setValues(results.map(r => [r.name, r.email, r.url]));
  idx.setFrozenRows(2);
}

/**
 * Step 2: Share the already-generated external view docs (comment-only to student emails).
 * Reads URLs from the index; does not create files or modify contents.
 */
function shareStudentViews() {
  const ss = SpreadsheetApp.getActive();
  const idx = ss.getSheetByName('Student Views');
  if (!idx) throw new Error('Run "Generate student views" first.');
  const last = idx.getLastRow();
  if (last < 3) return; // need at least one data row

  // Ensure status column
  const statusHeader = idx.getRange(2, 4).getValue();
  if (statusHeader !== 'Shared At') idx.getRange(2, 4).setValue('Shared At').setFontWeight('bold');

  const values = idx.getRange(3, 1, last - 2, 3).getValues();
  values.forEach((r, i) => {
    const [name, email, url] = r;
    if (!name || !url || !email) return;
    const fileId = extractFileIdFromUrl_(url);
    if (!fileId) return;
    try {
      const childFile = DriveApp.getFileById(fileId);
      childFile.addCommenter(email);
      idx.getRange(3 + i, 4).setValue(new Date());
    } catch (e) {
      if (console && console.warn) console.warn('Share add commenter warning', e);
    }
  });
}

function findChildFile_(name, inFolder) {
  const files = inFolder ? inFolder.getFilesByName(name) : DriveApp.getFilesByName(name);
  while (files.hasNext()) {
    const f = files.next();
    if (f.getMimeType() === MimeType.GOOGLE_SHEETS) return f;
  }
  return null;
}

function ensureOrCreateFolder_(parentFolder, folderName) {
  try {
    const it = parentFolder.getFoldersByName(folderName);
    if (it.hasNext()) return it.next();
    return parentFolder.createFolder(folderName);
  } catch (e) {
    if (console && console.warn) console.warn('ensureOrCreateFolder_ warning', e);
    return parentFolder;
  }
}

function extractFileIdFromUrl_(url) {
  if (!url) return '';
  const m = String(url).match(/\/d\/([a-zA-Z0-9_-]{10,})/);
  return m ? m[1] : '';
}

/**
 * Minimal formatting for external 'View' sheet to mirror Grade View readability.
 * We can't import styles via IMPORTRANGE, so set header banding, widths, and alignment.
 */
function formatExternalViewSheet_(sh) {
  if (!sh) return;
  // Set widths roughly matching Grade View
  sh.setColumnWidth(1, 90);  // Unit
  sh.setColumnWidth(2, 70);  // Skill #
  sh.setColumnWidth(3, 360); // Description
  sh.setColumnWidth(4, 110); // Grade
  // Attempts columns a bit wider for wrapped text if present
  for (let i = 5; i <= Math.min(9, sh.getMaxColumns()); i++) sh.setColumnWidth(i, 180);
  // Alignments
  const maxRows = Math.max(1, sh.getMaxRows() - 2);
  sh.getRange(3, 2, maxRows, 1).setHorizontalAlignment('center');
  sh.getRange(3, 4, maxRows, Math.max(1, Math.min(9, sh.getMaxColumns()) - 3)).setHorizontalAlignment('center');
  sh.getRange(3, 3, maxRows, 1).setWrap(true);

  // Apply subtle colors and striping similar to Grade View (best-effort; IMPORTRANGE blocks direct cell background copy)
  try {
    const neutralBg = (STYLE && STYLE.COLORS && STYLE.COLORS.UI && STYLE.COLORS.UI.NEUTRAL_BG) || '#f7f7f7';
    const neutralBgAlt = (STYLE && STYLE.COLORS && STYLE.COLORS.UI && STYLE.COLORS.UI.NEUTRAL_BG_ALT) || '#f0f0f0';
    const neutralText = (STYLE && STYLE.COLORS && STYLE.COLORS.UI && STYLE.COLORS.UI.NEUTRAL_TEXT) || '#333333';
    // Data area starts at row 3 (headers are 1 and 2)
    const dataStart = 3;
    const dataCount = Math.max(1, sh.getMaxRows() - 2);
    // Neutral stripe on A..C
    sh.getRange(dataStart, 1, dataCount, 3).setBackground(neutralBg).setFontColor(neutralText);
    let rules = sh.getConditionalFormatRules();
    const targets = [];
    const neutralRange = sh.getRange(dataStart, 1, dataCount, 3);
    targets.push(neutralRange.getA1Notation());
    const neutralStripe = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=ISEVEN(ROW())')
      .setBackground(neutralBgAlt)
      .setRanges([neutralRange])
      .build();
    // Attempts columns 5.. up to present
    const maxCols = Math.min(12, sh.getMaxColumns());
    for (let col = 5; col <= maxCols; col++) {
      // Use subtle neutral family since STYLE.LEVELS may not map here reliably through IMPORTRANGE
      const r = sh.getRange(dataStart, col, dataCount, 1);
      r.setBackground(neutralBg).setFontColor(neutralText);
      targets.push(r.getA1Notation());
      const stripe = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=ISEVEN(ROW())')
        .setBackground(neutralBgAlt)
        .setRanges([r])
        .build();
      rules.push(stripe);
    }
    // Gradient-like readability for Grade column (we cannot safely replicate exact thresholds without named ranges here)
    // So we just ensure white text for any non-empty cell in column 4 for contrast if a theme applies
    const gradeRange = sh.getRange(dataStart, 4, dataCount, 1);
    targets.push(gradeRange.getA1Notation());
    const gradeText = SpreadsheetApp.newConditionalFormatRule()
      .whenCellNotEmpty()
      .setFontColor((STYLE && STYLE.COLORS && STYLE.COLORS.GRADE_SCALE && STYLE.COLORS.GRADE_SCALE.TEXT) || '#ffffff')
      .setRanges([gradeRange])
      .build();
    // Dedup existing rules for these targets
    rules = rules.filter(r => !r.getRanges().some(rg => targets.includes(rg.getA1Notation())));
    rules.push(neutralStripe, gradeText);
    sh.setConditionalFormatRules(rules);
  } catch (e) {
    if (console && console.warn) console.warn('formatExternalViewSheet_ colors warn', e);
  }
}

/**
 * Copy formatting (styles, merges, widths, frozen rows, banding) from a source Grade View tab
 * to a target external 'View' sheet. Content is left as-is (IMPORTED). Best-effort.
 */
function mirrorFormattingFromSource_(srcSh, dstSh) {
  if (!srcSh || !dstSh) return;
  // Clear existing banding on destination to avoid duplicates
  try { dstSh.getBandings().forEach(b => b.remove()); } catch (e) { if (console && console.warn) console.warn('mirrorFormatting remove banding warn', e); }

  // Determine a reasonable block to copy (up to used rows/cols from source)
  const srcLastRow = Math.max(srcSh.getLastRow(), 50);
  const srcLastCol = Math.max(srcSh.getLastColumn(), 10);
  // Copy formatting only
  srcSh.getRange(1, 1, srcLastRow, srcLastCol).copyTo(dstSh.getRange(1, 1), { formatOnly: true });

  // Mirror frozen rows/cols
  try { dstSh.setFrozenRows(srcSh.getFrozenRows()); } catch (e) { if (console && console.warn) console.warn('mirrorFormatting frozen rows warn', e); }
  try { dstSh.setFrozenColumns(srcSh.getFrozenColumns()); } catch (e) { if (console && console.warn) console.warn('mirrorFormatting frozen cols warn', e); }

  // Mirror column widths
  for (let c = 1; c <= srcLastCol; c++) {
    try { dstSh.setColumnWidth(c, srcSh.getColumnWidth(c)); } catch (e) { if (console && console.warn) console.warn('mirrorFormatting set width warn', e); }
  }

  // Ensure description wraps and center certain columns similar to Grade View
  const maxRows = Math.max(1, dstSh.getMaxRows() - 2);
  try { dstSh.getRange(3, 2, maxRows, 1).setHorizontalAlignment('center'); } catch (e) { if (console && console.warn) console.warn('mirrorFormatting align skill# warn', e); }
  try { dstSh.getRange(3, 4, maxRows, Math.max(1, srcLastCol - 3)).setHorizontalAlignment('center'); } catch (e) { if (console && console.warn) console.warn('mirrorFormatting align grade/attempts warn', e); }
  try { dstSh.getRange(3, 3, maxRows, 1).setWrap(true); } catch (e) { if (console && console.warn) console.warn('mirrorFormatting wrap desc warn', e); }
}

/**
 * Create or refresh a child 'View' sheet by copying the fully formatted source tab,
 * then clearing values (keeping formats) and inserting the IMPORTRANGE.
 */
function ensureChildViewFromSource_(childSS, srcTab, parentUrl, tabName) {
  if (!childSS || !srcTab) return null;
  // Always insert a temporary visible sheet so we never delete the last visible sheet
  let temp = null;
  try { temp = childSS.insertSheet(); } catch (e) { if (console && console.warn) console.warn('insert temp sheet warn', e); }

  // Remove existing 'View' if present
  const existing = childSS.getSheetByName('View');
  if (existing) {
    try { childSS.deleteSheet(existing); } catch (e) { if (console && console.warn) console.warn('delete existing View warn', e); }
  }

  // Copy the source tab (formats, widths, frozen rows) into child
  const copied = srcTab.copyTo(childSS);
  try { copied.setName('View'); } catch (e) { if (console && console.warn) console.warn('rename copied to View warn', e); }
  try { copied.showSheet(); } catch (e) { if (console && console.warn) console.warn('show View sheet warn', e); }

  // Clean up any other sheets except the new 'View' and the temp one (delete temp last)
  const toDelete = childSS.getSheets().filter(s => s.getName() !== 'View' && (!temp || s.getSheetId() !== temp.getSheetId()));
  for (const sh of toDelete) {
    try { childSS.deleteSheet(sh); } catch (e) { if (console && console.warn) console.warn('delete extra sheet warn', e); }
  }
  // Now delete the temp sheet if it still exists
  if (temp) {
    try { childSS.deleteSheet(temp); } catch (e) { /* if this fails, we still have a valid View */ }
  }
  // Clear values only, keep formatting
  try { copied.clear({ contentsOnly: true }); } catch (e) { if (console && console.warn) console.warn('clear contents warn', e); }
  // Set IMPORTRANGE at A1 to bring in live data
  const importFormula = `=IMPORTRANGE("${parentUrl}", "${tabName}!A1:G")`;
  copied.getRange('A1').setFormula(importFormula);
  // Apply best-effort formatting to the external sheet
  try { formatExternalViewSheet_(copied); } catch (e) { if (console && console.warn) console.warn('format external after import warn', e); }
  return copied;
}
