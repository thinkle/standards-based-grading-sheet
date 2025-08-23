/* eslint-disable no-unused-vars */
/* exported generateStudentViews, shareStudentViews */
/* global SpreadsheetApp, DriveApp, Session, setupGradeViewSheet, RANGE_STUDENT_NAMES, RANGE_STUDENT_EMAILS, safeSheetName_ */

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
      const sh = childSS.getActiveSheet();
      sh.setName('View');
      // Import the local tab via IMPORTRANGE (teacher will authorize once when opening)
      const importFormula = `=IMPORTRANGE("${parentUrl}", "${tabName}!A1:G")`;
      sh.getRange('A1').setFormula(importFormula);
    } else {
      // Keep any existing file organized in the folder
      try { childFile.moveTo(childFolder); } catch (e) { if (console && console.warn) console.warn('Move existing child to folder warning', e); }
    }
    results.push({ name, email: p.email, url: childFile.getUrl() });
  });

  // Build/update index with teacher-facing warning and external links only
  const idxName = 'Student Views';
  let idx = ss.getSheetByName(idxName) || ss.insertSheet(idxName);
  idx.clear();
  idx.getRange('A1').setValue('List of student view links: give students/families READ or COMMENT ACCESS only. Do NOT grant edit access or they can potentially see the grades of peers. Open a view once to authorize IMPORTRANGE. You can share from the menu.');
  idx.getRange('A1:C1').merge().setFontWeight('bold').setBackground('#fff3cd').setWrap(true);
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
