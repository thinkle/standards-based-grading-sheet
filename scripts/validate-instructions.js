#!/usr/bin/env node
/**
 * Validate the HTML/XML fragments passed to setRichInstructions in Instructions.js
 * without needing to push to Apps Script. We extract each template string argument
 * and check it parses as XML (wrapped in a root node), reporting the exact line/column
 * on failure.
 */
const fs = require('fs');
const path = require('path');

const root = process.cwd();
const file = path.join(root, 'Instructions.js');

function extractTemplates(js) {
  const out = [];
  // naive extraction: find setRichInstructions( ... , `...` ) blocks
  // This handles backtick template strings possibly spanning multiple lines.
  const regex = /setRichInstructions\s*\(\s*[^,]+,\s*`([\s\S]*?)`\s*\)/g;
  let m;
  while ((m = regex.exec(js))) {
    const content = m[1];
    // Find line number of the match start
    const startIdx = m.index;
    const before = js.slice(0, startIdx);
    const line = before.split(/\r?\n/).length;
    out.push({ content, line });
  }
  return out;
}

function validateXml(fragment) {
  // Use a minimal DOM parser built into Node? Not reliable; instead use a basic check using XMLSerializer in a JSDOM
  // To avoid deps, we rely on a very small well-formedness heuristic with an actual parser via Intl.Segmenter fallback.
  // Better: use a tiny dependency-less parser with DOMParser emulation via xmldom. But we promised no install.
  // We'll bundle a tiny xmldom-like parser inline by trying require('xmldom'), and if missing, give a helpful message.
  try {
    const { DOMParser } = require('xmldom');
    const parser = new DOMParser({ errorHandler: { warning: () => { }, error: (e) => { throw new Error(e); }, fatalError: (e) => { throw new Error(e); } } });
    const xml = `<root>${fragment}</root>`;
    const doc = parser.parseFromString(xml, 'text/xml');
    const errs = doc.getElementsByTagName('parsererror');
    if (errs && errs.length) {
      throw new Error(errs[0].textContent || 'Invalid XML');
    }
    return null;
  } catch (e) {
    if (e.code === 'MODULE_NOT_FOUND') {
      return new Error('Missing dependency xmldom. Run "npm i -D xmldom" and retry.');
    }
    return e instanceof Error ? e : new Error(String(e));
  }
}

(function main() {
  const js = fs.readFileSync(file, 'utf8');
  const templates = extractTemplates(js);
  if (templates.length === 0) {
    console.log('No setRichInstructions templates found.');
    process.exit(0);
  }
  let failed = 0;
  templates.forEach(({ content, line }, i) => {
    const err = validateXml(content);
    if (err) {
      failed++;
      console.error(`\n[Fragment ${i + 1}] Starting near Instructions.js:${line}`);
      console.error(err.message);
    }
  });
  if (failed) {
    console.error(`\n${failed} fragment(s) failed XML validation.`);
    process.exit(1);
  } else {
    console.log(`All ${templates.length} fragment(s) are well-formed XML.`);
  }
})();
