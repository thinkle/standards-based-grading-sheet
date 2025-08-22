/** --------- STYLE SYSTEM (edit to taste) --------- */
const COLOR_PRIMARY = '#0033a0';

function styleBase_() {
  return SpreadsheetApp.newTextStyle().setFontFamily('Roboto').setFontSize(11).build();
}
function styleH1_() {
  return SpreadsheetApp.newTextStyle().setFontFamily('Roboto').setBold(true).setFontSize(18).setForegroundColor(COLOR_PRIMARY).build();
}
function styleH2_() {
  return SpreadsheetApp.newTextStyle().setFontFamily('Roboto').setBold(true).setFontSize(15).setForegroundColor(COLOR_PRIMARY).build();
}
function styleH3_() {
  return SpreadsheetApp.newTextStyle().setFontFamily('Roboto').setBold(true).setFontSize(13).setForegroundColor(COLOR_PRIMARY).build();
}
function styleLI_() {
  return SpreadsheetApp.newTextStyle().setFontFamily('Roboto').setFontSize(11).build();
}
function styleBold_() { return SpreadsheetApp.newTextStyle().setBold(true).build(); }
function styleItalic_() { return SpreadsheetApp.newTextStyle().setItalic(true).build(); }
function styleUnder_() { return SpreadsheetApp.newTextStyle().setUnderline(true).build(); }

/** Merge multiple TextStyles into one by re-applying props (simple overlay). */
function mergeStyles_() {
  const b = SpreadsheetApp.newTextStyle();
  for (var i = 0; i < arguments.length; i++) {
    var s = arguments[i];
    if (!s) continue;
    if (s.isBold && s.isBold()) b.setBold(true);
    if (s.isItalic && s.isItalic()) b.setItalic(true);
    if (s.isUnderline && s.isUnderline()) b.setUnderline(true);
    if (s.isStrikethrough && s.isStrikethrough()) b.setStrikethrough(true);
    var fam = s.getFontFamily && s.getFontFamily(); if (fam) b.setFontFamily(fam);
    var sz = s.getFontSize && s.getFontSize(); if (sz) b.setFontSize(sz);
    var col = s.getForegroundColor && s.getForegroundColor(); if (col) b.setForegroundColor(col);
  }
  return b.build();
}

/** --------- RENDERER ---------
 * Minimal HTML subset to RichTextValue:
 * Supports: <h1|h2|h3>, <li>, <b>, <i>, <u>, <a href="">
 * Unknown tags are ignored (content kept).
 */
function setRichInstructions(range, html) {
  try {
    const rt = htmlFragmentToRichText_(html);
    range.setRichTextValue(rt).setWrap(true);
  } catch (e) {
    throw new Error('Failed to render rich instructions. Ensure tags are well-formed XML. Original error: ' + e);
  }
}

function htmlFragmentToRichText_(frag) {
  // Wrap as XML for XmlService; ensure well-formed
  const xml = XmlService.parse('<root>' + frag + '</root>');
  const root = xml.getRootElement();
  const base = styleBase_();

  const builder = SpreadsheetApp.newRichTextValue();
  var text = '';
  var runs = []; // {start, end, style, link}

  function pushText(str, style, link) {
    if (!str) return;
    var start = text.length;
    text += str;
    var end = text.length;
    if (end > start) runs.push({ start: start, end: end, style: style, link: link });
  }

  function walkElement(el, inheritedStyle, currentLink) {
    var tag = (el.getName() || '').toLowerCase();

    var style = inheritedStyle || base;
    var link = currentLink;

    // Block tags
    if (tag === 'h1') style = mergeStyles_(base, styleH1_());
    else if (tag === 'h2') style = mergeStyles_(base, styleH2_());
    else if (tag === 'h3') style = mergeStyles_(base, styleH3_());
    else if (tag === 'li') {
      style = mergeStyles_(base, styleLI_());
      pushText('• ', style, link);
    }

    // Inline tags
    if (tag === 'b' || tag === 'strong') style = mergeStyles_(style, styleBold_());
    if (tag === 'i' || tag === 'em') style = mergeStyles_(style, styleItalic_());
    if (tag === 'u') style = mergeStyles_(style, styleUnder_());
    if (tag === 'a') {
      var hrefAttr = el.getAttribute('href');
      if (hrefAttr) link = hrefAttr.getValue();
    }

    // Children elements only (GAS lacks Element.getContent())
    var kids = el.getChildren();

    if (kids.length === 0) {
      // Leaf: just text content
      pushText(el.getText(), style, link);
    } else {
      // Walk child elements; NOTE: plain text between child tags is ignored.
      for (var i = 0; i < kids.length; i++) {
        walkElement(kids[i], style, link);
      }
    }

    // Block endings → newline
    if (tag === 'h1' || tag === 'h2' || tag === 'h3' || tag === 'li') pushText('\n', base, null);
  }

  // Walk top-level children (elements only)
  var top = root.getChildren();
  for (var i = 0; i < top.length; i++) walkElement(top[i], base, null);

  builder.setText(text || '');
  for (var j = 0; j < runs.length; j++) {
    var r = runs[j];
    if (r.style) builder.setTextStyle(r.start, r.end, r.style);
    if (r.link) builder.setLinkUrl(r.start, r.end, r.link);
  }
  return builder.build();
}

/** --------- EXAMPLES --------- */
function demo_setRichInstructions() {
  const sh = SpreadsheetApp.getActive().getActiveSheet();
  const html = (
    '<h2>Getting Started</h2>' +
    '<li>Open the <b>Grades</b> sheet</li>' +
    '<li>Click <i>Setup</i> → <u>Build</u></li>' +
    '<li>See the guide: <a href="https://example.com">Docs</a></li>' +
    '<h3>Notes</h3>' +
    '<li>Use <b>B/I/U</b> for emphasis</li>'
  );
  setRichInstructions(sh.getRange('A1'), html);
}