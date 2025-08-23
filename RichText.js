/** --------- STYLE SYSTEM (edit to taste) --------- */
const COLOR_PRIMARY = '#0033a0';
const COLOR_SECONDARY = '#464646';
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
  return SpreadsheetApp.newTextStyle().setFontFamily('Roboto').setBold(true).setFontSize(13).setForegroundColor(COLOR_SECONDARY).build();
}
function styleH4_() {
  return SpreadsheetApp.newTextStyle().setFontFamily('Roboto').setBold(true).setFontSize(11).setForegroundColor(COLOR_PRIMARY).build();
}
function styleH5_() {
  return SpreadsheetApp.newTextStyle().setFontFamily('Roboto').setBold(true).setFontSize(11).setForegroundColor(COLOR_SECONDARY).build();
}
function styleH6_() {
  return SpreadsheetApp.newTextStyle().setFontFamily('Roboto').setBold(true).setFontSize(10).setForegroundColor(COLOR_SECONDARY).build();
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
  function pushNL() {
    if (text.length === 0 || text.charAt(text.length - 1) !== '\n') text += '\n';
  }

  function walk(node, style, link) {
    // TEXT node
    if (node.getType && node.getType() === XmlService.ContentTypes.TEXT) {
      pushText(node.getValue(), style || base, link);
      return;
    }
    // Only ELEMENT nodes beyond this point
    if (!node.getType || node.getType() !== XmlService.ContentTypes.ELEMENT) return;

    /** @type {XmlService.Element} */
    var el = node;
    var tag = (el.getName() || '').toLowerCase();

    var curStyle = style || base;
    var curLink = link;

    var isBlock = (tag === 'h1' || tag === 'h2' || tag === 'h3' || tag === 'p' || tag === 'div');

    // Block presets
    if (tag === 'h1') curStyle = mergeStyles_(base, curStyle, styleH1_());
    else if (tag === 'h2') curStyle = mergeStyles_(base, curStyle, styleH2_());
    else if (tag === 'h3') curStyle = mergeStyles_(base, curStyle, styleH3_());
    else if (tag === 'h4') curStyle = mergeStyles_(base, curStyle, styleH4_());
    else if (tag === 'h5') curStyle = mergeStyles_(base, curStyle, styleH5_());
    else if (tag === 'h6') curStyle = mergeStyles_(base, curStyle, styleH6_());
    else if (tag === 'li') curStyle = mergeStyles_(base, curStyle, styleLI_());

    // Inline overlays
    if (tag === 'b' || tag === 'strong') curStyle = mergeStyles_(curStyle, styleBold_());
    if (tag === 'i' || tag === 'em') curStyle = mergeStyles_(curStyle, styleItalic_());
    if (tag === 'u') curStyle = mergeStyles_(curStyle, styleUnder_());
    if (tag === 'a') {
      var hrefAttr = el.getAttribute('href');
      if (hrefAttr) curLink = hrefAttr.getValue();
    }

    if (isBlock) pushNL();
    if (tag === 'li') pushText('• ', curStyle, curLink);

    // Mixed-content traversal using getContentSize/getContent(i)
    var size = el.getContentSize && el.getContentSize();
    if (size && size > 0) {
      for (var i = 0; i < size; i++) {
        var child = el.getContent(i);
        if (child.getType && child.getType() === XmlService.ContentTypes.TEXT) {
          var val = child.getValue();
          if (val) pushText(val, curStyle, curLink);
        } else if (child.getType && child.getType() === XmlService.ContentTypes.ELEMENT) {
          walk(child, curStyle, curLink);
        }
      }
    } else {
      // leaf element fallback
      var txt = el.getText();
      if (txt) pushText(txt, curStyle, curLink);
    }

    if (isBlock) pushNL();
  }

  // Walk root mixed content
  var rootSize = root.getContentSize();
  for (var i = 0; i < rootSize; i++) {
    walk(root.getContent(i), base, null);
  }

  builder.setText(text || '');
  for (var j = 0; j < runs.length; j++) {
    var r = runs[j];
    if (r.style) builder.setTextStyle(r.start, r.end, r.style);
    if (r.link) builder.setLinkUrl(r.start, r.end, r.link);
  }
  return builder.build();
}