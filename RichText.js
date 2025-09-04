/* RichText.js Last Update 2025-09-04 16:27 <59a35831b9eab9545765bddce8ea4f14897d0241d61a3d62453cffd1cc083630>
/* eslint-disable no-unused-vars */
/* exported setRichInstructions, getCellUrl, getSheetUrl, escapeBareAmpersands_ */
/* global SpreadsheetApp, XmlService, STYLE */
/** --------- STYLE SYSTEM (edit to taste) --------- */
const COLOR_PRIMARY = (typeof STYLE !== 'undefined' && STYLE.COLORS && STYLE.COLORS.BRAND_PRIMARY) ? STYLE.COLORS.BRAND_PRIMARY : '#0033a0';
const COLOR_SECONDARY = (typeof STYLE !== 'undefined' && STYLE.COLORS && STYLE.COLORS.BRAND_SECONDARY) ? STYLE.COLORS.BRAND_SECONDARY : '#464646';
const FONT_FAMILY = (typeof STYLE !== 'undefined' && STYLE.FONT_FAMILY) ? STYLE.FONT_FAMILY : 'Roboto';
const FONT_SIZE = (typeof STYLE !== 'undefined' && STYLE.FONT_SIZE) ? Number(STYLE.FONT_SIZE) : 11;
const FONT_SIZE_XL = (typeof STYLE !== 'undefined' && STYLE.FONT_SIZE_XLARGE) ? Number(STYLE.FONT_SIZE_XLARGE) : 18;
const FONT_SIZE_LG = (typeof STYLE !== 'undefined' && STYLE.FONT_SIZE_LARGE) ? Number(STYLE.FONT_SIZE_LARGE) : 15;
const FONT_SIZE_SM = (typeof STYLE !== 'undefined' && STYLE.FONT_SIZE_SMALL) ? Number(STYLE.FONT_SIZE_SMALL) : 8;
function styleBase_() { return SpreadsheetApp.newTextStyle().setFontFamily(FONT_FAMILY).setFontSize(FONT_SIZE).build(); }
function styleH1_() { return SpreadsheetApp.newTextStyle().setFontFamily(FONT_FAMILY).setBold(true).setFontSize(FONT_SIZE_XL).setForegroundColor(COLOR_PRIMARY).build(); }
function styleH2_() { return SpreadsheetApp.newTextStyle().setFontFamily(FONT_FAMILY).setBold(true).setFontSize(FONT_SIZE_LG).setForegroundColor(COLOR_PRIMARY).build(); }
function styleH3_() { return SpreadsheetApp.newTextStyle().setFontFamily(FONT_FAMILY).setBold(true).setFontSize(13).setForegroundColor(COLOR_SECONDARY).build(); }
function styleH4_() { return SpreadsheetApp.newTextStyle().setFontFamily(FONT_FAMILY).setBold(true).setFontSize(FONT_SIZE).setForegroundColor(COLOR_PRIMARY).build(); }
function styleH5_() { return SpreadsheetApp.newTextStyle().setFontFamily(FONT_FAMILY).setBold(true).setFontSize(FONT_SIZE).setForegroundColor(COLOR_SECONDARY).build(); }
function styleH6_() { return SpreadsheetApp.newTextStyle().setFontFamily(FONT_FAMILY).setBold(true).setFontSize(10).setForegroundColor(COLOR_SECONDARY).build(); }
function styleLI_() { return SpreadsheetApp.newTextStyle().setFontFamily(FONT_FAMILY).setFontSize(FONT_SIZE).build(); }
function styleBold_() { return SpreadsheetApp.newTextStyle().setBold(true).build(); }
function styleItalic_() { return SpreadsheetApp.newTextStyle().setItalic(true).build(); }
function styleUnder_() { return SpreadsheetApp.newTextStyle().setUnderline(true).build(); }
function styleSmall_() { return SpreadsheetApp.newTextStyle().setFontFamily(FONT_FAMILY).setFontSize(FONT_SIZE_SM).build(); }

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

/**
 * Escape bare ampersands that are not part of a valid XML/HTML entity.
 * This prevents XmlService.parse from choking when template interpolation
 * inserts URLs like ...&range=... without &amp;.
 *
 * Matches any '&' not followed by one of:
 *  - named entity: &word;
 *  - decimal entity: &#1234;
 *  - hex entity: &#x1F4A9;
 */
function escapeBareAmpersands_(s) {
  if (s == null) return '';
  return String(s).replace(/&(?!#\d+;|#x[0-9a-fA-F]+;|[a-zA-Z][a-zA-Z0-9]+;)/g, '&amp;');
}

function setRichInstructions(range, html) {
  try {
    // Pre-sanitize common entity issues from template interpolation
    const safe = escapeBareAmpersands_(html);
    const rt = htmlFragmentToRichText_(safe);
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
    if (tag === 'small') curStyle = mergeStyles_(curStyle, styleSmall_());
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

function getCellUrl(range, local = true) {
  // Get a Rich Text URL for referring to the current range
  const sheet = range.getSheet();
  const sheetId = sheet.getSheetId();
  const rangeA1 = range.getA1Notation();
  let hashUrl = `#gid=${sheetId}&amp;range=${rangeA1}`;
  if (local) {
    return hashUrl;
  } else {
    return `https://docs.google.com/spreadsheets/d/${SpreadsheetApp.getActiveSpreadsheet().getId()}/edit${hashUrl}`;
  }
}

function getSheetUrl(sheet, local = true) {
  // Get a Rich Text URL for referring to the current sheet
  const sheetId = sheet.getSheetId();
  if (local) {
    return `#gid=${sheetId}`;
  } else {
    return `https://docs.google.com/spreadsheets/d/${SpreadsheetApp.getActiveSpreadsheet().getId()}/edit#gid=${sheetId}`;
  }
}