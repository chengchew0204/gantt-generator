import { unzipSync, strFromU8 } from 'fflate';

/*
 * Read native Excel cell styles directly from the xlsx byte stream. Exists
 * because SheetJS 0.20.3 Community does not surface font booleans
 * (bold / italic / underline), alignment, or colour on its wsCell.s objects
 * - the same read-side Community limitation that motivated
 * XlsxStyleInjector on the write side.
 *
 * The extractor is best-effort: any parse failure returns an empty map,
 * never throws. It reports the subset of styles the app currently
 * understands: hAlign, vAlign, bold, italic, underline, text `color`,
 * background `bg`. Other native properties (font size, borders) still
 * round-trip only through GanttGen's gridCellStyles JSON blob in the
 * Settings sheet.
 */

// Legacy BIFF palette that OOXML still references via <color indexed="N"/>.
// Values lifted from ECMA-376 Part 1, 18.8.27 (Indexed Colors). Entries 64
// and 65 are reserved for "system foreground" and "system background" and
// resolve to null (the app falls back to its own defaults).
const INDEXED_PALETTE = [
  '#000000', '#ffffff', '#ff0000', '#00ff00', '#0000ff', '#ffff00', '#ff00ff', '#00ffff',
  '#000000', '#ffffff', '#ff0000', '#00ff00', '#0000ff', '#ffff00', '#ff00ff', '#00ffff',
  '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#c0c0c0', '#808080',
  '#9999ff', '#993366', '#ffffcc', '#ccffff', '#660066', '#ff8080', '#0066cc', '#ccccff',
  '#000080', '#ff00ff', '#ffff00', '#00ffff', '#800080', '#800000', '#008080', '#0000ff',
  '#00ccff', '#ccffff', '#ccffcc', '#ffff99', '#99ccff', '#ff99cc', '#cc99ff', '#ffcc99',
  '#3366ff', '#33cccc', '#99cc00', '#ffcc00', '#ff9900', '#ff6600', '#666699', '#969696',
  '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333',
];

// Built-in fallback for theme colour references. Real Office themes live
// in xl/theme/theme1.xml; parsing that plus the full tint algebra is out
// of scope for v1. These cover the defaults the Office 2007 palette ships
// with, which Excel writes whenever a user picks a colour from the "theme
// colours" row (index 0-11) without changing the workbook theme.
const THEME_FALLBACK = {
  0: '#ffffff', // lt1 / bg1
  1: '#000000', // dk1 / tx1
  2: '#eeece1', // lt2 / bg2
  3: '#1f497d', // dk2 / tx2
  4: '#4f81bd', // accent1
  5: '#c0504d', // accent2
  6: '#9bbb59', // accent3
  7: '#8064a2', // accent4
  8: '#4bacc6', // accent5
  9: '#f79646', // accent6
  10: '#0000ff', // hlink
  11: '#800080', // folHlink
};

// Parse the attribute blob of an OOXML <color .../> or <fgColor .../>
// element. Accepts rgb="AARRGGBB" (alpha-prefixed, sometimes 6 hex when a
// producer drops alpha), indexed="N", and theme="N". Returns a lowercase
// "#rrggbb" string or null for "unset / transparent / auto".
function parseOoxmlColor(attrs) {
  if (!attrs) return null;
  const autoMatch = attrs.match(/\bauto="(1|true)"/i);
  if (autoMatch) return null;
  const rgbMatch = attrs.match(/\brgb="([0-9a-fA-F]{6,8})"/);
  if (rgbMatch) {
    const hex = rgbMatch[1].toLowerCase();
    if (hex.length === 8) {
      const alpha = hex.slice(0, 2);
      if (alpha === '00') return null;
      return '#' + hex.slice(2);
    }
    return '#' + hex;
  }
  const indexedMatch = attrs.match(/\bindexed="(\d+)"/);
  if (indexedMatch) {
    const idx = parseInt(indexedMatch[1], 10);
    return INDEXED_PALETTE[idx] || null;
  }
  const themeMatch = attrs.match(/\btheme="(\d+)"/);
  if (themeMatch) {
    const idx = parseInt(themeMatch[1], 10);
    return THEME_FALLBACK[idx] || null;
  }
  return null;
}

function parseFonts(stylesXml) {
  const fontsBlock = stylesXml.match(/<fonts\b[^>]*>([\s\S]*?)<\/fonts>/);
  if (!fontsBlock) return [];
  const inner = fontsBlock[1];
  const fonts = [];
  const fontRegex = /<font\b[^>]*>([\s\S]*?)<\/font>/g;
  let m;
  while ((m = fontRegex.exec(inner)) !== null) {
    const fontInner = m[1];
    // <u/> or <u val="single"/> / <u val="double"/> / <u val="singleAccounting"/>
    // / <u val="doubleAccounting"/> all count as underlined; <u val="none"/>
    // does not.
    const uMatch = fontInner.match(/<u\b([^/>]*)\/?>/);
    const underline = !!uMatch && !/val="none"/i.test(uMatch[1]);
    // Font colour. Skip <color theme="1"/> without tint: that is Excel's
    // default text colour (tx1 = black) and appears on every styled font
    // the XlsxStyleInjector creates, so emitting '#000000' from it would
    // pollute every bold / italic / underlined cell in a GanttGen-written
    // file with an explicit black colour that the user never chose.
    const colorElMatch = fontInner.match(/<color\b([^/>]*)\/?>/);
    let color = null;
    if (colorElMatch) {
      const colorAttrs = colorElMatch[1];
      const isDefaultText =
        /\btheme="1"/.test(colorAttrs) && !/\btint="/.test(colorAttrs);
      if (!isDefaultText) color = parseOoxmlColor(colorAttrs);
    }
    fonts.push({
      bold: /<b\s*\/?>/.test(fontInner),
      italic: /<i\s*\/?>/.test(fontInner),
      underline,
      color,
    });
  }
  return fonts;
}

// Visible-pattern fills contribute a bg colour. `none` and `gray125` are
// the two required placeholders OOXML mandates at indices 0 and 1, plus
// some files use other patterns (darkGray etc.) that semantically imply a
// shaded background; we take the fgColor for any non-`none` pattern so
// the user's visual intent survives.
function parseFills(stylesXml) {
  const fillsBlock = stylesXml.match(/<fills\b[^>]*>([\s\S]*?)<\/fills>/);
  if (!fillsBlock) return [];
  const inner = fillsBlock[1];
  const fills = [];
  // Match both the nested form `<fill><patternFill ...>...</patternFill></fill>`
  // and any trailing gradientFill etc. A fill without a <patternFill> child
  // contributes nothing; we still push an entry so fillId indices line up.
  const fillRegex = /<fill\b[^>]*>([\s\S]*?)<\/fill>/g;
  let m;
  while ((m = fillRegex.exec(inner)) !== null) {
    const fillInner = m[1];
    const patternSelfClose = fillInner.match(/<patternFill\b([^/>]*)\/\s*>/);
    const patternBlock = fillInner.match(/<patternFill\b([^>]*)>([\s\S]*?)<\/patternFill>/);
    let attrs = '';
    let body = '';
    if (patternBlock) {
      attrs = patternBlock[1];
      body = patternBlock[2];
    } else if (patternSelfClose) {
      attrs = patternSelfClose[1];
    } else {
      fills.push({ color: null });
      continue;
    }
    const typeMatch = attrs.match(/\bpatternType="([^"]+)"/);
    const type = typeMatch ? typeMatch[1] : 'none';
    if (type === 'none' || type === 'gray125') {
      fills.push({ color: null });
      continue;
    }
    const fgMatch = body.match(/<fgColor\b([^/>]*)\/?>/);
    const color = fgMatch ? parseOoxmlColor(fgMatch[1]) : null;
    fills.push({ color });
  }
  return fills;
}

function parseCellXfs(stylesXml) {
  const xfsBlock = stylesXml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/);
  if (!xfsBlock) return [];
  const inner = xfsBlock[1];
  const xfs = [];
  // Match both self-closing (<xf .../>) and block (<xf ...>...</xf>) forms.
  const xfRegex = /<xf\b([^>]*?)(?:\/>|>([\s\S]*?)<\/xf>)/g;
  let m;
  while ((m = xfRegex.exec(inner)) !== null) {
    const attrs = m[1] || '';
    const body = m[2] || '';
    const fontIdMatch = attrs.match(/\bfontId="(\d+)"/);
    const fillIdMatch = attrs.match(/\bfillId="(\d+)"/);
    const applyFontAttr = attrs.match(/\bapplyFont="([^"]+)"/);
    const applyFillAttr = attrs.match(/\bapplyFill="([^"]+)"/);
    const applyAlignAttr = attrs.match(/\bapplyAlignment="([^"]+)"/);
    const alignEl = body.match(/<alignment\b([^>]*)\/?>/);
    let hAlign;
    let vAlign;
    const hasAlignmentElement = !!alignEl;
    if (alignEl) {
      const h = alignEl[1].match(/\bhorizontal="([^"]+)"/);
      const v = alignEl[1].match(/\bvertical="([^"]+)"/);
      if (h) hAlign = h[1];
      if (v) vAlign = v[1];
    }
    xfs.push({
      fontId: fontIdMatch ? parseInt(fontIdMatch[1], 10) : 0,
      fillId: fillIdMatch ? parseInt(fillIdMatch[1], 10) : 0,
      applyFont: applyFontAttr ? applyFontAttr[1] : undefined,
      applyFill: applyFillAttr ? applyFillAttr[1] : undefined,
      applyAlignment: applyAlignAttr ? applyAlignAttr[1] : undefined,
      hasAlignmentElement,
      hAlign,
      vAlign,
    });
  }
  return xfs;
}

function resolveXfStyle(xf, fonts, fills) {
  const out = {};
  const font = fonts[xf.fontId] || null;
  // Treat applyFont as true by default. Excel explicitly opts out with
  // applyFont="0" or "false"; absent means "apply the font referenced by
  // fontId". This matches how Excel itself treats <xf> entries and covers
  // the common case where writers omit applyFont.
  const applyFont = xf.applyFont !== '0' && xf.applyFont !== 'false';
  if (applyFont && font) {
    if (font.bold) out.bold = true;
    if (font.italic) out.italic = true;
    if (font.underline) out.underline = true;
    if (font.color) out.color = font.color;
  }
  // Fills: xf 0 always references fillId 0 (pattern "none") so default
  // cells correctly emit no bg. Same applyFill opt-out convention as
  // applyFont - absent means apply, "0"/"false" opt out.
  const applyFill = xf.applyFill !== '0' && xf.applyFill !== 'false';
  const fill = fills ? fills[xf.fillId] : null;
  if (applyFill && fill && fill.color) out.bg = fill.color;
  const applyAlignment = xf.applyAlignment !== '0' && xf.applyAlignment !== 'false';
  const applyAlignmentExplicit = xf.applyAlignment === '1' || xf.applyAlignment === 'true';
  if (applyAlignment) {
    if (xf.hAlign === 'left' || xf.hAlign === 'center' || xf.hAlign === 'right') {
      out.hAlign = xf.hAlign;
    }
    if (xf.vAlign === 'top' || xf.vAlign === 'bottom') {
      out.vAlign = xf.vAlign;
    } else if (xf.vAlign === 'center') {
      out.vAlign = 'middle';
    } else if (xf.hasAlignmentElement || applyAlignmentExplicit) {
      // OOXML's default value for <alignment vertical="..."> is "bottom".
      // Excel strips that attribute on save whenever it matches the
      // default, and will sometimes strip the entire <alignment> element
      // when every attribute inside it was a default (e.g. a cell where
      // the user only changed the font on top of an already-bottom xf).
      // When applyAlignment="1" is explicitly set (not absent, not "0"),
      // the xf is asserting alignment intent; if no alignment element is
      // present that intent must be the OOXML defaults, i.e. vertical
      // bottom. Falling back to bottom here matches how Excel itself
      // renders the cell and keeps the round-trip symmetric.
      out.vAlign = 'bottom';
    }
  }
  return out;
}

function parseSheetPaths(workbookXml, relsXml) {
  const nameToRid = new Map();
  const sheetRegex = /<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"[^>]*?\/>/g;
  let m;
  while ((m = sheetRegex.exec(workbookXml)) !== null) {
    nameToRid.set(m[1], m[2]);
  }
  const ridToTarget = new Map();
  const relRegex = /<Relationship\b[^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"[^>]*?\/>/g;
  while ((m = relRegex.exec(relsXml)) !== null) {
    ridToTarget.set(m[1], m[2]);
  }
  const out = {};
  for (const [name, rid] of nameToRid) {
    const target = ridToTarget.get(rid);
    if (!target) continue;
    const path = target.startsWith('/')
      ? target.slice(1)
      : 'xl/' + target.replace(/^\.\//, '');
    out[name] = path;
  }
  return out;
}

function extractSheetCellStyles(sheetXml, xfStyles) {
  const out = {};
  // <c r="A1" s="3"/>   or   <c r="A1" s="3" t="str"><v>...</v></c>
  // Attribute order is not guaranteed, so grab both r and s independently
  // from the opening-tag attribute blob.
  const cellRegex = /<c\b([^>]*?)(?:\/>|>[\s\S]*?<\/c>)/g;
  let m;
  while ((m = cellRegex.exec(sheetXml)) !== null) {
    const attrs = m[1];
    const refMatch = attrs.match(/\br="([^"]+)"/);
    if (!refMatch) continue;
    const sMatch = attrs.match(/\bs="(\d+)"/);
    if (!sMatch) continue;
    const xfId = parseInt(sMatch[1], 10);
    const style = xfStyles[xfId];
    if (!style || Object.keys(style).length === 0) continue;
    out[refMatch[1]] = { ...style };
  }
  return out;
}

/**
 * @param {ArrayBuffer | Uint8Array} xlsxBytes
 * @returns {Record<string, Record<string, {
 *   hAlign?: 'left'|'center'|'right',
 *   vAlign?: 'top'|'middle'|'bottom',
 *   bold?: boolean,
 *   italic?: boolean,
 *   underline?: boolean,
 *   color?: string,
 *   bg?: string,
 * }>>} sheetName -> cellRef -> resolved native style (empty if nothing to report)
 */
export function extractNativeCellStyles(xlsxBytes) {
  try {
    const source = xlsxBytes instanceof Uint8Array ? xlsxBytes : new Uint8Array(xlsxBytes);
    let unzipped;
    try { unzipped = unzipSync(source); } catch { return {}; }
    if (!unzipped['xl/styles.xml']) return {};

    const stylesXml = strFromU8(unzipped['xl/styles.xml']);
    const fonts = parseFonts(stylesXml);
    const fills = parseFills(stylesXml);
    const xfs = parseCellXfs(stylesXml);
    const xfStyles = xfs.map((xf) => resolveXfStyle(xf, fonts, fills));

    const workbookXml = unzipped['xl/workbook.xml']
      ? strFromU8(unzipped['xl/workbook.xml']) : '';
    const relsXml = unzipped['xl/_rels/workbook.xml.rels']
      ? strFromU8(unzipped['xl/_rels/workbook.xml.rels']) : '';
    if (!workbookXml || !relsXml) return {};
    const sheetPaths = parseSheetPaths(workbookXml, relsXml);

    const out = {};
    for (const [sheetName, path] of Object.entries(sheetPaths)) {
      if (!unzipped[path]) continue;
      const sheetXml = strFromU8(unzipped[path]);
      const cellStyles = extractSheetCellStyles(sheetXml, xfStyles);
      if (Object.keys(cellStyles).length > 0) {
        out[sheetName] = cellStyles;
      }
    }
    return out;
  } catch {
    return {};
  }
}
