import { unzipSync, strFromU8 } from 'fflate';

/*
 * Read native Excel cell styles directly from the xlsx byte stream. Exists
 * because SheetJS 0.20.3 Community does not surface font booleans
 * (bold / italic / underline) or alignment on its wsCell.s objects - the
 * same read-side Community limitation that motivated XlsxStyleInjector on
 * the write side.
 *
 * The extractor is best-effort: any parse failure returns an empty map,
 * never throws. It only reports the subset of styles the app currently
 * understands (hAlign, vAlign, bold, italic, underline); other native
 * properties (font colour, fill, borders, font size) still round-trip
 * through GanttGen's gridCellStyles JSON blob in the Settings sheet.
 */

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
    fonts.push({
      bold: /<b\s*\/?>/.test(fontInner),
      italic: /<i\s*\/?>/.test(fontInner),
      underline,
    });
  }
  return fonts;
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
    const applyFontAttr = attrs.match(/\bapplyFont="([^"]+)"/);
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
      applyFont: applyFontAttr ? applyFontAttr[1] : undefined,
      applyAlignment: applyAlignAttr ? applyAlignAttr[1] : undefined,
      hasAlignmentElement,
      hAlign,
      vAlign,
    });
  }
  return xfs;
}

function resolveXfStyle(xf, fonts) {
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
  }
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
    const xfs = parseCellXfs(stylesXml);
    const xfStyles = xfs.map((xf) => resolveXfStyle(xf, fonts));

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
