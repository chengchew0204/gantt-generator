import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';

/*
 * Post-process an xlsx byte stream produced by SheetJS Community to embed
 * native Excel cell styles that SheetJS Community does not write itself.
 *
 * Scope:
 *   - Horizontal / vertical alignment (via <alignment> on <xf>)
 *   - Bold / italic / underline (via additional <font> entries and fontId
 *     / applyFont on <xf>)
 *
 * Strategy: leave SheetJS's output intact, then append the minimum number
 * of <font> and <xf> entries to xl/styles.xml and rewrite the relevant
 * <c> elements in each xl/worksheets/sheetN.xml with the resolved
 * s="..." attribute. Fonts and xfs are deduplicated; a cell with no
 * effective style is left untouched.
 *
 * The injector is deliberately conservative: if sheetStyles is empty or
 * styles.xml is missing / unparseable, it returns the input bytes
 * unchanged.
 */

const H_VALID = new Set(['left', 'center', 'right']);
const V_APP_TO_OOXML = { top: 'top', middle: 'center', bottom: 'bottom' };

function escapeRegex(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function normalizeStyle(raw) {
  if (!raw) return null;
  const out = {};
  if (raw.bold) out.bold = true;
  if (raw.italic) out.italic = true;
  if (raw.underline) out.underline = true;
  if (raw.hAlign && H_VALID.has(raw.hAlign)) out.hAlign = raw.hAlign;
  if (raw.vAlign && V_APP_TO_OOXML[raw.vAlign]) out.vAlign = raw.vAlign;
  if (!out.bold && !out.italic && !out.underline && !out.hAlign && !out.vAlign) {
    return null;
  }
  return out;
}

function fontKey(s) {
  return `${s.bold ? 'b' : ''}${s.italic ? 'i' : ''}${s.underline ? 'u' : ''}`;
}

function hasFontStyle(s) {
  return !!(s.bold || s.italic || s.underline);
}

function alignKey(s) {
  return `${s.hAlign || ''}|${s.vAlign || ''}`;
}

function hasAlignment(s) {
  return !!(s.hAlign || s.vAlign);
}

function xfKey(s, fontId) {
  return `f${fontId}|${alignKey(s)}`;
}

function buildFontEntry(defaultInner, s) {
  const markers = [];
  if (s.bold) markers.push('<b/>');
  if (s.italic) markers.push('<i/>');
  if (s.underline) markers.push('<u/>');
  return `<font>${markers.join('')}${defaultInner}</font>`;
}

function buildXfEntry(s, fontId) {
  const applyFont = fontId > 0;
  const applyAlign = hasAlignment(s);
  const attrs = [
    'numFmtId="0"',
    `fontId="${fontId}"`,
    'fillId="0"',
    'borderId="0"',
    'xfId="0"',
  ];
  if (applyFont) attrs.push('applyFont="1"');
  if (applyAlign) attrs.push('applyAlignment="1"');
  const open = `<xf ${attrs.join(' ')}`;
  if (!applyAlign) return `${open}/>`;
  const alignAttrs = [];
  if (s.hAlign) alignAttrs.push(`horizontal="${s.hAlign}"`);
  // Always emit a vertical attribute when we emit <alignment>. OOXML's
  // default for vertical is "bottom", so Excel strips an explicit
  // vertical="bottom" on save whenever it matches the default. If we do
  // not write vertical at all, GanttGen renders the cell centered (the
  // Tailwind flex default) while Excel renders it bottom - and any
  // subsequent re-save by Excel would hide the user's bottom selection
  // in exactly the same way. Writing vertical="center" when vAlign is
  // undefined locks in GanttGen's visual centered default and keeps the
  // attribute out of the "gets-stripped-because-it-is-the-default" zone.
  const voox = s.vAlign && V_APP_TO_OOXML[s.vAlign]
    ? V_APP_TO_OOXML[s.vAlign]
    : 'center';
  alignAttrs.push(`vertical="${voox}"`);
  return `${open}><alignment ${alignAttrs.join(' ')}/></xf>`;
}

function parseSheetPaths(workbookXml, relsXml) {
  const nameToRid = new Map();
  // Use [^>]*? to skip between attributes: attribute values can contain `/`
  // (notably <Relationship Type="http://..."/>). Bounding by `>` keeps the
  // match inside one element.
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

function patchCellStyle(sheetXml, cellRef, xfId) {
  const safeRef = escapeRegex(cellRef);
  const cellRegex = new RegExp(`<c r="${safeRef}"([^/>]*)(/?)>`, 'g');
  return sheetXml.replace(cellRegex, (_match, attrs, slash) => {
    const cleanAttrs = attrs.replace(/\s+s="\d+"/, '');
    return `<c r="${cellRef}"${cleanAttrs} s="${xfId}"${slash}>`;
  });
}

function patchDefaultXfAlignment(stylesXml) {
  // Target only the first <xf ...> inside <cellXfs>. Styles.xml also
  // contains a <cellStyleXfs> block (usually right before <cellXfs>); we
  // must not touch that one. Using a bounded regex anchored on <cellXfs>
  // keeps the replacement safe.
  const cellXfsBlock = stylesXml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/);
  if (!cellXfsBlock) return stylesXml;
  const inner = cellXfsBlock[1];
  const firstXfMatch = inner.match(/<xf\b([^>]*?)(\/>|>[\s\S]*?<\/xf>)/);
  if (!firstXfMatch) return stylesXml;
  const attrs = firstXfMatch[1] || '';

  // Already has an <alignment> child - leave as-is (a previous patch run
  // or an external tool already set something meaningful).
  if (/<xf\b[^>]*>\s*<alignment\b/.test(firstXfMatch[0])) return stylesXml;

  // Rebuild attrs to ensure applyAlignment="1" is present without
  // duplicating it. Strip any existing applyAlignment= then append our
  // own value.
  const cleanedAttrs = attrs.replace(/\s+applyAlignment="[^"]*"/g, '').trimEnd();
  const newXf = `<xf${cleanedAttrs ? ' ' + cleanedAttrs.trim() : ''} applyAlignment="1"><alignment vertical="center"/></xf>`;
  const patchedInner = inner.replace(firstXfMatch[0], newXf);
  return stylesXml.replace(cellXfsBlock[0], `<cellXfs${cellXfsBlock[0].match(/<cellXfs\b([^>]*)>/)[1]}>${patchedInner}</cellXfs>`);
}

function extractDefaultFontInner(stylesXml) {
  // Grab the inner content of the FIRST <font> element so we can clone its
  // size / color / family / scheme when building bold / italic / underline
  // variants. Falls back to Excel-typical defaults if the styles.xml shape
  // is unexpected.
  const fontsBlock = stylesXml.match(/<fonts\b[^>]*>([\s\S]*?)<\/fonts>/);
  if (fontsBlock) {
    const firstFont = fontsBlock[1].match(/<font\b[^>]*>([\s\S]*?)<\/font>/);
    if (firstFont) return firstFont[1];
  }
  return '<sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/>';
}

/**
 * @param {ArrayBuffer | Uint8Array} xlsxBytes - raw xlsx bytes from XLSX.write
 * @param {Record<string, Record<string, {
 *   hAlign?: 'left'|'center'|'right',
 *   vAlign?: 'top'|'middle'|'bottom',
 *   bold?: boolean,
 *   italic?: boolean,
 *   underline?: boolean,
 * }>>} sheetStyles
 *   Keys are sheet display names (same as passed to XLSX.utils.book_append_sheet
 *   after sanitizeSheetName). Values are { [cellRef]: <style subset> }.
 * @returns {Uint8Array} the patched xlsx bytes (a new buffer when patched,
 *   the input buffer when nothing needs to change)
 */
export function injectCellStyles(xlsxBytes, sheetStyles) {
  const source = xlsxBytes instanceof Uint8Array ? xlsxBytes : new Uint8Array(xlsxBytes);

  let unzipped;
  try {
    unzipped = unzipSync(source);
  } catch {
    return source;
  }

  if (!unzipped['xl/styles.xml']) return source;

  // Collect effective per-cell styles bucketed by sheet. Drop cells with
  // no recognised style so we do not spend xfIds on no-ops.
  const perSheetEffective = {};
  if (sheetStyles) {
    for (const [sheetName, cellMap] of Object.entries(sheetStyles)) {
      const filtered = {};
      for (const [cellRef, raw] of Object.entries(cellMap)) {
        const norm = normalizeStyle(raw);
        if (!norm) continue;
        filtered[cellRef] = norm;
      }
      if (Object.keys(filtered).length > 0) perSheetEffective[sheetName] = filtered;
    }
  }

  let stylesXml = strFromU8(unzipped['xl/styles.xml']);

  // Patch the default cellXfs entry (xf index 0) so cells that SheetJS
  // writes without an `s` attribute render vertically centered in Excel,
  // matching GanttGen's Tailwind `items-center` default. OOXML's default
  // vertical alignment is bottom, so without this patch Excel would show
  // unstyled cells at the bottom even though GanttGen shows them
  // centered - a consistent visual mismatch the user reported. Patching
  // xf 0 keeps the change zero-cost on disk (no new xf entries) and
  // does not affect the import round-trip because the extractor skips
  // cells that have no `s` attribute (SheetJS never emits `s="0"`
  // explicitly for default-styled cells).
  stylesXml = patchDefaultXfAlignment(stylesXml);

  const cellXfsMatch = stylesXml.match(/<cellXfs\b[^>]*count="(\d+)"[^>]*>/);
  const fontsMatch = stylesXml.match(/<fonts\b[^>]*count="(\d+)"[^>]*>/);
  if (!cellXfsMatch || !fontsMatch) {
    // Default-xf patch still succeeded; keep it but skip further work.
    unzipped['xl/styles.xml'] = strToU8(stylesXml);
    return zipSync(unzipped);
  }

  const currentXfCount = parseInt(cellXfsMatch[1], 10);
  const currentFontCount = parseInt(fontsMatch[1], 10);
  const defaultFontInner = extractDefaultFontInner(stylesXml);

  // Font dedup: one entry per unique (bold, italic, underline) combination.
  // The all-false combination reuses fontId 0 (existing default).
  const fontRegistry = new Map();
  const newFontChunks = [];
  let nextFontId = currentFontCount;
  function resolveFontId(style) {
    if (!hasFontStyle(style)) return 0;
    const k = fontKey(style);
    if (fontRegistry.has(k)) return fontRegistry.get(k);
    const id = nextFontId++;
    fontRegistry.set(k, id);
    newFontChunks.push(buildFontEntry(defaultFontInner, style));
    return id;
  }

  // XF dedup: one entry per unique (fontId, hAlign, vAlign). Cells mapping
  // to xfId 0 (no font style and no alignment) need no patching, but we
  // already filtered those out via normalizeStyle.
  const xfRegistry = new Map();
  const newXfChunks = [];
  let nextXfId = currentXfCount;
  function resolveXfId(style) {
    const fid = resolveFontId(style);
    const k = xfKey(style, fid);
    if (xfRegistry.has(k)) return xfRegistry.get(k);
    const id = nextXfId++;
    xfRegistry.set(k, id);
    newXfChunks.push(buildXfEntry(style, fid));
    return id;
  }

  // Resolve each cell to a concrete xfId up front so we can stop early if
  // nothing was added (defensive; with the filter above this is unlikely).
  const perSheetXfs = {};
  for (const [sheetName, cellMap] of Object.entries(perSheetEffective)) {
    const mapped = {};
    for (const [cellRef, style] of Object.entries(cellMap)) {
      mapped[cellRef] = resolveXfId(style);
    }
    perSheetXfs[sheetName] = mapped;
  }

  if (newXfChunks.length === 0) {
    // No user-styled cells to inject, but we still patched the default
    // xf above to set vertical=center for unstyled cells. Persist that.
    unzipped['xl/styles.xml'] = strToU8(stylesXml);
    return zipSync(unzipped);
  }

  if (newFontChunks.length > 0) {
    const newFontCount = currentFontCount + newFontChunks.length;
    stylesXml = stylesXml.replace(
      /<fonts\b[^>]*count="\d+"[^>]*>([\s\S]*?)<\/fonts>/,
      (_m, inner) => `<fonts count="${newFontCount}">${inner}${newFontChunks.join('')}</fonts>`,
    );
  }

  const newXfCount = currentXfCount + newXfChunks.length;
  stylesXml = stylesXml.replace(
    /<cellXfs\b[^>]*count="\d+"[^>]*>([\s\S]*?)<\/cellXfs>/,
    (_m, inner) => `<cellXfs count="${newXfCount}">${inner}${newXfChunks.join('')}</cellXfs>`,
  );
  unzipped['xl/styles.xml'] = strToU8(stylesXml);

  const workbookXml = unzipped['xl/workbook.xml'] ? strFromU8(unzipped['xl/workbook.xml']) : '';
  const relsXml = unzipped['xl/_rels/workbook.xml.rels'] ? strFromU8(unzipped['xl/_rels/workbook.xml.rels']) : '';
  const sheetPaths = workbookXml && relsXml ? parseSheetPaths(workbookXml, relsXml) : {};

  for (const [sheetName, cellMap] of Object.entries(perSheetXfs)) {
    const path = sheetPaths[sheetName];
    if (!path || !unzipped[path]) continue;
    let sheetXml = strFromU8(unzipped[path]);
    for (const [cellRef, id] of Object.entries(cellMap)) {
      sheetXml = patchCellStyle(sheetXml, cellRef, id);
    }
    unzipped[path] = strToU8(sheetXml);
  }

  return zipSync(unzipped);
}

/**
 * Back-compat shim for callers that only cared about alignment. Prefer
 * injectCellStyles for new code.
 * @deprecated
 */
export function injectCellAlignments(xlsxBytes, sheetAlignments) {
  return injectCellStyles(xlsxBytes, sheetAlignments);
}
