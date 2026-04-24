import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { toExcelNumFmtCode } from './NumberFormat';

/*
 * Post-process an xlsx byte stream produced by SheetJS Community to embed
 * native Excel cell styles that SheetJS Community does not write itself.
 *
 * Scope:
 *   - Horizontal / vertical alignment (via <alignment> on <xf>)
 *   - Bold / italic / underline (via additional <font> entries and fontId
 *     / applyFont on <xf>)
 *   - Number formats (via <numFmts> entries for custom codes and
 *     numFmtId / applyNumberFormat on <xf>; built-in OOXML IDs are used
 *     where possible to avoid emitting a <numFmt> entry at all).
 *
 * Strategy: leave SheetJS's output intact, then append the minimum number
 * of <numFmt>, <font>, and <xf> entries to xl/styles.xml and rewrite the
 * relevant <c> elements in each xl/worksheets/sheetN.xml with the
 * resolved s="..." attribute. Fonts, numFmts, and xfs are deduplicated;
 * a cell with no effective style is left untouched.
 *
 * The injector is deliberately conservative: if sheetStyles is empty or
 * styles.xml is missing / unparseable, it returns the input bytes
 * unchanged.
 */

const H_VALID = new Set(['left', 'center', 'right']);
const V_APP_TO_OOXML = { top: 'top', middle: 'center', bottom: 'bottom' };
const NUM_FMT_VALID = new Set([
  'general',
  'number',
  'currency',
  'accounting',
  'shortDate',
  'longDate',
  'time',
  'percentage',
  'fraction',
  'scientific',
  'text',
]);

// App-side border style -> OOXML <border> style attribute. The app uses
// thin/medium/thick/double/dashed/dotted; OOXML uses the same names plus
// a handful we do not emit (hair, mediumDashed, slantDashDot, ...).
const BORDER_STYLE_TO_OOXML = {
  thin: 'thin',
  medium: 'medium',
  thick: 'thick',
  double: 'double',
  dashed: 'dashed',
  dotted: 'dotted',
};

function normalizeHexColor(v) {
  if (typeof v !== 'string') return null;
  const m = /^#?([0-9a-fA-F]{6})$/.exec(v.trim());
  if (!m) return null;
  return m[1].toUpperCase();
}

function normalizeBorderSide(b) {
  if (!b || !b.style || !BORDER_STYLE_TO_OOXML[b.style]) return null;
  const out = { style: b.style };
  const color = normalizeHexColor(b.color);
  if (color) out.color = color;
  return out;
}

// Excel reserves numFmtIds below 164 for built-ins. Custom entries must
// start at 164 per the OOXML spec.
const CUSTOM_NUM_FMT_START = 164;

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
  if (raw.numFmt && NUM_FMT_VALID.has(raw.numFmt) && raw.numFmt !== 'general') {
    out.numFmt = raw.numFmt;
    if (typeof raw.decimals === 'number') out.decimals = raw.decimals;
    if (typeof raw.currency === 'string' && raw.currency) out.currency = raw.currency;
    if (typeof raw.useThousands === 'boolean') out.useThousands = raw.useThousands;
    if (typeof raw.negativeStyle === 'string' && raw.negativeStyle) {
      out.negativeStyle = raw.negativeStyle;
    }
  }
  const color = normalizeHexColor(raw.color);
  if (color) out.color = color;
  const bg = normalizeHexColor(raw.bg);
  if (bg) out.bg = bg;
  if (raw.borders && typeof raw.borders === 'object') {
    const b = {};
    for (const side of ['top', 'right', 'bottom', 'left']) {
      const norm = normalizeBorderSide(raw.borders[side]);
      if (norm) b[side] = norm;
    }
    if (Object.keys(b).length > 0) out.borders = b;
  }
  if (
    !out.bold &&
    !out.italic &&
    !out.underline &&
    !out.hAlign &&
    !out.vAlign &&
    !out.numFmt &&
    !out.color &&
    !out.bg &&
    !out.borders
  ) {
    return null;
  }
  return out;
}

function fontKey(s) {
  return `${s.bold ? 'b' : ''}${s.italic ? 'i' : ''}${s.underline ? 'u' : ''}|${s.color || ''}`;
}

function hasFontStyle(s) {
  return !!(s.bold || s.italic || s.underline || s.color);
}

function alignKey(s) {
  return `${s.hAlign || ''}|${s.vAlign || ''}`;
}

function hasAlignment(s) {
  return !!(s.hAlign || s.vAlign);
}

function borderKey(s) {
  if (!s.borders) return '';
  const parts = [];
  for (const side of ['top', 'right', 'bottom', 'left']) {
    const b = s.borders[side];
    if (!b) { parts.push(''); continue; }
    parts.push(`${b.style}:${b.color || ''}`);
  }
  return parts.join('|');
}

function xfKey(s, fontId, numFmtId, fillId, borderId) {
  return `n${numFmtId}|f${fontId}|fill${fillId}|bd${borderId}|${alignKey(s)}`;
}

// Replace / inject the first <color .../> element in the cloned default
// font's inner XML so bold/italic/underline variants that also carry an
// explicit color emit a single consistent <color rgb="..."/>. The default
// font Excel writes references theme="1" (black); leaving that in place
// alongside a new <color rgb="..."/> would make Excel treat the font as
// multi-color, which is not valid.
function applyFontColorInner(defaultInner, colorHex) {
  const tag = `<color rgb="FF${colorHex}"/>`;
  if (/<color\b[^/>]*\/?>/.test(defaultInner)) {
    return defaultInner.replace(/<color\b[^/>]*\/?>/, tag);
  }
  return defaultInner + tag;
}

function buildFontEntry(defaultInner, s) {
  const markers = [];
  if (s.bold) markers.push('<b/>');
  if (s.italic) markers.push('<i/>');
  if (s.underline) markers.push('<u/>');
  const inner = s.color ? applyFontColorInner(defaultInner, s.color) : defaultInner;
  return `<font>${markers.join('')}${inner}</font>`;
}

function buildFillEntry(bgHex) {
  // OOXML pattern fill with a solid fgColor (Excel's "Fill Color" button
  // writes <patternFill patternType="solid"><fgColor rgb="..."/></patternFill>).
  // Alpha channel prefix FF makes the color fully opaque.
  return `<fill><patternFill patternType="solid"><fgColor rgb="FF${bgHex}"/><bgColor indexed="64"/></patternFill></fill>`;
}

function buildBorderEntry(borders) {
  const sideXml = (name, b) => {
    if (!b) return `<${name}/>`;
    const style = BORDER_STYLE_TO_OOXML[b.style] || 'thin';
    const color = b.color ? `<color rgb="FF${b.color}"/>` : '<color indexed="64"/>';
    return `<${name} style="${style}">${color}</${name}>`;
  };
  // OOXML mandates this child order: left, right, top, bottom, diagonal.
  return (
    '<border>' +
    sideXml('left', borders.left) +
    sideXml('right', borders.right) +
    sideXml('top', borders.top) +
    sideXml('bottom', borders.bottom) +
    '<diagonal/>' +
    '</border>'
  );
}

function escapeXmlAttr(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function buildNumFmtEntry(id, code) {
  return `<numFmt numFmtId="${id}" formatCode="${escapeXmlAttr(code)}"/>`;
}

function buildXfEntry(s, fontId, numFmtId, fillId, borderId) {
  const applyFont = fontId > 0;
  const applyAlign = hasAlignment(s);
  const applyNum = numFmtId > 0;
  const applyFill = fillId > 0;
  const applyBorder = borderId > 0;
  const attrs = [
    `numFmtId="${numFmtId}"`,
    `fontId="${fontId}"`,
    `fillId="${fillId}"`,
    `borderId="${borderId}"`,
    'xfId="0"',
  ];
  if (applyNum) attrs.push('applyNumberFormat="1"');
  if (applyFont) attrs.push('applyFont="1"');
  if (applyFill) attrs.push('applyFill="1"');
  if (applyBorder) attrs.push('applyBorder="1"');
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

function parseExistingNumFmts(stylesXml) {
  // Build a map of existing formatCode -> numFmtId from any <numFmts>
  // block SheetJS or an earlier injector run wrote. We dedup against
  // this to avoid emitting duplicates on re-save.
  const out = new Map();
  const block = stylesXml.match(/<numFmts\b[^>]*>([\s\S]*?)<\/numFmts>/);
  if (!block) return out;
  const entryRegex = /<numFmt\b[^>]*?numFmtId="(\d+)"[^>]*?formatCode="([^"]*)"[^>]*?\/?>/g;
  let m;
  while ((m = entryRegex.exec(block[1])) !== null) {
    const id = parseInt(m[1], 10);
    const code = m[2]
      .replace(/&quot;/g, '"')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>');
    if (!Number.isFinite(id) || id < CUSTOM_NUM_FMT_START) continue;
    out.set(code, id);
  }
  return out;
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
  const fillsMatch = stylesXml.match(/<fills\b[^>]*count="(\d+)"[^>]*>/);
  const bordersMatch = stylesXml.match(/<borders\b[^>]*count="(\d+)"[^>]*>/);
  const currentFillCount = fillsMatch ? parseInt(fillsMatch[1], 10) : 0;
  const currentBorderCount = bordersMatch ? parseInt(bordersMatch[1], 10) : 0;
  const defaultFontInner = extractDefaultFontInner(stylesXml);

  // numFmt dedup: resolve the style's format code via NumberFormat. Built-in
  // IDs (0 = General, 9 = 0%, 10 = 0.00%, 14 = m/d/yyyy, 19 = h:mm:ss AM/PM,
  // 49 = @, ...) require no <numFmt> entry. Custom codes register under
  // ids starting at CUSTOM_NUM_FMT_START and reuse entries already present
  // in the workbook on re-save.
  const existingCustomNumFmts = parseExistingNumFmts(stylesXml);
  const codeToNumFmtId = new Map(existingCustomNumFmts);
  const newNumFmtEntries = [];
  let nextCustomNumFmtId = CUSTOM_NUM_FMT_START;
  for (const id of existingCustomNumFmts.values()) {
    if (id >= nextCustomNumFmtId) nextCustomNumFmtId = id + 1;
  }
  function resolveNumFmtId(style) {
    if (!style.numFmt || style.numFmt === 'general') return 0;
    const excel = toExcelNumFmtCode(style);
    if (!excel || !excel.code) return 0;
    if (typeof excel.id === 'number') return excel.id;
    const cached = codeToNumFmtId.get(excel.code);
    if (cached != null) return cached;
    const id = nextCustomNumFmtId++;
    codeToNumFmtId.set(excel.code, id);
    newNumFmtEntries.push(buildNumFmtEntry(id, excel.code));
    return id;
  }

  // Font dedup: one entry per unique (bold, italic, underline, color)
  // combination. The all-default combination reuses fontId 0.
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

  // Fill dedup: one <fill> entry per unique solid colour. OOXML reserves
  // fillIds 0 and 1 for the `none` and `gray125` placeholders that every
  // xlsx must emit, so the first user fill lands at index currentFillCount
  // (which will be >= 2 for any workbook SheetJS writes).
  const fillRegistry = new Map();
  const newFillChunks = [];
  let nextFillId = currentFillCount;
  function resolveFillId(style) {
    if (!style.bg) return 0;
    const k = style.bg;
    if (fillRegistry.has(k)) return fillRegistry.get(k);
    const id = nextFillId++;
    fillRegistry.set(k, id);
    newFillChunks.push(buildFillEntry(style.bg));
    return id;
  }

  // Border dedup: one <border> entry per unique four-side combination.
  // OOXML reserves borderId 0 for the "no borders" placeholder; we never
  // touch that entry.
  const borderRegistry = new Map();
  const newBorderChunks = [];
  let nextBorderId = currentBorderCount;
  function resolveBorderId(style) {
    if (!style.borders) return 0;
    const k = borderKey(style);
    if (!k) return 0;
    if (borderRegistry.has(k)) return borderRegistry.get(k);
    const id = nextBorderId++;
    borderRegistry.set(k, id);
    newBorderChunks.push(buildBorderEntry(style.borders));
    return id;
  }

  // XF dedup: one entry per unique (numFmtId, fontId, fillId, borderId,
  // hAlign, vAlign). Cells mapping to xfId 0 (no font style, no alignment,
  // no number format, no fill, no border) need no patching, but we
  // already filtered those out via normalizeStyle.
  const xfRegistry = new Map();
  const newXfChunks = [];
  let nextXfId = currentXfCount;
  function resolveXfId(style) {
    const nfid = resolveNumFmtId(style);
    const fid = resolveFontId(style);
    const fillId = resolveFillId(style);
    const borderId = resolveBorderId(style);
    const k = xfKey(style, fid, nfid, fillId, borderId);
    if (xfRegistry.has(k)) return xfRegistry.get(k);
    const id = nextXfId++;
    xfRegistry.set(k, id);
    newXfChunks.push(buildXfEntry(style, fid, nfid, fillId, borderId));
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

  if (newNumFmtEntries.length > 0) {
    const existingBlock = stylesXml.match(/<numFmts\b[^>]*>([\s\S]*?)<\/numFmts>/);
    if (existingBlock) {
      // Splice new entries into the existing block and refresh the count.
      const existingEntryCount = existingCustomNumFmts.size;
      const newCount = existingEntryCount + newNumFmtEntries.length;
      stylesXml = stylesXml.replace(
        /<numFmts\b[^>]*>([\s\S]*?)<\/numFmts>/,
        (_m, inner) => `<numFmts count="${newCount}">${inner}${newNumFmtEntries.join('')}</numFmts>`,
      );
    } else {
      // Insert a fresh <numFmts> block as the first child of <styleSheet>.
      // OOXML requires numFmts to precede fonts / fills / borders / cellXfs.
      const block = `<numFmts count="${newNumFmtEntries.length}">${newNumFmtEntries.join('')}</numFmts>`;
      const styleSheetMatch = stylesXml.match(/<styleSheet\b[^>]*>/);
      if (styleSheetMatch) {
        const insertAt = styleSheetMatch.index + styleSheetMatch[0].length;
        stylesXml = stylesXml.slice(0, insertAt) + block + stylesXml.slice(insertAt);
      }
    }
  }

  if (newFontChunks.length > 0) {
    const newFontCount = currentFontCount + newFontChunks.length;
    stylesXml = stylesXml.replace(
      /<fonts\b[^>]*count="\d+"[^>]*>([\s\S]*?)<\/fonts>/,
      (_m, inner) => `<fonts count="${newFontCount}">${inner}${newFontChunks.join('')}</fonts>`,
    );
  }

  if (newFillChunks.length > 0) {
    const newFillCount = currentFillCount + newFillChunks.length;
    if (fillsMatch) {
      stylesXml = stylesXml.replace(
        /<fills\b[^>]*count="\d+"[^>]*>([\s\S]*?)<\/fills>/,
        (_m, inner) => `<fills count="${newFillCount}">${inner}${newFillChunks.join('')}</fills>`,
      );
    } else {
      // No <fills> block yet (rare: SheetJS normally emits one). Create
      // one with the required gray125/none placeholders at indices 0/1
      // so our user fills start at index 2 as assumed by the registry.
      const placeholder =
        '<fill><patternFill patternType="none"/></fill>' +
        '<fill><patternFill patternType="gray125"/></fill>';
      const block = `<fills count="${2 + newFillChunks.length}">${placeholder}${newFillChunks.join('')}</fills>`;
      // Insert just after </fonts> if present; otherwise right before <cellXfs>.
      if (/<\/fonts>/.test(stylesXml)) {
        stylesXml = stylesXml.replace('</fonts>', `</fonts>${block}`);
      } else {
        stylesXml = stylesXml.replace(/<cellXfs\b/, `${block}<cellXfs`);
      }
    }
  }

  if (newBorderChunks.length > 0) {
    const newBorderCount = currentBorderCount + newBorderChunks.length;
    if (bordersMatch) {
      stylesXml = stylesXml.replace(
        /<borders\b[^>]*count="\d+"[^>]*>([\s\S]*?)<\/borders>/,
        (_m, inner) => `<borders count="${newBorderCount}">${inner}${newBorderChunks.join('')}</borders>`,
      );
    } else {
      // No <borders> block yet. OOXML requires borderId 0 to be a default
      // empty border; emit that placeholder before the user borders so
      // cell references align with the registry's borderId allocation.
      const placeholder = '<border><left/><right/><top/><bottom/><diagonal/></border>';
      const block = `<borders count="${1 + newBorderChunks.length}">${placeholder}${newBorderChunks.join('')}</borders>`;
      if (/<\/fills>/.test(stylesXml)) {
        stylesXml = stylesXml.replace('</fills>', `</fills>${block}`);
      } else {
        stylesXml = stylesXml.replace(/<cellXfs\b/, `${block}<cellXfs`);
      }
    }
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
