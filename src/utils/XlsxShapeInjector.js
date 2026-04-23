import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { getShapePreset, isLineShape } from './ShapePresets';

/*
 * Post-process an xlsx byte stream produced by SheetJS + XlsxStyleInjector
 * to embed native DrawingML shapes. Mirrors the approach in ADR 003 /
 * ADR 004: SheetJS Community does not emit drawing parts, so we append
 * them ourselves in-memory via fflate.
 *
 * For each sheet that owns shapes, the injector:
 *   - emits xl/drawings/drawingN.xml (DrawingML <xdr:wsDr>)
 *   - emits xl/drawings/_rels/drawingN.xml.rels (empty, required by OOXML)
 *   - updates xl/worksheets/_rels/sheetN.xml.rels (creates if missing)
 *   - inserts <drawing r:id="..."/> into xl/worksheets/sheetN.xml
 *   - registers the new drawing part in [Content_Types].xml
 *
 * Coordinates are EMU (1 px @ 96 DPI = 9525 EMU). The anchor form is
 * <xdr:twoCellAnchor editAs="absolute">, which Excel treats as "Don't
 * move or size with cells" — shapes stay where the user placed them
 * regardless of column / row resizes elsewhere.
 */

const EMU_PER_PX = 9525;
const DEG_TO_OOXML_ANGLE = 60000; // 1 degree = 60000 OOXML rotation units

const NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function colLabel(index) {
  let s = '';
  let n = index;
  while (n >= 0) {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

function escapeXml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
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

function sheetRelsPath(sheetPath) {
  const lastSlash = sheetPath.lastIndexOf('/');
  const dir = sheetPath.slice(0, lastSlash);
  const file = sheetPath.slice(lastSlash + 1);
  return `${dir}/_rels/${file}.rels`;
}

function emptyRelsXml() {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
  );
}

function upsertDrawingRel(existingXml, drawingFileName) {
  const target = `../drawings/${drawingFileName}`;
  const type = `${NS_R}/drawing`;
  let rels = existingXml || emptyRelsXml();
  const idRegex = /Id="rId(\d+)"/g;
  let maxId = 0;
  let m;
  while ((m = idRegex.exec(rels)) !== null) {
    const n = parseInt(m[1], 10);
    if (n > maxId) maxId = n;
  }
  const rId = `rId${maxId + 1}`;
  const rel = `<Relationship Id="${rId}" Type="${type}" Target="${target}"/>`;
  if (/<Relationships\b[^>]*\/>/.test(rels)) {
    rels = rels.replace(
      /<Relationships\b([^>]*)\/>/,
      `<Relationships$1>${rel}</Relationships>`,
    );
  } else if (rels.includes('</Relationships>')) {
    rels = rels.replace('</Relationships>', `${rel}</Relationships>`);
  } else {
    rels =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel}</Relationships>`;
  }
  return { relsXml: rels, rId };
}

function ensureRelationshipsXmlns(sheetXml) {
  const wsMatch = sheetXml.match(/<worksheet\b[^>]*>/);
  if (!wsMatch) return sheetXml;
  const open = wsMatch[0];
  if (/xmlns:r="/.test(open)) return sheetXml;
  const patched = open.replace(/>$/, ` xmlns:r="${NS_R}">`);
  return sheetXml.replace(open, patched);
}

function insertDrawingRef(sheetXml, rId) {
  if (new RegExp(`<drawing\\s+r:id="${rId}"\\s*/>`).test(sheetXml)) return sheetXml;
  return sheetXml.replace('</worksheet>', `<drawing r:id="${rId}"/></worksheet>`);
}

function getColWPx(idx, widths) {
  const key = colLabel(idx);
  return widths[key] || 80;
}

function getRowHPx(idx, heights) {
  return heights[idx] || 26;
}

function pxToAnchor(px, py, info) {
  const { colWidths, rowHeights, cols, rows } = info;
  let remX = Math.max(0, px);
  let col = 0;
  for (let i = 0; i < cols; i++) {
    const w = getColWPx(i, colWidths);
    if (remX < w) { col = i; break; }
    remX -= w;
    col = i + 1;
  }
  const colOff = Math.round(remX * EMU_PER_PX);

  let remY = Math.max(0, py);
  let row = 0;
  for (let i = 0; i < rows; i++) {
    const h = getRowHPx(i, rowHeights);
    if (remY < h) { row = i; break; }
    remY -= h;
    row = i + 1;
  }
  const rowOff = Math.round(remY * EMU_PER_PX);
  return { col, row, colOff, rowOff };
}

function hexToSrgb(color) {
  if (!color) return null;
  const m = /^#([0-9a-f]{6})$/i.exec(color.trim());
  if (!m) return null;
  return m[1].toUpperCase();
}

function alphaAttr(alpha) {
  if (alpha == null || alpha >= 1) return '';
  const pct = Math.max(0, Math.min(100, alpha * 100));
  return `<a:alpha val="${Math.round(pct * 1000)}"/>`;
}

function solidFillXml(color, alpha) {
  const srgb = hexToSrgb(color);
  if (!srgb) return '<a:noFill/>';
  return `<a:solidFill><a:srgbClr val="${srgb}">${alphaAttr(alpha)}</a:srgbClr></a:solidFill>`;
}

function buildFillXml(fill, isLine) {
  if (isLine) return '<a:noFill/>';
  if (!fill || fill.type === 'none') return '<a:noFill/>';
  return solidFillXml(fill.color, fill.alpha);
}

function buildOutlineXml(outline) {
  if (!outline) return '<a:ln><a:noFill/></a:ln>';
  const width = outline.width != null ? outline.width : 1;
  const widthEmu = Math.max(0, Math.round(width * EMU_PER_PX));
  const color = outline.color;
  const alpha = outline.alpha;
  const dash = outline.dash || 'solid';
  const fillInner = color ? solidFillXml(color, alpha) : '<a:noFill/>';
  return `<a:ln w="${widthEmu}">${fillInner}<a:prstDash val="${dash}"/></a:ln>`;
}

function buildEffectsXml(effects) {
  if (!effects || (!effects.shadow && !effects.glow && !effects.softEdge)) return '';
  const parts = [];
  if (effects.shadow) {
    const sh = effects.shadow;
    const blurRad = Math.round((sh.blur != null ? sh.blur : 4) * EMU_PER_PX);
    const dx = sh.offsetX != null ? sh.offsetX : 3;
    const dy = sh.offsetY != null ? sh.offsetY : 3;
    const dist = Math.round(Math.sqrt(dx * dx + dy * dy) * EMU_PER_PX);
    // OOXML dir: 0 is east, angles are counter-clockwise in 60000-per-degree.
    const dirDeg = (Math.atan2(dy, dx) * 180) / Math.PI;
    const dir = Math.round(((dirDeg + 360) % 360) * DEG_TO_OOXML_ANGLE);
    const color = hexToSrgb(sh.color || '#000000') || '000000';
    const alpha = alphaAttr(sh.alpha != null ? sh.alpha : 0.35);
    parts.push(
      `<a:outerShdw blurRad="${blurRad}" dist="${dist}" dir="${dir}" algn="tl" rotWithShape="0">` +
      `<a:srgbClr val="${color}">${alpha}</a:srgbClr>` +
      `</a:outerShdw>`,
    );
  }
  if (effects.glow) {
    const gl = effects.glow;
    const rad = Math.round((gl.radius != null ? gl.radius : 4) * EMU_PER_PX);
    const color = hexToSrgb(gl.color || '#FFEA00') || 'FFEA00';
    const alpha = alphaAttr(gl.alpha != null ? gl.alpha : 0.6);
    parts.push(`<a:glow rad="${rad}"><a:srgbClr val="${color}">${alpha}</a:srgbClr></a:glow>`);
  }
  if (effects.softEdge) {
    const rad = Math.round((effects.softEdge.radius != null ? effects.softEdge.radius : 2) * EMU_PER_PX);
    parts.push(`<a:softEdge rad="${rad}"/>`);
  }
  return `<a:effectLst>${parts.join('')}</a:effectLst>`;
}

function buildTextBodyXml(shape) {
  const preset = getShapePreset(shape.type);
  const vertAttr = preset?.vertical ? ' vert="vert"' : '';
  if (!shape.text || !shape.text.value) {
    if (!shape.textStyle) return '';
    return (
      '<xdr:txBody>' +
      `<a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="ctr"${vertAttr}/>` +
      '<a:lstStyle/>' +
      '<a:p><a:endParaRPr lang="en-US"/></a:p>' +
      '</xdr:txBody>'
    );
  }
  const ts = shape.textStyle || {};
  const hAlign = ts.hAlign === 'left' ? 'l' : ts.hAlign === 'right' ? 'r' : 'ctr';
  const vAlign = ts.vAlign === 'top' ? 't' : ts.vAlign === 'bottom' ? 'b' : 'ctr';
  const fontSize = ts.fontSize || 12;
  const sz = Math.round(fontSize * 100);
  const boldAttr = ts.bold ? ' b="1"' : '';
  const italicAttr = ts.italic ? ' i="1"' : '';
  const uAttr = ts.underline ? ' u="sng"' : '';

  const fillColor = ts.fill?.color;
  const fillAlpha = ts.fill?.alpha;
  const textFill = fillColor ? solidFillXml(fillColor, fillAlpha) : '';

  const outline = ts.outline;
  const textOutline = outline && outline.color
    ? `<a:ln w="${Math.max(0, Math.round((outline.width || 1) * EMU_PER_PX))}">${solidFillXml(outline.color, outline.alpha)}</a:ln>`
    : '';

  const effects = ts.effects || {};
  const textEffects = buildEffectsXml(effects);

  const fontFamily = ts.fontFamily || 'Calibri';
  const latin = `<a:latin typeface="${escapeXml(fontFamily)}"/>`;

  const lines = String(shape.text.value).split('\n');
  const pBlocks = lines.map((line) => {
    const text = escapeXml(line);
    return (
      `<a:p>` +
      `<a:pPr algn="${hAlign}"/>` +
      `<a:r>` +
      `<a:rPr lang="en-US" sz="${sz}"${boldAttr}${italicAttr}${uAttr}>` +
      textFill +
      textOutline +
      textEffects +
      latin +
      `</a:rPr>` +
      `<a:t>${text}</a:t>` +
      `</a:r>` +
      `</a:p>`
    );
  });

  return (
    `<xdr:txBody>` +
    `<a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="${vAlign}"${vertAttr}/>` +
    `<a:lstStyle/>` +
    pBlocks.join('') +
    `</xdr:txBody>`
  );
}

function buildSpXml(shape, localId) {
  const preset = getShapePreset(shape.type);
  const prstGeom = preset ? preset.prstGeom : 'rect';
  const isLine = preset ? preset.kind === 'line' : false;
  const w = Math.max(1, Math.round(shape.w * EMU_PER_PX));
  const h = Math.max(1, Math.round(shape.h * EMU_PER_PX));
  const rot = Math.round(((shape.rot || 0) % 360) * DEG_TO_OOXML_ANGLE);
  const rotAttr = rot !== 0 ? ` rot="${rot}"` : '';
  const flipH = shape.flipH ? ' flipH="1"' : '';
  const flipV = shape.flipV ? ' flipV="1"' : '';

  const fillXml = buildFillXml(shape.style?.fill, isLine);
  const lnXml = buildOutlineXml(shape.style?.outline);
  const effectsXml = buildEffectsXml(shape.style?.effects);
  const txBodyXml = buildTextBodyXml(shape);

  const name = escapeXml(preset?.label || 'Shape');

  return (
    `<xdr:sp macro="" textlink="">` +
    `<xdr:nvSpPr>` +
    `<xdr:cNvPr id="${localId + 1}" name="${name} ${localId}"/>` +
    `<xdr:cNvSpPr/>` +
    `</xdr:nvSpPr>` +
    `<xdr:spPr>` +
    `<a:xfrm${rotAttr}${flipH}${flipV}>` +
    `<a:off x="0" y="0"/>` +
    `<a:ext cx="${w}" cy="${h}"/>` +
    `</a:xfrm>` +
    `<a:prstGeom prst="${prstGeom}"><a:avLst/></a:prstGeom>` +
    fillXml +
    lnXml +
    effectsXml +
    `</xdr:spPr>` +
    `<xdr:style>` +
    `<a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>` +
    `<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>` +
    `<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>` +
    `<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>` +
    `</xdr:style>` +
    txBodyXml +
    `</xdr:sp>`
  );
}

function buildTwoCellAnchor(shape, localId, info) {
  const from = pxToAnchor(shape.x, shape.y, info);
  const to = pxToAnchor(shape.x + shape.w, shape.y + shape.h, info);
  const sp = buildSpXml(shape, localId);
  return (
    `<xdr:twoCellAnchor editAs="absolute">` +
    `<xdr:from>` +
    `<xdr:col>${from.col}</xdr:col>` +
    `<xdr:colOff>${from.colOff}</xdr:colOff>` +
    `<xdr:row>${from.row}</xdr:row>` +
    `<xdr:rowOff>${from.rowOff}</xdr:rowOff>` +
    `</xdr:from>` +
    `<xdr:to>` +
    `<xdr:col>${to.col}</xdr:col>` +
    `<xdr:colOff>${to.colOff}</xdr:colOff>` +
    `<xdr:row>${to.row}</xdr:row>` +
    `<xdr:rowOff>${to.rowOff}</xdr:rowOff>` +
    `</xdr:to>` +
    sp +
    `<xdr:clientData/>` +
    `</xdr:twoCellAnchor>`
  );
}

function buildDrawingXml(info) {
  const shapes = [...info.shapes].sort((a, b) => (a.z || 0) - (b.z || 0));
  const anchors = shapes.map((s, idx) => buildTwoCellAnchor(s, idx + 1, info));
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<xdr:wsDr xmlns:xdr="${NS_XDR}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">` +
    anchors.join('') +
    '</xdr:wsDr>'
  );
}

function ensureContentType(contentTypesXml, drawingPath) {
  const override = `<Override PartName="/${drawingPath}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`;
  if (contentTypesXml.includes(`PartName="/${drawingPath}"`)) return { xml: contentTypesXml, changed: false };
  const patched = contentTypesXml.replace('</Types>', `${override}</Types>`);
  return { xml: patched, changed: true };
}

/**
 * @param {ArrayBuffer | Uint8Array} xlsxBytes
 * @param {Record<string, {
 *   shapes: object[],
 *   colWidths: Record<string, number>,
 *   rowHeights: Record<number, number>,
 *   cols: number,
 *   rows: number,
 * }>} sheetShapes
 * @returns {Uint8Array}
 */
export function injectShapes(xlsxBytes, sheetShapes) {
  const source = xlsxBytes instanceof Uint8Array ? xlsxBytes : new Uint8Array(xlsxBytes);
  if (!sheetShapes || Object.keys(sheetShapes).length === 0) return source;

  let unzipped;
  try { unzipped = unzipSync(source); } catch { return source; }

  const workbookXml = unzipped['xl/workbook.xml'] ? strFromU8(unzipped['xl/workbook.xml']) : '';
  const relsXml = unzipped['xl/_rels/workbook.xml.rels'] ? strFromU8(unzipped['xl/_rels/workbook.xml.rels']) : '';
  if (!workbookXml || !relsXml) return source;
  const sheetPaths = parseSheetPaths(workbookXml, relsXml);

  let drawingCounter = 1;
  while (unzipped[`xl/drawings/drawing${drawingCounter}.xml`]) drawingCounter++;

  let contentTypesXml = unzipped['[Content_Types].xml']
    ? strFromU8(unzipped['[Content_Types].xml'])
    : '';
  if (!contentTypesXml) return source;
  let contentTypesChanged = false;
  let touched = false;

  for (const [sheetName, info] of Object.entries(sheetShapes)) {
    if (!info || !Array.isArray(info.shapes) || info.shapes.length === 0) continue;
    const sheetPath = sheetPaths[sheetName];
    if (!sheetPath || !unzipped[sheetPath]) continue;

    const drawingFileName = `drawing${drawingCounter}.xml`;
    const drawingPath = `xl/drawings/${drawingFileName}`;
    const drawingRelsPath = `xl/drawings/_rels/${drawingFileName}.rels`;

    unzipped[drawingPath] = strToU8(buildDrawingXml(info));
    unzipped[drawingRelsPath] = strToU8(emptyRelsXml());

    const relsPath = sheetRelsPath(sheetPath);
    const existingRelsXml = unzipped[relsPath] ? strFromU8(unzipped[relsPath]) : null;
    const { relsXml: newRelsXml, rId } = upsertDrawingRel(existingRelsXml, drawingFileName);
    unzipped[relsPath] = strToU8(newRelsXml);

    let sheetXml = strFromU8(unzipped[sheetPath]);
    sheetXml = ensureRelationshipsXmlns(sheetXml);
    sheetXml = insertDrawingRef(sheetXml, rId);
    unzipped[sheetPath] = strToU8(sheetXml);

    const ctResult = ensureContentType(contentTypesXml, drawingPath);
    contentTypesXml = ctResult.xml;
    if (ctResult.changed) contentTypesChanged = true;

    drawingCounter++;
    touched = true;
  }

  if (!touched) return source;
  if (contentTypesChanged) {
    unzipped['[Content_Types].xml'] = strToU8(contentTypesXml);
  }
  return zipSync(unzipped);
}
