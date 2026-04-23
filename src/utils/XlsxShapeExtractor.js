import { unzipSync, strFromU8 } from 'fflate';
import { getPresetByPrstGeom } from './ShapePresets';

/*
 * Read native DrawingML shapes from the xlsx byte stream. Mirror of
 * XlsxShapeInjector. Returns per-sheet arrays of shape objects matching
 * GanttGen's runtime shape schema so the DataGrid can render them
 * directly.
 *
 * Best-effort: any parse failure returns an empty map, never throws.
 */

const EMU_PER_PX = 9525;
const DEG_TO_OOXML_ANGLE = 60000;
const DEFAULT_COL_WIDTH_PX = 80;
const DEFAULT_ROW_HEIGHT_PX = 26;

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

function parseDrawingTarget(sheetRelsXml) {
  if (!sheetRelsXml) return null;
  const regex = /<Relationship\b[^>]*?Type="[^"]*\/drawing"[^>]*?Target="([^"]+)"[^>]*?\/>/;
  const m = sheetRelsXml.match(regex);
  if (!m) return null;
  const target = m[1];
  if (target.startsWith('/')) return target.slice(1);
  // target is usually ../drawings/drawingN.xml relative to xl/worksheets/_rels/
  if (target.startsWith('../')) return 'xl/' + target.slice(3);
  return target;
}

// SheetJS's wpx-to-width encoding empirically uses MDW=6 for the default
// Calibri 11 font (wpx=80 -> width="13.332..."). The OOXML spec says
// MDW is font-dependent; for GanttGen-produced files this matches, and
// it's close enough for files authored in Excel too.
function parseColWidthsFromSheet(sheetXml) {
  const map = {};
  const colsBlock = sheetXml.match(/<cols\b[^>]*>([\s\S]*?)<\/cols>/);
  if (!colsBlock) return map;
  const colRegex = /<col\b([^/>]*)\/>/g;
  let m;
  while ((m = colRegex.exec(colsBlock[1])) !== null) {
    const attrs = m[1];
    const minM = attrs.match(/\bmin="(\d+)"/);
    const maxM = attrs.match(/\bmax="(\d+)"/);
    const widthM = attrs.match(/\bwidth="([\d.]+)"/);
    if (!minM || !maxM || !widthM) continue;
    const min = parseInt(minM[1], 10);
    const max = parseInt(maxM[1], 10);
    const widthChar = parseFloat(widthM[1]);
    const widthPx = Math.round(widthChar * 6);
    for (let i = min; i <= max; i++) {
      map[i - 1] = widthPx;
    }
  }
  return map;
}

// GanttGen exports pass both { hpx, hpt } set to the same px value so
// SheetJS's OOXML output has ht=<px-value>. Reading ht back as pixels
// gives an exact round-trip for GanttGen files; canonical Excel files
// (where ht is strictly points) will appear ~33% shorter, but the
// shape anchors remain near the right row so shapes are still usable.
function parseRowHeightsFromSheet(sheetXml) {
  const map = {};
  const rowRegex = /<row\b([^/>]*)(?:\/>|>[\s\S]*?<\/row>)/g;
  let m;
  while ((m = rowRegex.exec(sheetXml)) !== null) {
    const attrs = m[1];
    const rM = attrs.match(/\br="(\d+)"/);
    const htM = attrs.match(/\bht="([\d.]+)"/);
    if (!rM || !htM) continue;
    const rowIdx = parseInt(rM[1], 10) - 1;
    const ht = parseFloat(htM[1]);
    map[rowIdx] = Math.round(ht);
  }
  return map;
}

function anchorToPx(col, colOff, row, rowOff, colWidths, rowHeights) {
  let x = 0;
  for (let i = 0; i < col; i++) {
    x += colWidths[i] != null ? colWidths[i] : DEFAULT_COL_WIDTH_PX;
  }
  x += colOff / EMU_PER_PX;

  let y = 0;
  for (let i = 0; i < row; i++) {
    y += rowHeights[i] != null ? rowHeights[i] : DEFAULT_ROW_HEIGHT_PX;
  }
  y += rowOff / EMU_PER_PX;
  return { x, y };
}

function getIntAttr(attrs, name) {
  const m = attrs.match(new RegExp(`\\b${name}="(-?\\d+)"`));
  return m ? parseInt(m[1], 10) : 0;
}

function getIntElement(block, tag) {
  const m = block.match(new RegExp(`<xdr:${tag}>([\\s\\S]*?)<\\/xdr:${tag}>`));
  return m ? parseInt(m[1].trim(), 10) : 0;
}

function parseSrgbFromColor(xml) {
  const m = xml.match(/<a:srgbClr\b[^>]*val="([0-9A-Fa-f]{6})"/);
  if (!m) return null;
  return '#' + m[1].toLowerCase();
}

function parseAlphaFromColor(xml) {
  const m = xml.match(/<a:alpha\b[^>]*val="(\d+)"/);
  if (!m) return 1;
  const n = parseInt(m[1], 10);
  return Math.max(0, Math.min(1, n / 100000));
}

function parseFill(spPrXml) {
  const noFill = /<a:noFill\s*\/>/.test(spPrXml);
  if (noFill) return { type: 'none' };
  const fillMatch = spPrXml.match(/<a:solidFill>([\s\S]*?)<\/a:solidFill>/);
  if (!fillMatch) return undefined;
  const color = parseSrgbFromColor(fillMatch[1]);
  const alpha = parseAlphaFromColor(fillMatch[1]);
  if (!color) return undefined;
  return { type: 'solid', color, alpha };
}

function parseOutline(spPrXml) {
  const lnMatch = spPrXml.match(/<a:ln\b([^>]*)>([\s\S]*?)<\/a:ln>|<a:ln\b([^>]*)\/>/);
  if (!lnMatch) return undefined;
  const attrs = lnMatch[1] || lnMatch[3] || '';
  const inner = lnMatch[2] || '';
  const wM = attrs.match(/\bw="(\d+)"/);
  const width = wM ? parseInt(wM[1], 10) / EMU_PER_PX : 1;
  const dashM = inner.match(/<a:prstDash\b[^>]*val="([^"]+)"/);
  const dash = dashM ? dashM[1] : 'solid';
  const fillMatch = inner.match(/<a:solidFill>([\s\S]*?)<\/a:solidFill>/);
  const color = fillMatch ? parseSrgbFromColor(fillMatch[1]) : undefined;
  const alpha = fillMatch ? parseAlphaFromColor(fillMatch[1]) : 1;
  if (/<a:noFill\s*\/>/.test(inner)) return { color: undefined, alpha: 1, width, dash };
  return { color, alpha, width, dash };
}

function parseEffects(spPrXml) {
  const effLst = spPrXml.match(/<a:effectLst>([\s\S]*?)<\/a:effectLst>/);
  if (!effLst) return {};
  const inner = effLst[1];
  const out = {};

  const shadowMatch = inner.match(/<a:outerShdw\b([^>]*)>([\s\S]*?)<\/a:outerShdw>/);
  if (shadowMatch) {
    const a = shadowMatch[1];
    const body = shadowMatch[2];
    const blurRad = getIntAttr(a, 'blurRad') / EMU_PER_PX;
    const dist = getIntAttr(a, 'dist') / EMU_PER_PX;
    const dirM = a.match(/\bdir="(\d+)"/);
    const dirDeg = dirM ? parseInt(dirM[1], 10) / DEG_TO_OOXML_ANGLE : 0;
    const rad = (dirDeg * Math.PI) / 180;
    const offsetX = Math.round(Math.cos(rad) * dist * 100) / 100;
    const offsetY = Math.round(Math.sin(rad) * dist * 100) / 100;
    const color = parseSrgbFromColor(body) || '#000000';
    const alpha = parseAlphaFromColor(body);
    out.shadow = { color, alpha, offsetX, offsetY, blur: blurRad };
  }

  const glowMatch = inner.match(/<a:glow\b([^>]*)>([\s\S]*?)<\/a:glow>/);
  if (glowMatch) {
    const a = glowMatch[1];
    const body = glowMatch[2];
    const rad = getIntAttr(a, 'rad') / EMU_PER_PX;
    const color = parseSrgbFromColor(body) || '#ffea00';
    const alpha = parseAlphaFromColor(body);
    out.glow = { color, alpha, radius: rad };
  }

  const softMatch = inner.match(/<a:softEdge\b([^>]*)\/?>/);
  if (softMatch) {
    const rad = getIntAttr(softMatch[1], 'rad') / EMU_PER_PX;
    out.softEdge = { radius: rad };
  }

  return out;
}

function parseTextBody(txBodyXml) {
  if (!txBodyXml) return null;
  const bodyPr = txBodyXml.match(/<a:bodyPr\b([^>]*)\/?>/);
  const vAttr = bodyPr ? bodyPr[1].match(/\banchor="([^"]+)"/) : null;
  const vAlign = vAttr ? (vAttr[1] === 't' ? 'top' : vAttr[1] === 'b' ? 'bottom' : 'middle') : 'middle';
  const vertM = bodyPr ? bodyPr[1].match(/\bvert="([^"]+)"/) : null;
  const vertical = vertM && vertM[1] !== 'horz';

  const paragraphs = [];
  const pRegex = /<a:p\b[^>]*>([\s\S]*?)<\/a:p>/g;
  let firstRPr = null;
  let hAlign = 'center';
  let m;
  while ((m = pRegex.exec(txBodyXml)) !== null) {
    const pInner = m[1];
    const algnM = pInner.match(/<a:pPr\b[^>]*algn="([^"]+)"/);
    if (algnM) hAlign = algnM[1] === 'l' ? 'left' : algnM[1] === 'r' ? 'right' : 'center';
    const runs = [];
    const rRegex = /<a:r\b[^>]*>([\s\S]*?)<\/a:r>/g;
    let rm;
    while ((rm = rRegex.exec(pInner)) !== null) {
      const rBody = rm[1];
      const tM = rBody.match(/<a:t>([\s\S]*?)<\/a:t>/);
      if (tM) runs.push(xmlUnescape(tM[1]));
      if (!firstRPr) {
        const rPrM = rBody.match(/<a:rPr\b([^>]*)(?:\/>|>[\s\S]*?<\/a:rPr>)/);
        if (rPrM) firstRPr = { attrs: rPrM[1], body: rPrM[0] };
      }
    }
    paragraphs.push(runs.join(''));
  }
  const text = paragraphs.join('\n');

  const textStyle = { hAlign, vAlign };
  if (firstRPr) {
    const a = firstRPr.attrs;
    const szM = a.match(/\bsz="(\d+)"/);
    if (szM) textStyle.fontSize = parseInt(szM[1], 10) / 100;
    if (/\bb="1"/.test(a)) textStyle.bold = true;
    if (/\bi="1"/.test(a)) textStyle.italic = true;
    if (/\bu="sng"|u="dbl"/.test(a)) textStyle.underline = true;
    const body = firstRPr.body;
    const latinM = body.match(/<a:latin\b[^>]*typeface="([^"]+)"/);
    if (latinM) textStyle.fontFamily = latinM[1];
    const fillM = body.match(/<a:solidFill>([\s\S]*?)<\/a:solidFill>/);
    if (fillM) {
      const color = parseSrgbFromColor(fillM[1]);
      const alpha = parseAlphaFromColor(fillM[1]);
      if (color) textStyle.fill = { color, alpha };
    }
    const lnM = body.match(/<a:ln\b([^>]*)>([\s\S]*?)<\/a:ln>/);
    if (lnM) {
      const wM = lnM[1].match(/\bw="(\d+)"/);
      const width = wM ? parseInt(wM[1], 10) / EMU_PER_PX : 1;
      const innerFill = lnM[2].match(/<a:solidFill>([\s\S]*?)<\/a:solidFill>/);
      if (innerFill) {
        const color = parseSrgbFromColor(innerFill[1]);
        const alpha = parseAlphaFromColor(innerFill[1]);
        if (color) textStyle.outline = { color, alpha, width };
      }
    }
  }

  return { text, textStyle, vertical };
}

function xmlUnescape(s) {
  return String(s)
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

function parseSingleAnchor(anchorXml, colWidths, rowHeights, idSeed) {
  const fromBlock = anchorXml.match(/<xdr:from>([\s\S]*?)<\/xdr:from>/);
  const toBlock = anchorXml.match(/<xdr:to>([\s\S]*?)<\/xdr:to>/);
  if (!fromBlock || !toBlock) return null;

  const fCol = getIntElement(fromBlock[1], 'col');
  const fColOff = getIntElement(fromBlock[1], 'colOff');
  const fRow = getIntElement(fromBlock[1], 'row');
  const fRowOff = getIntElement(fromBlock[1], 'rowOff');
  const tCol = getIntElement(toBlock[1], 'col');
  const tColOff = getIntElement(toBlock[1], 'colOff');
  const tRow = getIntElement(toBlock[1], 'row');
  const tRowOff = getIntElement(toBlock[1], 'rowOff');

  const fromPx = anchorToPx(fCol, fColOff, fRow, fRowOff, colWidths, rowHeights);
  const toPx = anchorToPx(tCol, tColOff, tRow, tRowOff, colWidths, rowHeights);

  const x = Math.min(fromPx.x, toPx.x);
  const y = Math.min(fromPx.y, toPx.y);
  const w = Math.abs(toPx.x - fromPx.x);
  const h = Math.abs(toPx.y - fromPx.y);

  // Shape element
  const spMatch = anchorXml.match(/<xdr:sp\b[^>]*>([\s\S]*?)<\/xdr:sp>/);
  if (!spMatch) return null;
  const spInner = spMatch[1];

  const prstM = spInner.match(/<a:prstGeom\b[^>]*prst="([^"]+)"/);
  const prstGeom = prstM ? prstM[1] : 'rect';
  const preset = getPresetByPrstGeom(prstGeom);
  let presetId = preset ? preset.id : 'rect';

  const spPrM = spInner.match(/<xdr:spPr\b[^>]*>([\s\S]*?)<\/xdr:spPr>/);
  const spPr = spPrM ? spPrM[1] : '';

  const xfrmM = spPr.match(/<a:xfrm\b([^>]*)>/);
  const xfrmAttrs = xfrmM ? xfrmM[1] : '';
  const rotM = xfrmAttrs.match(/\brot="(-?\d+)"/);
  const rot = rotM ? (parseInt(rotM[1], 10) / DEG_TO_OOXML_ANGLE) % 360 : 0;
  const flipH = /\bflipH="1"/.test(xfrmAttrs);
  const flipV = /\bflipV="1"/.test(xfrmAttrs);

  const fill = parseFill(spPr);
  const outline = parseOutline(spPr);
  const effects = parseEffects(spPr);

  const txBodyM = spInner.match(/<xdr:txBody\b[^>]*>([\s\S]*?)<\/xdr:txBody>/);
  const textInfo = txBodyM ? parseTextBody(txBodyM[0]) : null;

  // Promote rect with no fill + no outline to textBox; if the body has a
  // non-horizontal `vert` attribute, promote to verticalTextBox.
  if (presetId === 'rect') {
    const noFill = !fill || fill.type === 'none';
    const noOutline = !outline || !outline.color;
    if (noFill && noOutline) {
      const vertical = textInfo?.vertical;
      presetId = vertical ? 'verticalTextBox' : 'textBox';
    }
  }

  const shape = {
    id: `shp_imp_${idSeed}_${Math.random().toString(36).slice(2, 8)}`,
    type: presetId,
    x,
    y,
    w,
    h,
    rot: Math.round(((rot % 360) + 360) % 360),
    z: idSeed + 1,
    flipH,
    flipV,
    style: {
      fill: fill || { type: 'solid', color: '#2383e2', alpha: 1 },
      outline: outline || { color: '#0b5394', alpha: 1, width: 1, dash: 'solid' },
      effects,
    },
    textStyle: textInfo?.textStyle || { fontFamily: 'Calibri', fontSize: 12, fill: { color: '#1f1f1f', alpha: 1 }, hAlign: 'center', vAlign: 'middle' },
  };
  if (textInfo && textInfo.text) {
    shape.text = { value: textInfo.text };
  } else if (preset && preset.kind !== 'line') {
    shape.text = { value: '' };
  }
  return shape;
}

function parseDrawingShapes(drawingXml, colWidths, rowHeights) {
  const out = [];
  const anchorRegex = /<xdr:twoCellAnchor\b[^>]*>[\s\S]*?<\/xdr:twoCellAnchor>|<xdr:oneCellAnchor\b[^>]*>[\s\S]*?<\/xdr:oneCellAnchor>|<xdr:absoluteAnchor\b[^>]*>[\s\S]*?<\/xdr:absoluteAnchor>/g;
  let m;
  let idx = 0;
  while ((m = anchorRegex.exec(drawingXml)) !== null) {
    const shape = parseSingleAnchor(m[0], colWidths, rowHeights, idx++);
    if (shape) out.push(shape);
  }
  return out;
}

/**
 * @param {ArrayBuffer | Uint8Array} xlsxBytes
 * @returns {Record<string, object[]>} sheetName -> shape objects
 */
export function extractShapes(xlsxBytes) {
  try {
    const source = xlsxBytes instanceof Uint8Array ? xlsxBytes : new Uint8Array(xlsxBytes);
    let unzipped;
    try { unzipped = unzipSync(source); } catch { return {}; }

    const workbookXml = unzipped['xl/workbook.xml'] ? strFromU8(unzipped['xl/workbook.xml']) : '';
    const relsXml = unzipped['xl/_rels/workbook.xml.rels'] ? strFromU8(unzipped['xl/_rels/workbook.xml.rels']) : '';
    if (!workbookXml || !relsXml) return {};
    const sheetPaths = parseSheetPaths(workbookXml, relsXml);

    const out = {};
    for (const [sheetName, sheetPath] of Object.entries(sheetPaths)) {
      if (!unzipped[sheetPath]) continue;
      const sheetXml = strFromU8(unzipped[sheetPath]);
      const relsP = sheetRelsPath(sheetPath);
      const sheetRelsXml = unzipped[relsP] ? strFromU8(unzipped[relsP]) : null;
      const drawingPath = parseDrawingTarget(sheetRelsXml);
      if (!drawingPath || !unzipped[drawingPath]) continue;
      const drawingXml = strFromU8(unzipped[drawingPath]);
      const colWidths = parseColWidthsFromSheet(sheetXml);
      const rowHeights = parseRowHeightsFromSheet(sheetXml);
      const shapes = parseDrawingShapes(drawingXml, colWidths, rowHeights);
      if (shapes.length > 0) out[sheetName] = shapes;
    }
    return out;
  } catch {
    return {};
  }
}
