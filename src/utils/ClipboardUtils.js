// Excel-compatible clipboard codec for the DataGrid.
//
// Copy path writes two representations atomically:
//   - text/plain: a tab-separated grid that Excel, LibreOffice, Google
//     Sheets, and every text editor understand.
//   - text/html: a <table> whose root element carries a data-ganttgen-clip
//     attribute holding the base64-encoded high-fidelity payload (cell
//     values, formulas, per-cell style object, merge offsets, column
//     widths, and row heights). External apps see a regular HTML table
//     (so <td rowspan>/<td colspan>, <col width>, and <tr height>
//     reconstruct merges and layout natively); GanttGen reads the marker
//     attribute and restores the original cell objects verbatim.
//
// Paste path inspects whichever formats the clipboard carries and returns
// one of three shapes for the caller to apply: 'rich' (our own payload),
// 'html' (generic spreadsheet HTML), or 'tsv' (fallback plain text). For
// the 'html' path, cell styles, merges, column widths, and row heights
// are extracted so that an Excel -> GanttGen paste preserves the original
// layout in addition to values.

const CLIP_MARKER_ATTR = 'data-ganttgen-clip';
const CLIP_VERSION = 1;

function colLabel(index) {
  let label = '';
  let n = index;
  while (n >= 0) {
    label = String.fromCharCode(65 + (n % 26)) + label;
    n = Math.floor(n / 26) - 1;
  }
  return label;
}

function cellKey(row, col) {
  return `${colLabel(col)}${row + 1}`;
}

// ---------- TSV encode / decode ----------------------------------------

function encodeTsvField(text) {
  const s = text == null ? '' : String(text);
  if (s.includes('\t') || s.includes('\n') || s.includes('\r') || s.includes('"')) {
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

function tsvEncode(grid) {
  return grid.map((row) => row.map(encodeTsvField).join('\t')).join('\r\n');
}

// CSV/TSV state machine with quoted-field support. Rows separated by \r\n,
// \n, or \r; fields separated by \t; fields may be wrapped in double
// quotes to embed tabs or newlines, with "" escaping a literal quote.
function tsvDecode(text) {
  if (!text) return [[]];
  const rows = [];
  let row = [];
  let field = '';
  let inQuotes = false;
  let i = 0;
  while (i < text.length) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        inQuotes = false;
        i++;
        continue;
      }
      field += ch;
      i++;
      continue;
    }
    if (ch === '"') {
      if (field.length === 0) {
        inQuotes = true;
        i++;
        continue;
      }
      field += ch;
      i++;
      continue;
    }
    if (ch === '\t') {
      row.push(field);
      field = '';
      i++;
      continue;
    }
    if (ch === '\r' || ch === '\n') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
      if (ch === '\r' && text[i + 1] === '\n') i += 2;
      else i++;
      continue;
    }
    field += ch;
    i++;
  }
  row.push(field);
  rows.push(row);
  if (
    rows.length > 1 &&
    rows[rows.length - 1].length === 1 &&
    rows[rows.length - 1][0] === ''
  ) {
    rows.pop();
  }
  return rows;
}

// ---------- HTML escape / base64 --------------------------------------

function escapeHtml(text) {
  return String(text == null ? '' : text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function utf8ToBase64(str) {
  const bytes = new TextEncoder().encode(str);
  let binary = '';
  const CHUNK = 0x8000;
  for (let i = 0; i < bytes.length; i += CHUNK) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + CHUNK));
  }
  return btoa(binary);
}

function base64ToUtf8(b64) {
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return new TextDecoder().decode(bytes);
}

// ---------- CSS parsing helpers ---------------------------------------

const NAMED_COLORS = {
  black: '#000000', white: '#ffffff', red: '#ff0000', green: '#008000',
  blue: '#0000ff', yellow: '#ffff00', cyan: '#00ffff', magenta: '#ff00ff',
  gray: '#808080', grey: '#808080', silver: '#c0c0c0', maroon: '#800000',
  olive: '#808000', lime: '#00ff00', navy: '#000080', purple: '#800080',
  teal: '#008080', aqua: '#00ffff', fuchsia: '#ff00ff',
};

function normalizeColor(v) {
  if (!v) return null;
  const s = String(v).trim().toLowerCase();
  if (!s) return null;
  if (s === 'windowtext' || s === 'transparent' || s === 'none' || s === 'auto') return null;
  if (/^#[0-9a-f]{6}$/.test(s)) return s;
  if (/^#[0-9a-f]{3}$/.test(s)) {
    return '#' + s[1] + s[1] + s[2] + s[2] + s[3] + s[3];
  }
  const rgb = /^rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$/i.exec(s);
  if (rgb) {
    return (
      '#' +
      [rgb[1], rgb[2], rgb[3]]
        .map((n) => Math.max(0, Math.min(255, parseInt(n, 10))).toString(16).padStart(2, '0'))
        .join('')
    );
  }
  const rgba = /^rgba\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*([\d.]+)\s*\)$/i.exec(s);
  if (rgba) {
    return (
      '#' +
      [rgba[1], rgba[2], rgba[3]]
        .map((n) => Math.max(0, Math.min(255, parseInt(n, 10))).toString(16).padStart(2, '0'))
        .join('')
    );
  }
  if (NAMED_COLORS[s]) return NAMED_COLORS[s];
  return null;
}

function parseStyleDecl(text) {
  const out = {};
  if (!text) return out;
  for (const part of String(text).split(';')) {
    const idx = part.indexOf(':');
    if (idx < 0) continue;
    const k = part.slice(0, idx).trim().toLowerCase();
    const v = part.slice(idx + 1).trim();
    if (k) out[k] = v;
  }
  return out;
}

// Parse all `.className { ... }` rules from one or more <style> blocks.
// Excel clipboard HTML leans on class-based styling (e.g. class="xl65"),
// so without this step most of the formatting stays invisible to us.
function parseCssClassRules(cssText) {
  const map = {};
  if (!cssText) return map;
  // Strip HTML comments commonly wrapping Office clipboard CSS (<!-- ... -->).
  const cleaned = cssText.replace(/<!--|-->/g, '');
  // Match each rule block. Very permissive: class name, optional extra
  // selectors joined with comma, then { declarations }. We only keep the
  // class name at the start of a simple class selector; more complex
  // selectors (e.g. ".xl65 td") fall through unmatched, which is fine.
  const ruleRe = /\.([A-Za-z_][\w-]*)\s*(?:,\s*\.[A-Za-z_][\w-]*\s*)*\{([^}]*)\}/g;
  let m;
  while ((m = ruleRe.exec(cleaned)) !== null) {
    const cls = m[1];
    const decl = parseStyleDecl(m[2]);
    map[cls] = { ...(map[cls] || {}), ...decl };
  }
  return map;
}

function parseBorderShorthand(v) {
  if (!v) return null;
  const s = String(v).trim().toLowerCase();
  if (s === 'none' || s === '0' || s === '0pt' || s === '0px' || s === 'hidden') return null;
  // Pull out width + style + color in any order.
  const parts = s.split(/\s+/);
  let widthPt = null;
  let style = 'solid';
  let color = null;
  for (const p of parts) {
    if (/^[\d.]+pt$/.test(p)) widthPt = parseFloat(p);
    else if (/^[\d.]+px$/.test(p)) widthPt = parseFloat(p) * 0.75;
    else if (/^(solid|dashed|dotted|double|groove|ridge|inset|outset)$/.test(p)) {
      style = p;
    } else {
      const c = normalizeColor(p);
      if (c) color = c;
    }
  }
  if (style === 'double') return color ? { style: 'double', color } : { style: 'double' };
  if (style === 'dashed') return color ? { style: 'dashed', color } : { style: 'dashed' };
  if (style === 'dotted') return color ? { style: 'dotted', color } : { style: 'dotted' };
  // Everything else maps to the app's thin/medium/thick based on width.
  let weight = 'thin';
  if (widthPt != null) {
    if (widthPt >= 2) weight = 'thick';
    else if (widthPt >= 1.25) weight = 'medium';
  }
  return color ? { style: weight, color } : { style: weight };
}

// Build the app-side `cell.s` style object from the combined CSS
// declaration (class rules + inline style) plus HTML attributes.
function extractCellStyle(decl, bgcolor) {
  const s = {};

  const fw = decl['font-weight'];
  if (fw) {
    const n = parseInt(fw, 10);
    if (fw === 'bold' || fw === 'bolder' || (Number.isFinite(n) && n >= 600)) s.bold = true;
  }

  if (decl['font-style'] === 'italic') s.italic = true;

  const td = decl['text-decoration'] || decl['text-decoration-line'];
  if (td && /underline/.test(td)) s.underline = true;

  const fg = normalizeColor(decl['color']);
  if (fg && fg !== '#000000') s.color = fg;

  // background-color and background shorthand both appear in Excel HTML.
  const bgDecl =
    normalizeColor(decl['background-color']) ||
    normalizeColor((decl['background'] || '').split(/\s+/)[0]) ||
    normalizeColor(bgcolor);
  if (bgDecl && bgDecl !== '#ffffff') s.bg = bgDecl;

  const ta = (decl['text-align'] || '').toLowerCase();
  if (ta === 'left' || ta === 'center' || ta === 'right') s.hAlign = ta;

  const va = (decl['vertical-align'] || '').toLowerCase();
  if (va === 'top') s.vAlign = 'top';
  else if (va === 'middle' || va === 'center') s.vAlign = 'middle';
  else if (va === 'bottom') s.vAlign = 'bottom';

  const fs = decl['font-size'];
  if (fs) {
    const m = /^([\d.]+)\s*(pt|px)?$/.exec(String(fs).trim().toLowerCase());
    if (m) {
      const n = parseFloat(m[1]);
      const unit = m[2] || 'pt';
      const px = unit === 'pt' ? Math.round((n * 4) / 3) : Math.round(n);
      if (px > 0 && px !== 12) s.fontSize = px;
    }
  }

  const borders = {};
  // Short-hand first; per-side overrides it below.
  const shared = parseBorderShorthand(decl['border']);
  if (shared) {
    borders.top = shared;
    borders.right = shared;
    borders.bottom = shared;
    borders.left = shared;
  }
  for (const side of ['top', 'right', 'bottom', 'left']) {
    const b = parseBorderShorthand(decl['border-' + side]);
    if (b) borders[side] = b;
    else if (decl['border-' + side + '-style'] === 'none') delete borders[side];
  }
  if (Object.keys(borders).length > 0) s.borders = borders;

  return s;
}

// ---------- buildGridClipboard -----------------------------------------

/**
 * Build the clipboard payload for a rectangular selection.
 *
 * @param {object} args
 * @param {object} args.cells         per-cell map keyed by "A1" etc.
 * @param {Array}  args.merges        per-tab merges list
 * @param {object} args.rect          { r1, c1, r2, c2 } inclusive
 * @param {object} args.displayValues computed display values for formulas
 * @param {object} [args.colWidths]   per-column pixel widths keyed by label
 * @param {object} [args.rowHeights]  per-row pixel heights keyed by row index
 * @returns {{ tsv: string, html: string, json: object }}
 */
export function buildGridClipboard({
  cells,
  merges,
  rect,
  displayValues,
  colWidths,
  rowHeights,
}) {
  const rowCount = rect.r2 - rect.r1 + 1;
  const colCount = rect.c2 - rect.c1 + 1;

  // Only merges fully contained inside the rect are emitted. Partial
  // overlaps are out of scope for copy (Excel rejects them too).
  const containedMerges = (merges || [])
    .filter((m) => m.r1 >= rect.r1 && m.r2 <= rect.r2 && m.c1 >= rect.c1 && m.c2 <= rect.c2)
    .map((m) => ({
      r1: m.r1 - rect.r1,
      c1: m.c1 - rect.c1,
      r2: m.r2 - rect.r1,
      c2: m.c2 - rect.c1,
    }));

  const covered = new Set();
  for (const m of containedMerges) {
    for (let rr = m.r1; rr <= m.r2; rr++) {
      for (let cc = m.c1; cc <= m.c2; cc++) {
        if (rr === m.r1 && cc === m.c1) continue;
        covered.add(`${rr},${cc}`);
      }
    }
  }
  const anchorMap = new Map();
  for (const m of containedMerges) anchorMap.set(`${m.r1},${m.c1}`, m);

  // Rect-local column width / row height subsets. Keys are 0..N-1 offsets
  // so the rich-paste side can translate them to absolute positions.
  const pickedColWidths = {};
  if (colWidths) {
    for (let c = 0; c < colCount; c++) {
      const label = colLabel(rect.c1 + c);
      const w = colWidths[label];
      if (Number.isFinite(w) && w > 0) pickedColWidths[c] = w;
    }
  }
  const pickedRowHeights = {};
  if (rowHeights) {
    for (let r = 0; r < rowCount; r++) {
      const h = rowHeights[rect.r1 + r];
      if (Number.isFinite(h) && h > 0) pickedRowHeights[r] = h;
    }
  }

  // High-fidelity payload for GanttGen-to-GanttGen paste.
  const jsonCells = [];
  for (let r = 0; r < rowCount; r++) {
    const row = [];
    for (let c = 0; c < colCount; c++) {
      const key = cellKey(rect.r1 + r, rect.c1 + c);
      const cell = cells[key];
      if (!cell) { row.push(null); continue; }
      const out = {};
      if (cell.v !== undefined) out.v = cell.v;
      if (cell.f !== undefined) out.f = cell.f;
      if (cell.s && Object.keys(cell.s).length > 0) out.s = cell.s;
      row.push(Object.keys(out).length > 0 ? out : null);
    }
    jsonCells.push(row);
  }

  const json = {
    app: 'GanttGen',
    version: CLIP_VERSION,
    rows: rowCount,
    cols: colCount,
    cells: jsonCells,
    merges: containedMerges,
    colWidths: pickedColWidths,
    rowHeights: pickedRowHeights,
  };

  // TSV: display value for formulas, skip cells covered by a merge anchor.
  const tsvGrid = [];
  for (let r = 0; r < rowCount; r++) {
    const row = [];
    for (let c = 0; c < colCount; c++) {
      if (covered.has(`${r},${c}`)) { row.push(''); continue; }
      const key = cellKey(rect.r1 + r, rect.c1 + c);
      const dv = displayValues ? displayValues[key] : undefined;
      const cell = cells[key];
      const v = dv != null && dv !== '' ? dv : cell ? cell.v : '';
      row.push(v == null ? '' : v);
    }
    tsvGrid.push(row);
  }
  const tsv = tsvEncode(tsvGrid);

  // HTML representation. Modelled on the clipboard HTML Microsoft Excel
  // emits itself when copying a range. Full <html>/<head>/<body>
  // envelope with Office namespaces, `<meta name="ProgId"
  // content="Excel.Sheet">` to route the paste through Excel's
  // spreadsheet filter, Office-specific `mso-width-source:userset` and
  // `mso-width-alt:N` hints on each `<col>`, and per-<td> `width` /
  // `height` attributes plus matching `style="width:Npt;height:Npt"`
  // declarations. Without those hints Excel's filter treats the content
  // as generic HTML and silently drops column widths / row heights on
  // paste into an existing sheet.
  const marker = utf8ToBase64(JSON.stringify(json));
  // Excel's mso-width-alt unit is 1/256 of a character; at the default
  // Calibri 11pt ~ 7 px/char that maps to px * 256 / 7.
  const pxToMsoAlt = (px) => Math.round((px * 256) / 7);

  // Compute the table's own width attribute (Excel picks up the table
  // box size from this, then applies per-col widths inside it).
  let tableWidthPx = 0;
  for (let c = 0; c < colCount; c++) {
    tableWidthPx += pickedColWidths[c] || 80; // 80 = DataGrid DEFAULT_COL_WIDTH
  }
  const tableWidthPt = (tableWidthPx * 0.75).toFixed(2);

  let html = '';
  html +=
    '<html xmlns:v="urn:schemas-microsoft-com:vml" ' +
    'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
    'xmlns:x="urn:schemas-microsoft-com:office:excel" ' +
    'xmlns="http://www.w3.org/TR/REC-html40">';
  html += '<head>';
  html += '<meta http-equiv="Content-Type" content="text/html; charset=utf-8">';
  html += '<meta name="ProgId" content="Excel.Sheet">';
  html += '<meta name="Generator" content="GanttGen">';
  // Office conditional-comment hint block. Excel's own clipboard output
  // always includes this; it is the trigger some Excel versions use to
  // switch their HTML-paste filter from the generic web-table codepath
  // (which ignores column widths / row heights) to the spreadsheet
  // fragment codepath (which honours them). Every other recipient
  // (browsers, LibreOffice, Google Sheets) ignores the comment block.
  html += '<!--[if gte mso 9]><xml>';
  html += '<x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet>';
  html += '<x:Name>Sheet1</x:Name>';
  html += '<x:WorksheetOptions>';
  html += '<x:ProtectContents>False</x:ProtectContents>';
  html += '<x:ProtectObjects>False</x:ProtectObjects>';
  html += '<x:ProtectScenarios>False</x:ProtectScenarios>';
  html += '</x:WorksheetOptions>';
  html += '</x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook>';
  html += '</xml><![endif]-->';
  html += '</head>';
  html += '<body>';
  html +=
    `<table ${CLIP_MARKER_ATTR}="${marker}" border="0" cellspacing="0" cellpadding="0" ` +
    `width="${tableWidthPx}" style="border-collapse:collapse;width:${tableWidthPt}pt">`;

  // <colgroup> ordering matters even though we repeat widths on each td;
  // Excel reads mso-width-alt from here as the authoritative width in
  // its internal 1/256-character units.
  html += '<colgroup>';
  for (let c = 0; c < colCount; c++) {
    const w = pickedColWidths[c];
    if (w) {
      const pt = (w * 0.75).toFixed(2);
      const alt = pxToMsoAlt(w);
      html +=
        `<col width="${w}" ` +
        `style="mso-width-source:userset;mso-width-alt:${alt};width:${pt}pt">`;
    } else {
      html += '<col>';
    }
  }
  html += '</colgroup>';

  for (let r = 0; r < rowCount; r++) {
    const rh = pickedRowHeights[r];
    if (rh) {
      const pt = (rh * 0.75).toFixed(2);
      html += `<tr height="${rh}" style="height:${pt}pt">`;
    } else {
      html += '<tr>';
    }
    let firstTdOfRow = true;
    for (let c = 0; c < colCount; c++) {
      if (covered.has(`${r},${c}`)) continue;
      const anchor = anchorMap.get(`${r},${c}`);
      const rs = anchor ? anchor.r2 - anchor.r1 + 1 : 1;
      const cs = anchor ? anchor.c2 - anchor.c1 + 1 : 1;
      const attrs = [];
      if (rs > 1) attrs.push(`rowspan="${rs}"`);
      if (cs > 1) attrs.push(`colspan="${cs}"`);

      // Per-cell width attribute (sum of spanned columns, in px).
      // Excel's paste applies width from the first <td> it sees for a
      // given column; redundant with <col> but both sides reinforce.
      const cellColWidthPx = sumColWidths(pickedColWidths, c, cs);
      if (cellColWidthPx) attrs.push(`width="${cellColWidthPx}"`);

      // First <td> of a row is where Excel picks up row height.
      if (firstTdOfRow && rh) attrs.push(`height="${rh}"`);

      const key = cellKey(rect.r1 + r, rect.c1 + c);
      const dv = displayValues ? displayValues[key] : undefined;
      const cell = cells[key];
      const v = dv != null && dv !== '' ? dv : cell ? cell.v : '';
      const text = v == null ? '' : String(v);

      const styleParts = [];
      if (firstTdOfRow && rh) {
        styleParts.push(`height:${(rh * 0.75).toFixed(2)}pt`);
      }
      if (cellColWidthPx) {
        styleParts.push(`width:${(cellColWidthPx * 0.75).toFixed(2)}pt`);
      }
      if (cell && cell.s) {
        const s = cellStyleToInlineCss(cell.s);
        if (s) styleParts.push(s);
      }
      if (styleParts.length > 0) attrs.push(`style="${styleParts.join(';')}"`);
      if (cell && cell.s && cell.s.bg) attrs.push(`bgcolor="${cell.s.bg}"`);

      html += `<td${attrs.length ? ' ' + attrs.join(' ') : ''}>${escapeHtml(text)}</td>`;
      firstTdOfRow = false;
    }
    html += '</tr>';
  }
  html += '</table></body></html>';

  return { tsv, html, json };
}

// Sum the rect-local pixel widths of `cs` consecutive columns starting
// at offset `c`. Returns null if none of the covered columns carry an
// explicit width (in which case we don't emit a td-level width at all,
// letting Excel fall back to the default column width).
function sumColWidths(pickedColWidths, c, cs) {
  let total = 0;
  let any = false;
  for (let i = 0; i < cs; i++) {
    const w = pickedColWidths[c + i];
    if (Number.isFinite(w) && w > 0) {
      total += w;
      any = true;
    } else {
      total += 80; // default
    }
  }
  return any ? total : null;
}

// Convert the app-side cell.s style object back into an inline CSS
// declaration that Excel, LibreOffice, and Google Sheets all recognise
// when pasting. Width / height are emitted separately in the td-building
// loop so the caller can thread mso-* hints alongside them.
function cellStyleToInlineCss(s) {
  if (!s) return '';
  const parts = [];
  // 700 is the numeric weight Office HTML uses on its own output; `bold`
  // works too but numeric is the form Excel round-trips most reliably.
  if (s.bold) parts.push('font-weight:700');
  if (s.italic) parts.push('font-style:italic');
  if (s.underline) parts.push('text-decoration:underline');
  if (s.color) parts.push(`color:${s.color}`);
  if (s.bg) {
    // Excel's HTML filter reads the `background` shorthand; the standard
    // `background-color` also sticks and keeps LibreOffice / Sheets happy.
    parts.push(`background:${s.bg}`);
    parts.push(`background-color:${s.bg}`);
  }
  if (s.fontSize) {
    const pt = (s.fontSize * 0.75).toFixed(2);
    parts.push(`font-size:${pt}pt`);
  }
  if (s.hAlign) parts.push(`text-align:${s.hAlign}`);
  if (s.vAlign) {
    const v = s.vAlign === 'middle' ? 'middle' : s.vAlign;
    parts.push(`vertical-align:${v}`);
  }
  if (s.borders) {
    for (const side of ['top', 'right', 'bottom', 'left']) {
      const b = s.borders[side];
      if (!b) continue;
      const w =
        b.style === 'thick' ? '2.0pt' : b.style === 'medium' ? '1.5pt' : '0.5pt';
      const css =
        b.style === 'double' || b.style === 'dashed' || b.style === 'dotted'
          ? b.style
          : 'solid';
      const color = b.color || 'windowtext';
      parts.push(`border-${side}:${w} ${css} ${color}`);
    }
  }
  return parts.join(';');
}

// ---------- parseClipboardInput ----------------------------------------

function parseMarkerPayload(htmlText) {
  if (!htmlText || typeof htmlText !== 'string') return null;
  if (htmlText.indexOf(CLIP_MARKER_ATTR) === -1) return null;
  let doc;
  try {
    doc = new DOMParser().parseFromString(htmlText, 'text/html');
  } catch {
    return null;
  }
  const table = doc.querySelector(`table[${CLIP_MARKER_ATTR}]`);
  if (!table) return null;
  const raw = table.getAttribute(CLIP_MARKER_ATTR);
  if (!raw) return null;
  try {
    const payload = JSON.parse(base64ToUtf8(raw));
    if (!payload || payload.app !== 'GanttGen') return null;
    return payload;
  } catch {
    return null;
  }
}

// Pull visible text out of a cell. textContent drops <br>; iterate
// childNodes to preserve line breaks in Excel-copied multi-line cells.
function extractCellText(node) {
  let out = '';
  for (const child of node.childNodes) {
    if (child.nodeType === 3) {
      out += child.textContent;
    } else if (child.nodeType === 1) {
      const tag = child.tagName;
      if (tag === 'BR') out += '\n';
      else out += extractCellText(child);
    }
  }
  return out.replace(/\u00a0/g, ' ');
}

// Parse a generic spreadsheet-style HTML table (Excel, LibreOffice,
// Google Sheets all emit a roughly similar <table> shape). Honours
// rowspan / colspan for merges, and extracts cell-level styles, column
// widths, and row heights so the full visual layout round-trips.
function parseHtmlTable(htmlText) {
  if (!htmlText || typeof htmlText !== 'string') return null;
  let doc;
  try {
    doc = new DOMParser().parseFromString(htmlText, 'text/html');
  } catch {
    return null;
  }
  const table = doc.querySelector('table');
  if (!table) return null;

  // Aggregate every <style> block so class rules resolve. Office HTML
  // puts most formatting in a single block at the top of the fragment.
  const cssText = Array.from(doc.querySelectorAll('style'))
    .map((s) => s.textContent || '')
    .join('\n');
  const classMap = parseCssClassRules(cssText);

  // Column widths from <col width="...">. Excel emits px values here.
  const colWidthsByIdx = {};
  const colEls = Array.from(table.querySelectorAll('colgroup col, col'));
  let colAccumulator = 0;
  for (const col of colEls) {
    const span = parseInt(col.getAttribute('span') || '1', 10) || 1;
    const widthAttr = col.getAttribute('width');
    const wAttr = parseInt(widthAttr, 10);
    let width = Number.isFinite(wAttr) && wAttr > 0 ? wAttr : null;
    if (width == null) {
      const styleW = parseStyleDecl(col.getAttribute('style') || '')['width'];
      if (styleW) {
        const m = /^([\d.]+)\s*(px|pt)?$/.exec(String(styleW).trim().toLowerCase());
        if (m) {
          const n = parseFloat(m[1]);
          width = m[2] === 'pt' ? Math.round((n * 4) / 3) : Math.round(n);
        }
      }
    }
    for (let i = 0; i < span; i++) {
      if (width != null) colWidthsByIdx[colAccumulator] = width;
      colAccumulator++;
    }
  }

  const rowHeightsByIdx = {};

  // Keep every <tr>, including those with no <td> / <th>. Excel emits an
  // empty <tr> for rows that are fully covered by a multi-row merge in an
  // earlier row; dropping them compresses the vertical extent and pushes
  // later cells sideways into the columns that the merge "owns". The
  // `occupied` map below handles the empty rows correctly on its own.
  const rowEls = Array.from(table.querySelectorAll('tr'));
  if (rowEls.length === 0) return null;

  // Seed the grid width from the colgroup (authoritative when present),
  // else from the widest row's td count. The per-row loop widens this if
  // a later row turns out to be wider (e.g. header row with fewer cells
  // than a body row containing colspan expansions).
  let totalCols = colEls.reduce(
    (acc, col) => acc + (parseInt(col.getAttribute('span') || '1', 10) || 1),
    0,
  );
  if (totalCols === 0) {
    for (const tr of rowEls) {
      let w = 0;
      for (const td of tr.querySelectorAll('td, th')) {
        w += parseInt(td.getAttribute('colspan') || '1', 10) || 1;
      }
      if (w > totalCols) totalCols = w;
    }
  }

  const grid = [];
  const cellStyles = [];
  const merges = [];
  const occupied = new Map();

  for (let r = 0; r < rowEls.length; r++) {
    const tr = rowEls[r];
    const hAttr = parseInt(tr.getAttribute('height'), 10);
    if (Number.isFinite(hAttr) && hAttr > 0) {
      rowHeightsByIdx[r] = hAttr;
    } else {
      const styleH = parseStyleDecl(tr.getAttribute('style') || '')['height'];
      if (styleH) {
        const m = /^([\d.]+)\s*(px|pt)?$/.exec(String(styleH).trim().toLowerCase());
        if (m) {
          const n = parseFloat(m[1]);
          const px = m[2] === 'pt' ? Math.round((n * 4) / 3) : Math.round(n);
          if (px > 0) rowHeightsByIdx[r] = px;
        }
      }
    }

    grid.push([]);
    cellStyles.push([]);
    const cells = Array.from(tr.querySelectorAll('td, th'));
    let c = 0;
    let cellIdx = 0;
    while (cellIdx < cells.length) {
      while (occupied.get(`${r},${c}`)) {
        grid[r][c] = '';
        cellStyles[r][c] = null;
        c++;
      }
      const cell = cells[cellIdx];
      const rs = parseInt(cell.getAttribute('rowspan') || '1', 10) || 1;
      const cs = parseInt(cell.getAttribute('colspan') || '1', 10) || 1;

      const classNames = (cell.getAttribute('class') || '').split(/\s+/).filter(Boolean);
      let decl = {};
      for (const cn of classNames) {
        if (classMap[cn]) decl = { ...decl, ...classMap[cn] };
      }
      const inlineDecl = parseStyleDecl(cell.getAttribute('style') || '');
      decl = { ...decl, ...inlineDecl };
      const bgcolor = cell.getAttribute('bgcolor');
      const style = extractCellStyle(decl, bgcolor);

      const text = extractCellText(cell);
      grid[r][c] = text;
      cellStyles[r][c] = Object.keys(style).length > 0 ? style : null;

      if (rs > 1 || cs > 1) {
        merges.push({ r1: r, c1: c, r2: r + rs - 1, c2: c + cs - 1 });
      }
      for (let rr = 0; rr < rs; rr++) {
        for (let cc = 0; cc < cs; cc++) {
          if (rr === 0 && cc === 0) continue;
          occupied.set(`${r + rr},${c + cc}`, true);
          if (rr === 0) {
            grid[r][c + cc] = '';
            cellStyles[r][c + cc] = null;
          }
        }
      }
      c += cs;
      cellIdx++;
      if (c >= totalCols && cellIdx < cells.length) {
        totalCols = c + cs;
      }
    }
    while (grid[r].length < totalCols) {
      grid[r].push('');
      cellStyles[r].push(null);
    }
  }

  let width = 0;
  for (const row of grid) if (row.length > width) width = row.length;
  for (let i = 0; i < grid.length; i++) {
    while (grid[i].length < width) {
      grid[i].push('');
      cellStyles[i].push(null);
    }
  }

  return {
    grid,
    merges,
    cellStyles,
    colWidths: colWidthsByIdx,
    rowHeights: rowHeightsByIdx,
  };
}

/**
 * Inspect the clipboard and return the best-fidelity representation.
 *
 * @param {object} args
 * @param {string} args.htmlText  value of clipboardData.getData('text/html')
 * @param {string} args.plainText value of clipboardData.getData('text/plain')
 * @returns {null | { kind: 'rich', payload: object }
 *                 | { kind: 'html', grid: string[][], merges: object[],
 *                     cellStyles: object[][], colWidths: object,
 *                     rowHeights: object }
 *                 | { kind: 'tsv',  grid: string[][] }}
 */
export function parseClipboardInput({ htmlText, plainText }) {
  const rich = parseMarkerPayload(htmlText);
  if (rich) return { kind: 'rich', payload: rich };

  const html = parseHtmlTable(htmlText);
  if (html && html.grid.length > 0 && html.grid[0].length > 0) {
    return {
      kind: 'html',
      grid: html.grid,
      merges: html.merges,
      cellStyles: html.cellStyles,
      colWidths: html.colWidths,
      rowHeights: html.rowHeights,
    };
  }

  if (plainText && plainText.length > 0) {
    const grid = tsvDecode(plainText);
    if (grid.length > 0 && grid[0].length > 0) {
      return { kind: 'tsv', grid };
    }
  }

  return null;
}

/**
 * Coerce a string to a number when it unambiguously represents one.
 * Mirrors Excel's paste-time type inference for TSV / HTML input.
 */
export function coerceNumeric(text) {
  if (text == null) return '';
  const s = String(text).trim();
  if (s === '') return '';
  const n = Number(s);
  if (Number.isFinite(n) && String(n) === s.replace(/^\+/, '')) return n;
  return text;
}

/**
 * Convert a numeric column index (0-based) to an A1-style label.
 * Exposed so the DataGrid can translate rect-local colWidths offsets
 * into app-side labels when applying a rich / html paste.
 */
export function colLabelFromIndex(index) {
  return colLabel(index);
}
