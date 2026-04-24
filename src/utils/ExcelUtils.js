import * as XLSX from 'xlsx';
import { injectCellStyles } from './XlsxStyleInjector';
import { extractNativeCellStyles } from './XlsxStyleExtractor';
import { injectShapes } from './XlsxShapeInjector';
import { extractShapes } from './XlsxShapeExtractor';

// Bumped whenever the set of styles XlsxStyleInjector writes natively
// changes. Files tagged with version >= 1 treat the native xlsx bytes as
// authoritative for the keys listed in NATIVE_STYLE_KEYS; older files
// fall back to the gridCellStyles JSON blob. See ADR 003.
const NATIVE_CELL_STYLES_VERSION = '2';
const NATIVE_STYLE_KEYS = [
  'hAlign',
  'vAlign',
  'bold',
  'italic',
  'underline',
  'numFmt',
  'decimals',
  'currency',
  'useThousands',
  'negativeStyle',
];

// Mirror of NATIVE_CELL_STYLES_VERSION for the DrawingML shape codec.
// Files tagged with version >= 1 treat the embedded DrawingML parts as
// authoritative for shape geometry / styling. See ADR 004.
const NATIVE_SHAPES_VERSION = '1';

const TASK_COLUMNS = [
  'Task ID',
  'Task Name',
  'Dependency',
  'Task Category',
  'Start Date',
  'End Date',
  'Duration',
  'Progress (%)',
  'Status',
  'Owner',
  'Remarks',
  'Baseline Start',
  'Baseline End',
  'Parent ID',
];

const SETTINGS_FIELDS = [
  { key: 'projectName', label: 'Project Name', defaultValue: '' },
  { key: 'themeName', label: 'Theme Name', defaultValue: 'Notion Light' },
  { key: 'themeBgPrimary', label: 'Theme BG Primary', defaultValue: '#ffffff' },
  { key: 'themeBgSecondary', label: 'Theme BG Secondary', defaultValue: '#f7f7f5' },
  { key: 'themeBgTertiary', label: 'Theme BG Tertiary', defaultValue: '#edece9' },
  { key: 'themeBgHover', label: 'Theme BG Hover', defaultValue: '#e8e7e4' },
  { key: 'themeBorder', label: 'Theme Border', defaultValue: '#e0dfdc' },
  { key: 'themeBorderSubtle', label: 'Theme Border Subtle', defaultValue: '#ebebea' },
  { key: 'themeTextPrimary', label: 'Theme Text Primary', defaultValue: '#1f1f1f' },
  { key: 'themeTextSecondary', label: 'Theme Text Secondary', defaultValue: '#6b6b6b' },
  { key: 'themeTextMuted', label: 'Theme Text Muted', defaultValue: '#9b9b9b' },
  { key: 'themeAccent', label: 'Theme Accent', defaultValue: '#2383e2' },
  { key: 'themeAccentHover', label: 'Theme Accent Hover', defaultValue: '#1a6dbe' },
  { key: 'themeSuccess', label: 'Theme Success', defaultValue: '#0f7b0f' },
  { key: 'themeWarning', label: 'Theme Warning', defaultValue: '#d97706' },
  { key: 'themeDanger', label: 'Theme Danger', defaultValue: '#e03e3e' },
  { key: 'themeInfo', label: 'Theme Info', defaultValue: '#2383e2' },
  { key: 'themeCriticalPath', label: 'Theme Critical Path', defaultValue: '#e03e3e' },
  { key: 'showCriticalPath', label: 'Show Critical Path', defaultValue: 'false' },
  { key: 'showSlack', label: 'Show Slack', defaultValue: 'false' },
  { key: 'showDependencies', label: 'Show Dependencies', defaultValue: 'false' },
  { key: 'showTodayLine', label: 'Show Today Line', defaultValue: 'true' },
  { key: 'showBaseline', label: 'Show Baseline', defaultValue: 'true' },
  { key: 'skipWeekends', label: 'Skip Weekends', defaultValue: 'true' },
  { key: 'showScaleButtons', label: 'Show Scale Buttons', defaultValue: 'true' },
  { key: 'showWeekLabels', label: 'Show Week Labels', defaultValue: 'false' },
  { key: 'showMonthLabels', label: 'Show Month Labels', defaultValue: 'true' },
  { key: 'showDayLabels', label: 'Show Day Labels', defaultValue: 'true' },
  { key: 'visibleColumns', label: 'Visible Columns', defaultValue: 'id,name,duration,startDate,endDate,progress,status,owner' },
  { key: 'categoryColors', label: 'Category Colors', defaultValue: '{}' },
  { key: 'customColumns', label: 'Custom Columns', defaultValue: '[]' },
  { key: 'columnOrder', label: 'Column Order', defaultValue: '' },
  { key: 'splitRatio', label: 'Split Ratio', defaultValue: '0.38' },
  { key: 'collapsedParents', label: 'Collapsed Parents', defaultValue: '' },
  { key: 'colWidths', label: 'Column Widths', defaultValue: '{}' },
  { key: 'ganttScale', label: 'Gantt Scale', defaultValue: 'day' },
  { key: 'ganttZoom', label: 'Gantt Zoom', defaultValue: '100' },
  { key: 'tabs', label: 'Tabs', defaultValue: '[]' },
  { key: 'activeTab', label: 'Active Tab', defaultValue: 'gantt' },
  { key: 'gridCellStyles', label: 'Grid Cell Styles', defaultValue: '{}' },
  { key: 'nativeCellStylesVersion', label: 'Native Cell Styles Version', defaultValue: '0' },
  { key: 'gridShapes', label: 'Grid Shapes', defaultValue: '{}' },
  { key: 'nativeShapesVersion', label: 'Native Shapes Version', defaultValue: '0' },
];

export function downloadTemplate() {
  const wb = XLSX.utils.book_new();

  const taskHeader = [TASK_COLUMNS];
  const taskSheet = XLSX.utils.aoa_to_sheet(taskHeader);

  const colWidths = TASK_COLUMNS.map((col) => ({
    wch: Math.max(col.length + 2, 14),
  }));
  taskSheet['!cols'] = colWidths;

  XLSX.utils.book_append_sheet(wb, taskSheet, 'Tasks');

  const settingsData = SETTINGS_FIELDS.map((f) => [f.label, f.defaultValue]);
  settingsData.unshift(['Setting', 'Value']);
  const settingsSheet = XLSX.utils.aoa_to_sheet(settingsData);
  XLSX.utils.book_append_sheet(wb, settingsSheet, 'Settings');

  applyHiddenSheetFlags(wb);

  XLSX.writeFile(wb, 'GanttGen_Template.xlsx');
}

const REQUIRED_HEADERS = ['Task ID', 'Task Name'];

const ACCEPTED_EXTENSIONS = ['.xlsx', '.xls'];

function validateFileType(file) {
  const name = (file.name || '').toLowerCase();
  const ext = name.slice(name.lastIndexOf('.'));
  if (!ACCEPTED_EXTENSIONS.includes(ext)) {
    throw new Error(
      `Invalid file type "${ext}". Please import an Excel file (.xlsx or .xls).`,
    );
  }
}

function validateTaskHeaders(wb) {
  const sheetName = wb.SheetNames.find((n) => n.toLowerCase() === 'tasks');
  if (!sheetName) {
    throw new Error(
      'The Excel file has no "Tasks" sheet. Please use the GanttGen template.',
    );
  }
  const sheet = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  const headerRow = rows[0] || [];
  const headers = headerRow.map((h) => String(h).trim());

  const missing = REQUIRED_HEADERS.filter((h) => !headers.includes(h));
  if (missing.length > 0) {
    throw new Error(
      `Missing required columns: ${missing.join(', ')}. Please use the GanttGen template.`,
    );
  }
}

/**
 * @param {File} file
 * @returns {Promise<{ tasks: object[], settings: object }>}
 */
const RESERVED_SHEETS = new Set(['tasks', 'settings']);

export async function importExcel(file) {
  validateFileType(file);

  const data = await file.arrayBuffer();
  // cellStyles: true is required for SheetJS to parse <cols>/<row ht=...>
  // metadata back into sheet['!cols'] / sheet['!rows']. Without it the
  // reader discards column widths and row heights even though they are
  // present in the .xlsx file.
  const wb = XLSX.read(data, { type: 'array', cellDates: true, cellStyles: true });

  validateTaskHeaders(wb);

  const settings = parseSettingsSheet(wb);

  let customCols = [];
  try { customCols = JSON.parse(settings.customColumns || '[]'); } catch { /* ignore */ }

  const tasks = parseTasksSheet(wb, customCols);

  let tabs = [];
  try { tabs = JSON.parse(settings.tabs || '[]'); } catch { /* ignore */ }
  if (!Array.isArray(tabs)) tabs = [];

  // Auto-adopt any non-reserved worksheet the user created directly in
  // Excel (or any other OOXML tool) that the Settings-sheet tab list
  // has not caught up with. Without this the sheet's data is parsed
  // but the tab never appears in the UI because parseGridSheets keys
  // gridData by tab.id and App.jsx reads its tab list from settings.tabs.
  const knownTabNames = new Set(tabs.map((t) => t && t.name).filter(Boolean));
  const adoptStamp = Date.now().toString(36);
  let adoptIndex = 0;
  for (const sheetName of wb.SheetNames) {
    if (RESERVED_SHEETS.has(sheetName.toLowerCase())) continue;
    if (knownTabNames.has(sheetName)) continue;
    tabs.push({ id: `tab_${adoptStamp}_${adoptIndex++}`, name: sheetName });
    knownTabNames.add(sheetName);
  }

  let gridCellStyles = {};
  try { gridCellStyles = JSON.parse(settings.gridCellStyles || '{}'); } catch { /* ignore */ }

  let gridShapes = {};
  try { gridShapes = JSON.parse(settings.gridShapes || '{}'); } catch { /* ignore */ }

  // Read font toggles (bold / italic / underline) and alignment straight
  // from the xlsx XML. SheetJS Community does not expose these on
  // wsCell.s, so this is how Excel-native edits round-trip back into the
  // app. Extractor is best-effort: returns {} if the archive is unusable.
  const nativeSheetStyles = extractNativeCellStyles(data);
  const fileVersion = parseInt(settings.nativeCellStylesVersion || '0', 10) || 0;
  const nativeAuthoritative = fileVersion >= 1;

  const shapesVersion = parseInt(settings.nativeShapesVersion || '0', 10) || 0;
  const nativeShapesAuthoritative = shapesVersion >= 1;
  const nativeSheetShapes = nativeShapesAuthoritative ? extractShapes(data) : {};

  const gridData = parseGridSheets(
    wb,
    tabs,
    gridCellStyles,
    nativeSheetStyles,
    nativeAuthoritative,
    gridShapes,
    nativeSheetShapes,
    nativeShapesAuthoritative,
  );

  return { tasks, settings, gridData, tabs };
}

/**
 * Build an xlsx workbook from tasks + settings + optional grid data.
 */
function buildWorkbook(tasks, settings, gridData) {
  const wb = XLSX.utils.book_new();

  let customCols = [];
  try { customCols = JSON.parse(settings?.customColumns || '[]'); } catch { /* ignore */ }

  const allHeaders = [...TASK_COLUMNS, ...customCols.map((c) => c.label)];

  const taskRows = tasks.map((t) => {
    const row = [
      t.id ?? '',
      t.name ?? '',
      t.dependency ?? '',
      t.category ?? '',
      formatDate(t.startDate),
      formatDate(t.endDate),
      t.duration ?? '',
      t.progress ?? '',
      t.status ?? '',
      t.owner ?? '',
      t.remarks ?? '',
      formatDate(t.baselineStart),
      formatDate(t.baselineEnd),
      t.parentId ?? '',
    ];
    for (const cc of customCols) row.push(t[cc.key] ?? '');
    return row;
  });
  taskRows.unshift(allHeaders);

  const taskSheet = XLSX.utils.aoa_to_sheet(taskRows);
  const colWidths = allHeaders.map((col) => ({
    wch: Math.max(col.length + 2, 14),
  }));
  taskSheet['!cols'] = colWidths;
  XLSX.utils.book_append_sheet(wb, taskSheet, 'Tasks');

  const settingsRows = [['Setting', 'Value']];
  for (const field of SETTINGS_FIELDS) {
    // The injector runs unconditionally on export, so every file produced
    // by this build advertises the current native-styles format version.
    // Any value the caller passes for this key is intentionally ignored.
    let value;
    if (field.key === 'nativeCellStylesVersion') {
      value = NATIVE_CELL_STYLES_VERSION;
    } else if (field.key === 'nativeShapesVersion') {
      value = NATIVE_SHAPES_VERSION;
    } else {
      value =
        settings && settings[field.key] != null
          ? String(settings[field.key])
          : field.defaultValue;
    }
    settingsRows.push([field.label, value]);
  }
  const settingsSheet = XLSX.utils.aoa_to_sheet(settingsRows);
  XLSX.utils.book_append_sheet(wb, settingsSheet, 'Settings');

  let tabs = [];
  try { tabs = JSON.parse(settings?.tabs || '[]'); } catch { /* ignore */ }
  if (gridData && tabs.length > 0) {
    for (const tab of tabs) {
      const tabData = gridData[tab.id];
      if (!tabData) continue;
      const sheetName = sanitizeSheetName(tab.name);
      const gridSheet = gridDataToSheet(tabData);
      // Persist column widths in pixels (wpx) so the round-trip is exact.
      // Only entries that differ from the app default are written; this lets
      // unresized columns keep their default size on re-import. Unresized
      // columns are emitted as null (not {}) so SheetJS skips them entirely.
      // An empty `{}` produces a <col min="X" max="X"/> element with no width,
      // which Excel renders as a hidden column, causing A-Z to disappear when
      // no columns had been manually resized (issue #15).
      const colWidthEntries = tabData.colWidths || {};
      if (Object.keys(colWidthEntries).length > 0) {
        const maxCol = tabData.cols || 26;
        gridSheet['!cols'] = Array.from({ length: maxCol }, (_, i) => {
          const key = colLabelFromIndex(i);
          const w = colWidthEntries[key];
          return w ? { wpx: w } : null;
        });
      }
      const rowHeightEntries = tabData.rowHeights || {};
      if (Object.keys(rowHeightEntries).length > 0) {
        const maxRow = tabData.rows || 50;
        gridSheet['!rows'] = Array.from({ length: maxRow }, (_, i) => {
          const h = rowHeightEntries[i];
          return h ? { hpx: h, hpt: h } : null;
        });
      }
      XLSX.utils.book_append_sheet(wb, gridSheet, sheetName);
    }
  }

  applyHiddenSheetFlags(wb);

  return wb;
}

// The Settings sheet stores internal app metadata (theme, view toggles,
// gridCellStyles / gridShapes JSON blobs, native-format version markers,
// tab list) that users should not see on the Excel tab strip. OOXML
// `state="veryHidden"` keeps the sheet in the workbook so all round-trip
// code (parseSettingsSheet, XlsxStyleExtractor, XlsxShapeExtractor) keeps
// working, while making the tab invisible in Excel's UI - it can only be
// surfaced via VBA, not via a right-click Unhide. See issue #36.
function applyHiddenSheetFlags(wb) {
  if (!wb || !Array.isArray(wb.SheetNames)) return;
  wb.Workbook = wb.Workbook || {};
  wb.Workbook.Sheets = wb.SheetNames.map((name) => ({
    Hidden: name.toLowerCase() === 'settings' ? 2 : 0,
  }));
}

/**
 * Collect the cells in each grid tab that carry Excel-native-visible
 * formatting so the xlsx byte stream can be post-processed to embed them.
 * SheetJS Community does not emit cell styles on write, so formatting
 * applied via the toolbar is lost to other spreadsheet apps (Excel,
 * LibreOffice, Google Sheets) without this step. The GanttGen-internal
 * round-trip via `gridCellStyles` is unaffected.
 *
 * Returns only the subset of `cell.s` that XlsxStyleInjector currently
 * knows how to write: horizontal / vertical alignment, the three font
 * toggles (bold, italic, underline), and the Excel-style number format
 * fields (numFmt, decimals, currency, useThousands, negativeStyle).
 */
function collectSheetStyles(settings, gridData) {
  if (!gridData) return {};
  let tabs = [];
  try { tabs = JSON.parse(settings?.tabs || '[]'); } catch { /* ignore */ }
  const out = {};
  for (const tab of tabs) {
    const tabData = gridData[tab.id];
    if (!tabData?.cells) continue;
    const sheetName = sanitizeSheetName(tab.name);
    const cellStyles = {};
    for (const [cellRef, cell] of Object.entries(tabData.cells)) {
      const s = cell?.s;
      if (!s) continue;
      const picked = {};
      if (s.hAlign) picked.hAlign = s.hAlign;
      if (s.vAlign) picked.vAlign = s.vAlign;
      if (s.bold) picked.bold = true;
      if (s.italic) picked.italic = true;
      if (s.underline) picked.underline = true;
      if (s.numFmt && s.numFmt !== 'general') picked.numFmt = s.numFmt;
      if (typeof s.decimals === 'number') picked.decimals = s.decimals;
      if (typeof s.currency === 'string' && s.currency) picked.currency = s.currency;
      if (typeof s.useThousands === 'boolean') picked.useThousands = s.useThousands;
      if (typeof s.negativeStyle === 'string' && s.negativeStyle) {
        picked.negativeStyle = s.negativeStyle;
      }
      if (Object.keys(picked).length === 0) continue;
      cellStyles[cellRef] = picked;
    }
    if (Object.keys(cellStyles).length > 0) {
      out[sheetName] = cellStyles;
    }
  }
  return out;
}

/**
 * Collect shapes per sheet for the OOXML DrawingML injector. The shapes
 * carry dimensions in pixels in app-space; the injector converts to EMU
 * and emits `<xdr:twoCellAnchor editAs="absolute">` entries anchored via
 * the DataGrid's column / row offsets.
 */
function collectSheetShapes(settings, gridData) {
  if (!gridData) return {};
  let tabs = [];
  try { tabs = JSON.parse(settings?.tabs || '[]'); } catch { /* ignore */ }
  const out = {};
  for (const tab of tabs) {
    const tabData = gridData[tab.id];
    if (!tabData || !Array.isArray(tabData.shapes) || tabData.shapes.length === 0) continue;
    const sheetName = sanitizeSheetName(tab.name);
    out[sheetName] = {
      shapes: tabData.shapes,
      colWidths: tabData.colWidths || {},
      rowHeights: tabData.rowHeights || {},
      cols: tabData.cols || 26,
      rows: tabData.rows || 50,
    };
  }
  return out;
}

function uint8ToBase64(bytes) {
  const CHUNK = 0x8000;
  let binary = '';
  for (let i = 0; i < bytes.length; i += CHUNK) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + CHUNK));
  }
  return btoa(binary);
}

function downloadXlsxBytes(bytes, filename) {
  const blob = new Blob([bytes], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  // Delay revoke so Safari/Firefox have time to start the download.
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/**
 * @param {object[]} tasks
 * @param {object} settings
 * @param {string} [filename='GanttGen_Project.xlsx']
 * @param {object} [gridData]
 */
export function exportExcel(tasks, settings, filename = 'GanttGen_Project.xlsx', gridData) {
  const wb = buildWorkbook(tasks, settings, gridData);
  const rawBytes = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const sheetStyles = collectSheetStyles(settings, gridData);
  let patched = injectCellStyles(rawBytes, sheetStyles);
  const sheetShapes = collectSheetShapes(settings, gridData);
  patched = injectShapes(patched, sheetShapes);
  downloadXlsxBytes(patched, filename);
}

/**
 * Build an xlsx workbook from tasks + settings and return it as a base64 string.
 * Used by the Share feature to embed project data into a standalone HTML file.
 * @param {object[]} tasks
 * @param {object} settings
 * @param {object} [gridData]
 * @returns {string} base64-encoded xlsx data
 */
export function exportExcelToBase64(tasks, settings, gridData) {
  const wb = buildWorkbook(tasks, settings, gridData);
  const rawBytes = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const sheetStyles = collectSheetStyles(settings, gridData);
  let patched = injectCellStyles(rawBytes, sheetStyles);
  const sheetShapes = collectSheetShapes(settings, gridData);
  patched = injectShapes(patched, sheetShapes);
  return uint8ToBase64(patched);
}


/**
 * Build a settings object mapping CSS variable suffixes to hex values from
 * the flat settings map. Used by App.jsx to apply theme on import.
 */
export function extractThemeColors(settings) {
  return {
    'bg-primary': settings.themeBgPrimary,
    'bg-secondary': settings.themeBgSecondary,
    'bg-tertiary': settings.themeBgTertiary,
    'bg-hover': settings.themeBgHover,
    'border': settings.themeBorder,
    'border-subtle': settings.themeBorderSubtle,
    'text-primary': settings.themeTextPrimary,
    'text-secondary': settings.themeTextSecondary,
    'text-muted': settings.themeTextMuted,
    'accent': settings.themeAccent,
    'accent-hover': settings.themeAccentHover,
    'success': settings.themeSuccess,
    'warning': settings.themeWarning,
    'danger': settings.themeDanger,
    'info': settings.themeInfo,
    'critical-path': settings.themeCriticalPath,
  };
}

/**
 * Write a theme color map back into the flat settings object.
 */
export function writeThemeColors(settings, colors) {
  return {
    ...settings,
    themeBgPrimary: colors['bg-primary'],
    themeBgSecondary: colors['bg-secondary'],
    themeBgTertiary: colors['bg-tertiary'],
    themeBgHover: colors['bg-hover'],
    themeBorder: colors['border'],
    themeBorderSubtle: colors['border-subtle'],
    themeTextPrimary: colors['text-primary'],
    themeTextSecondary: colors['text-secondary'],
    themeTextMuted: colors['text-muted'],
    themeAccent: colors['accent'],
    themeAccentHover: colors['accent-hover'],
    themeSuccess: colors['success'],
    themeWarning: colors['warning'],
    themeDanger: colors['danger'],
    themeInfo: colors['info'],
    themeCriticalPath: colors['critical-path'],
  };
}

function parseTasksSheet(wb, customCols = []) {
  const sheetName = wb.SheetNames.find(
    (n) => n.toLowerCase() === 'tasks',
  );
  if (!sheetName) return [];

  const sheet = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

  return rows.map((row) => {
    const task = {
      id: String(row['Task ID'] ?? ''),
      name: row['Task Name'] ?? '',
      dependency: String(row['Dependency'] ?? ''),
      category: row['Task Category'] ?? '',
      startDate: parseExcelDate(row['Start Date']),
      endDate: parseExcelDate(row['End Date']),
      duration: row['Duration'] !== '' ? Number(row['Duration']) : 0,
      progress: row['Progress (%)'] !== '' ? Number(row['Progress (%)']) : 0,
      status: row['Status'] ?? '',
      owner: row['Owner'] ?? '',
      remarks: row['Remarks'] ?? '',
      baselineStart: parseExcelDate(row['Baseline Start']),
      baselineEnd: parseExcelDate(row['Baseline End']),
      parentId: String(row['Parent ID'] ?? ''),
    };
    for (const cc of customCols) {
      task[cc.key] = row[cc.label] ?? '';
    }
    return task;
  });
}

function parseSettingsSheet(wb) {
  const sheetName = wb.SheetNames.find(
    (n) => n.toLowerCase() === 'settings',
  );
  if (!sheetName) return buildDefaultSettings();

  const sheet = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

  const labelToKey = new Map();
  for (const field of SETTINGS_FIELDS) {
    labelToKey.set(field.label, field.key);
  }

  const settings = buildDefaultSettings();
  for (const row of rows) {
    const label = row['Setting'];
    const value = row['Value'];
    const key = labelToKey.get(label);
    if (key) {
      settings[key] = value;
    }
  }
  return settings;
}

function buildDefaultSettings() {
  const settings = {};
  for (const field of SETTINGS_FIELDS) {
    settings[field.key] = field.defaultValue;
  }
  return settings;
}

function localIso(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function parseExcelDate(value) {
  if (!value) return null;
  if (value instanceof Date) return localIso(value);
  const str = String(value).trim();
  if (!str) return null;
  const parts = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (parts) return `${parts[1]}-${parts[2]}-${parts[3]}`;
  const d = new Date(str);
  return isNaN(d.getTime()) ? null : localIso(d);
}

function formatDate(value) {
  if (!value) return '';
  return String(value);
}

function colLabelFromIndex(index) {
  let label = '';
  let n = index;
  while (n >= 0) {
    label = String.fromCharCode(65 + (n % 26)) + label;
    n = Math.floor(n / 26) - 1;
  }
  return label;
}

function colIndexFromLabel(label) {
  let idx = 0;
  for (let i = 0; i < label.length; i++) {
    idx = idx * 26 + (label.charCodeAt(i) - 64);
  }
  return idx - 1;
}

function parseCellRef(ref) {
  const match = ref.match(/^([A-Z]+)(\d+)$/);
  if (!match) return null;
  return { col: colIndexFromLabel(match[1]), row: parseInt(match[2], 10) - 1 };
}

function sanitizeSheetName(name) {
  let safe = String(name).replace(/[\\/*?[\]:]/g, '_').slice(0, 31);
  if (!safe) safe = 'Sheet';
  return safe;
}

function gridDataToSheet(tabData) {
  const cells = tabData.cells || {};
  const rowCount = tabData.rows || 50;
  const colCount = tabData.cols || 26;
  const tabColWidths = tabData.colWidths || {};
  const tabRowHeights = tabData.rowHeights || {};
  const tabMerges = Array.isArray(tabData.merges) ? tabData.merges : [];

  let maxRow = 0;
  let maxCol = 0;
  for (const key of Object.keys(cells)) {
    const ref = parseCellRef(key);
    if (!ref) continue;
    if (ref.row > maxRow) maxRow = ref.row;
    if (ref.col > maxCol) maxCol = ref.col;
  }

  // Extend range to cover any columns/rows that have custom widths/heights.
  // SheetJS drops !cols / !rows entries outside the sheet's !ref, so the
  // width/height of a resized but otherwise empty column or row would
  // otherwise be silently discarded at write time.
  for (const key of Object.keys(tabColWidths)) {
    const idx = colIndexFromLabel(key);
    if (Number.isFinite(idx) && idx > maxCol) maxCol = idx;
  }
  for (const key of Object.keys(tabRowHeights)) {
    const idx = parseInt(key, 10);
    if (Number.isFinite(idx) && idx > maxRow) maxRow = idx;
  }

  // Merges extending past the last populated cell must also be inside
  // !ref or SheetJS will drop them at write time.
  for (const m of tabMerges) {
    if (Number.isFinite(m?.r2) && m.r2 > maxRow) maxRow = m.r2;
    if (Number.isFinite(m?.c2) && m.c2 > maxCol) maxCol = m.c2;
  }

  const rows = Math.min(Math.max(maxRow + 1, 1), rowCount);
  const cols = Math.min(Math.max(maxCol + 1, 1), colCount);
  const aoa = [];

  for (let r = 0; r < rows; r++) {
    const row = [];
    for (let c = 0; c < cols; c++) {
      const key = `${colLabelFromIndex(c)}${r + 1}`;
      const cell = cells[key];
      row.push(cell ? (cell.v ?? '') : '');
    }
    aoa.push(row);
  }

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // Force !ref to span the full range we intend to persist. aoa_to_sheet
  // tightens !ref to the region containing non-empty values, which can
  // strand !cols / !rows metadata at the trailing edge.
  ws['!ref'] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: rows - 1, c: cols - 1 },
  });

  // Overwrite cells that have formulas so xlsx writes the `f` field.
  for (const key of Object.keys(cells)) {
    const cell = cells[key];
    if (cell && cell.f) {
      const ref = parseCellRef(key);
      if (!ref) continue;
      const wsKey = XLSX.utils.encode_cell({ r: ref.row, c: ref.col });
      if (ws[wsKey]) {
        ws[wsKey].f = cell.f.startsWith('=') ? cell.f.slice(1) : cell.f;
      } else {
        ws[wsKey] = { t: 'n', v: cell.v ?? 0, f: cell.f.startsWith('=') ? cell.f.slice(1) : cell.f };
      }
    }
  }

  // Merged ranges are a first-class concept in OOXML (<mergeCells>) and
  // SheetJS Community reads / writes them natively via ws['!merges'].
  // Skip degenerate single-cell entries; Excel rejects those.
  if (tabMerges.length > 0) {
    const normalized = tabMerges
      .filter((m) => m && (m.r2 > m.r1 || m.c2 > m.c1))
      .map((m) => ({ s: { r: m.r1, c: m.c1 }, e: { r: m.r2, c: m.c2 } }));
    if (normalized.length > 0) ws['!merges'] = normalized;
  }

  return ws;
}

function parseGridSheets(
  wb,
  tabs,
  gridCellStyles = {},
  nativeSheetStyles = {},
  nativeAuthoritative = false,
  gridShapes = {},
  nativeSheetShapes = {},
  nativeShapesAuthoritative = false,
) {
  const gridData = {};
  if (!tabs || tabs.length === 0) return gridData;

  const tabByName = new Map();
  for (const tab of tabs) {
    tabByName.set(tab.name, tab);
  }

  for (const sheetName of wb.SheetNames) {
    if (RESERVED_SHEETS.has(sheetName.toLowerCase())) continue;
    const tab = tabByName.get(sheetName);
    if (!tab) continue;

    const sheet = wb.Sheets[sheetName];
    const ref = sheet['!ref'];
    if (!ref) continue;

    const range = XLSX.utils.decode_range(ref);
    const cells = {};
    let maxRow = 0;
    let maxCol = 0;

    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const wsKey = XLSX.utils.encode_cell({ r, c });
        const wsCell = sheet[wsKey];
        if (!wsCell) continue;

        const key = `${colLabelFromIndex(c)}${r + 1}`;
        const cellObj = { v: wsCell.v ?? '' };

        if (wsCell.f) {
          cellObj.f = '=' + wsCell.f;
        }

        if (cellObj.v !== '' || cellObj.f) {
          cells[key] = cellObj;
          if (r > maxRow) maxRow = r;
          if (c > maxCol) maxCol = c;
        }
      }
    }

    // Merge per-cell styles from two sources:
    //   1. gridCellStyles JSON blob written to the Settings sheet by
    //      earlier app versions. Authoritative for keys NOT covered by
    //      the native injector (color, bg, fontSize, borders).
    //   2. Native xlsx styles extracted directly from the XML by
    //      XlsxStyleExtractor. Authoritative for NATIVE_STYLE_KEYS when
    //      the file was produced by an injector-aware build
    //      (nativeAuthoritative === true).
    //
    // Cells whose ONLY effect is a style (no value, no formula) are
    // materialised here with v: '' so the grid still renders the
    // formatting when the user scrolls to them.
    const tabStyles = gridCellStyles[tab.id] || {};
    const sheetNative = nativeSheetStyles[sheetName] || {};
    const styledRefs = new Set([
      ...Object.keys(tabStyles),
      ...Object.keys(sheetNative),
    ]);
    for (const cellRef of styledRefs) {
      const json = tabStyles[cellRef] || {};
      const native = sheetNative[cellRef] || {};
      const merged = { ...json };
      if (nativeAuthoritative) {
        // Excel XML is the source of truth for the keys it covers. Drop
        // any stale values the JSON blob held on to so that styles the
        // user removed in Excel do not re-materialise on import.
        for (const k of NATIVE_STYLE_KEYS) delete merged[k];
      }
      Object.assign(merged, native);
      if (Object.keys(merged).length === 0) continue;
      if (!cells[cellRef]) cells[cellRef] = { v: '' };
      cells[cellRef].s = merged;
    }

    const importedColWidths = {};
    if (sheet['!cols']) {
      sheet['!cols'].forEach((col, i) => {
        if (!col) return;
        // Prefer pixel width (wpx) for lossless round-trip; fall back to
        // character width (wch) for files produced by other tools.
        const px = col.wpx ?? (col.wch != null ? Math.round(col.wch * 7) : null);
        if (px) importedColWidths[colLabelFromIndex(i)] = px;
      });
    }

    const importedRowHeights = {};
    if (sheet['!rows']) {
      sheet['!rows'].forEach((row, i) => {
        if (!row) return;
        const px = row.hpx ?? row.hpt;
        if (px) importedRowHeights[i] = px;
      });
    }

    const importedMerges = [];
    if (Array.isArray(sheet['!merges'])) {
      for (const m of sheet['!merges']) {
        const r1 = m?.s?.r, c1 = m?.s?.c, r2 = m?.e?.r, c2 = m?.e?.c;
        if (!Number.isFinite(r1) || !Number.isFinite(c1) || !Number.isFinite(r2) || !Number.isFinite(c2)) continue;
        if (r1 === r2 && c1 === c2) continue;
        importedMerges.push({
          r1: Math.min(r1, r2),
          c1: Math.min(c1, c2),
          r2: Math.max(r1, r2),
          c2: Math.max(c1, c2),
        });
        if (Math.max(r1, r2) > maxRow) maxRow = Math.max(r1, r2);
        if (Math.max(c1, c2) > maxCol) maxCol = Math.max(c1, c2);
      }
    }

    // Shapes: prefer native DrawingML when the file is tagged authoritative
    // (v >= 1). Otherwise fall back to the JSON blob in the Settings sheet.
    let shapesForTab = [];
    if (nativeShapesAuthoritative && nativeSheetShapes[sheetName]) {
      shapesForTab = nativeSheetShapes[sheetName];
    } else if (Array.isArray(gridShapes[tab.id])) {
      shapesForTab = gridShapes[tab.id];
    }

    gridData[tab.id] = {
      rows: Math.max(maxRow + 10, 50),
      cols: Math.max(maxCol + 5, 26),
      cells,
      colWidths: importedColWidths,
      rowHeights: importedRowHeights,
      merges: importedMerges,
      shapes: shapesForTab,
      showGridLines: true,
    };
  }

  return gridData;
}
