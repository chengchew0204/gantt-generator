import * as XLSX from 'xlsx';

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
  { key: 'visibleColumns', label: 'Visible Columns', defaultValue: 'id,name,duration,startDate,endDate,progress,status,owner,category' },
  { key: 'categoryColors', label: 'Category Colors', defaultValue: '{}' },
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

  XLSX.writeFile(wb, 'GanttGen_Template.xlsx');
}

/**
 * @param {File} file
 * @returns {Promise<{ tasks: object[], settings: object }>}
 */
export async function importExcel(file) {
  const data = await file.arrayBuffer();
  const wb = XLSX.read(data, { type: 'array', cellDates: true });

  const tasks = parseTasksSheet(wb);
  const settings = parseSettingsSheet(wb);

  return { tasks, settings };
}

/**
 * Build an xlsx workbook from tasks + settings, returned as a Uint8Array.
 */
function buildWorkbook(tasks, settings) {
  const wb = XLSX.utils.book_new();

  const taskRows = tasks.map((t) => [
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
  ]);
  taskRows.unshift(TASK_COLUMNS);

  const taskSheet = XLSX.utils.aoa_to_sheet(taskRows);
  const colWidths = TASK_COLUMNS.map((col) => ({
    wch: Math.max(col.length + 2, 14),
  }));
  taskSheet['!cols'] = colWidths;
  XLSX.utils.book_append_sheet(wb, taskSheet, 'Tasks');

  const settingsRows = [['Setting', 'Value']];
  for (const field of SETTINGS_FIELDS) {
    const value =
      settings && settings[field.key] != null
        ? String(settings[field.key])
        : field.defaultValue;
    settingsRows.push([field.label, value]);
  }
  const settingsSheet = XLSX.utils.aoa_to_sheet(settingsRows);
  XLSX.utils.book_append_sheet(wb, settingsSheet, 'Settings');

  return wb;
}

/**
 * @param {object[]} tasks
 * @param {object} settings
 * @param {string} [filename='GanttGen_Project.xlsx']
 */
export function exportExcel(tasks, settings, filename = 'GanttGen_Project.xlsx') {
  const wb = buildWorkbook(tasks, settings);
  XLSX.writeFile(wb, filename);
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

function parseTasksSheet(wb) {
  const sheetName = wb.SheetNames.find(
    (n) => n.toLowerCase() === 'tasks',
  );
  if (!sheetName) return [];

  const sheet = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

  return rows.map((row) => ({
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
  }));
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
