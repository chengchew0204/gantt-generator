import { useState, useMemo, useCallback, useRef, useEffect } from 'react';
import { toPng } from 'html-to-image';
import Dashboard from './components/Dashboard';
import DataTable from './components/DataTable';
import GanttChart from './components/GanttChart';
import ThemePanel from './components/ThemePanel';
import GuideOverlay from './components/GuideOverlay';
import {
  downloadTemplate, importExcel, exportExcel,
  extractThemeColors, writeThemeColors,
} from './utils/ExcelUtils';
import { computeCpm } from './utils/CpmEngine';
import { addWorkingDays, workingDaysBetween } from './utils/DateUtils';
import useUndoRedo from './hooks/useUndoRedo';

const MIN_LEFT_WIDTH = 280;
const MIN_RIGHT_WIDTH = 300;

/**
 * Find the scrollable body element within a flex-column pane root by inspecting
 * computed CSS overflow (avoids fragile class-name matching across Tailwind versions).
 */
function findScrollEl(rootEl) {
  for (const child of rootEl.children) {
    const s = window.getComputedStyle(child);
    if (s.overflowY === 'auto' || s.overflow === 'auto') return child;
  }
  return null;
}

/**
 * Compute the export height for a single flex-column pane (DataTable or GanttChart).
 *
 * Strategy:
 *  1. Find the scrollable body child via computed overflow style.
 *  2. Measure the body's INNER content element height via getBoundingClientRect() —
 *     this reflects the real rendered SVG/rows height, not the stretched flex container.
 *  3. Sum that with all sibling chrome elements (header, toolbar, footer bars).
 */
function measureFlexColumnExportHeight(rootEl) {
  if (!rootEl) return null;
  const children = [...rootEl.children];
  const scrollEl = findScrollEl(rootEl);
  if (!scrollEl) return null;
  const scrollIdx = children.indexOf(scrollEl);

  const contentEl = scrollEl.firstElementChild;
  const contentH = contentEl
    ? Math.round(contentEl.getBoundingClientRect().height)
    : Math.round(scrollEl.getBoundingClientRect().height);

  let h = 0;
  for (let i = 0; i < scrollIdx; i++) h += children[i].offsetHeight;
  h += contentH;
  for (let i = scrollIdx + 1; i < children.length; i++) {
    if (!children[i].hasAttribute('data-export-exclude')) h += children[i].offsetHeight;
  }
  return Math.ceil(h);
}

function measureSplitPaneExportHeight(containerEl) {
  const leftRoot = containerEl.querySelector('[data-guide="data-table"]')?.firstElementChild;
  const rightRoot = containerEl.querySelector('[data-guide="gantt-chart"]')?.firstElementChild;
  const leftH = measureFlexColumnExportHeight(leftRoot);
  const rightH = measureFlexColumnExportHeight(rightRoot);
  if (leftH == null || rightH == null) return null;
  return Math.max(leftH, rightH);
}

const DEFAULT_CATEGORY_PALETTE = [
  '#6366f1', '#0ea5e9', '#10b981', '#f59e0b',
  '#ec4899', '#8b5cf6', '#14b8a6', '#f97316',
];

function assignCategoryColors(tasks, existing) {
  const colors = { ...existing };
  let idx = Object.keys(colors).length;
  for (const t of tasks) {
    const cat = (t.category || '').trim();
    if (cat && !colors[cat]) {
      colors[cat] = DEFAULT_CATEGORY_PALETTE[idx % DEFAULT_CATEGORY_PALETTE.length];
      idx++;
    }
  }
  return colors;
}

const ALL_COLUMNS = [
  { key: 'id', label: 'ID', type: 'text' },
  { key: 'name', label: 'Task Name', type: 'text' },
  { key: 'dependency', label: 'Dependency', type: 'text' },
  { key: 'category', label: 'Category', type: 'text' },
  { key: 'startDate', label: 'Start Date', type: 'date' },
  { key: 'endDate', label: 'End Date', type: 'date' },
  { key: 'duration', label: 'Duration', type: 'number' },
  { key: 'progress', label: 'Progress (%)', type: 'number' },
  { key: 'status', label: 'Status', type: 'status' },
  { key: 'owner', label: 'Owner', type: 'text' },
  { key: 'remarks', label: 'Remarks', type: 'text' },
  { key: 'baselineStart', label: 'Baseline Start', type: 'date' },
  { key: 'baselineEnd', label: 'Baseline End', type: 'date' },
  { key: 'parentId', label: 'Parent ID', type: 'text' },
];

const DEFAULT_VISIBLE = new Set([
  'id', 'name', 'duration', 'startDate', 'endDate', 'progress', 'status',
]);

const DEFAULT_VIEW_OPTIONS = {
  showCriticalPath: false,
  showSlack: false,
  showDependencies: true,
  showTodayLine: true,
  showBaseline: true,
  skipWeekends: true,
  showScaleButtons: true,
  showWeekLabels: false,
  showMonthLabels: true,
  showDayLabels: true,
};

function toBool(v) {
  if (typeof v === 'boolean') return v;
  return v !== 'false' && v !== false;
}

function nextId(tasks) {
  let max = 0;
  for (const t of tasks) {
    const n = parseInt(t.id, 10);
    if (n > max) max = n;
  }
  return String(max + 1);
}

function buildWbsTree(tasks) {
  const parentIds = new Set();
  for (const t of tasks) {
    if (t.parentId) parentIds.add(String(t.parentId));
  }

  return tasks.map((task) => {
    const id = String(task.id);
    if (!parentIds.has(id)) return task;

    const children = tasks.filter((c) => String(c.parentId) === id);
    let minStart = null;
    let maxEnd = null;
    let totalProgress = 0;
    let childCount = 0;

    for (const child of children) {
      if (child.startDate && (!minStart || child.startDate < minStart)) minStart = child.startDate;
      if (child.endDate && (!maxEnd || child.endDate > maxEnd)) maxEnd = child.endDate;
      totalProgress += child.progress || 0;
      childCount++;
    }

    return {
      ...task,
      isParent: true,
      startDate: minStart || task.startDate,
      endDate: maxEnd || task.endDate,
      progress: childCount > 0 ? Math.round(totalProgress / childCount) : task.progress,
    };
  });
}

function applyThemeToDOM(colors) {
  const root = document.documentElement;
  for (const [key, value] of Object.entries(colors)) {
    if (value) root.style.setProperty(`--color-${key}`, value);
  }
  const accent = colors['accent'];
  if (accent) {
    const r = parseInt(accent.slice(1, 3), 16);
    const g = parseInt(accent.slice(3, 5), 16);
    const b = parseInt(accent.slice(5, 7), 16);
    root.style.setProperty('--color-accent-muted', `rgba(${r}, ${g}, ${b}, 0.12)`);
  }
}

export default function App() {
  const [tasks, setTasks] = useState([]);
  const [settings, setSettings] = useState({});
  const [splitRatio, setSplitRatio] = useState(0.38);
  const [visibleColumns, setVisibleColumns] = useState(DEFAULT_VISIBLE);
  const [collapsedParents, setCollapsedParents] = useState(new Set());
  const [activeTheme, setActiveTheme] = useState('Notion Light');
  const [themePanelOpen, setThemePanelOpen] = useState(false);
  const [guideOpen, setGuideOpen] = useState(false);
  const [viewOptionsOpen, setViewOptionsOpen] = useState(false);
  const [viewOptions, setViewOptions] = useState(DEFAULT_VIEW_OPTIONS);
  const [syncScrollTop, setSyncScrollTop] = useState(0);
  const [selectedTaskId, setSelectedTaskId] = useState(null);
  const [categoryColors, setCategoryColors] = useState({});
  const [projectName, setProjectName] = useState('');
  const [datePickField, setDatePickField] = useState(null);
  const containerRef = useRef(null);
  const chartRef = useRef(null);
  const viewBtnRef = useRef(null);
  const isDragging = useRef(false);
  const handleExportRef = useRef(null);
  const tasksRef = useRef(tasks);
  tasksRef.current = tasks;

  const { pushState, undo, redo, canUndo, canRedo, beginBatch, endBatch } = useUndoRedo();

  const recordAndSetTasks = useCallback((updater) => {
    pushState(tasksRef.current);
    setTasks(updater);
  }, [pushState]);

  const handleUndo = useCallback(() => {
    const prev = undo(tasksRef.current);
    if (prev) setTasks(prev);
  }, [undo]);

  const handleRedo = useCallback(() => {
    const next = redo(tasksRef.current);
    if (next) setTasks(next);
  }, [redo]);

  const handleBeginDrag = useCallback(() => {
    beginBatch(tasksRef.current);
  }, [beginBatch]);

  const handleEndDrag = useCallback(() => {
    endBatch();
  }, [endBatch]);

  const skipWeekends = toBool(viewOptions.skipWeekends);
  const prevSkipWeekends = useRef(skipWeekends);
  const showScaleButtons = toBool(viewOptions.showScaleButtons ?? true);
  const showMonthLabels = toBool(viewOptions.showMonthLabels ?? true);
  const showWeekLabels = toBool(viewOptions.showWeekLabels ?? false);
  const showDayLabels = toBool(viewOptions.showDayLabels ?? true);
  const effectiveHeaderHeight =
    (showScaleButtons ? 28 : 0) +
    (showMonthLabels ? 16 : 0) +
    (showWeekLabels ? 18 : 0) +
    (showDayLabels ? 18 : 0);

  const wbsTasks = useMemo(() => buildWbsTree(tasks), [tasks]);

  const enrichedTasks = useMemo(() => {
    if (wbsTasks.length === 0) return [];
    const cpmResults = computeCpm(wbsTasks);
    const cpmMap = new Map(cpmResults.map((r) => [r.id, r]));

    return wbsTasks.map((task) => {
      const cpm = cpmMap.get(String(task.id));
      return {
        ...task,
        earlyStart: cpm?.earlyStart ?? 0,
        earlyFinish: cpm?.earlyFinish ?? 0,
        lateStart: cpm?.lateStart ?? 0,
        lateFinish: cpm?.lateFinish ?? 0,
        totalFloat: cpm?.totalFloat ?? 0,
        isCritical: cpm?.isCritical ?? false,
      };
    });
  }, [wbsTasks]);

  const displayTasks = useMemo(() => {
    if (collapsedParents.size === 0) return enrichedTasks;
    return enrichedTasks.filter((t) => {
      if (!t.parentId) return true;
      return !collapsedParents.has(String(t.parentId));
    });
  }, [enrichedTasks, collapsedParents]);

  const handleUpdateTask = useCallback((taskId, field, value) => {
    recordAndSetTasks((prev) =>
      prev.map((t) => {
        if (String(t.id) !== String(taskId)) return t;
        const updated = { ...t, [field]: value };

        const dur = Number(updated.duration);
        if (field === 'duration' && updated.startDate && Number.isFinite(Number(value)) && Number(value) > 0) {
          updated.endDate = addWorkingDays(updated.startDate, Number(value), skipWeekends);
        } else if (field === 'startDate' && value && Number.isFinite(dur) && dur > 0) {
          updated.endDate = addWorkingDays(value, dur, skipWeekends);
        } else if (field === 'endDate' && value && updated.startDate) {
          updated.duration = workingDaysBetween(updated.startDate, value, skipWeekends);
        }

        return updated;
      }),
    );
  }, [skipWeekends, recordAndSetTasks]);

  // Batch-update multiple fields at once with no auto-calc side effects.
  // Used by Gantt bar drags where dates are computed directly from pixel offsets.
  const handleUpdateTaskFields = useCallback((taskId, fields) => {
    recordAndSetTasks((prev) =>
      prev.map((t) => {
        if (String(t.id) !== String(taskId)) return t;
        const updated = { ...t, ...fields };
        if (updated.startDate && updated.endDate) {
          updated.duration = workingDaysBetween(updated.startDate, updated.endDate, skipWeekends);
        }
        return updated;
      }),
    );
  }, [skipWeekends, recordAndSetTasks]);

  const handleAddTask = useCallback(() => {
    recordAndSetTasks((prev) => [
      ...prev,
      {
        id: nextId(prev),
        name: '',
        dependency: '',
        category: '',
        startDate: null,
        endDate: null,
        duration: 0,
        progress: 0,
        status: 'Not Started',
        owner: '',
        remarks: '',
        baselineStart: null,
        baselineEnd: null,
        parentId: '',
      },
    ]);
  }, [recordAndSetTasks]);

  const handleDeleteTask = useCallback((taskId) => {
    recordAndSetTasks((prev) => prev.filter((t) => String(t.id) !== String(taskId)));
  }, [recordAndSetTasks]);

  const handleReorderTask = useCallback((fromIndex, toIndex) => {
    recordAndSetTasks((prev) => {
      const next = [...prev];
      const [moved] = next.splice(fromIndex, 1);
      next.splice(toIndex, 0, moved);
      return next;
    });
  }, [recordAndSetTasks]);

  const handleToggleCollapse = useCallback((parentId) => {
    setCollapsedParents((prev) => {
      const next = new Set(prev);
      if (next.has(parentId)) next.delete(parentId);
      else next.add(parentId);
      return next;
    });
  }, []);

  const handleApplyPreset = useCallback((name, colors) => {
    setActiveTheme(name);
    applyThemeToDOM(colors);
    setSettings((prev) => writeThemeColors({ ...prev, themeName: name }, colors));
  }, []);

  const handleApplyCustomColor = useCallback((key, hex) => {
    setActiveTheme('Custom');
    document.documentElement.style.setProperty(`--color-${key}`, hex);
    if (key === 'accent') {
      const r = parseInt(hex.slice(1, 3), 16);
      const g = parseInt(hex.slice(3, 5), 16);
      const b = parseInt(hex.slice(5, 7), 16);
      document.documentElement.style.setProperty('--color-accent-muted', `rgba(${r}, ${g}, ${b}, 0.12)`);
    }
    setSettings((prev) => {
      const current = extractThemeColors(prev);
      current[key] = hex;
      return writeThemeColors({ ...prev, themeName: 'Custom' }, current);
    });
  }, []);

  const handleChangeCategoryColor = useCallback((category, hex) => {
    setCategoryColors((prev) => ({ ...prev, [category]: hex }));
  }, []);

  const handleOpenTheme = useCallback(() => {
    setThemePanelOpen(true);
  }, []);

  const handleOpenGuide = useCallback(() => {
    setGuideOpen(true);
  }, []);

  const handleToggleViewOptions = useCallback(() => {
    setViewOptionsOpen((v) => !v);
  }, []);

  const handleToggleViewOption = useCallback((key) => {
    setViewOptions((prev) => {
      const next = { ...prev, [key]: !toBool(prev[key]) };
      return next;
    });
  }, []);

  useEffect(() => {
    setCategoryColors((prev) => assignCategoryColors(tasks, prev));
  }, [tasks]);

  useEffect(() => {
    if (prevSkipWeekends.current === skipWeekends) return;
    prevSkipWeekends.current = skipWeekends;
    setTasks((prev) =>
      prev.map((t) => {
        const dur = Number(t.duration);
        if (t.startDate && Number.isFinite(dur) && dur > 0) {
          return { ...t, endDate: addWorkingDays(t.startDate, dur, skipWeekends) };
        }
        return t;
      }),
    );
  }, [skipWeekends]);

  useEffect(() => {
    setSettings((prev) => {
      const next = { ...prev };
      for (const [k, v] of Object.entries(viewOptions)) {
        next[k] = String(v);
      }
      next.visibleColumns = [...visibleColumns].join(',');
      next.categoryColors = JSON.stringify(categoryColors);
      next.projectName = projectName;
      return next;
    });
  }, [viewOptions, visibleColumns, categoryColors, projectName]);

  useEffect(() => {
    const handler = (e) => {
      const mod = e.metaKey || e.ctrlKey;
      if (!mod) return;

      if (e.key === 's' || e.key === 'S') {
        e.preventDefault();
        handleExportRef.current();
        return;
      }

      if (e.key === 'z' && !e.shiftKey) {
        e.preventDefault();
        handleUndo();
        return;
      }

      if ((e.key === 'z' && e.shiftKey) || (e.key === 'Z' && e.shiftKey) || e.key === 'y') {
        e.preventDefault();
        handleRedo();
      }
    };
    document.addEventListener('keydown', handler);
    return () => document.removeEventListener('keydown', handler);
  }, [handleUndo, handleRedo]);

  const handleImport = useCallback(async (file) => {
    try {
      const result = await importExcel(file);
      setTasks(result.tasks);
      setSettings(result.settings);
      setCollapsedParents(new Set());

      const s = result.settings;
      setProjectName(s.projectName || '');
      const colors = extractThemeColors(s);
      applyThemeToDOM(colors);
      setActiveTheme(s.themeName || 'Notion Light');

      setViewOptions({
        showCriticalPath: toBool(s.showCriticalPath ?? true),
        showSlack: toBool(s.showSlack ?? true),
        showDependencies: toBool(s.showDependencies ?? true),
        showTodayLine: toBool(s.showTodayLine ?? true),
        showBaseline: toBool(s.showBaseline ?? true),
        skipWeekends: toBool(s.skipWeekends ?? true),
        showScaleButtons: toBool(s.showScaleButtons ?? true),
        showWeekLabels: toBool(s.showWeekLabels ?? false),
        showMonthLabels: toBool(s.showMonthLabels ?? true),
        showDayLabels: toBool(s.showDayLabels ?? true),
      });

      if (s.visibleColumns) {
        const cols = s.visibleColumns.split(',').map((c) => c.trim()).filter(Boolean);
        if (cols.length > 0) setVisibleColumns(new Set(cols));
      }

      let restored = {};
      if (s.categoryColors) {
        try { restored = JSON.parse(s.categoryColors); } catch (_) { /* ignore */ }
      }
      setCategoryColors(assignCategoryColors(result.tasks, restored));
    } catch (err) {
      console.error('Import failed:', err);
    }
  }, []);

  const handleExport = useCallback(() => {
    const filename = 'GanttGen-' + (projectName.trim() || 'Project').replace(/[^\w\s-]/g, '') + '.xlsx';
    exportExcel(enrichedTasks, settings, filename);
  }, [enrichedTasks, settings, projectName]);
  handleExportRef.current = handleExport;

  const handleExportPng = useCallback(async (mode) => {
    const target = mode === 'full' ? containerRef.current : chartRef.current;
    if (!target) return;

    const scrollbarStyle = document.createElement('style');
    scrollbarStyle.textContent = '::-webkit-scrollbar{display:none!important}*{scrollbar-width:none!important}';
    document.head.appendChild(scrollbarStyle);

    try {
      const bgColor = getComputedStyle(document.documentElement)
        .getPropertyValue('--color-bg-primary').trim() || '#0f0f12';

      let exportHeight = null;
      if (mode === 'full') {
        exportHeight = measureSplitPaneExportHeight(target);
      } else {
        exportHeight = measureFlexColumnExportHeight(target.firstElementChild);
      }

      const opts = { backgroundColor: bgColor };
      if (exportHeight != null && exportHeight > 0 && exportHeight < target.offsetHeight) {
        opts.height = exportHeight;
      }

      const dataUrl = await toPng(target, opts);
      const link = document.createElement('a');
      link.download = mode === 'full' ? 'GanttGen_Full.png' : 'GanttGen_Chart.png';
      link.href = dataUrl;
      link.click();
    } catch (err) {
      console.error('PNG export failed:', err);
    } finally {
      document.head.removeChild(scrollbarStyle);
    }
  }, []);

  const handleMouseDown = useCallback((e) => {
    e.preventDefault();
    isDragging.current = true;

    const onMouseMove = (e) => {
      if (!isDragging.current || !containerRef.current) return;
      const rect = containerRef.current.getBoundingClientRect();
      const totalWidth = rect.width;
      const x = e.clientX - rect.left;
      const ratio = x / totalWidth;
      const minLeftRatio = MIN_LEFT_WIDTH / totalWidth;
      const maxLeftRatio = 1 - MIN_RIGHT_WIDTH / totalWidth;
      setSplitRatio(Math.min(maxLeftRatio, Math.max(minLeftRatio, ratio)));
    };

    const onMouseUp = () => {
      isDragging.current = false;
      document.removeEventListener('mousemove', onMouseMove);
      document.removeEventListener('mouseup', onMouseUp);
    };

    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp);
  }, []);

  return (
    <div className="flex flex-col h-screen" style={{ backgroundColor: 'var(--color-bg-primary)' }}>
      <Dashboard
        tasks={enrichedTasks}
        projectName={projectName}
        onChangeProjectName={setProjectName}
        onImport={handleImport}
        onDownloadTemplate={downloadTemplate}
        onExport={handleExport}
        onExportPng={handleExportPng}
        onOpenGuide={handleOpenGuide}
        onOpenTheme={handleOpenTheme}
        onUndo={handleUndo}
        onRedo={handleRedo}
        canUndo={canUndo}
        canRedo={canRedo}
        viewOptionsOpen={viewOptionsOpen}
        onToggleViewOptions={handleToggleViewOptions}
        viewBtnRef={viewBtnRef}
        viewOptions={viewOptions}
        onToggleViewOption={handleToggleViewOption}
        columns={ALL_COLUMNS}
        visibleColumns={visibleColumns}
        onToggleColumn={(key) => {
          setVisibleColumns((prev) => {
            const next = new Set(prev);
            if (next.has(key)) next.delete(key);
            else next.add(key);
            return next;
          });
        }}
      />

      <div ref={containerRef} className="flex flex-1 min-h-0">
        <div
          className="overflow-hidden"
          data-guide="data-table"
          style={{
            width: `${splitRatio * 100}%`,
            borderRight: '1px solid var(--color-border)',
          }}
        >
          <DataTable
            tasks={displayTasks}
            allTasks={enrichedTasks}
            columns={ALL_COLUMNS}
            visibleColumns={visibleColumns}
            onToggleColumn={(key) => {
              setVisibleColumns((prev) => {
                const next = new Set(prev);
                if (next.has(key)) next.delete(key);
                else next.add(key);
                return next;
              });
            }}
            collapsedParents={collapsedParents}
            onToggleCollapse={handleToggleCollapse}
            onUpdateTask={handleUpdateTask}
            onAddTask={handleAddTask}
            onDeleteTask={handleDeleteTask}
            onReorderTask={handleReorderTask}
            scrollTop={syncScrollTop}
            onScroll={setSyncScrollTop}
            selectedTaskId={selectedTaskId}
            onSelectTask={setSelectedTaskId}
            headerHeight={effectiveHeaderHeight}
            datePickField={datePickField}
            onDatePickField={setDatePickField}
            onBeginDrag={handleBeginDrag}
            onEndDrag={handleEndDrag}
          />
        </div>

        <div
          onMouseDown={handleMouseDown}
          className="w-1 flex-shrink-0 cursor-col-resize transition-colors"
          style={{ backgroundColor: 'var(--color-border)' }}
          onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = 'var(--color-accent)')}
          onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'var(--color-border)')}
        />

        <div ref={chartRef} className="flex-1 min-w-0 overflow-hidden" data-guide="gantt-chart">
          <GanttChart
            tasks={displayTasks}
            allTasks={enrichedTasks}
            viewOptions={viewOptions}
            scrollTop={syncScrollTop}
            onScroll={setSyncScrollTop}
            onUpdateTask={handleUpdateTask}
            onUpdateTaskFields={handleUpdateTaskFields}
            selectedTaskId={selectedTaskId}
            onSelectTask={setSelectedTaskId}
            categoryColors={categoryColors}
            datePickField={datePickField}
            onDatePickField={setDatePickField}
            onBeginDrag={handleBeginDrag}
            onEndDrag={handleEndDrag}
          />
        </div>
      </div>

      <ThemePanel
        open={themePanelOpen}
        onClose={() => setThemePanelOpen(false)}
        activeTheme={activeTheme}
        onApplyPreset={handleApplyPreset}
        onApplyCustomColor={handleApplyCustomColor}
        categoryColors={categoryColors}
        onChangeCategoryColor={handleChangeCategoryColor}
      />

      <GuideOverlay
        open={guideOpen}
        onClose={() => setGuideOpen(false)}
      />
    </div>
  );
}
