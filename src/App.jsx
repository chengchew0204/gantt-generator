import { useState, useMemo, useCallback, useRef, useEffect } from 'react';
import { toPng } from 'html-to-image';
import { FileUp, X, AlertCircle, Share2, Globe, MousePointerClick, RotateCcw } from 'lucide-react';
import Dashboard from './components/Dashboard';
import DataTable from './components/DataTable';
import GanttChart from './components/GanttChart';
import DataGrid from './components/DataGrid';
import ThemePanel from './components/ThemePanel';
import GuideOverlay from './components/GuideOverlay';
import StatusBar from './components/StatusBar';
import {
  downloadTemplate, importExcel, exportExcel, exportExcelToBase64,
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
  'id', 'name', 'duration', 'startDate', 'endDate', 'progress', 'status', 'owner',
]);

const DEFAULT_VIEW_OPTIONS = {
  showCriticalPath: false,
  showSlack: false,
  showDependencies: false,
  showTodayLine: true,
  showBaseline: true,
  showTaskNames: true,
  showProgressPercent: true,
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

function buildWbsTree(tasks, skipWeekends) {
  const childrenMap = new Map();
  for (const t of tasks) {
    if (t.parentId) {
      const pid = String(t.parentId);
      if (!childrenMap.has(pid)) childrenMap.set(pid, []);
      childrenMap.get(pid).push(t);
    }
  }

  const resultMap = new Map();
  for (const t of tasks) {
    resultMap.set(String(t.id), { ...t });
  }

  const resolved = new Set();

  function resolve(taskId) {
    // Guard against orphan parentIds (dangling references). Happens when a
    // user types a non-existent id into the Parent ID column or an imported
    // xlsx has malformed hierarchy data. Matches the defensive early-return
    // pattern used by assignDepth(). Returning a neutral aggregate keeps the
    // caller's start/end/progress folding sane without polluting resultMap.
    if (!resultMap.has(taskId)) {
      return { startDate: null, endDate: null, progress: 0 };
    }
    if (resolved.has(taskId)) {
      const t = resultMap.get(taskId);
      return { startDate: t.startDate, endDate: t.endDate, progress: t.progress || 0 };
    }
    resolved.add(taskId);

    const children = childrenMap.get(taskId);
    if (!children || children.length === 0) {
      const t = resultMap.get(taskId);
      return { startDate: t.startDate, endDate: t.endDate, progress: t.progress || 0 };
    }

    let minStart = null;
    let maxEnd = null;
    let totalProgress = 0;

    for (const child of children) {
      const r = resolve(String(child.id));
      if (r.startDate && (!minStart || r.startDate < minStart)) minStart = r.startDate;
      if (r.endDate && (!maxEnd || r.endDate > maxEnd)) maxEnd = r.endDate;
      totalProgress += r.progress;
    }

    const task = resultMap.get(taskId);
    task.isParent = true;
    task.startDate = minStart || task.startDate;
    task.endDate = maxEnd || task.endDate;
    task.progress = children.length > 0 ? Math.round(totalProgress / children.length) : task.progress;
    task.duration = task.startDate && task.endDate
      ? workingDaysBetween(task.startDate, task.endDate, skipWeekends)
      : 0;

    return { startDate: task.startDate, endDate: task.endDate, progress: task.progress };
  }

  for (const parentId of childrenMap.keys()) {
    resolve(parentId);
  }

  assignDepth(resultMap, childrenMap);

  return tasks.map((t) => resultMap.get(String(t.id)));
}

// Walk the WBS tree from every root downward, stamping task.depth for each
// node. Cycle-safe: a `seen` set breaks the walk if a user sets two tasks'
// parentId to each other. Orphans (parentId pointing at a missing task) are
// treated as roots at depth 0 on the final sweep.
function assignDepth(resultMap, childrenMap) {
  const seen = new Set();
  function walk(id, d) {
    if (seen.has(id)) return;
    seen.add(id);
    const t = resultMap.get(id);
    if (!t) return;
    t.depth = d;
    const kids = childrenMap.get(id) || [];
    for (const k of kids) walk(String(k.id), d + 1);
  }
  const rootIds = [];
  for (const [id, t] of resultMap.entries()) {
    const pid = t.parentId ? String(t.parentId) : '';
    if (!pid || !resultMap.has(pid)) rootIds.push(id);
  }
  for (const id of rootIds) walk(id, 0);
  for (const t of resultMap.values()) {
    if (t.depth == null) t.depth = 0;
  }
}

function parseThemeHex(hex) {
  if (!hex || typeof hex !== 'string') return null;
  const h = hex.trim();
  if (!/^#[0-9a-fA-F]{6}$/.test(h)) return null;
  return {
    r: parseInt(h.slice(1, 3), 16),
    g: parseInt(h.slice(3, 5), 16),
    b: parseInt(h.slice(5, 7), 16),
  };
}

function contrastingTextOnHex(hex) {
  const rgb = parseThemeHex(hex);
  if (!rgb) return '#ffffff';
  const yiq = (rgb.r * 299 + rgb.g * 587 + rgb.b * 114) / 1000;
  return yiq >= 150 ? '#121212' : '#ffffff';
}

function refreshAccentContrastVars(root) {
  const style = getComputedStyle(root);
  const accent = style.getPropertyValue('--color-accent').trim();
  const accentHover = style.getPropertyValue('--color-accent-hover').trim();
  const parsed = parseThemeHex(accent);
  if (parsed) {
    root.style.setProperty(
      '--color-accent-muted',
      `rgba(${parsed.r}, ${parsed.g}, ${parsed.b}, 0.12)`,
    );
    root.style.setProperty('--color-on-accent', contrastingTextOnHex(accent));
  }
  const hoverHex = parseThemeHex(accentHover) ? accentHover : accent;
  root.style.setProperty('--color-on-accent-hover', contrastingTextOnHex(hoverHex));
}

function applyThemeToDOM(colors) {
  const root = document.documentElement;
  for (const [key, value] of Object.entries(colors)) {
    if (value) root.style.setProperty(`--color-${key}`, value);
  }
  refreshAccentContrastVars(root);
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
  const [customColumns, setCustomColumns] = useState([]);
  const [columnOrder, setColumnOrder] = useState(() => ALL_COLUMNS.map((c) => c.key));
  const [colWidths, setColWidths] = useState({});
  const [ganttScale, setGanttScale] = useState('day');
  const [ganttZoom, setGanttZoom] = useState(100);
  const [gridZoom, setGridZoom] = useState(100);
  const [activeTab, setActiveTab] = useState('gantt');
  const [tabs, setTabs] = useState([]);
  const [gridData, setGridData] = useState({});
  const [datePickField, setDatePickField] = useState(null);
  const [lastSavedAt, setLastSavedAt] = useState(null);
  const [isDirty, setIsDirty] = useState(false);
  const isDirtyRef = useRef(false);
  const [dragActive, setDragActive] = useState(false);
  const [importError, setImportError] = useState(null);
  const [shareNotice, setShareNotice] = useState(null);
  const [showFireworks, setShowFireworks] = useState(false);
  const prevAllCompleteRef = useRef(false);
  const dragCounter = useRef(0);
  const containerRef = useRef(null);
  const chartRef = useRef(null);
  const viewBtnRef = useRef(null);
  const isDragging = useRef(false);
  const handleExportRef = useRef(null);
  const tasksRef = useRef(tasks);
  tasksRef.current = tasks;

  const { pushState, undo, redo, canUndo, canRedo, beginBatch, endBatch } = useUndoRedo();

  const {
    pushState: pushGridState,
    undo: undoGrid,
    redo: redoGrid,
    canUndo: canUndoGrid,
    canRedo: canRedoGrid,
  } = useUndoRedo();

  const gridDataRef = useRef(gridData);
  gridDataRef.current = gridData;
  const activeTabRef = useRef(activeTab);
  activeTabRef.current = activeTab;

  const markDirty = useCallback(() => {
    setIsDirty(true);
    isDirtyRef.current = true;
  }, []);

  const recordAndSetTasks = useCallback((updater) => {
    pushState(tasksRef.current);
    setTasks(updater);
    markDirty();
  }, [pushState, markDirty]);

  const handleUndo = useCallback(() => {
    if (activeTabRef.current !== 'gantt') {
      const prev = undoGrid(gridDataRef.current);
      if (prev) setGridData(prev);
      return;
    }
    const prev = undo(tasksRef.current);
    if (prev) setTasks(prev);
  }, [undo, undoGrid]);

  const handleRedo = useCallback(() => {
    if (activeTabRef.current !== 'gantt') {
      const next = redoGrid(gridDataRef.current);
      if (next) setGridData(next);
      return;
    }
    const next = redo(tasksRef.current);
    if (next) setTasks(next);
  }, [redo, redoGrid]);

  const handleBeginDrag = useCallback(() => {
    beginBatch(tasksRef.current);
  }, [beginBatch]);

  const handleEndDrag = useCallback(() => {
    endBatch();
  }, [endBatch]);

  const skipWeekends = toBool(viewOptions.skipWeekends);
  const prevSkipWeekends = useRef(skipWeekends);
  const showMonthLabels = toBool(viewOptions.showMonthLabels ?? true);
  const showWeekLabels = toBool(viewOptions.showWeekLabels ?? false);
  const showDayLabels = toBool(viewOptions.showDayLabels ?? true);
  const effectiveHeaderHeight =
    (showMonthLabels ? 16 : 0) +
    (showWeekLabels ? 18 : 0) +
    (showDayLabels ? 18 : 0);

  const allColumnsOrdered = useMemo(() => {
    const colMap = new Map([
      ...ALL_COLUMNS.map((c) => [c.key, c]),
      ...customColumns.map((c) => [c.key, c]),
    ]);
    return columnOrder.filter((k) => colMap.has(k)).map((k) => colMap.get(k));
  }, [columnOrder, customColumns]);

  const wbsTasks = useMemo(() => buildWbsTree(tasks, skipWeekends), [tasks, skipWeekends]);

  const enrichedTasks = useMemo(() => {
    if (wbsTasks.length === 0) return [];
    const cpmResults = computeCpm(wbsTasks);
    const cpmMap = new Map(cpmResults.map((r) => [r.id, r]));

    const successorMap = new Map();
    for (const task of wbsTasks) {
      if (!task.dependency) continue;
      const preds = String(task.dependency).split(',').map((s) => s.trim()).filter(Boolean);
      for (const predId of preds) {
        if (!successorMap.has(predId)) successorMap.set(predId, []);
        successorMap.get(predId).push(task);
      }
    }

    return wbsTasks.map((task) => {
      const cpm = cpmMap.get(String(task.id));
      const totalFloat = cpm?.totalFloat ?? 0;

      let freeFloat = 0;
      if (task.endDate) {
        const succs = successorMap.get(String(task.id));
        if (succs && succs.length > 0) {
          let minSuccStart = null;
          for (const s of succs) {
            if (s.startDate && (!minSuccStart || s.startDate < minSuccStart)) {
              minSuccStart = s.startDate;
            }
          }
          if (minSuccStart) {
            const gapMs = new Date(minSuccStart + 'T12:00:00') - new Date(task.endDate + 'T12:00:00');
            const gapDays = Math.round(gapMs / 86400000) - 1;
            freeFloat = Math.max(gapDays, 0);
          }
        }
      }

      const slackDays = Math.max(totalFloat, freeFloat);

      return {
        ...task,
        earlyStart: cpm?.earlyStart ?? 0,
        earlyFinish: cpm?.earlyFinish ?? 0,
        lateStart: cpm?.lateStart ?? 0,
        lateFinish: cpm?.lateFinish ?? 0,
        totalFloat: slackDays,
        isCritical: cpm?.isCritical ?? false,
      };
    });
  }, [wbsTasks]);

  const displayTasks = useMemo(() => {
    if (collapsedParents.size === 0) return enrichedTasks;
    // Transitive walk: a task is hidden if ANY ancestor (not only its direct
    // parent) is collapsed. Needed so collapsing a grandparent actually hides
    // grandchildren, matching MS Project / Excel grouping behaviour.
    const parentOf = new Map();
    for (const t of enrichedTasks) {
      parentOf.set(String(t.id), t.parentId ? String(t.parentId) : null);
    }
    const guard = new Set();
    return enrichedTasks.filter((t) => {
      let pid = parentOf.get(String(t.id));
      guard.clear();
      while (pid) {
        if (guard.has(pid)) break;
        guard.add(pid);
        if (collapsedParents.has(pid)) return false;
        pid = parentOf.get(pid);
      }
      return true;
    });
  }, [enrichedTasks, collapsedParents]);

  useEffect(() => {
    const leafTasks = enrichedTasks.filter((t) => !t.isParent);
    const allComplete = leafTasks.length > 0 && leafTasks.every((t) => t.status === 'Completed');
    if (allComplete && !prevAllCompleteRef.current) {
      const allTasks = enrichedTasks;
      const taskCount = leafTasks.length;
      const ownerSet = new Set(leafTasks.map((t) => t.owner).filter(Boolean));
      const ownerCount = ownerSet.size;
      const criticalCount = leafTasks.filter((t) => t.isCritical).length;

      let minStart = null;
      let maxEnd = null;
      for (const t of allTasks) {
        if (t.startDate && (!minStart || t.startDate < minStart)) minStart = t.startDate;
        if (t.endDate && (!maxEnd || t.endDate > maxEnd)) maxEnd = t.endDate;
      }

      let calendarDays = null;
      let workingDays = null;
      if (minStart && maxEnd) {
        const ms = new Date(maxEnd) - new Date(minStart);
        calendarDays = Math.round(ms / 86400000) + 1;
        workingDays = workingDaysBetween(minStart, maxEnd, skipWeekends);
      }

      setShowFireworks({ taskCount, ownerCount, criticalCount, calendarDays, workingDays, projectName });
    }
    prevAllCompleteRef.current = allComplete;
  }, [enrichedTasks, projectName, skipWeekends]);

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

        if (field === 'progress' && Number(value) >= 100) {
          updated.status = 'Completed';
        }
        if (field === 'status' && value === 'Completed') {
          updated.progress = 100;
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
        ...Object.fromEntries(customColumns.map((cc) => [cc.key, ''])),
      },
    ]);
  }, [recordAndSetTasks, customColumns]);

  // Deleting a parent task must cascade to every descendant in a single step.
  // Without the cascade the surviving children kept a dangling parentId, which
  // made buildWbsTree's resolve() walk a missing resultMap entry and throw
  // (TypeError: Cannot set properties of undefined) during render -> white
  // screen. The same recordAndSetTasks call carries one undo snapshot so
  // Ctrl+Z restores the parent plus every descendant atomically.
  const handleDeleteTask = useCallback((taskId) => {
    const targetId = String(taskId);
    const current = tasksRef.current;
    const toRemove = new Set([targetId]);
    let grew = true;
    while (grew) {
      grew = false;
      for (const t of current) {
        const id = String(t.id);
        if (toRemove.has(id)) continue;
        const pid = t.parentId ? String(t.parentId) : '';
        if (pid && toRemove.has(pid)) {
          toRemove.add(id);
          grew = true;
        }
      }
    }
    recordAndSetTasks((prev) => prev.filter((t) => !toRemove.has(String(t.id))));
    setCollapsedParents((prev) => {
      let changed = false;
      const next = new Set(prev);
      for (const id of toRemove) {
        if (next.delete(id)) changed = true;
      }
      return changed ? next : prev;
    });
  }, [recordAndSetTasks]);

  // Move a task (and all its descendants) so they appear before `beforeId` in the raw list.
  // `beforeId` = null means append at the end.
  // When a collapsed parent is the drop target, the DataTable passes the ID of the next
  // visible task (which is already after all hidden children), so the group lands correctly.
  const handleReorderTask = useCallback((draggedId, beforeId) => {
    recordAndSetTasks((prev) => {
      // Collect the dragged task plus every descendant (preserving internal order).
      const toMoveIds = new Set();
      const collectGroup = (id) => {
        toMoveIds.add(String(id));
        for (const t of prev) {
          if (String(t.parentId) === String(id)) collectGroup(t.id);
        }
      };
      collectGroup(draggedId);

      // If the anchor is inside the moving group, treat as no-op.
      if (beforeId != null && toMoveIds.has(String(beforeId))) return prev;

      const movedTasks = prev.filter((t) => toMoveIds.has(String(t.id)));
      const remaining  = prev.filter((t) => !toMoveIds.has(String(t.id)));

      if (beforeId == null) return [...remaining, ...movedTasks];

      const insertAt = remaining.findIndex((t) => String(t.id) === String(beforeId));
      if (insertAt === -1) return [...remaining, ...movedTasks];

      return [
        ...remaining.slice(0, insertAt),
        ...movedTasks,
        ...remaining.slice(insertAt),
      ];
    });
  }, [recordAndSetTasks]);

  const handleToggleCollapse = useCallback((parentId) => {
    setCollapsedParents((prev) => {
      const next = new Set(prev);
      if (next.has(parentId)) next.delete(parentId);
      else next.add(parentId);
      return next;
    });
    markDirty();
  }, [markDirty]);

  const handleColWidthChange = useCallback((colKey, newW) => {
    setColWidths((prev) => ({ ...prev, [colKey]: newW }));
    markDirty();
  }, [markDirty]);

  const handleGanttScaleChange = useCallback((v) => {
    setGanttScale(v);
    markDirty();
  }, [markDirty]);

  const handleGanttZoomChange = useCallback((v) => {
    setGanttZoom(v);
    markDirty();
  }, [markDirty]);

  const handleGridZoomChange = useCallback((v) => {
    setGridZoom(v);
    markDirty();
  }, [markDirty]);

  const handleAddColumn = useCallback((label) => {
    const key = 'custom_' + Date.now().toString(36);
    const col = { key, label, type: 'text' };
    setCustomColumns((prev) => [...prev, col]);
    setColumnOrder((prev) => [...prev, key]);
    setVisibleColumns((prev) => {
      const next = new Set(prev);
      next.add(key);
      return next;
    });
    markDirty();
  }, [markDirty]);

  const handleReorderColumn = useCallback((fromKey, targetKey) => {
    if (fromKey === targetKey) return;
    setColumnOrder((prev) => {
      const arr = [...prev];
      const fromIdx = arr.indexOf(fromKey);
      if (fromIdx === -1) return prev;
      arr.splice(fromIdx, 1);
      if (targetKey == null) {
        arr.push(fromKey);
      } else {
        const insertIdx = arr.indexOf(targetKey);
        if (insertIdx === -1) {
          arr.push(fromKey);
        } else {
          arr.splice(insertIdx, 0, fromKey);
        }
      }
      return arr;
    });
    markDirty();
  }, [markDirty]);

  const handleDeleteColumn = useCallback((key) => {
    setCustomColumns((prev) => prev.filter((c) => c.key !== key));
    setColumnOrder((prev) => prev.filter((k) => k !== key));
    setVisibleColumns((prev) => {
      const next = new Set(prev);
      next.delete(key);
      return next;
    });
    setTasks((prev) =>
      prev.map((t) => {
        const updated = { ...t };
        delete updated[key];
        return updated;
      }),
    );
    markDirty();
  }, [markDirty]);

  const handleAddTab = useCallback(() => {
    const id = 'tab_' + Date.now().toString(36);
    const existingNames = tabs.map((t) => t.name);
    let name = 'Sheet1';
    let counter = 1;
    while (existingNames.includes(name)) {
      counter++;
      name = `Sheet${counter}`;
    }
    setTabs((prev) => [...prev, { id, name }]);
    setGridData((prev) => ({
      ...prev,
      [id]: { rows: 50, cols: 26, cells: {}, colWidths: {}, merges: [], shapes: [] },
    }));
    setActiveTab(id);
    markDirty();
  }, [tabs, markDirty]);

  const handleRenameTab = useCallback((tabId, newName) => {
    const reserved = ['tasks', 'settings'];
    const trimmed = newName.trim();
    if (!trimmed || reserved.includes(trimmed.toLowerCase())) return;
    setTabs((prev) => prev.map((t) => t.id === tabId ? { ...t, name: trimmed } : t));
    markDirty();
  }, [markDirty]);

  const handleDeleteTab = useCallback((tabId) => {
    setTabs((prev) => prev.filter((t) => t.id !== tabId));
    setGridData((prev) => {
      const next = { ...prev };
      delete next[tabId];
      return next;
    });
    if (activeTab === tabId) setActiveTab('gantt');
    markDirty();
  }, [activeTab, markDirty]);

  const handleSelectTab = useCallback((tabId) => {
    setActiveTab(tabId);
  }, []);

  const handleGridCellChange = useCallback((tabId, newData) => {
    pushGridState(gridDataRef.current);
    setGridData((prev) => ({ ...prev, [tabId]: newData }));
    markDirty();
  }, [pushGridState, markDirty]);

  const handleApplyPreset = useCallback((name, colors) => {
    setActiveTheme(name);
    applyThemeToDOM(colors);
    setSettings((prev) => writeThemeColors({ ...prev, themeName: name }, colors));
    markDirty();
  }, [markDirty]);

  const handleApplyCustomColor = useCallback((key, hex) => {
    setActiveTheme('Custom');
    document.documentElement.style.setProperty(`--color-${key}`, hex);
    if (key === 'accent' || key === 'accent-hover') {
      refreshAccentContrastVars(document.documentElement);
    }
    setSettings((prev) => {
      const current = extractThemeColors(prev);
      current[key] = hex;
      return writeThemeColors({ ...prev, themeName: 'Custom' }, current);
    });
    markDirty();
  }, [markDirty]);

  const handleChangeCategoryColor = useCallback((category, hex) => {
    setCategoryColors((prev) => ({ ...prev, [category]: hex }));
    markDirty();
  }, [markDirty]);

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
    markDirty();
  }, [markDirty]);

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
      next.customColumns = JSON.stringify(customColumns);
      next.columnOrder = columnOrder.join(',');
      next.splitRatio = String(splitRatio);
      next.collapsedParents = [...collapsedParents].join(',');
      next.colWidths = JSON.stringify(colWidths);
      next.ganttScale = ganttScale;
      next.ganttZoom = String(ganttZoom);
      next.gridZoom = String(gridZoom);
      next.tabs = JSON.stringify(tabs);
      next.activeTab = activeTab;
      const stylesMap = {};
      const shapesMap = {};
      for (const tab of tabs) {
        const tabData = gridData[tab.id];
        if (!tabData) continue;
        if (tabData.cells) {
          const tabStyles = {};
          for (const [cellRef, cell] of Object.entries(tabData.cells)) {
            if (cell.s && Object.keys(cell.s).length > 0) {
              tabStyles[cellRef] = cell.s;
            }
          }
          if (Object.keys(tabStyles).length > 0) stylesMap[tab.id] = tabStyles;
        }
        if (Array.isArray(tabData.shapes) && tabData.shapes.length > 0) {
          shapesMap[tab.id] = tabData.shapes;
        }
      }
      next.gridCellStyles = JSON.stringify(stylesMap);
      next.gridShapes = JSON.stringify(shapesMap);
      return next;
    });
  }, [viewOptions, visibleColumns, categoryColors, projectName, customColumns, columnOrder, splitRatio, collapsedParents, colWidths, ganttScale, ganttZoom, gridZoom, tabs, activeTab, gridData]);

  useEffect(() => {
    const handler = (e) => {
      const mod = e.metaKey || e.ctrlKey;
      if (!mod) return;

      if (e.key === 's' || e.key === 'S') {
        e.preventDefault();
        handleExportRef.current();
        return;
      }

      const tag = e.target?.tagName;
      const isTextInput = tag === 'INPUT' || tag === 'TEXTAREA';

      if (e.key === 'z' && !e.shiftKey) {
        if (isTextInput) return;
        e.preventDefault();
        handleUndo();
        return;
      }

      if ((e.key === 'z' && e.shiftKey) || (e.key === 'Z' && e.shiftKey) || e.key === 'y') {
        if (isTextInput) return;
        e.preventDefault();
        handleRedo();
      }
    };
    document.addEventListener('keydown', handler);
    return () => document.removeEventListener('keydown', handler);
  }, [handleUndo, handleRedo]);

  useEffect(() => {
    const handler = (e) => {
      if (!isDirtyRef.current) return;
      e.preventDefault();
    };
    window.addEventListener('beforeunload', handler);
    return () => window.removeEventListener('beforeunload', handler);
  }, []);

  const handleImport = useCallback(async (file) => {
    setImportError(null);
    try {
      const result = await importExcel(file);
      setTasks(result.tasks);
      setSettings(result.settings);
      setGridData(result.gridData || {});

      const s = result.settings;
      setProjectName(s.projectName || '');

      const ratio = parseFloat(s.splitRatio);
      if (Number.isFinite(ratio) && ratio > 0 && ratio < 1) {
        setSplitRatio(ratio);
      } else {
        setSplitRatio(0.38);
      }

      const collapsedList = (s.collapsedParents || '')
        .split(',')
        .map((x) => x.trim())
        .filter(Boolean);
      setCollapsedParents(new Set(collapsedList));

      let restoredColWidths = {};
      try { restoredColWidths = JSON.parse(s.colWidths || '{}'); } catch { /* ignore */ }
      setColWidths(restoredColWidths && typeof restoredColWidths === 'object' ? restoredColWidths : {});

      setGanttScale(s.ganttScale === 'week' ? 'week' : 'day');
      const zoom = parseInt(s.ganttZoom, 10);
      setGanttZoom(Number.isFinite(zoom) && zoom >= 5 && zoom <= 300 ? zoom : 100);
      const gz = parseInt(s.gridZoom, 10);
      setGridZoom(Number.isFinite(gz) && gz >= 5 && gz <= 300 ? gz : 100);
      const colors = extractThemeColors(s);
      applyThemeToDOM(colors);
      setActiveTheme(s.themeName || 'Notion Light');

      setViewOptions({
        showCriticalPath: toBool(s.showCriticalPath ?? true),
        showSlack: toBool(s.showSlack ?? true),
        showDependencies: toBool(s.showDependencies ?? true),
        showTodayLine: toBool(s.showTodayLine ?? true),
        showBaseline: toBool(s.showBaseline ?? true),
        showTaskNames: toBool(s.showTaskNames ?? true),
        showProgressPercent: toBool(s.showProgressPercent ?? true),
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

      let restoredCustomCols = [];
      try { restoredCustomCols = JSON.parse(s.customColumns || '[]'); } catch { /* ignore */ }
      setCustomColumns(restoredCustomCols);

      const restoredOrder = s.columnOrder
        ? s.columnOrder.split(',').map((c) => c.trim()).filter(Boolean)
        : [];
      setColumnOrder(restoredOrder.length > 0 ? restoredOrder : ALL_COLUMNS.map((c) => c.key));

      // Use the tabs list returned by importExcel - it reflects both
      // the persisted tab order from the Settings sheet AND any new
      // non-reserved worksheets the user added directly in Excel.
      const restoredTabs = Array.isArray(result.tabs) ? result.tabs : [];
      setTabs(restoredTabs);
      setActiveTab(s.activeTab || 'gantt');

      let restored = {};
      if (s.categoryColors) {
        try { restored = JSON.parse(s.categoryColors); } catch (_) { /* ignore */ }
      }
      setCategoryColors(assignCategoryColors(result.tasks, restored));

      setLastSavedAt(null);
      setIsDirty(false);
      isDirtyRef.current = false;
    } catch (err) {
      console.error('Import failed:', err);
      setImportError(err.message || 'Import failed. Please check the file and try again.');
    }
  }, []);

  useEffect(() => {
    if (!importError) return;
    const timer = setTimeout(() => setImportError(null), 6000);
    return () => clearTimeout(timer);
  }, [importError]);

  // shareNotice is dismissed manually by the user (no auto-timer).

  // Auto-load project data embedded at share time (window.__GANTTGEN_INITIAL_XLSX__).
  useEffect(() => {
    const rawData = window.__GANTTGEN_INITIAL_XLSX__;
    if (!rawData) return;
    try { delete window.__GANTTGEN_INITIAL_XLSX__; } catch { /* ignore on strict environments */ }

    const binary = atob(rawData);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    const blob = new Blob([bytes], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const file = new File([blob], 'shared-project.xlsx', { type: blob.type });
    handleImport(file);
  }, [handleImport]);

  const handleShare = useCallback(() => {
    // Share only works from the production build where all JS/CSS is inlined in <head>.
    // In dev mode, <head> contains Vite HMR stubs that won't work standalone.
    const isDevMode = !!document.querySelector('script[src*="/@vite/client"]');
    if (isDevMode) {
      setImportError('Share is only available from the production build (dist/index.html). Run "npm run build" first.');
      return;
    }

    const base64Data = exportExcelToBase64(enrichedTasks, settings, gridData);

    // Reconstruct the HTML from the live DOM's <head> (which retains the original inlined
    // <script> and <style> blocks untouched by React) and a clean <body> with an empty #root.
    // This avoids fetch/XHR which fail on file:// protocol in most browsers.
    let headContent = document.head.innerHTML;

    // Strip any previously embedded data (re-sharing a shared file).
    headContent = headContent.replace(
      /<script>\s*window\.__GANTTGEN_INITIAL_XLSX__\s*=\s*"[^"]*";\s*<\/script>\s*/,
      '',
    );

    // Strip Vite dev-server injected styles (data-vite-dev-id) as a safety net.
    headContent = headContent.replace(/<style[^>]*data-vite-dev-id[^>]*>[\s\S]*?<\/style>\s*/g, '');

    const dataScript = '<script>window.__GANTTGEN_INITIAL_XLSX__="' + base64Data + '";</' + 'script>';

    const html = '<!DOCTYPE html>\n<html lang="en">\n<head>\n'
      + dataScript + '\n'
      + headContent
      + '\n</head>\n<body>\n<div id="root"></div>\n</body>\n</html>';

    const safeProjectName = (projectName.trim() || 'Project').replace(/[^\w\s-]/g, '');
    const filename = `GanttGen-${safeProjectName}.html`;

    const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    link.click();
    setTimeout(() => URL.revokeObjectURL(url), 30000);

    setShareNotice(filename);
  }, [enrichedTasks, settings, projectName, gridData]);

  const handleDragEnter = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
    dragCounter.current += 1;
    if (dragCounter.current === 1) setDragActive(true);
  }, []);

  const handleDragLeave = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
    dragCounter.current -= 1;
    if (dragCounter.current <= 0) {
      dragCounter.current = 0;
      setDragActive(false);
    }
  }, []);

  const handleDragOver = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
  }, []);

  const handleDrop = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
    dragCounter.current = 0;
    setDragActive(false);

    const file = e.dataTransfer?.files?.[0];
    if (file) handleImport(file);
  }, [handleImport]);

  const handleExport = useCallback(() => {
    const filename = 'GanttGen-' + (projectName.trim() || 'Project').replace(/[^\w\s-]/g, '') + '.xlsx';
    exportExcel(enrichedTasks, settings, filename, gridData);
    setLastSavedAt(new Date());
    setIsDirty(false);
    isDirtyRef.current = false;
  }, [enrichedTasks, settings, projectName, gridData]);
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

      const opts = {
        backgroundColor: bgColor,
        filter: (node) => !node?.hasAttribute?.('data-export-exclude'),
      };
      if (exportHeight != null && exportHeight > 0 && exportHeight < target.offsetHeight) {
        opts.height = exportHeight;
      }

      const dataUrl = await toPng(target, opts);
      const safeName = (projectName.trim() || 'Project').replace(/[^\w\s-]/g, '');
      const suffix = mode === 'full' ? 'Full' : 'Chart';
      const link = document.createElement('a');
      link.download = `GanttGen-${safeName}_${suffix}.png`;
      link.href = dataUrl;
      link.click();
    } catch (err) {
      console.error('PNG export failed:', err);
    } finally {
      document.head.removeChild(scrollbarStyle);
    }
  }, [projectName]);

  const handleMouseDown = useCallback((e) => {
    e.preventDefault();
    isDragging.current = true;
    let moved = false;

    const onMouseMove = (e) => {
      if (!isDragging.current || !containerRef.current) return;
      const rect = containerRef.current.getBoundingClientRect();
      const totalWidth = rect.width;
      const x = e.clientX - rect.left;
      const ratio = x / totalWidth;
      const minLeftRatio = MIN_LEFT_WIDTH / totalWidth;
      const maxLeftRatio = 1 - MIN_RIGHT_WIDTH / totalWidth;
      setSplitRatio(Math.min(maxLeftRatio, Math.max(minLeftRatio, ratio)));
      moved = true;
    };

    const onMouseUp = () => {
      isDragging.current = false;
      document.removeEventListener('mousemove', onMouseMove);
      document.removeEventListener('mouseup', onMouseUp);
      if (moved) markDirty();
    };

    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp);
  }, [markDirty]);

  return (
    <div
      className="flex flex-col h-screen relative"
      style={{ backgroundColor: 'var(--color-bg-primary)' }}
      onDragEnter={handleDragEnter}
      onDragLeave={handleDragLeave}
      onDragOver={handleDragOver}
      onDrop={handleDrop}
    >
      {dragActive && <DropOverlay />}
      {importError && <ImportErrorToast message={importError} onDismiss={() => setImportError(null)} />}
      {showFireworks && <FireworksOverlay stats={showFireworks} onDismiss={() => setShowFireworks(false)} />}
      {shareNotice && <ShareNoticeToast message={shareNotice} onDismiss={() => setShareNotice(null)} />}
      <Dashboard
        tasks={enrichedTasks}
        projectName={projectName}
        onChangeProjectName={(v) => { setProjectName(v); markDirty(); }}
        lastSavedAt={lastSavedAt}
        isDirty={isDirty}
        onImport={handleImport}
        onDownloadTemplate={downloadTemplate}
        onExport={handleExport}
        onExportPng={handleExportPng}
        onShare={handleShare}
        onOpenGuide={handleOpenGuide}
        onOpenTheme={handleOpenTheme}
        onUndo={handleUndo}
        onRedo={handleRedo}
        canUndo={activeTab !== 'gantt' ? canUndoGrid : canUndo}
        canRedo={activeTab !== 'gantt' ? canRedoGrid : canRedo}
        viewOptionsOpen={viewOptionsOpen}
        onToggleViewOptions={handleToggleViewOptions}
        viewBtnRef={viewBtnRef}
        viewOptions={viewOptions}
        onToggleViewOption={handleToggleViewOption}
        columns={allColumnsOrdered}
        visibleColumns={visibleColumns}
        onToggleColumn={(key) => {
          setVisibleColumns((prev) => {
            const next = new Set(prev);
            if (next.has(key)) next.delete(key);
            else next.add(key);
            return next;
          });
          markDirty();
        }}
      />

      {activeTab === 'gantt' ? (
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
              columns={allColumnsOrdered}
              visibleColumns={visibleColumns}
              onAddColumn={handleAddColumn}
              onReorderColumn={handleReorderColumn}
              onDeleteColumn={handleDeleteColumn}
              onToggleColumn={(key) => {
                setVisibleColumns((prev) => {
                  const next = new Set(prev);
                  if (next.has(key)) next.delete(key);
                  else next.add(key);
                  return next;
                });
                markDirty();
              }}
              collapsedParents={collapsedParents}
              onToggleCollapse={handleToggleCollapse}
              onUpdateTask={handleUpdateTask}
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
              colWidths={colWidths}
              onColWidthChange={handleColWidthChange}
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
              scale={ganttScale}
              zoomPct={ganttZoom}
            />
          </div>
        </div>
      ) : (
        <div className="flex-1 min-h-0">
          <DataGrid
            data={gridData[activeTab]}
            onChange={(newData) => handleGridCellChange(activeTab, newData)}
            zoomPct={gridZoom}
          />
        </div>
      )}

      <StatusBar
        onAddTask={handleAddTask}
        scale={ganttScale}
        onScaleChange={handleGanttScaleChange}
        zoomPct={ganttZoom}
        onZoomChange={handleGanttZoomChange}
        gridZoomPct={gridZoom}
        onGridZoomChange={handleGridZoomChange}
        activeTab={activeTab}
        tabs={tabs}
        onSelectTab={handleSelectTab}
        onAddTab={handleAddTab}
        onRenameTab={handleRenameTab}
        onDeleteTab={handleDeleteTab}
      />

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

function DropOverlay() {
  return (
    <div
      className="absolute inset-0 z-[9999] flex items-center justify-center pointer-events-none"
      style={{ backgroundColor: 'rgba(0, 0, 0, 0.55)', backdropFilter: 'blur(4px)' }}
    >
      <div
        className="flex flex-col items-center gap-3 px-10 py-8 rounded-xl border-2 border-dashed"
        style={{
          borderColor: 'var(--color-accent)',
          backgroundColor: 'var(--color-bg-secondary)',
        }}
      >
        <FileUp size={36} style={{ color: 'var(--color-accent)' }} />
        <span className="text-base font-semibold" style={{ color: 'var(--color-text-primary)' }}>
          Drop Excel file to import
        </span>
        <span className="text-sm" style={{ color: 'var(--color-text-muted)' }}>
          .xlsx or .xls
        </span>
      </div>
    </div>
  );
}

function ImportErrorToast({ message, onDismiss }) {
  return (
    <div
      className="absolute top-4 left-1/2 -translate-x-1/2 z-[9999] flex items-start gap-2.5 px-4 py-3 rounded-lg shadow-xl max-w-md"
      style={{
        backgroundColor: 'var(--color-bg-secondary)',
        border: '1px solid var(--color-danger)',
      }}
    >
      <AlertCircle size={18} className="flex-shrink-0 mt-0.5" style={{ color: 'var(--color-danger)' }} />
      <span className="text-[13px] leading-snug" style={{ color: 'var(--color-text-primary)' }}>
        {message}
      </span>
      <button
        onClick={onDismiss}
        className="flex-shrink-0 ml-2 p-0.5 rounded transition-colors cursor-pointer"
        style={{ color: 'var(--color-text-muted)' }}
        onMouseEnter={(e) => { e.currentTarget.style.color = 'var(--color-text-primary)'; }}
        onMouseLeave={(e) => { e.currentTarget.style.color = 'var(--color-text-muted)'; }}
      >
        <X size={14} />
      </button>
    </div>
  );
}

function ShareNoticeToast({ message: filename, onDismiss }) {
  return (
    <div
      className="absolute inset-0 z-[9999] flex items-start justify-center pt-16 px-4"
      style={{
        backgroundColor: 'rgba(15, 15, 18, 0.45)',
        WebkitBackdropFilter: 'blur(10px)',
        backdropFilter: 'blur(10px)',
      }}
      aria-modal="true"
      role="dialog"
      aria-labelledby="share-notice-title"
    >
      <div
        className="rounded-xl shadow-2xl max-w-full"
        style={{
          backgroundColor: 'var(--color-bg-secondary)',
          border: '1px solid var(--color-border)',
          width: 360,
        }}
        onClick={(e) => e.stopPropagation()}
      >
      <div
        className="flex items-center justify-between px-4 py-2.5 rounded-t-xl"
        style={{ backgroundColor: 'var(--color-accent)', color: 'var(--color-on-accent)' }}
      >
        <div className="flex items-center gap-2">
          <Share2 size={15} />
          <span id="share-notice-title" className="text-[13px] font-semibold">File ready to share</span>
        </div>
        <button
          onClick={onDismiss}
          className="p-0.5 rounded transition-opacity cursor-pointer opacity-70"
          onMouseEnter={(e) => { e.currentTarget.style.opacity = '1'; }}
          onMouseLeave={(e) => { e.currentTarget.style.opacity = '0.7'; }}
        >
          <X size={14} />
        </button>
      </div>

      <div className="px-4 py-3 flex flex-col gap-2.5">
        <div
          className="text-[12px] font-medium px-2.5 py-1.5 rounded-md truncate"
          style={{
            backgroundColor: 'var(--color-bg-tertiary)',
            color: 'var(--color-text-primary)',
            border: '1px solid var(--color-border-subtle)',
          }}
          title={filename}
        >
          {filename}
        </div>

        <div className="flex flex-col gap-1.5">
          <HintRow icon={Share2} text="Send this file to anyone via chat, email, or USB" />
          <HintRow icon={MousePointerClick} text="Recipient double-clicks to open instantly" />
          <HintRow icon={Globe} text="Works in any modern browser, fully offline" />
        </div>
      </div>

      <div className="px-4 pb-3">
        <button
          onClick={onDismiss}
          className="w-full py-1.5 rounded-md text-[12px] font-medium transition-colors cursor-pointer"
          style={{
            backgroundColor: 'var(--color-bg-tertiary)',
            color: 'var(--color-text-secondary)',
            border: '1px solid var(--color-border)',
          }}
          onMouseEnter={(e) => { e.currentTarget.style.backgroundColor = 'var(--color-bg-hover)'; }}
          onMouseLeave={(e) => { e.currentTarget.style.backgroundColor = 'var(--color-bg-tertiary)'; }}
        >
          Got it
        </button>
      </div>
      </div>
    </div>
  );
}

function HintRow({ icon: Icon, text }) {
  return (
    <div className="flex items-center gap-2">
      <Icon size={13} className="flex-shrink-0" style={{ color: 'var(--color-accent)', opacity: 0.8 }} />
      <span className="text-[12px] leading-snug" style={{ color: 'var(--color-text-secondary)' }}>
        {text}
      </span>
    </div>
  );
}

const FW_COLORS = [
  '#ff1744','#ff5252','#ff9100','#ffea00','#76ff03','#00e676',
  '#00bcd4','#2979ff','#651fff','#d500f9','#f50057','#ff6d00',
  '#69f0ae','#40c4ff','#e040fb','#ffff00',
];
const GRAVITY = 0.035;
const FRICTION = 0.985;
const PARTICLE_COUNT = 120;
const WAVE_COUNT = 4;
const ROCKETS_PER_WAVE = 5;
const WAVE_INTERVAL = 1800;
const DURATION_MS = 9000;

function FireworksOverlay({ stats, onDismiss }) {
  const [replayKey, setReplayKey] = useState(0);
  const canvasRef = useRef(null);

  const handleReplay = () => setReplayKey((k) => k + 1);

  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    let raf;
    let particles = [];
    let rockets = [];
    let elapsed = 0;
    let last = performance.now();

    const resize = () => {
      canvas.width = canvas.offsetWidth * devicePixelRatio;
      canvas.height = canvas.offsetHeight * devicePixelRatio;
      ctx.scale(devicePixelRatio, devicePixelRatio);
    };
    resize();
    window.addEventListener('resize', resize);

    const w = () => canvas.offsetWidth;
    const h = () => canvas.offsetHeight;

    const spawn = (x, y) => {
      const base = FW_COLORS[Math.floor(Math.random() * FW_COLORS.length)];
      const accent = FW_COLORS[Math.floor(Math.random() * FW_COLORS.length)];
      for (let i = 0; i < PARTICLE_COUNT; i++) {
        const angle = Math.random() * Math.PI * 2;
        const speed = Math.random() * 7 + 2;
        const isTrail = Math.random() < 0.3;
        particles.push({
          x, y,
          vx: Math.cos(angle) * speed * (isTrail ? 0.4 : 1),
          vy: Math.sin(angle) * speed * (isTrail ? 0.4 : 1),
          alpha: 1,
          color: Math.random() < 0.7 ? base : accent,
          size: isTrail ? Math.random() * 1.5 + 0.5 : Math.random() * 3.5 + 1.5,
          decay: 0.004 + Math.random() * 0.006,
          glow: !isTrail && Math.random() < 0.4,
        });
      }
    };

    const launchRocket = () => {
      rockets.push({
        x: w() * (0.08 + Math.random() * 0.84),
        y: h(),
        vy: -(h() * 0.015 + Math.random() * 4),
        targetY: h() * (0.1 + Math.random() * 0.3),
        trail: [],
        exploded: false,
      });
    };

    const timeouts = [];
    for (let wave = 0; wave < WAVE_COUNT; wave++) {
      for (let r = 0; r < ROCKETS_PER_WAVE; r++) {
        timeouts.push(setTimeout(launchRocket, wave * WAVE_INTERVAL + r * 200 + Math.random() * 150));
      }
    }

    const tick = (now) => {
      const dt = now - last;
      last = now;
      elapsed += dt;

      ctx.clearRect(0, 0, w(), h());

      for (let i = rockets.length - 1; i >= 0; i--) {
        const r = rockets[i];
        r.y += r.vy;
        r.trail.push({ x: r.x, y: r.y, alpha: 1 });
        if (r.trail.length > 12) r.trail.shift();

        if (r.y <= r.targetY && !r.exploded) {
          r.exploded = true;
          spawn(r.x, r.y);
          rockets.splice(i, 1);
        } else if (!r.exploded) {
          for (const tp of r.trail) {
            tp.alpha -= 0.08;
            if (tp.alpha > 0) {
              ctx.globalAlpha = tp.alpha * 0.6;
              ctx.beginPath();
              ctx.arc(tp.x, tp.y, 1.5, 0, Math.PI * 2);
              ctx.fillStyle = '#ffe0b2';
              ctx.fill();
            }
          }
          ctx.globalAlpha = 1;
          ctx.beginPath();
          ctx.arc(r.x, r.y, 3, 0, Math.PI * 2);
          ctx.fillStyle = '#fff';
          ctx.shadowColor = '#fff';
          ctx.shadowBlur = 8;
          ctx.fill();
          ctx.shadowBlur = 0;
        }
      }

      for (let i = particles.length - 1; i >= 0; i--) {
        const p = particles[i];
        p.vx *= FRICTION;
        p.vy *= FRICTION;
        p.vy += GRAVITY;
        p.x += p.vx;
        p.y += p.vy;
        p.alpha -= p.decay;
        if (p.alpha <= 0) { particles.splice(i, 1); continue; }
        ctx.globalAlpha = p.alpha;
        if (p.glow) {
          ctx.shadowColor = p.color;
          ctx.shadowBlur = 10;
        }
        ctx.beginPath();
        ctx.arc(p.x, p.y, p.size, 0, Math.PI * 2);
        ctx.fillStyle = p.color;
        ctx.fill();
        if (p.glow) ctx.shadowBlur = 0;
      }
      ctx.globalAlpha = 1;

      if (elapsed < DURATION_MS || particles.length > 0 || rockets.length > 0) {
        raf = requestAnimationFrame(tick);
      }
    };
    raf = requestAnimationFrame(tick);

    return () => {
      cancelAnimationFrame(raf);
      timeouts.forEach(clearTimeout);
      window.removeEventListener('resize', resize);
    };
  }, [replayKey]);

  const { taskCount, ownerCount, criticalCount, calendarDays, workingDays, projectName: pName } = stats || {};

  const rewindItems = [
    calendarDays != null && { value: calendarDays, unit: calendarDays === 1 ? 'day' : 'days', label: 'calendar span' },
    workingDays != null && { value: workingDays, unit: workingDays === 1 ? 'working day' : 'working days', label: 'excl. weekends' },
    taskCount > 0 && { value: taskCount, unit: taskCount === 1 ? 'task' : 'tasks', label: 'completed' },
    criticalCount > 0 && { value: criticalCount, unit: criticalCount === 1 ? 'critical task' : 'critical tasks', label: 'on the critical path' },
    ownerCount > 0 && { value: ownerCount, unit: ownerCount === 1 ? 'team member' : 'team members', label: 'delivered this' },
  ].filter(Boolean);

  return (
    <div
      className="absolute inset-0 z-[9999] flex flex-col items-center justify-center gap-10"
      style={{ backgroundColor: 'rgba(5, 5, 12, 0.65)', backdropFilter: 'blur(8px)', WebkitBackdropFilter: 'blur(8px)' }}
      onClick={onDismiss}
    >
      <canvas
        ref={canvasRef}
        className="absolute inset-0 w-full h-full pointer-events-none"
      />

      <div className="relative text-center select-none pointer-events-none px-8">
        {pName && (
          <div className="text-[15px] font-semibold uppercase tracking-widest mb-3" style={{ color: 'rgba(255,220,100,0.75)' }}>
            {pName}
          </div>
        )}
        <div
          className="text-[56px] font-extrabold tracking-tight leading-tight"
          style={{ color: '#fff', textShadow: '0 0 40px rgba(255,200,50,0.55), 0 2px 16px rgba(0,0,0,0.6)' }}
        >
          Mission Accomplished!
        </div>
        <div className="text-[20px] font-medium mt-4" style={{ color: 'rgba(255,255,255,0.7)' }}>
          Every task, every milestone -- done. You crushed it.
        </div>
      </div>

      {rewindItems.length > 0 && (
        <div
          className="relative flex flex-wrap justify-center gap-12 px-10 pointer-events-none select-none"
          style={{ maxWidth: 780 }}
        >
          {rewindItems.map((item) => (
            <div key={item.label} className="flex flex-col items-center">
              <div className="text-[52px] font-extrabold tabular-nums leading-none" style={{ color: '#fff' }}>
                {item.value}
              </div>
              <div className="text-[16px] font-semibold mt-2" style={{ color: 'rgba(255,220,100,0.9)' }}>
                {item.unit}
              </div>
              <div className="text-[13px] mt-1" style={{ color: 'rgba(255,255,255,0.4)' }}>
                {item.label}
              </div>
            </div>
          ))}
        </div>
      )}

      <div className="relative flex items-center gap-8 mt-2">
        <button
          title="Replay"
          onClick={(e) => { e.stopPropagation(); handleReplay(); }}
          className="flex items-center justify-center cursor-pointer transition-opacity"
          style={{ color: 'rgba(255,255,255,0.55)', background: 'none', border: 'none', padding: 0 }}
          onMouseEnter={(e) => { e.currentTarget.style.color = '#fff'; }}
          onMouseLeave={(e) => { e.currentTarget.style.color = 'rgba(255,255,255,0.55)'; }}
        >
          <RotateCcw size={28} />
        </button>
        <button
          title="Dismiss"
          onClick={(e) => { e.stopPropagation(); onDismiss(); }}
          className="flex items-center justify-center cursor-pointer transition-opacity"
          style={{ color: 'rgba(255,255,255,0.55)', background: 'none', border: 'none', padding: 0 }}
          onMouseEnter={(e) => { e.currentTarget.style.color = '#fff'; }}
          onMouseLeave={(e) => { e.currentTarget.style.color = 'rgba(255,255,255,0.55)'; }}
        >
          <X size={28} />
        </button>
      </div>
    </div>
  );
}
