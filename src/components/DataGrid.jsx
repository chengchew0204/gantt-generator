import { useState, useRef, useCallback, useEffect, useMemo } from 'react';
import { createFormulaEngine } from '../utils/FormulaEngine';
import {
  buildGridClipboard,
  parseClipboardInput,
  coerceNumeric,
  colLabelFromIndex,
} from '../utils/ClipboardUtils';
import {
  formatCellValue,
  parseCellInput,
  bumpDecimals as bumpDecimalsStyle,
} from '../utils/NumberFormat';
import SpreadsheetToolbar from './SpreadsheetToolbar';
import ShapeLayer, {
  nudgeShapes,
  deleteShapes,
  duplicateShapes,
  reorderShape,
} from './ShapeLayer';

const DEFAULT_ROW_HEIGHT = 26;
const HEADER_HEIGHT = 28;
const DEFAULT_COL_WIDTH = 80;
const MIN_COL_WIDTH = 30;
const MIN_ROW_HEIGHT = 16;
const ROW_HEADER_WIDTH = 44;

const REF_COLORS = ['#4285f4', '#ea4335', '#34a853', '#a142f4', '#ff6d01', '#46bdc6'];

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

const H_ALIGN_TO_JUSTIFY = {
  left: 'flex-start',
  center: 'center',
  right: 'flex-end',
};

const V_ALIGN_TO_ITEMS = {
  top: 'flex-start',
  middle: 'center',
  bottom: 'flex-end',
};

function cellStyleToCss(s) {
  if (!s) return {};
  const css = {};
  if (s.bold) css.fontWeight = 700;
  if (s.italic) css.fontStyle = 'italic';
  if (s.underline) css.textDecoration = 'underline';
  if (s.fontSize) css.fontSize = s.fontSize;
  if (s.color) css.color = s.color;
  if (s.bg) css.backgroundColor = s.bg;
  if (s.hAlign && H_ALIGN_TO_JUSTIFY[s.hAlign]) {
    css.justifyContent = H_ALIGN_TO_JUSTIFY[s.hAlign];
    css.textAlign = s.hAlign;
  }
  if (s.vAlign && V_ALIGN_TO_ITEMS[s.vAlign]) {
    css.alignItems = V_ALIGN_TO_ITEMS[s.vAlign];
  }
  return css;
}

function borderSideToCss(b) {
  if (!b || !b.style || b.style === 'none') return null;
  const color = b.color || 'var(--color-text-primary)';
  switch (b.style) {
    case 'thin':   return `1px solid ${color}`;
    case 'medium': return `2px solid ${color}`;
    case 'thick':  return `3px solid ${color}`;
    case 'double': return `3px double ${color}`;
    case 'dashed': return `1px dashed ${color}`;
    case 'dotted': return `1px dotted ${color}`;
    default:       return `1px solid ${color}`;
  }
}

const B_THIN = { style: 'thin' };
const B_THICK = { style: 'thick' };
const B_DOUBLE = { style: 'double' };

function bordersForPreset(id, r, c, r1, r2, c1, c2) {
  const atT = r === r1, atB = r === r2, atL = c === c1, atR = c === c2;
  switch (id) {
    case 'none': return null;
    case 'all':  return { top: B_THIN, bottom: B_THIN, left: B_THIN, right: B_THIN };
    case 'outside': {
      const b = {};
      if (atT) b.top = B_THIN; if (atB) b.bottom = B_THIN;
      if (atL) b.left = B_THIN; if (atR) b.right = B_THIN;
      return Object.keys(b).length ? b : null;
    }
    case 'thick-outside': {
      const b = {};
      if (atT) b.top = B_THICK; if (atB) b.bottom = B_THICK;
      if (atL) b.left = B_THICK; if (atR) b.right = B_THICK;
      return Object.keys(b).length ? b : null;
    }
    case 'bottom':        return atB ? { bottom: B_THIN } : null;
    case 'top':           return atT ? { top: B_THIN } : null;
    case 'left':          return atL ? { left: B_THIN } : null;
    case 'right':         return atR ? { right: B_THIN } : null;
    case 'bottom-double': return atB ? { bottom: B_DOUBLE } : null;
    case 'thick-bottom':  return atB ? { bottom: B_THICK } : null;
    case 'top-bottom': {
      const b = {};
      if (atT) b.top = B_THIN; if (atB) b.bottom = B_THIN;
      return Object.keys(b).length ? b : null;
    }
    case 'top-thick-bottom': {
      const b = {};
      if (atT) b.top = B_THIN; if (atB) b.bottom = B_THICK;
      return Object.keys(b).length ? b : null;
    }
    case 'top-double-bottom': {
      const b = {};
      if (atT) b.top = B_THIN; if (atB) b.bottom = B_DOUBLE;
      return Object.keys(b).length ? b : null;
    }
    default: return null;
  }
}

function rectsOverlap(a, b) {
  return !(a.r2 < b.r1 || b.r2 < a.r1 || a.c2 < b.c1 || b.c2 < a.c1);
}

function rectEquals(a, b) {
  return a.r1 === b.r1 && a.c1 === b.c1 && a.r2 === b.r2 && a.c2 === b.c2;
}

function mergeShapeStyle(current, partial) {
  const base = current || {};
  const next = { ...base };
  if (partial.fill !== undefined) {
    next.fill = { ...(base.fill || {}), ...partial.fill };
  }
  if (partial.outline !== undefined) {
    next.outline = { ...(base.outline || {}), ...partial.outline };
  }
  if (partial.effects !== undefined) {
    next.effects = partial.effects;
  }
  return next;
}

function mergeTextStyle(current, partial) {
  const base = current || {};
  const next = { ...base };
  for (const k of Object.keys(partial)) {
    if (k === 'fill' || k === 'outline' || k === 'effects') {
      if (k === 'effects') next[k] = partial[k];
      else next[k] = { ...(base[k] || {}), ...partial[k] };
    } else {
      next[k] = partial[k];
    }
  }
  return next;
}

function parseFormulaRefs(text) {
  if (!text || !text.startsWith('=')) return {};
  const highlights = {};
  const regex = /([A-Z]{1,3})(\d{1,7})/g;
  let colorIdx = 0;
  const colorMap = new Map();
  let match;
  while ((match = regex.exec(text)) !== null) {
    const ref = match[0];
    if (!colorMap.has(ref)) {
      colorMap.set(ref, REF_COLORS[colorIdx % REF_COLORS.length]);
      colorIdx++;
    }
    highlights[ref] = colorMap.get(ref);
  }
  return highlights;
}

export default function DataGrid({ data, onChange }) {
  const rows = data?.rows ?? 50;
  const cols = data?.cols ?? 26;
  const cells = data?.cells ?? {};
  const gridColWidths = data?.colWidths ?? {};
  const gridRowHeights = data?.rowHeights ?? {};
  const showGridLines = data?.showGridLines !== false;
  const merges = useMemo(() => data?.merges ?? [], [data?.merges]);
  const shapes = useMemo(() => data?.shapes ?? [], [data?.shapes]);

  const [editingCell, setEditingCell] = useState(null);
  const [selectedCell, setSelectedCell] = useState(null);
  const [selectionEnd, setSelectionEnd] = useState(null);
  const [editText, setEditText] = useState('');
  const [colWidths, setColWidths] = useState(gridColWidths);
  const [rowHeights, setRowHeights] = useState(gridRowHeights);
  const [selectedShapeIds, setSelectedShapeIds] = useState([]);
  const [shapeMode, setShapeMode] = useState(null);
  const [editingShapeTextId, setEditingShapeTextId] = useState(null);

  const cellInputRef = useRef(null);
  const barInputRef = useRef(null);
  const containerRef = useRef(null);
  const scrollRef = useRef(null);
  const editSourceRef = useRef('cell');
  const editModeRef = useRef('enter');
  const isDraggingRef = useRef(false);

  const engineRef = useRef(null);
  const [displayValues, setDisplayValues] = useState({});

  const isFormulaEditing = editingCell != null && editText.startsWith('=');

  const mergeLookup = useMemo(() => {
    const anchorMap = new Map();
    const coveredMap = new Map();
    for (const m of merges) {
      anchorMap.set(`${m.r1},${m.c1}`, m);
      for (let r = m.r1; r <= m.r2; r++) {
        for (let c = m.c1; c <= m.c2; c++) {
          if (r === m.r1 && c === m.c1) continue;
          coveredMap.set(`${r},${c}`, m);
        }
      }
    }
    return { anchorMap, coveredMap };
  }, [merges]);

  const getMergeAt = useCallback((r, c) => {
    const k = `${r},${c}`;
    return mergeLookup.anchorMap.get(k) || mergeLookup.coveredMap.get(k) || null;
  }, [mergeLookup]);

  const expandRectWithMerges = useCallback((rect) => {
    if (!rect || merges.length === 0) return rect;
    let out = { ...rect };
    let changed = true;
    let iter = 0;
    while (changed && iter < 50) {
      changed = false;
      for (const m of merges) {
        if (rectsOverlap(out, m)) {
          if (m.r1 < out.r1) { out.r1 = m.r1; changed = true; }
          if (m.r2 > out.r2) { out.r2 = m.r2; changed = true; }
          if (m.c1 < out.c1) { out.c1 = m.c1; changed = true; }
          if (m.c2 > out.c2) { out.c2 = m.c2; changed = true; }
        }
      }
      iter++;
    }
    return out;
  }, [merges]);

  const selRect = useMemo(() => {
    if (!selectedCell) return null;
    const end = selectionEnd || selectedCell;
    const base = {
      r1: Math.min(selectedCell.row, end.row),
      r2: Math.max(selectedCell.row, end.row),
      c1: Math.min(selectedCell.col, end.col),
      c2: Math.max(selectedCell.col, end.col),
    };
    return expandRectWithMerges(base);
  }, [selectedCell, selectionEnd, expandRectWithMerges]);

  const isMultiSel = selRect && (selRect.r1 !== selRect.r2 || selRect.c1 !== selRect.c2);

  const mergedSelection = useMemo(() => {
    if (!selRect) return null;
    for (const m of merges) {
      if (rectEquals(selRect, m)) return m;
    }
    return null;
  }, [selRect, merges]);

  const refHighlights = useMemo(() => {
    if (!isFormulaEditing) return {};
    return parseFormulaRefs(editText);
  }, [isFormulaEditing, editText]);

  useEffect(() => {
    setColWidths(gridColWidths);
  }, [gridColWidths]);

  useEffect(() => {
    setRowHeights(gridRowHeights);
  }, [gridRowHeights]);

  useEffect(() => {
    if (engineRef.current) {
      engineRef.current.destroy();
    }
    const engine = createFormulaEngine();
    engineRef.current = engine;

    const maxRows = Math.max(rows, 100);
    const maxCols = Math.max(cols, 26);
    engine.hydrate(cells, maxRows, maxCols);

    const dv = {};
    for (const key of Object.keys(cells)) {
      const match = key.match(/^([A-Z]+)(\d+)$/);
      if (!match) continue;
      let colIdx = 0;
      for (let i = 0; i < match[1].length; i++) colIdx = colIdx * 26 + (match[1].charCodeAt(i) - 64);
      colIdx -= 1;
      const rowIdx = parseInt(match[2], 10) - 1;
      const val = engine.getDisplayValue(rowIdx, colIdx);
      if (val != null && val !== '') dv[key] = val;
    }
    setDisplayValues(dv);

    return () => { engine.destroy(); };
  }, [cells, rows, cols]);

  useEffect(() => {
    if (!editingCell) return;
    requestAnimationFrame(() => {
      const target = editSourceRef.current === 'bar' ? barInputRef.current : cellInputRef.current;
      if (target) {
        target.focus();
        const len = target.value?.length ?? 0;
        target.selectionStart = len;
        target.selectionEnd = len;
      }
    });
  }, [editingCell]);

  // Focus the container on mount so keyboard shortcuts (arrows, printable keys
  // -> start-editing) work immediately when a data-sheet tab is opened, without
  // requiring the user to click a cell first.
  useEffect(() => {
    containerRef.current?.focus({ preventScroll: true });
  }, []);

  const getColW = useCallback((colIdx) => {
    const key = colLabel(colIdx);
    return colWidths[key] ?? DEFAULT_COL_WIDTH;
  }, [colWidths]);

  const getRowH = useCallback((rowIdx) => {
    return rowHeights[rowIdx] ?? DEFAULT_ROW_HEIGHT;
  }, [rowHeights]);

  const startEditing = useCallback((row, col, initialText, source = 'cell', mode = 'enter') => {
    editSourceRef.current = source;
    editModeRef.current = mode;
    setSelectedCell({ row, col });
    setEditingCell({ row, col });
    setEditText(initialText);
  }, []);

  const cancelEditing = useCallback(() => {
    setEditingCell(null);
    setEditText('');
    requestAnimationFrame(() => {
      containerRef.current?.focus();
    });
  }, []);

  const commitEdit = useCallback((row, col, rawInput) => {
    const key = cellKey(row, col);
    const engine = engineRef.current;

    if (onChange) {
      const updated = { ...cells };
      const existingStyle = updated[key]?.s;
      // Route user input through the format-aware parser so Percentage
      // strips '%', Currency strips the symbol, Text stores verbatim
      // (including any leading '='), and dates normalise to ISO.
      const parsed = parseCellInput(rawInput, existingStyle);

      if (rawInput === '' || rawInput == null) {
        delete updated[key];
        if (engine) engine.setCellValue(row, col, null);
      } else if (parsed.formula) {
        if (engine) engine.setCellValue(row, col, parsed.formula);
        const computed = engine ? engine.getDisplayValue(row, col) : parsed.formula;
        const oldCell = updated[key];
        updated[key] = { ...oldCell, f: parsed.formula, v: computed };
      } else {
        const val = parsed.value;
        if (engine) engine.setCellValue(row, col, val);
        const oldCell = updated[key];
        updated[key] = { ...oldCell, v: val };
        delete updated[key].f;
      }

      const dv = { ...displayValues };
      if (engine) {
        for (const k of Object.keys(updated)) {
          const m = k.match(/^([A-Z]+)(\d+)$/);
          if (!m) continue;
          let ci = 0;
          for (let i = 0; i < m[1].length; i++) ci = ci * 26 + (m[1].charCodeAt(i) - 64);
          ci -= 1;
          const ri = parseInt(m[2], 10) - 1;
          const val = engine.getDisplayValue(ri, ci);
          if (val != null && val !== '') dv[k] = val;
          else delete dv[k];
        }
      }
      setDisplayValues(dv);

      let newRows = rows;
      let newCols = cols;
      if (row >= rows - 3) newRows = Math.max(rows, row + 10);
      if (col >= cols - 2) newCols = Math.max(cols, col + 5);

      onChange({ ...data, cells: updated, rows: newRows, cols: newCols });
    }
    setEditingCell(null);
    setEditText('');
    requestAnimationFrame(() => {
      containerRef.current?.focus();
    });
  }, [cells, onChange, data, rows, cols, displayValues]);

  const insertReference = useCallback((ref) => {
    const input = editSourceRef.current === 'bar' ? barInputRef.current : cellInputRef.current;
    const start = input?.selectionStart ?? editText.length;
    const end = input?.selectionEnd ?? editText.length;
    const newText = editText.slice(0, start) + ref + editText.slice(end);
    const newCursor = start + ref.length;
    setEditText(newText);
    requestAnimationFrame(() => {
      if (input) {
        input.focus();
        input.selectionStart = newCursor;
        input.selectionEnd = newCursor;
      }
    });
  }, [editText]);

  const handleShapesChange = useCallback(
    (nextShapes) => {
      if (!onChange) return;
      onChange({ ...data, shapes: nextShapes });
    },
    [onChange, data],
  );

  const handleSelectShape = useCallback((ids) => {
    setSelectedShapeIds(Array.isArray(ids) ? ids : []);
    if (ids && ids.length > 0) {
      setSelectedCell(null);
      setSelectionEnd(null);
      setEditingShapeTextId(null);
      containerRef.current?.focus({ preventScroll: true });
    }
  }, []);

  const clearShapeSelection = useCallback(() => {
    setSelectedShapeIds([]);
    setEditingShapeTextId(null);
  }, []);

  const handleExitShapeMode = useCallback(() => {
    setShapeMode(null);
  }, []);

  const handleSetShapeMode = useCallback((id) => {
    setShapeMode(id || null);
    if (id) {
      setSelectedShapeIds([]);
      setEditingShapeTextId(null);
    }
  }, []);

  const handleBeginTextEdit = useCallback((id) => {
    setEditingShapeTextId(id);
  }, []);

  const handleCommitShapeText = useCallback(
    (id, text) => {
      setEditingShapeTextId(null);
      const next = shapes.map((s) =>
        s.id === id ? { ...s, text: { ...(s.text || {}), value: text } } : s,
      );
      handleShapesChange(next);
      requestAnimationFrame(() => containerRef.current?.focus({ preventScroll: true }));
    },
    [shapes, handleShapesChange],
  );

  const applyShapeStyle = useCallback(
    (partial) => {
      if (selectedShapeIds.length === 0) return;
      const idSet = new Set(selectedShapeIds);
      const next = shapes.map((s) => {
        if (!idSet.has(s.id)) return s;
        return { ...s, style: mergeShapeStyle(s.style, partial) };
      });
      handleShapesChange(next);
    },
    [selectedShapeIds, shapes, handleShapesChange],
  );

  const applyTextStyle = useCallback(
    (partial) => {
      if (selectedShapeIds.length === 0) return;
      const idSet = new Set(selectedShapeIds);
      const next = shapes.map((s) => {
        if (!idSet.has(s.id)) return s;
        return { ...s, textStyle: mergeTextStyle(s.textStyle, partial) };
      });
      handleShapesChange(next);
    },
    [selectedShapeIds, shapes, handleShapesChange],
  );

  const handleToggleShapeTextFormat = useCallback(
    (key) => {
      if (selectedShapeIds.length === 0) return;
      const target = shapes.find((s) => s.id === selectedShapeIds[0]);
      const current = target?.textStyle?.[key];
      applyTextStyle({ [key]: !current });
    },
    [selectedShapeIds, shapes, applyTextStyle],
  );

  const handleArrange = useCallback(
    (direction) => {
      if (selectedShapeIds.length === 0) return;
      let next = shapes;
      for (const id of selectedShapeIds) {
        next = reorderShape(next, id, direction);
      }
      handleShapesChange(next);
    },
    [selectedShapeIds, shapes, handleShapesChange],
  );

  const handleDeleteSelectedShapes = useCallback(() => {
    if (selectedShapeIds.length === 0) return;
    handleShapesChange(deleteShapes(shapes, selectedShapeIds));
    setSelectedShapeIds([]);
  }, [selectedShapeIds, shapes, handleShapesChange]);

  const handleCellMouseDown = useCallback((e, row, col) => {
    clearShapeSelection();
    if (isFormulaEditing) {
      const clickedKey = cellKey(row, col);
      const editingKey = editingCell ? cellKey(editingCell.row, editingCell.col) : null;
      if (clickedKey !== editingKey) {
        e.preventDefault();
        insertReference(clickedKey);
        return;
      }
    }

    e.preventDefault();

    // Move keyboard focus to the container so onKeyDown fires for direct typing.
    // mousedown+preventDefault suppresses the default focus shift, leaving focus
    // on whatever was focused before (e.g. the tab button in StatusBar).
    containerRef.current?.focus({ preventScroll: true });

    if (editingCell) {
      commitEdit(editingCell.row, editingCell.col, editText);
    }

    const hitMerge = getMergeAt(row, col);
    const anchorRow = hitMerge ? hitMerge.r1 : row;
    const anchorCol = hitMerge ? hitMerge.c1 : col;
    setSelectedCell({ row: anchorRow, col: anchorCol });
    setSelectionEnd(null);
    isDraggingRef.current = true;

    const onMove = (ev) => {
      if (!isDraggingRef.current) return;
      const scrollEl = scrollRef.current;
      if (!scrollEl) return;
      const rect = scrollEl.getBoundingClientRect();
      const x = ev.clientX - rect.left + scrollEl.scrollLeft - ROW_HEADER_WIDTH;
      const y = ev.clientY - rect.top + scrollEl.scrollTop - HEADER_HEIGHT;

      let cumW = 0;
      let endCol = 0;
      for (let ci = 0; ci < cols; ci++) {
        const w = getColW(ci);
        if (x < cumW + w) { endCol = ci; break; }
        cumW += w;
        endCol = ci;
      }
      let cumH = 0;
      let endRow = 0;
      for (let ri = 0; ri < rows; ri++) {
        const h = getRowH(ri);
        if (y < cumH + h) { endRow = ri; break; }
        cumH += h;
        endRow = ri;
      }
      setSelectionEnd({ row: endRow, col: endCol });
    };

    const onUp = () => {
      isDraggingRef.current = false;
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
    };

    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  }, [isFormulaEditing, editingCell, editText, insertReference, commitEdit, cols, rows, getColW, getRowH, getMergeAt, clearShapeSelection]);

  const handleCellClick = useCallback(() => {
    // Selection is handled entirely in handleCellMouseDown now.
  }, []);

  const handleCellDoubleClick = useCallback((row, col) => {
    const hitMerge = getMergeAt(row, col);
    const effRow = hitMerge ? hitMerge.r1 : row;
    const effCol = hitMerge ? hitMerge.c1 : col;
    if (isFormulaEditing) {
      const clickedKey = cellKey(effRow, effCol);
      const editingKey = editingCell ? cellKey(editingCell.row, editingCell.col) : null;
      if (clickedKey !== editingKey) {
        insertReference(clickedKey);
        return;
      }
    }
    setSelectionEnd(null);
    const key = cellKey(effRow, effCol);
    const cd = cells[key];
    startEditing(effRow, effCol, cd?.f || String(cd?.v ?? ''), 'cell', 'edit');
  }, [isFormulaEditing, editingCell, cells, startEditing, insertReference, getMergeAt]);

  const handleCellBlur = useCallback((e) => {
    if (e.relatedTarget && containerRef.current?.contains(e.relatedTarget)) return;
    if (editingCell) {
      commitEdit(editingCell.row, editingCell.col, editText);
    }
  }, [editingCell, editText, commitEdit]);

  const handleEditKeyDown = useCallback((e) => {
    if (!editingCell) return;
    const { row, col } = editingCell;

    if (e.key === 'Enter') {
      e.preventDefault();
      commitEdit(row, col, editText);
      const dir = e.shiftKey ? -1 : 1;
      setSelectedCell({ row: Math.max(0, Math.min(rows - 1, row + dir)), col });
      setSelectionEnd(null);
      return;
    }
    if (e.key === 'Tab') {
      e.preventDefault();
      commitEdit(row, col, editText);
      const nextCol = e.shiftKey ? Math.max(0, col - 1) : Math.min(cols - 1, col + 1);
      setSelectedCell({ row, col: nextCol });
      setSelectionEnd(null);
      return;
    }
    if (e.key === 'Escape') {
      cancelEditing();
      return;
    }

    if (isFormulaEditing) return;

    const input = e.target;
    const cursorPos = input.selectionStart ?? 0;
    const cursorEnd = input.selectionEnd ?? 0;
    const textLen = editText.length;
    const isEnterMode = editModeRef.current === 'enter';

    if (e.key === 'ArrowUp') {
      e.preventDefault();
      commitEdit(row, col, editText);
      setSelectedCell({ row: Math.max(0, row - 1), col });
      setSelectionEnd(null);
      return;
    }
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      commitEdit(row, col, editText);
      setSelectedCell({ row: Math.min(rows - 1, row + 1), col });
      setSelectionEnd(null);
      return;
    }
    if (e.key === 'ArrowLeft') {
      if (isEnterMode || (cursorPos === 0 && cursorPos === cursorEnd)) {
        e.preventDefault();
        commitEdit(row, col, editText);
        setSelectedCell({ row, col: Math.max(0, col - 1) });
        setSelectionEnd(null);
        return;
      }
    }
    if (e.key === 'ArrowRight') {
      if (isEnterMode || (cursorPos === textLen && cursorPos === cursorEnd)) {
        e.preventDefault();
        commitEdit(row, col, editText);
        setSelectedCell({ row, col: Math.min(cols - 1, col + 1) });
        setSelectionEnd(null);
        return;
      }
    }
  }, [editingCell, editText, commitEdit, cancelEditing, rows, cols, isFormulaEditing]);

  const applyStyleToSelected = useCallback((partialStyle) => {
    // Route font toggles + font size + font color + text bg to the
    // selected shape's textStyle when a shape is the active selection;
    // otherwise update cell styles as before. Matches Excel's dual-role
    // behaviour where the format ribbon targets the shape text when a
    // shape is selected.
    if (selectedShapeIds.length > 0) {
      const textPartial = {};
      if (partialStyle.bold !== undefined) textPartial.bold = partialStyle.bold;
      if (partialStyle.italic !== undefined) textPartial.italic = partialStyle.italic;
      if (partialStyle.underline !== undefined) textPartial.underline = partialStyle.underline;
      if (partialStyle.fontSize !== undefined) textPartial.fontSize = partialStyle.fontSize;
      if (partialStyle.color !== undefined) textPartial.fill = { color: partialStyle.color, alpha: 1 };
      if (partialStyle.hAlign !== undefined) textPartial.hAlign = partialStyle.hAlign;
      if (partialStyle.vAlign !== undefined) textPartial.vAlign = partialStyle.vAlign;
      // bg is ignored on shapes (use Shape Fill instead).
      if (Object.keys(textPartial).length === 0) return;
      const idSet = new Set(selectedShapeIds);
      const next = shapes.map((s) => {
        if (!idSet.has(s.id)) return s;
        return { ...s, textStyle: mergeTextStyle(s.textStyle, textPartial) };
      });
      handleShapesChange(next);
      return;
    }

    if (!selRect || !onChange) return;
    const updated = { ...cells };
    for (let r = selRect.r1; r <= selRect.r2; r++) {
      for (let c = selRect.c1; c <= selRect.c2; c++) {
        const key = cellKey(r, c);
        const existing = updated[key] || {};
        updated[key] = { ...existing, s: { ...(existing.s || {}), ...partialStyle } };
      }
    }
    onChange({ ...data, cells: updated });
  }, [selectedShapeIds, shapes, handleShapesChange, selRect, cells, onChange, data]);

  // Apply a number-format partial (numFmt / decimals / currency / ...)
  // across the current selection. Unlike applyStyleToSelected this path
  // is cell-only - shapes ignore number formatting.
  const applyNumberFormatToSelected = useCallback((partial) => {
    if (!partial || !selRect || !onChange) return;
    const updated = { ...cells };
    for (let r = selRect.r1; r <= selRect.r2; r++) {
      for (let c = selRect.c1; c <= selRect.c2; c++) {
        const key = cellKey(r, c);
        const existing = updated[key] || {};
        const nextStyle = { ...(existing.s || {}), ...partial };
        // Clearing back to 'general' discards the other number-format
        // fields so stale decimals / currency / negativeStyle don't
        // leak into downstream Excel exports as dead format entries.
        if (partial.numFmt === 'general' || partial.numFmt === 'text') {
          delete nextStyle.decimals;
          delete nextStyle.currency;
          delete nextStyle.useThousands;
          delete nextStyle.negativeStyle;
        }
        updated[key] = { ...existing, s: nextStyle };
      }
    }
    onChange({ ...data, cells: updated });
  }, [selRect, cells, onChange, data]);

  const bumpDecimalsInSelection = useCallback((delta) => {
    if (!selRect || !onChange) return;
    const updated = { ...cells };
    for (let r = selRect.r1; r <= selRect.r2; r++) {
      for (let c = selRect.c1; c <= selRect.c2; c++) {
        const key = cellKey(r, c);
        const existing = updated[key] || {};
        const partial = bumpDecimalsStyle(existing.s, delta);
        updated[key] = { ...existing, s: { ...(existing.s || {}), ...partial } };
      }
    }
    onChange({ ...data, cells: updated });
  }, [selRect, cells, onChange, data]);

  // Apply a parsed clipboard payload at the given anchor. Handles the three
  // input shapes produced by parseClipboardInput (rich / html / tsv) with a
  // single onChange call so the whole paste becomes one undo step.
  const applyPaste = useCallback((parsed, anchor) => {
    if (!parsed || !onChange) return;

    let R;
    let C;
    let writeCell;
    let pastedMerges = [];
    let pastedColWidths = {};
    let pastedRowHeights = {};

    if (parsed.kind === 'rich') {
      const p = parsed.payload;
      R = p.rows;
      C = p.cols;
      writeCell = (r, c) => {
        const src = p.cells[r]?.[c];
        if (!src) return null;
        const next = {};
        if (src.v !== undefined) next.v = src.v;
        if (src.f !== undefined) next.f = src.f;
        if (src.s !== undefined) next.s = src.s;
        return Object.keys(next).length > 0 ? next : null;
      };
      pastedMerges = (p.merges || []).map((m) => ({
        r1: anchor.row + m.r1,
        c1: anchor.col + m.c1,
        r2: anchor.row + m.r2,
        c2: anchor.col + m.c2,
      }));
      pastedColWidths = p.colWidths || {};
      pastedRowHeights = p.rowHeights || {};
    } else if (parsed.kind === 'html') {
      R = parsed.grid.length;
      C = R > 0 ? parsed.grid[0].length : 0;
      const styleGrid = parsed.cellStyles || [];
      writeCell = (r, c) => {
        const text = parsed.grid[r]?.[c];
        const style = styleGrid[r]?.[c];
        const hasText = text != null && text !== '';
        const hasStyle = style && Object.keys(style).length > 0;
        if (!hasText && !hasStyle) return null;
        const out = {};
        if (hasText) out.v = coerceNumeric(text);
        if (hasStyle) out.s = style;
        return out;
      };
      pastedMerges = (parsed.merges || []).map((m) => ({
        r1: anchor.row + m.r1,
        c1: anchor.col + m.c1,
        r2: anchor.row + m.r2,
        c2: anchor.col + m.c2,
      }));
      pastedColWidths = parsed.colWidths || {};
      pastedRowHeights = parsed.rowHeights || {};
    } else if (parsed.kind === 'tsv') {
      R = parsed.grid.length;
      C = R > 0 ? parsed.grid[0].length : 0;
      writeCell = (r, c) => {
        const text = parsed.grid[r]?.[c];
        if (text == null || text === '') return null;
        return { v: coerceNumeric(text) };
      };
    } else {
      return;
    }

    if (R <= 0 || C <= 0) return;

    const targetRect = {
      r1: anchor.row,
      c1: anchor.col,
      r2: anchor.row + R - 1,
      c2: anchor.col + C - 1,
    };

    // Drop any existing merges that overlap the paste rect. Matches Google
    // Sheets / most spreadsheet clones; avoids Excel's modal-dialog block.
    const preservedMerges = (merges || []).filter((m) => !rectsOverlap(m, targetRect));

    const updated = { ...cells };
    for (let r = 0; r < R; r++) {
      for (let c = 0; c < C; c++) {
        const dstR = anchor.row + r;
        const dstC = anchor.col + c;
        const key = cellKey(dstR, dstC);
        const next = writeCell(r, c);
        if (next == null) {
          // Rich paste: empty source cells overwrite the destination.
          // Plain / html paste: empty cells leave the destination alone so
          //   that pasting "A\t\tC" does not wipe cell B.
          if (parsed.kind === 'rich' && updated[key]) {
            delete updated[key];
          }
          continue;
        }
        if (parsed.kind === 'rich') {
          updated[key] = next;
        } else {
          // HTML paste replaces destination cells that were blank and
          // merges into styled ones; tsv paste does the same. Using a
          // spread preserves any existing attributes on the destination
          // that the source doesn't specify (e.g. untouched borders on
          // neighbour cells).
          const existing = updated[key];
          updated[key] = { ...(existing || {}), ...next };
        }
      }
    }

    // Grow the sheet if the paste spills past the current extents, matching
    // the "auto-expand" rule used by commitEdit.
    let newRows = rows;
    let newCols = cols;
    if (anchor.row + R > rows - 3) newRows = Math.max(rows, anchor.row + R + 10);
    if (anchor.col + C > cols - 2) newCols = Math.max(cols, anchor.col + C + 5);

    const mergedMerges = pastedMerges.length > 0
      ? [...preservedMerges, ...pastedMerges]
      : preservedMerges;

    // Apply column widths / row heights. Paste-time layout lookups are
    // keyed on rect-local offsets (0..N-1); translate to absolute app-side
    // keys (column label / row index) here so the DataGrid renders the
    // imported block with its source sizing.
    let nextColWidths = data.colWidths;
    if (pastedColWidths && Object.keys(pastedColWidths).length > 0) {
      nextColWidths = { ...(data.colWidths || {}) };
      for (const [iStr, w] of Object.entries(pastedColWidths)) {
        const i = parseInt(iStr, 10);
        if (!Number.isFinite(i)) continue;
        const absCol = anchor.col + i;
        if (absCol < 0) continue;
        nextColWidths[colLabelFromIndex(absCol)] = w;
      }
    }
    let nextRowHeights = data.rowHeights;
    if (pastedRowHeights && Object.keys(pastedRowHeights).length > 0) {
      nextRowHeights = { ...(data.rowHeights || {}) };
      for (const [rStr, h] of Object.entries(pastedRowHeights)) {
        const r = parseInt(rStr, 10);
        if (!Number.isFinite(r)) continue;
        const absRow = anchor.row + r;
        if (absRow < 0) continue;
        nextRowHeights[absRow] = h;
      }
    }

    onChange({
      ...data,
      cells: updated,
      merges: mergedMerges,
      rows: newRows,
      cols: newCols,
      colWidths: nextColWidths,
      rowHeights: nextRowHeights,
    });

    // Move selection to cover the pasted region so subsequent Ctrl+V
    // repeats or styling operations apply to the new block.
    setSelectedCell({ row: anchor.row, col: anchor.col });
    setSelectionEnd({ row: anchor.row + R - 1, col: anchor.col + C - 1 });
  }, [onChange, data, cells, merges, rows, cols]);

  const handleCopy = useCallback((e, isCut) => {
    if (editingCell || editingShapeTextId) return;
    const tag = e.target?.tagName;
    if (tag === 'INPUT' || tag === 'TEXTAREA') return;
    if (!selRect || !e.clipboardData) return;

    const { tsv, html } = buildGridClipboard({
      cells,
      merges,
      rect: selRect,
      displayValues,
      colWidths,
      rowHeights,
    });
    e.clipboardData.setData('text/plain', tsv);
    e.clipboardData.setData('text/html', html);
    e.preventDefault();

    if (isCut && onChange) {
      const updated = { ...cells };
      for (let r = selRect.r1; r <= selRect.r2; r++) {
        for (let c = selRect.c1; c <= selRect.c2; c++) {
          const k = cellKey(r, c);
          const existing = updated[k];
          if (!existing) continue;
          const rest = { ...existing };
          delete rest.v;
          delete rest.f;
          if (Object.keys(rest).length === 0) delete updated[k];
          else updated[k] = rest;
        }
      }
      onChange({ ...data, cells: updated });
    }
  }, [editingCell, editingShapeTextId, selRect, cells, merges, displayValues, colWidths, rowHeights, onChange, data]);

  const handlePaste = useCallback((e) => {
    if (editingCell || editingShapeTextId) return;
    const tag = e.target?.tagName;
    if (tag === 'INPUT' || tag === 'TEXTAREA') return;
    if (!selectedCell || !selRect || !e.clipboardData) return;

    const htmlText = e.clipboardData.getData('text/html') || '';
    const plainText = e.clipboardData.getData('text/plain') || '';
    const parsed = parseClipboardInput({ htmlText, plainText });
    if (!parsed) return;
    e.preventDefault();
    applyPaste(parsed, { row: selRect.r1, col: selRect.c1 });
  }, [editingCell, editingShapeTextId, selectedCell, selRect, applyPaste]);

  const handleGridKeyDown = useCallback((e) => {
    if (editingShapeTextId) return;

    if (selectedShapeIds.length > 0) {
      const step = e.shiftKey ? 10 : 1;
      if ((e.ctrlKey || e.metaKey)) {
        if (e.key === 'b' || e.key === 'B') {
          e.preventDefault();
          const cur = shapes.find((s) => s.id === selectedShapeIds[0])?.textStyle?.bold;
          applyStyleToSelected({ bold: !cur });
          return;
        }
        if ((e.key === 'i' || e.key === 'I') && !e.shiftKey) {
          e.preventDefault();
          const cur = shapes.find((s) => s.id === selectedShapeIds[0])?.textStyle?.italic;
          applyStyleToSelected({ italic: !cur });
          return;
        }
        if (e.key === 'u' || e.key === 'U') {
          e.preventDefault();
          const cur = shapes.find((s) => s.id === selectedShapeIds[0])?.textStyle?.underline;
          applyStyleToSelected({ underline: !cur });
          return;
        }
      }
      if (e.key === 'Delete' || e.key === 'Backspace') {
        e.preventDefault();
        handleDeleteSelectedShapes();
        return;
      }
      if (e.key === 'Escape') {
        e.preventDefault();
        clearShapeSelection();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && (e.key === 'd' || e.key === 'D')) {
        e.preventDefault();
        const { shapes: next, newIds } = duplicateShapes(shapes, selectedShapeIds);
        handleShapesChange(next);
        setSelectedShapeIds(newIds);
        return;
      }
      if (e.key === 'ArrowUp')    { e.preventDefault(); handleShapesChange(nudgeShapes(shapes, selectedShapeIds, 0, -step)); return; }
      if (e.key === 'ArrowDown')  { e.preventDefault(); handleShapesChange(nudgeShapes(shapes, selectedShapeIds, 0, step));  return; }
      if (e.key === 'ArrowLeft')  { e.preventDefault(); handleShapesChange(nudgeShapes(shapes, selectedShapeIds, -step, 0)); return; }
      if (e.key === 'ArrowRight') { e.preventDefault(); handleShapesChange(nudgeShapes(shapes, selectedShapeIds, step, 0));  return; }
      if (e.key === 'Enter' || e.key === 'F2') {
        e.preventDefault();
        if (selectedShapeIds.length === 1) setEditingShapeTextId(selectedShapeIds[0]);
        return;
      }
      return;
    }

    if (shapeMode && e.key === 'Escape') {
      e.preventDefault();
      setShapeMode(null);
      return;
    }

    if (!selectedCell) return;
    const { row, col } = selectedCell;

    if ((e.ctrlKey || e.metaKey) && !editingCell) {
      if (e.key === 'b' || e.key === 'B') {
        e.preventDefault();
        applyStyleToSelected({ bold: !(cells[cellKey(row, col)]?.s?.bold) });
        return;
      }
      if (e.key === 'i' && !e.shiftKey) {
        e.preventDefault();
        applyStyleToSelected({ italic: !(cells[cellKey(row, col)]?.s?.italic) });
        return;
      }
      if (e.key === 'u' || e.key === 'U') {
        e.preventDefault();
        applyStyleToSelected({ underline: !(cells[cellKey(row, col)]?.s?.underline) });
        return;
      }

      // Excel number-format shortcuts (Ctrl+Shift+1..6 / ~).
      // We key off e.code so the mapping works regardless of the
      // active keyboard layout's Shift+digit character.
      if (e.shiftKey) {
        let partial = null;
        switch (e.code) {
          case 'Backquote':
            partial = { numFmt: 'general' };
            break;
          case 'Digit1':
            partial = { numFmt: 'number', decimals: 2, useThousands: true, negativeStyle: 'parens' };
            break;
          case 'Digit2':
            partial = { numFmt: 'time' };
            break;
          case 'Digit3':
            partial = { numFmt: 'shortDate' };
            break;
          case 'Digit4':
            partial = { numFmt: 'currency', decimals: 2, currency: cells[cellKey(row, col)]?.s?.currency || '$', negativeStyle: 'parens' };
            break;
          case 'Digit5':
            partial = { numFmt: 'percentage', decimals: 0 };
            break;
          case 'Digit6':
            partial = { numFmt: 'scientific', decimals: 2 };
            break;
          default:
            break;
        }
        if (partial) {
          e.preventDefault();
          applyNumberFormatToSelected(partial);
          return;
        }
      }
    }

    if (editingCell) return;

    const stepCell = (dr, dc) => {
      const cur = getMergeAt(row, col);
      const sr = cur ? cur.r1 : row;
      const sc = cur ? cur.c1 : col;
      const er = cur ? cur.r2 : row;
      const ec = cur ? cur.c2 : col;
      let nr = sr, nc = sc;
      if (dr > 0) nr = er + 1;
      else if (dr < 0) nr = sr - 1;
      if (dc > 0) nc = ec + 1;
      else if (dc < 0) nc = sc - 1;
      nr = Math.max(0, Math.min(rows - 1, nr));
      nc = Math.max(0, Math.min(cols - 1, nc));
      const target = getMergeAt(nr, nc);
      if (target) return { row: target.r1, col: target.c1 };
      return { row: nr, col: nc };
    };

    if (e.key === 'Enter' || e.key === 'F2') {
      e.preventDefault();
      setSelectionEnd(null);
      const key = cellKey(row, col);
      const cd = cells[key];
      startEditing(row, col, cd?.f || String(cd?.v ?? ''), 'cell', 'edit');
      return;
    }
    if (e.key === 'Tab') {
      e.preventDefault();
      const next = stepCell(0, e.shiftKey ? -1 : 1);
      setSelectedCell(next);
      setSelectionEnd(null);
      return;
    }
    if (e.key === 'ArrowUp')    { e.preventDefault(); setSelectedCell(stepCell(-1, 0)); setSelectionEnd(null); return; }
    if (e.key === 'ArrowDown')  { e.preventDefault(); setSelectedCell(stepCell(1, 0));  setSelectionEnd(null); return; }
    if (e.key === 'ArrowLeft')  { e.preventDefault(); setSelectedCell(stepCell(0, -1)); setSelectionEnd(null); return; }
    if (e.key === 'ArrowRight') { e.preventDefault(); setSelectedCell(stepCell(0, 1));  setSelectionEnd(null); return; }
    if (e.key === 'Delete' || e.key === 'Backspace') {
      e.preventDefault();
      if (selRect && onChange) {
        const updated = { ...cells };
        for (let r = selRect.r1; r <= selRect.r2; r++) {
          for (let c = selRect.c1; c <= selRect.c2; c++) {
            const k = cellKey(r, c);
            if (updated[k]) {
              const { v, f, ...rest } = updated[k];
              if (Object.keys(rest).length > 0) updated[k] = rest;
              else delete updated[k];
            }
          }
        }
        onChange({ ...data, cells: updated });
      }
      return;
    }
    if (e.key.length === 1 && !e.ctrlKey && !e.metaKey) {
      e.preventDefault();
      setSelectionEnd(null);
      startEditing(row, col, e.key, 'cell');
    }
  }, [
    selectedCell,
    editingCell,
    rows,
    cols,
    cells,
    startEditing,
    applyStyleToSelected,
    applyNumberFormatToSelected,
    selRect,
    onChange,
    data,
    getMergeAt,
    selectedShapeIds,
    shapes,
    shapeMode,
    editingShapeTextId,
    handleDeleteSelectedShapes,
    clearShapeSelection,
    handleShapesChange,
  ]);

  const handleStartBarEdit = useCallback(() => {
    if (editingCell) {
      editSourceRef.current = 'bar';
      return;
    }
    if (!selectedCell) return;
    const { row, col } = selectedCell;
    const key = cellKey(row, col);
    const cd = cells[key];
    startEditing(row, col, cd?.f || String(cd?.v ?? ''), 'bar', 'edit');
  }, [editingCell, selectedCell, cells, startEditing]);

  const handleResizeCol = useCallback((colIdx, startX, startWidth) => {
    const key = colLabel(colIdx);
    let latestW = startWidth;
    const onMove = (e) => {
      latestW = Math.max(MIN_COL_WIDTH, startWidth + (e.clientX - startX));
      setColWidths((prev) => ({ ...prev, [key]: latestW }));
    };
    const onUp = () => {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
      if (onChange) {
        setColWidths((prev) => {
          const final = { ...prev, [key]: latestW };
          onChange({ ...data, colWidths: final });
          return final;
        });
      }
    };
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  }, [onChange, data]);

  const handleAutoFitCol = useCallback((colIdx) => {
    const colKey = colLabel(colIdx);
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    let maxW = MIN_COL_WIDTH;
    for (let ri = 0; ri < rows; ri++) {
      const key = cellKey(ri, colIdx);
      const cd = cells[key];
      const raw = displayValues[key] ?? cd?.v ?? '';
      const text = formatCellValue(raw, cd?.s);
      if (!text) continue;
      const fontSize = cd?.s?.fontSize || 12;
      const bold = cd?.s?.bold ? '700' : '400';
      ctx.font = `${bold} ${fontSize}px sans-serif`;
      const measured = ctx.measureText(text).width + 16;
      if (measured > maxW) maxW = measured;
    }
    maxW = Math.ceil(Math.min(maxW, 400));
    if (onChange) {
      setColWidths((prev) => {
        const final = { ...prev, [colKey]: maxW };
        onChange({ ...data, colWidths: final });
        return final;
      });
    }
  }, [rows, cells, displayValues, onChange, data]);

  const handleResizeRow = useCallback((rowIdx, startY, startHeight) => {
    let latestH = startHeight;
    const onMove = (e) => {
      latestH = Math.max(MIN_ROW_HEIGHT, startHeight + (e.clientY - startY));
      setRowHeights((prev) => ({ ...prev, [rowIdx]: latestH }));
    };
    const onUp = () => {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
      if (onChange) {
        setRowHeights((prev) => {
          const final = { ...prev, [rowIdx]: latestH };
          onChange({ ...data, rowHeights: final });
          return final;
        });
      }
    };
    document.body.style.cursor = 'row-resize';
    document.body.style.userSelect = 'none';
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  }, [onChange, data]);

  const handleApplyBorderPreset = useCallback((presetId) => {
    if (!selRect || !onChange) return;
    const updated = { ...cells };
    const { r1, r2, c1, c2 } = selRect;
    const isClear = presetId === 'none';
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        const key = cellKey(r, c);
        const existing = updated[key] || {};
        const existingBorders = existing.s?.borders;
        if (isClear) {
          updated[key] = {
            ...existing,
            s: { ...(existing.s || {}), borders: undefined },
          };
          continue;
        }
        const presetBorders = bordersForPreset(presetId, r, c, r1, r2, c1, c2);
        if (!presetBorders) continue;
        const merged = { ...(existingBorders || {}), ...presetBorders };
        updated[key] = {
          ...existing,
          s: { ...(existing.s || {}), borders: merged },
        };
      }
    }
    onChange({ ...data, cells: updated });
  }, [selRect, cells, onChange, data]);

  const handleToggleGridLines = useCallback(() => {
    if (!onChange) return;
    onChange({ ...data, showGridLines: !showGridLines });
  }, [onChange, data, showGridLines]);

  const [pendingMerge, setPendingMerge] = useState(null);

  const applyMerge = useCallback((kind, rect) => {
    if (!onChange || !rect) return;
    let toAdd = [];
    if (kind === 'across') {
      for (let r = rect.r1; r <= rect.r2; r++) {
        if (rect.c2 > rect.c1) {
          toAdd.push({ r1: r, c1: rect.c1, r2: r, c2: rect.c2 });
        }
      }
    } else {
      if (rect.r2 > rect.r1 || rect.c2 > rect.c1) {
        toAdd.push({ r1: rect.r1, c1: rect.c1, r2: rect.r2, c2: rect.c2 });
      }
    }
    if (toAdd.length === 0) return;

    const preserved = merges.filter((m) => !toAdd.some((am) => rectsOverlap(m, am)));

    const updatedCells = { ...cells };
    for (const m of toAdd) {
      for (let r = m.r1; r <= m.r2; r++) {
        for (let c = m.c1; c <= m.c2; c++) {
          if (r === m.r1 && c === m.c1) continue;
          const k = cellKey(r, c);
          const existing = updatedCells[k];
          if (!existing) continue;
          const rest = { ...existing };
          delete rest.v;
          delete rest.f;
          if (Object.keys(rest).length === 0) {
            delete updatedCells[k];
          } else {
            updatedCells[k] = rest;
          }
        }
      }
    }

    if (kind === 'center') {
      for (const m of toAdd) {
        const anchorKey = cellKey(m.r1, m.c1);
        const existing = updatedCells[anchorKey] || {};
        updatedCells[anchorKey] = {
          ...existing,
          s: { ...(existing.s || {}), hAlign: 'center', vAlign: 'middle' },
        };
      }
    }

    onChange({ ...data, cells: updatedCells, merges: [...preserved, ...toAdd] });
  }, [merges, cells, data, onChange]);

  const countExtraNonEmpty = useCallback((kind, rect) => {
    if (!rect) return 0;
    let n = 0;
    const isAnchor = (r, c) => {
      if (kind === 'across') return c === rect.c1;
      return r === rect.r1 && c === rect.c1;
    };
    for (let r = rect.r1; r <= rect.r2; r++) {
      for (let c = rect.c1; c <= rect.c2; c++) {
        if (isAnchor(r, c)) continue;
        const cd = cells[cellKey(r, c)];
        if (cd && ((cd.v !== '' && cd.v != null) || cd.f)) n++;
      }
    }
    return n;
  }, [cells]);

  const handleMerge = useCallback((kind) => {
    if (!selRect) return;
    if (kind !== 'across' && selRect.r1 === selRect.r2 && selRect.c1 === selRect.c2) return;
    if (kind === 'across' && selRect.c1 === selRect.c2) return;

    const extras = countExtraNonEmpty(kind, selRect);
    if (extras > 0) {
      setPendingMerge({ kind, rect: selRect });
      return;
    }
    applyMerge(kind, selRect);
  }, [selRect, countExtraNonEmpty, applyMerge]);

  const handleUnmerge = useCallback(() => {
    if (!selRect || !onChange) return;
    const remaining = merges.filter((m) => !rectsOverlap(m, selRect));
    if (remaining.length === merges.length) return;
    onChange({ ...data, merges: remaining });
  }, [selRect, merges, data, onChange]);

  const confirmPendingMerge = useCallback(() => {
    if (!pendingMerge) return;
    const { kind, rect } = pendingMerge;
    setPendingMerge(null);
    applyMerge(kind, rect);
  }, [pendingMerge, applyMerge]);

  const cancelPendingMerge = useCallback(() => {
    setPendingMerge(null);
  }, []);

  const totalWidth = useMemo(() => {
    return ROW_HEADER_WIDTH + Array.from({ length: cols }, (_, i) => getColW(i)).reduce((a, b) => a + b, 0);
  }, [cols, getColW]);

  const totalHeight = useMemo(() => {
    return HEADER_HEIGHT + Array.from({ length: rows }, (_, i) => getRowH(i)).reduce((a, b) => a + b, 0);
  }, [rows, getRowH]);

  const colOffsets = useMemo(() => {
    const arr = new Array(cols + 1);
    arr[0] = 0;
    for (let i = 0; i < cols; i++) arr[i + 1] = arr[i] + getColW(i);
    return arr;
  }, [cols, getColW]);

  const rowOffsets = useMemo(() => {
    const arr = new Array(rows + 1);
    arr[0] = 0;
    for (let i = 0; i < rows; i++) arr[i + 1] = arr[i] + getRowH(i);
    return arr;
  }, [rows, getRowH]);

  const selectedKey = selectedCell ? cellKey(selectedCell.row, selectedCell.col) : null;
  const selectedCellData = selectedKey ? cells[selectedKey] : null;
  const toolbarDisplayValue = selectedCellData?.f || String(displayValues[selectedKey] ?? selectedCellData?.v ?? '');
  const selectedCellRawValue = selectedKey != null
    ? (displayValues[selectedKey] != null ? displayValues[selectedKey] : selectedCellData?.v)
    : undefined;
  const toolbarCellRef = isMultiSel
    ? `${cellKey(selRect.r1, selRect.c1)}:${cellKey(selRect.r2, selRect.c2)}`
    : selectedKey;

  const primarySelectedShape = useMemo(() => {
    if (selectedShapeIds.length === 0) return null;
    return shapes.find((s) => s.id === selectedShapeIds[0]) || null;
  }, [selectedShapeIds, shapes]);

  // When a shape is selected, present its textStyle to the shared
  // Bold/Italic/Underline/fontSize/font-colour controls. Otherwise fall
  // back to the cell style as usual.
  const toolbarStyle = useMemo(() => {
    if (primarySelectedShape) {
      const ts = primarySelectedShape.textStyle || {};
      return {
        bold: !!ts.bold,
        italic: !!ts.italic,
        underline: !!ts.underline,
        fontSize: ts.fontSize,
        color: ts.fill?.color,
        hAlign: ts.hAlign,
        vAlign: ts.vAlign,
      };
    }
    return selectedCellData?.s;
  }, [primarySelectedShape, selectedCellData]);

  const borderStyle = showGridLines ? '1px solid var(--color-border-subtle)' : '1px solid transparent';

  return (
    <div
      ref={containerRef}
      className="flex flex-col h-full outline-none"
      tabIndex={0}
      onKeyDown={handleGridKeyDown}
      onCopy={(e) => handleCopy(e, false)}
      onCut={(e) => handleCopy(e, true)}
      onPaste={handlePaste}
      style={{ backgroundColor: 'var(--color-bg-primary)' }}
    >
      <SpreadsheetToolbar
        cellRef={toolbarCellRef}
        displayValue={toolbarDisplayValue}
        editText={editText}
        isEditing={editingCell != null}
        onEditTextChange={setEditText}
        onStartBarEdit={handleStartBarEdit}
        onBarKeyDown={handleEditKeyDown}
        barInputRef={barInputRef}
        cellStyle={toolbarStyle}
        selectedCellRawValue={selectedCellRawValue}
        onApplyStyle={applyStyleToSelected}
        onApplyNumberFormat={applyNumberFormatToSelected}
        onBumpDecimals={bumpDecimalsInSelection}
        onApplyBorderPreset={handleApplyBorderPreset}
        onToggleGridLines={handleToggleGridLines}
        showGridLines={showGridLines}
        onMerge={handleMerge}
        onUnmerge={handleUnmerge}
        isMergedSelection={!!mergedSelection}
        shapeMode={shapeMode}
        onSetShapeMode={handleSetShapeMode}
        selectedShape={primarySelectedShape}
        onApplyShapeStyle={applyShapeStyle}
        onApplyTextStyle={applyTextStyle}
        onArrange={handleArrange}
        onDeleteShape={handleDeleteSelectedShapes}
      />

      {/* Grid */}
      <div ref={scrollRef} className="flex-1 overflow-auto">
        <div style={{ minWidth: totalWidth, position: 'relative' }}>
          {/* Column headers */}
          <div className="flex sticky top-0 z-10" style={{ backgroundColor: 'var(--color-bg-secondary)' }}>
            <div
              className="flex-shrink-0 flex items-center justify-center text-[10px] font-semibold"
              style={{
                width: ROW_HEADER_WIDTH,
                height: HEADER_HEIGHT,
                borderBottom: '1px solid var(--color-border)',
                borderRight: '1px solid var(--color-border-subtle)',
                color: 'var(--color-text-muted)',
              }}
            />
            {Array.from({ length: cols }, (_, ci) => {
              const w = getColW(ci);
              const label = colLabel(ci);
              const isSelected = selRect && ci >= selRect.c1 && ci <= selRect.c2;
              return (
                <div
                  key={ci}
                  className="flex-shrink-0 relative flex items-center justify-center text-[10px] font-semibold select-none"
                  style={{
                    width: w,
                    height: HEADER_HEIGHT,
                    borderBottom: '1px solid var(--color-border)',
                    borderRight: '1px solid var(--color-border-subtle)',
                    color: isSelected ? 'var(--color-accent)' : 'var(--color-text-muted)',
                    backgroundColor: isSelected ? 'var(--color-accent-muted)' : 'var(--color-bg-secondary)',
                  }}
                >
                  {label}
                  <div
                    onMouseDown={(e) => { e.preventDefault(); handleResizeCol(ci, e.clientX, w); }}
                    onDoubleClick={(e) => { e.preventDefault(); e.stopPropagation(); handleAutoFitCol(ci); }}
                    className="absolute right-0 top-0 h-full w-[4px] cursor-col-resize z-10"
                  />
                </div>
              );
            })}
          </div>

          {/* Rows */}
          {Array.from({ length: rows }, (_, ri) => {
            const rh = getRowH(ri);
            const isRowSelected = selRect && ri >= selRect.r1 && ri <= selRect.r2;
            return (
              <div key={ri} className="flex" style={{ height: rh }}>
                <div
                  className="flex-shrink-0 relative flex items-center justify-center text-[10px] font-medium select-none sticky left-0 z-[5]"
                  style={{
                    width: ROW_HEADER_WIDTH,
                    borderBottom: borderStyle,
                    borderRight: '1px solid var(--color-border-subtle)',
                    color: isRowSelected ? 'var(--color-accent)' : 'var(--color-text-muted)',
                    backgroundColor: isRowSelected ? 'var(--color-accent-muted)' : 'var(--color-bg-secondary)',
                  }}
                >
                  {ri + 1}
                  <div
                    onMouseDown={(e) => { e.preventDefault(); handleResizeRow(ri, e.clientY, rh); }}
                    className="absolute bottom-0 left-0 w-full h-[3px] cursor-row-resize z-10"
                  />
                </div>
                {Array.from({ length: cols }, (_, ci) => {
                  const mergeKey = `${ri},${ci}`;
                  const isMergeAnchor = mergeLookup.anchorMap.has(mergeKey);
                  const isMergeCovered = mergeLookup.coveredMap.has(mergeKey);
                  if (isMergeAnchor || isMergeCovered) {
                    return (
                      <div
                        key={ci}
                        className="flex-shrink-0"
                        style={{ width: getColW(ci), height: rh }}
                      />
                    );
                  }
                  const key = cellKey(ri, ci);
                  const cellData = cells[key];
                  const dv = displayValues[key] ?? cellData?.v ?? '';
                  const displayStr = formatCellValue(dv, cellData?.s);
                  const w = getColW(ci);
                  const isAnchor = selectedCell && selectedCell.row === ri && selectedCell.col === ci;
                  const inRange = selRect && ri >= selRect.r1 && ri <= selRect.r2 && ci >= selRect.c1 && ci <= selRect.c2;
                  const isEditing = editingCell && editingCell.row === ri && editingCell.col === ci;
                  const cellCss = cellStyleToCss(cellData?.s);
                  const cellBorders = cellData?.s?.borders;
                  const highlightColor = refHighlights[key];

                  const aboveBottom = ri > 0 ? cells[cellKey(ri - 1, ci)]?.s?.borders?.bottom : null;
                  const leftRight = ci > 0 ? cells[cellKey(ri, ci - 1)]?.s?.borders?.right : null;

                  const effBottom = borderSideToCss(cellBorders?.bottom) || borderStyle;
                  const effRight = borderSideToCss(cellBorders?.right) || borderStyle;
                  const effTop = cellBorders?.top && !aboveBottom ? borderSideToCss(cellBorders.top) : undefined;
                  const effLeft = cellBorders?.left && !leftRight ? borderSideToCss(cellBorders.left) : undefined;

                  let outlineStyle = 'none';
                  let outlineColor;
                  if (isEditing) {
                    outlineStyle = '2px solid var(--color-accent)';
                  } else if (highlightColor) {
                    outlineStyle = `2px solid ${highlightColor}`;
                    outlineColor = highlightColor;
                  } else if (isAnchor && !isMultiSel) {
                    outlineStyle = '2px solid var(--color-accent)';
                  }

                  return (
                    <div
                      key={ci}
                      className="flex-shrink-0 relative"
                      style={{
                        width: w,
                        height: rh,
                        borderBottom: effBottom,
                        borderRight: effRight,
                        borderTop: effTop,
                        borderLeft: effLeft,
                      }}
                      onMouseDown={(e) => handleCellMouseDown(e, ri, ci)}
                      onClick={() => handleCellClick(ri, ci)}
                      onDoubleClick={() => handleCellDoubleClick(ri, ci)}
                    >
                      {isEditing ? (
                        <input
                          ref={cellInputRef}
                          type="text"
                          value={editText}
                          onChange={(e) => setEditText(e.target.value)}
                          onBlur={handleCellBlur}
                          onKeyDown={handleEditKeyDown}
                          onFocus={() => { editSourceRef.current = 'cell'; }}
                          className="w-full h-full px-1 text-[12px] outline-none"
                          style={{
                            backgroundColor: 'var(--color-bg-primary)',
                            color: 'var(--color-text-primary)',
                            border: '2px solid var(--color-accent)',
                          }}
                        />
                      ) : (
                        <div
                          className="w-full h-full px-1 flex items-center text-[12px] truncate select-none"
                          style={{
                            color: 'var(--color-text-primary)',
                            outline: outlineStyle,
                            outlineOffset: -2,
                            ...cellCss,
                          }}
                        >
                          {highlightColor && (
                            <div
                              className="absolute inset-0 pointer-events-none"
                              style={{
                                backgroundColor: outlineColor,
                                opacity: 0.06,
                              }}
                            />
                          )}
                          {inRange && isMultiSel && !isAnchor && (
                            <div
                              className="absolute inset-0 pointer-events-none z-[1]"
                              style={{ backgroundColor: 'var(--color-accent)', opacity: 0.10 }}
                            />
                          )}
                          {inRange && isMultiSel && (
                            <div
                              className="absolute inset-0 pointer-events-none z-[3]"
                              style={{
                                borderTop: ri === selRect.r1 ? '2px solid var(--color-accent)' : 'none',
                                borderBottom: ri === selRect.r2 ? '2px solid var(--color-accent)' : 'none',
                                borderLeft: ci === selRect.c1 ? '2px solid var(--color-accent)' : 'none',
                                borderRight: ci === selRect.c2 ? '2px solid var(--color-accent)' : 'none',
                              }}
                            />
                          )}
                          {displayStr}
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            );
          })}

          {/* Merge overlay layer: one absolute-positioned div per merge range */}
          {merges.map((m) => {
            const left = ROW_HEADER_WIDTH + (colOffsets[m.c1] || 0);
            const top = HEADER_HEIGHT + (rowOffsets[m.r1] || 0);
            const width = (colOffsets[m.c2 + 1] || 0) - (colOffsets[m.c1] || 0);
            const height = (rowOffsets[m.r2 + 1] || 0) - (rowOffsets[m.r1] || 0);

            const key = cellKey(m.r1, m.c1);
            const cellData = cells[key];
            const dv = displayValues[key] ?? cellData?.v ?? '';
            const displayStr = formatCellValue(dv, cellData?.s);
            const cellCss = cellStyleToCss(cellData?.s);
            const aB = cellData?.s?.borders;
            const cornerTR = cells[cellKey(m.r1, m.c2)];
            const cornerBL = cells[cellKey(m.r2, m.c1)];
            const cornerBR = cells[cellKey(m.r2, m.c2)];

            const effTop = borderSideToCss(aB?.top) || borderStyle;
            const effLeft = borderSideToCss(aB?.left) || borderStyle;
            const effRight =
              borderSideToCss(aB?.right) ||
              borderSideToCss(cornerTR?.s?.borders?.right) ||
              borderSideToCss(cornerBR?.s?.borders?.right) ||
              borderStyle;
            const effBottom =
              borderSideToCss(aB?.bottom) ||
              borderSideToCss(cornerBL?.s?.borders?.bottom) ||
              borderSideToCss(cornerBR?.s?.borders?.bottom) ||
              borderStyle;

            const isAnchor = selectedCell && selectedCell.row === m.r1 && selectedCell.col === m.c1;
            const isEditing = editingCell && editingCell.row === m.r1 && editingCell.col === m.c1;
            const inRange = selRect && rectsOverlap(selRect, m);
            const isExactMergeSel = selRect && rectEquals(selRect, m);
            const highlightColor = refHighlights[key];

            let outlineStyle = 'none';
            let outlineColor;
            if (isEditing) {
              outlineStyle = '2px solid var(--color-accent)';
            } else if (highlightColor) {
              outlineStyle = `2px solid ${highlightColor}`;
              outlineColor = highlightColor;
            } else if (isExactMergeSel || (isAnchor && !isMultiSel)) {
              outlineStyle = '2px solid var(--color-accent)';
            }

            return (
              <div
                key={`merge-${m.r1}-${m.c1}-${m.r2}-${m.c2}`}
                className="absolute"
                style={{
                  left,
                  top,
                  width,
                  height,
                  zIndex: 1,
                  borderTop: effTop,
                  borderLeft: effLeft,
                  borderRight: effRight,
                  borderBottom: effBottom,
                  backgroundColor: 'var(--color-bg-primary)',
                }}
                onMouseDown={(e) => handleCellMouseDown(e, m.r1, m.c1)}
                onDoubleClick={() => handleCellDoubleClick(m.r1, m.c1)}
              >
                {isEditing ? (
                  <input
                    ref={cellInputRef}
                    type="text"
                    value={editText}
                    onChange={(e) => setEditText(e.target.value)}
                    onBlur={handleCellBlur}
                    onKeyDown={handleEditKeyDown}
                    onFocus={() => { editSourceRef.current = 'cell'; }}
                    className="w-full h-full px-1 text-[12px] outline-none"
                    style={{
                      backgroundColor: 'var(--color-bg-primary)',
                      color: 'var(--color-text-primary)',
                      border: '2px solid var(--color-accent)',
                    }}
                  />
                ) : (
                  <div
                    className="w-full h-full px-1 flex items-center text-[12px] truncate select-none relative"
                    style={{
                      color: 'var(--color-text-primary)',
                      outline: outlineStyle,
                      outlineOffset: -2,
                      ...cellCss,
                    }}
                  >
                    {highlightColor && (
                      <div
                        className="absolute inset-0 pointer-events-none"
                        style={{ backgroundColor: outlineColor, opacity: 0.06 }}
                      />
                    )}
                    {inRange && isMultiSel && !isExactMergeSel && (
                      <div
                        className="absolute inset-0 pointer-events-none z-[1]"
                        style={{ backgroundColor: 'var(--color-accent)', opacity: 0.10 }}
                      />
                    )}
                    {inRange && isMultiSel && !isExactMergeSel && (
                      <div
                        className="absolute inset-0 pointer-events-none z-[3]"
                        style={{
                          borderTop: m.r1 === selRect.r1 ? '2px solid var(--color-accent)' : 'none',
                          borderBottom: m.r2 === selRect.r2 ? '2px solid var(--color-accent)' : 'none',
                          borderLeft: m.c1 === selRect.c1 ? '2px solid var(--color-accent)' : 'none',
                          borderRight: m.c2 === selRect.c2 ? '2px solid var(--color-accent)' : 'none',
                        }}
                      />
                    )}
                    {displayStr}
                  </div>
                )}
              </div>
            );
          })}

          <ShapeLayer
            shapes={shapes}
            selectedIds={selectedShapeIds}
            shapeMode={shapeMode}
            colOffsets={colOffsets}
            rowOffsets={rowOffsets}
            rowHeaderWidth={ROW_HEADER_WIDTH}
            headerHeight={HEADER_HEIGHT}
            totalWidth={totalWidth}
            totalHeight={totalHeight}
            onChange={handleShapesChange}
            onSelect={handleSelectShape}
            onExitShapeMode={handleExitShapeMode}
            editingTextId={editingShapeTextId}
            onBeginTextEdit={handleBeginTextEdit}
            onCommitText={handleCommitShapeText}
            onToggleTextFormat={handleToggleShapeTextFormat}
          />
        </div>
      </div>

      {pendingMerge && (
        <MergeConfirmDialog
          onConfirm={confirmPendingMerge}
          onCancel={cancelPendingMerge}
        />
      )}
    </div>
  );
}

function MergeConfirmDialog({ onConfirm, onCancel }) {
  return (
    <div
      className="absolute inset-0 z-[9999] flex items-center justify-center"
      style={{ backgroundColor: 'rgba(0, 0, 0, 0.35)', backdropFilter: 'blur(2px)' }}
      onMouseDown={(e) => { e.stopPropagation(); onCancel(); }}
    >
      <div
        className="rounded-lg shadow-xl max-w-sm"
        style={{
          backgroundColor: 'var(--color-bg-secondary)',
          border: '1px solid var(--color-border)',
          width: 360,
        }}
        onMouseDown={(e) => e.stopPropagation()}
      >
        <div className="px-4 py-3" style={{ borderBottom: '1px solid var(--color-border-subtle)' }}>
          <div className="text-[13px] font-semibold" style={{ color: 'var(--color-text-primary)' }}>
            Merge cells
          </div>
        </div>
        <div className="px-4 py-3 text-[12px] leading-relaxed" style={{ color: 'var(--color-text-secondary)' }}>
          Merging cells only keeps the upper-left cell value, and discards other values. Continue?
        </div>
        <div className="flex justify-end gap-2 px-4 py-3" style={{ borderTop: '1px solid var(--color-border-subtle)' }}>
          <button
            onClick={onCancel}
            onMouseDown={(e) => e.preventDefault()}
            className="px-3 py-1 rounded text-[12px] font-medium cursor-pointer transition-colors"
            style={{
              backgroundColor: 'transparent',
              color: 'var(--color-text-secondary)',
              border: '1px solid var(--color-border)',
            }}
            onMouseEnter={(e) => { e.currentTarget.style.backgroundColor = 'var(--color-bg-hover)'; }}
            onMouseLeave={(e) => { e.currentTarget.style.backgroundColor = 'transparent'; }}
          >
            Cancel
          </button>
          <button
            onClick={onConfirm}
            onMouseDown={(e) => e.preventDefault()}
            className="px-3 py-1 rounded text-[12px] font-medium cursor-pointer transition-colors"
            style={{
              backgroundColor: 'var(--color-accent)',
              color: 'var(--color-on-accent)',
              border: '1px solid var(--color-accent)',
            }}
            onMouseEnter={(e) => { e.currentTarget.style.backgroundColor = 'var(--color-accent-hover)'; }}
            onMouseLeave={(e) => { e.currentTarget.style.backgroundColor = 'var(--color-accent)'; }}
          >
            OK
          </button>
        </div>
      </div>
    </div>
  );
}
