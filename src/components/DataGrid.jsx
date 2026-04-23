import { useState, useRef, useCallback, useEffect, useMemo } from 'react';
import { createFormulaEngine } from '../utils/FormulaEngine';
import SpreadsheetToolbar from './SpreadsheetToolbar';

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

function cellStyleToCss(s) {
  if (!s) return {};
  const css = {};
  if (s.bold) css.fontWeight = 700;
  if (s.italic) css.fontStyle = 'italic';
  if (s.underline) css.textDecoration = 'underline';
  if (s.fontSize) css.fontSize = s.fontSize;
  if (s.color) css.color = s.color;
  if (s.bg) css.backgroundColor = s.bg;
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

  const [editingCell, setEditingCell] = useState(null);
  const [selectedCell, setSelectedCell] = useState(null);
  const [selectionEnd, setSelectionEnd] = useState(null);
  const [editText, setEditText] = useState('');
  const [colWidths, setColWidths] = useState(gridColWidths);
  const [rowHeights, setRowHeights] = useState(gridRowHeights);

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

  const selRect = useMemo(() => {
    if (!selectedCell) return null;
    const end = selectionEnd || selectedCell;
    return {
      r1: Math.min(selectedCell.row, end.row),
      r2: Math.max(selectedCell.row, end.row),
      c1: Math.min(selectedCell.col, end.col),
      c2: Math.max(selectedCell.col, end.col),
    };
  }, [selectedCell, selectionEnd]);

  const isMultiSel = selRect && (selRect.r1 !== selRect.r2 || selRect.c1 !== selRect.c2);

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
      const isFormula = typeof rawInput === 'string' && rawInput.startsWith('=');

      if (rawInput === '' || rawInput == null) {
        delete updated[key];
        if (engine) engine.setCellValue(row, col, null);
      } else if (isFormula) {
        if (engine) engine.setCellValue(row, col, rawInput);
        const computed = engine ? engine.getDisplayValue(row, col) : rawInput;
        const oldCell = updated[key];
        updated[key] = { ...oldCell, f: rawInput, v: computed };
      } else {
        const numVal = Number(rawInput);
        const val = rawInput === '' ? '' : (Number.isFinite(numVal) ? numVal : rawInput);
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

  const handleCellMouseDown = useCallback((e, row, col) => {
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

    setSelectedCell({ row, col });
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
  }, [isFormulaEditing, editingCell, editText, insertReference, commitEdit, cols, rows, getColW]);

  const handleCellClick = useCallback(() => {
    // Selection is handled entirely in handleCellMouseDown now.
  }, []);

  const handleCellDoubleClick = useCallback((row, col) => {
    if (isFormulaEditing) {
      const clickedKey = cellKey(row, col);
      const editingKey = editingCell ? cellKey(editingCell.row, editingCell.col) : null;
      if (clickedKey !== editingKey) {
        insertReference(clickedKey);
        return;
      }
    }
    setSelectionEnd(null);
    const key = cellKey(row, col);
    const cd = cells[key];
    startEditing(row, col, cd?.f || String(cd?.v ?? ''), 'cell', 'edit');
  }, [isFormulaEditing, editingCell, cells, startEditing, insertReference]);

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
  }, [selRect, cells, onChange, data]);

  const handleGridKeyDown = useCallback((e) => {
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
    }

    if (editingCell) return;

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
      const nextCol = e.shiftKey ? Math.max(0, col - 1) : Math.min(cols - 1, col + 1);
      setSelectedCell({ row, col: nextCol });
      setSelectionEnd(null);
      return;
    }
    if (e.key === 'ArrowUp') { e.preventDefault(); setSelectedCell({ row: Math.max(0, row - 1), col }); setSelectionEnd(null); return; }
    if (e.key === 'ArrowDown') { e.preventDefault(); setSelectedCell({ row: Math.min(rows - 1, row + 1), col }); setSelectionEnd(null); return; }
    if (e.key === 'ArrowLeft') { e.preventDefault(); setSelectedCell({ row, col: Math.max(0, col - 1) }); setSelectionEnd(null); return; }
    if (e.key === 'ArrowRight') { e.preventDefault(); setSelectedCell({ row, col: Math.min(cols - 1, col + 1) }); setSelectionEnd(null); return; }
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
    selRect,
    onChange,
    data,
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
      const text = String(displayValues[key] ?? cd?.v ?? '');
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

  const totalWidth = useMemo(() => {
    return ROW_HEADER_WIDTH + Array.from({ length: cols }, (_, i) => getColW(i)).reduce((a, b) => a + b, 0);
  }, [cols, getColW]);

  const selectedKey = selectedCell ? cellKey(selectedCell.row, selectedCell.col) : null;
  const selectedCellData = selectedKey ? cells[selectedKey] : null;
  const toolbarDisplayValue = selectedCellData?.f || String(displayValues[selectedKey] ?? selectedCellData?.v ?? '');
  const toolbarCellRef = isMultiSel
    ? `${cellKey(selRect.r1, selRect.c1)}:${cellKey(selRect.r2, selRect.c2)}`
    : selectedKey;

  const borderStyle = showGridLines ? '1px solid var(--color-border-subtle)' : '1px solid transparent';

  return (
    <div
      ref={containerRef}
      className="flex flex-col h-full outline-none"
      tabIndex={0}
      onKeyDown={handleGridKeyDown}
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
        cellStyle={selectedCellData?.s}
        onApplyStyle={applyStyleToSelected}
        onApplyBorderPreset={handleApplyBorderPreset}
        onToggleGridLines={handleToggleGridLines}
        showGridLines={showGridLines}
      />

      {/* Grid */}
      <div ref={scrollRef} className="flex-1 overflow-auto">
        <div style={{ minWidth: totalWidth }}>
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
                  const key = cellKey(ri, ci);
                  const cellData = cells[key];
                  const dv = displayValues[key] ?? cellData?.v ?? '';
                  const displayStr = dv != null ? String(dv) : '';
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
        </div>
      </div>
    </div>
  );
}
