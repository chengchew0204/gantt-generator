import { useState, useRef, useEffect, useCallback } from 'react';
import {
  Table2,
  Plus,
  Trash2,
  ChevronRight,
  ChevronDown,
  Columns3,
  GripVertical,
} from 'lucide-react';

const ROW_HEIGHT = 32;
const HEADER_HEIGHT = 56;
const MIN_COL_WIDTH = 32;

const STATUS_OPTIONS = [
  'Not Started',
  'In Progress',
  'Completed',
  'Blocked',
  'Delayed',
  'Pending',
];

const DEFAULT_COL_WIDTHS = {
  id: 32,
  name: 140,
  dependency: 70,
  category: 70,
  startDate: 88,
  endDate: 88,
  duration: 56,
  progress: 80,
  status: 76,
  owner: 70,
  remarks: 80,
  baselineStart: 88,
  baselineEnd: 88,
  parentId: 56,
};

function getColW(widths, key) {
  return widths[key] ?? DEFAULT_COL_WIDTHS[key] ?? 70;
}

export default function DataTable({
  tasks,
  allTasks,
  columns,
  visibleColumns,
  onToggleColumn,
  collapsedParents,
  onToggleCollapse,
  onUpdateTask,
  onAddTask,
  onDeleteTask,
  onReorderTask,
  scrollTop,
  onScroll,
  selectedTaskId,
  onSelectTask,
  headerHeight: headerHeightProp,
  datePickField,
  onDatePickField,
  onBeginDrag,
  onEndDrag,
}) {
  const headerH = headerHeightProp || HEADER_HEIGHT;
  const [showColumnPicker, setShowColumnPicker] = useState(false);
  const [hScrollLeft, setHScrollLeft] = useState(0);
  const [colWidths, setColWidths] = useState(DEFAULT_COL_WIDTHS);
  const [dragIndex, setDragIndex] = useState(null);
  const [dropIndex, setDropIndex] = useState(null);
  const [editingCell, setEditingCell] = useState(null);
  const dragRef = useRef(null);
  const dropRef = useRef(null);
  const pickerRef = useRef(null);
  const scrollRef = useRef(null);
  const suppressScroll = useRef(false);

  const cols = columns.filter((c) => visibleColumns.has(c.key));

  const navigateCell = useCallback((rowIndex, colKey, direction) => {
    const editableCols = cols.filter((c) => c.key !== 'id');
    const colIdx = editableCols.findIndex((c) => c.key === colKey);

    let nextRow = rowIndex;
    let nextColIdx = colIdx;

    const step = () => {
      if (direction === 'up') nextRow--;
      else if (direction === 'down') nextRow++;
      else if (direction === 'left') nextColIdx--;
      else if (direction === 'right') nextColIdx++;
    };

    const isCellEditable = (r, c) => {
      if (r < 0 || r >= tasks.length) return false;
      if (c < 0 || c >= editableCols.length) return false;
      const ck = editableCols[c].key;
      if (ck === 'status') return false;
      const t = tasks[r];
      if (t.isParent && (ck === 'startDate' || ck === 'endDate' || ck === 'progress' || ck === 'duration')) return false;
      return true;
    };

    const maxSteps = Math.max(tasks.length, editableCols.length);
    for (let i = 0; i < maxSteps; i++) {
      step();
      if (nextRow < 0 || nextRow >= tasks.length || nextColIdx < 0 || nextColIdx >= editableCols.length) {
        setEditingCell(null);
        return;
      }
      if (isCellEditable(nextRow, nextColIdx)) {
        setEditingCell({ rowIndex: nextRow, colKey: editableCols[nextColIdx].key });
        if (onSelectTask) onSelectTask(String(tasks[nextRow].id));
        return;
      }
    }
    setEditingCell(null);
  }, [cols, tasks, onSelectTask]);

  useEffect(() => {
    if (!showColumnPicker) return;
    const handler = (e) => {
      if (pickerRef.current && !pickerRef.current.contains(e.target)) {
        setShowColumnPicker(false);
      }
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [showColumnPicker]);

  useEffect(() => {
    if (scrollRef.current && scrollTop != null) {
      if (Math.abs(scrollRef.current.scrollTop - scrollTop) > 1) {
        suppressScroll.current = true;
        scrollRef.current.scrollTop = scrollTop;
      }
    }
  }, [scrollTop]);

  const handleScroll = useCallback(() => {
    if (scrollRef.current) {
      setHScrollLeft(scrollRef.current.scrollLeft);
    }
    if (suppressScroll.current) {
      suppressScroll.current = false;
      return;
    }
    if (scrollRef.current && onScroll) {
      onScroll(scrollRef.current.scrollTop);
    }
  }, [onScroll]);

  const handleResizeCol = useCallback((colKey, startX, startWidth) => {
    const onMove = (e) => {
      const newW = Math.max(MIN_COL_WIDTH, startWidth + (e.clientX - startX));
      setColWidths((prev) => ({ ...prev, [colKey]: newW }));
    };
    const onUp = () => {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
    };
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  }, []);

  const handleDragStart = useCallback((index, e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragIndex(index);
    setDropIndex(index);
    dragRef.current = index;
    dropRef.current = index;

    const len = tasks.length;

    const onMove = (ev) => {
      if (!scrollRef.current) return;
      const rect = scrollRef.current.getBoundingClientRect();
      const y = ev.clientY - rect.top + scrollRef.current.scrollTop;
      const newDrop = Math.max(0, Math.min(len, Math.round(y / ROW_HEIGHT)));
      dropRef.current = newDrop;
      setDropIndex(newDrop);
    };

    const onUp = () => {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
      const from = dragRef.current;
      const drop = dropRef.current;
      if (from != null && drop != null && from !== drop && onReorderTask) {
        const to = drop > from ? drop - 1 : drop;
        if (to !== from) onReorderTask(from, to);
      }
      dragRef.current = null;
      dropRef.current = null;
      setDragIndex(null);
      setDropIndex(null);
    };

    document.body.style.cursor = 'grabbing';
    document.body.style.userSelect = 'none';
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  }, [tasks.length, onReorderTask]);

  if (tasks.length === 0 && (!allTasks || allTasks.length === 0)) {
    return (
      <div className="flex flex-col h-full">
        <div className="flex-1 flex flex-col items-center justify-center gap-3 px-6">
          <div
            className="flex items-center justify-center w-12 h-12 rounded-xl"
            style={{ backgroundColor: 'var(--color-accent-muted)' }}
          >
            <Table2 size={22} style={{ color: 'var(--color-accent)' }} />
          </div>
          <p className="text-sm font-medium" style={{ color: 'var(--color-text-secondary)' }}>
            No tasks loaded
          </p>
          <p className="text-xs text-center max-w-[200px]" style={{ color: 'var(--color-text-muted)' }}>
            Import an Excel file or download a template to get started.
          </p>
        </div>
        <AddRowBar onAddTask={onAddTask} />
      </div>
    );
  }

  const totalMinWidth = 6 + 32 + cols.reduce((sum, c) => sum + getColW(colWidths, c.key), 0);

  return (
    <div className="flex flex-col h-full">
      {/* Header */}
      <div
        className="flex-shrink-0 overflow-hidden"
        style={{
          height: headerH,
          backgroundColor: 'var(--color-bg-secondary)',
          borderBottom: '1px solid var(--color-border)',
        }}
      >
        <div
          className="flex items-center text-xs font-medium whitespace-nowrap h-full"
          style={{ minWidth: totalMinWidth, color: 'var(--color-text-muted)', transform: `translateX(${-hScrollLeft}px)` }}
        >
          <div className="w-6 flex-shrink-0 px-1" />
          {cols.map((col) => (
            <ResizableHeaderCell
              key={col.key}
              col={col}
              width={getColW(colWidths, col.key)}
              onResizeStart={handleResizeCol}
            />
          ))}
          <div className="flex-1" />
          <div className="w-8 flex-shrink-0 px-1 relative" ref={pickerRef}>
            <button
              onClick={() => setShowColumnPicker((v) => !v)}
              className="p-0.5 rounded cursor-pointer transition-colors"
              style={{ color: 'var(--color-text-muted)' }}
              title="Toggle columns"
            >
              <Columns3 size={12} />
            </button>
            {showColumnPicker && (
              <ColumnPicker columns={columns} visibleColumns={visibleColumns} onToggle={onToggleColumn} />
            )}
          </div>
        </div>
      </div>

      {/* Body */}
      <div ref={scrollRef} className="flex-1 overflow-auto" onScroll={handleScroll}
        onClick={(e) => { if (e.target === e.currentTarget && onSelectTask) onSelectTask(null); }}>
        <div style={{ minWidth: totalMinWidth, position: 'relative' }}
          onClick={(e) => { if (e.target === e.currentTarget && onSelectTask) onSelectTask(null); }}>
          {tasks.map((task, i) => (
            <TaskRow
              key={task.id}
              task={task}
              index={i}
              cols={cols}
              colWidths={colWidths}
              isParent={!!task.isParent}
              isCollapsed={collapsedParents.has(String(task.id))}
              hasParent={!!task.parentId}
              onToggleCollapse={onToggleCollapse}
              onUpdateTask={onUpdateTask}
              onDeleteTask={onDeleteTask}
              selected={String(task.id) === String(selectedTaskId)}
              onSelect={onSelectTask}
              datePickField={datePickField}
              onDatePickField={onDatePickField}
              onDragStart={handleDragStart}
              isDragging={dragIndex === i}
              editingCell={editingCell}
              onSetEditingCell={setEditingCell}
              onNavigateCell={navigateCell}
            />
          ))}
          {dragIndex != null && dropIndex != null && dropIndex !== dragIndex && dropIndex !== dragIndex + 1 && (
            <div style={{
              position: 'absolute',
              left: 0,
              right: 0,
              top: dropIndex * ROW_HEIGHT - 1,
              height: 2,
              backgroundColor: 'var(--color-accent)',
              zIndex: 20,
              pointerEvents: 'none',
            }} />
          )}
        </div>
      </div>
      <AddRowBar onAddTask={onAddTask} />
    </div>
  );
}

function ResizableHeaderCell({ col, width, onResizeStart }) {
  const handleMouseDown = (e) => {
    e.preventDefault();
    e.stopPropagation();
    onResizeStart(col.key, e.clientX, width);
  };

  return (
    <div
      className="flex-shrink-0 relative flex items-center"
      style={{ width, maxWidth: width }}
    >
      <span className="px-2 truncate flex-1">{col.label}</span>
      <div
        onMouseDown={handleMouseDown}
        className="absolute right-0 top-0 h-full w-[5px] cursor-col-resize z-10"
        style={{ borderRight: '1px solid var(--color-border-subtle)' }}
      />
    </div>
  );
}

function TaskRow({
  task,
  index,
  cols,
  colWidths,
  isParent,
  isCollapsed,
  hasParent,
  onToggleCollapse,
  onUpdateTask,
  onDeleteTask,
  selected,
  onSelect,
  datePickField,
  onDatePickField,
  onDragStart,
  isDragging,
  editingCell,
  onSetEditingCell,
  onNavigateCell,
}) {
  const [hovering, setHovering] = useState(false);

  const bgColor = selected
    ? 'var(--color-accent-muted)'
    : hovering
      ? 'var(--color-bg-hover)'
      : 'transparent';

  return (
    <div
      className="flex items-center transition-colors w-full"
      style={{
        height: ROW_HEIGHT,
        minHeight: ROW_HEIGHT,
        maxHeight: ROW_HEIGHT,
        borderBottom: '1px solid var(--color-border-subtle)',
        backgroundColor: bgColor,
        borderLeft: selected ? '2px solid var(--color-accent)' : '2px solid transparent',
        opacity: isDragging ? 0.4 : 1,
      }}
      onMouseEnter={() => setHovering(true)}
      onMouseLeave={() => setHovering(false)}
      onClick={() => { if (onSelect) onSelect(String(task.id)); }}
    >
      <div className="w-6 flex-shrink-0 px-1 flex items-center justify-center">
        {isParent ? (
          <button
            onClick={(e) => { e.stopPropagation(); onToggleCollapse(String(task.id)); }}
            className="p-0.5 cursor-pointer rounded transition-colors"
            style={{ color: 'var(--color-text-muted)' }}
          >
            {isCollapsed ? <ChevronRight size={12} /> : <ChevronDown size={12} />}
          </button>
        ) : hovering ? (
          <div
            onMouseDown={(e) => onDragStart(index, e)}
            style={{ cursor: 'grab', color: 'var(--color-text-muted)' }}
            className="p-0.5"
          >
            <GripVertical size={12} />
          </div>
        ) : hasParent ? (
          <span className="text-[10px]" style={{ color: 'var(--color-border)' }}>&mdash;</span>
        ) : null}
      </div>
      {cols.map((col) => {
        const w = getColW(colWidths, col.key);
        const isCellEditing = editingCell && editingCell.rowIndex === index && editingCell.colKey === col.key;
        return (
          <div key={col.key} className="flex-shrink-0 overflow-hidden" style={{ width: w, maxWidth: w }}>
            <EditableCell
              task={task}
              col={col}
              isParent={isParent}
              onUpdateTask={onUpdateTask}
              selected={selected}
              datePickField={datePickField}
              onDatePickField={onDatePickField}
              editing={isCellEditing}
              onStartEditing={() => onSetEditingCell({ rowIndex: index, colKey: col.key })}
              onStopEditing={() => onSetEditingCell(null)}
              onNavigate={(dir) => onNavigateCell(index, col.key, dir)}
            />
          </div>
        );
      })}
      <div className="flex-1" />
      <div className="w-8 flex-shrink-0 px-1 flex items-center justify-center">
        {hovering && (
          <button
            onClick={() => onDeleteTask(task.id)}
            className="p-0.5 rounded cursor-pointer transition-colors"
            style={{ color: 'var(--color-text-muted)' }}
            title="Delete task"
          >
            <Trash2 size={12} />
          </button>
        )}
      </div>
    </div>
  );
}

function EditableCell({ task, col, isParent, onUpdateTask, selected, datePickField, onDatePickField, editing, onStartEditing, onStopEditing, onNavigate }) {
  const inputRef = useRef(null);
  const navigatingRef = useRef(false);

  const value = task[col.key];
  const isDateCol = col.type === 'date' && (col.key === 'startDate' || col.key === 'endDate');

  useEffect(() => {
    if (editing && inputRef.current) {
      inputRef.current.focus();
      if (inputRef.current.select) inputRef.current.select();
    }
  }, [editing]);

  const commitValue = useCallback((newVal) => {
    if (col.type === 'number') {
      const n = parseFloat(newVal);
      onUpdateTask(task.id, col.key, Number.isFinite(n) ? n : 0);
    } else {
      onUpdateTask(task.id, col.key, newVal);
    }
  }, [task.id, col.key, col.type, onUpdateTask]);

  const commitAndStop = useCallback((newVal, clearPick) => {
    if (navigatingRef.current) return;
    onStopEditing();
    if (clearPick && onDatePickField && isDateCol) onDatePickField(null);
    commitValue(newVal);
  }, [onStopEditing, onDatePickField, isDateCol, commitValue]);

  const commitAndNavigate = useCallback((newVal, direction) => {
    navigatingRef.current = true;
    commitValue(newVal);
    if (isDateCol && onDatePickField) onDatePickField(null);
    onNavigate(direction);
    requestAnimationFrame(() => { navigatingRef.current = false; });
  }, [commitValue, isDateCol, onDatePickField, onNavigate]);

  const handleStartEditing = () => {
    onStartEditing();
    if (isDateCol && onDatePickField) {
      onDatePickField(col.key);
    }
  };

  const cancelEditing = useCallback(() => {
    onStopEditing();
    if (isDateCol && onDatePickField) onDatePickField(null);
  }, [onStopEditing, isDateCol, onDatePickField]);

  const handleKeyDown = useCallback((e, getValue) => {
    if (e.key === 'ArrowUp') {
      e.preventDefault();
      commitAndNavigate(getValue(), 'up');
      return;
    }
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      commitAndNavigate(getValue(), 'down');
      return;
    }
    if (e.key === 'ArrowLeft') {
      const input = e.target;
      const atStart = input.selectionStart === 0 && input.selectionEnd === 0;
      if (col.type === 'text' && !atStart) return;
      e.preventDefault();
      commitAndNavigate(getValue(), 'left');
      return;
    }
    if (e.key === 'ArrowRight') {
      const input = e.target;
      const len = (input.value || '').length;
      const atEnd = input.selectionStart === len && input.selectionEnd === len;
      if (col.type === 'text' && !atEnd) return;
      e.preventDefault();
      commitAndNavigate(getValue(), 'right');
      return;
    }
    if (e.key === 'Tab') {
      e.preventDefault();
      commitAndNavigate(getValue(), e.shiftKey ? 'left' : 'right');
      return;
    }
    if (e.key === 'Enter') {
      e.preventDefault();
      commitAndNavigate(getValue(), 'down');
      return;
    }
    if (e.key === 'Escape') {
      cancelEditing();
    }
  }, [commitAndNavigate, cancelEditing, col.type]);

  if (col.key === 'id') {
    return (
      <span
        className="block px-2 tabular-nums text-xs truncate"
        style={{ color: 'var(--color-text-muted)', fontWeight: isParent ? 700 : 400, lineHeight: `${ROW_HEIGHT}px` }}
      >
        {value}
      </span>
    );
  }

  if (isParent && (col.key === 'startDate' || col.key === 'endDate' || col.key === 'progress' || col.key === 'duration')) {
    return <ReadOnlyDisplay col={col} value={value} isParent />;
  }

  if (col.key === 'status' || col.type === 'status') {
    return <StatusSelect value={value} onChange={(v) => onUpdateTask(task.id, col.key, v)} />;
  }

  if (!editing) {
    return (
      <button
        onClick={handleStartEditing}
        className="block w-full text-left px-2 text-xs cursor-text truncate"
        style={{
          color: value ? 'var(--color-text-primary)' : 'var(--color-text-muted)',
          fontWeight: col.key === 'name' && isParent ? 700 : col.key === 'name' ? 500 : 400,
          lineHeight: `${ROW_HEIGHT}px`,
          height: ROW_HEIGHT,
        }}
        title={String(value || '')}
      >
        {col.type === 'number'
          ? col.key === 'progress'
            ? <ProgressBar value={value} />
            : `${value ?? 0}`
          : value || '\u00A0'}
      </button>
    );
  }

  if (col.type === 'date') {
    return (
      <input
        ref={inputRef}
        type="date"
        defaultValue={value || ''}
        onBlur={(e) => commitAndStop(e.target.value || null, false)}
        onKeyDown={(e) => handleKeyDown(e, () => e.target.value || null)}
        className="w-full px-2 text-xs outline-none"
        style={{
          height: ROW_HEIGHT - 2,
          backgroundColor: 'var(--color-bg-tertiary)',
          color: 'var(--color-text-primary)',
          border: '1px solid var(--color-accent)',
          borderRadius: 4,
        }}
      />
    );
  }

  return (
    <input
      ref={inputRef}
      type={col.type === 'number' ? 'number' : 'text'}
      defaultValue={value ?? ''}
      onBlur={(e) => commitAndStop(e.target.value)}
      onKeyDown={(e) => handleKeyDown(e, () => e.target.value)}
      className="w-full px-2 text-xs outline-none"
      style={{
        height: ROW_HEIGHT - 2,
        backgroundColor: 'var(--color-bg-tertiary)',
        color: 'var(--color-text-primary)',
        border: '1px solid var(--color-accent)',
        borderRadius: 4,
      }}
    />
  );
}

function ReadOnlyDisplay({ col, value, isParent }) {
  const pct = col.key === 'progress' ? Math.min(100, Math.max(0, value || 0)) : null;
  return (
    <span
      className="block px-2 tabular-nums text-xs truncate"
      style={{
        color: 'var(--color-text-secondary)',
        fontWeight: isParent ? 700 : 400,
        lineHeight: `${ROW_HEIGHT}px`,
      }}
    >
      {pct != null ? `${pct}%` : col.type === 'number' ? `${value ?? 0}` : value || '-'}
    </span>
  );
}

function StatusSelect({ value, onChange }) {
  return (
    <select
      value={value || ''}
      onChange={(e) => onChange(e.target.value)}
      className="w-full text-[10px] outline-none cursor-pointer appearance-none truncate"
      style={{
        height: 20,
        padding: '0 4px',
        backgroundColor: 'var(--color-bg-tertiary)',
        color: 'var(--color-text-primary)',
        border: '1px solid var(--color-border)',
        borderRadius: 3,
      }}
    >
      <option value="">--</option>
      {STATUS_OPTIONS.map((s) => (
        <option key={s} value={s}>{s}</option>
      ))}
    </select>
  );
}

function ProgressBar({ value }) {
  const pct = Math.min(100, Math.max(0, value || 0));
  return (
    <div className="flex items-center gap-2">
      <div className="h-1.5 w-14 rounded-full overflow-hidden" style={{ backgroundColor: 'var(--color-bg-tertiary)' }}>
        <div
          className="h-full rounded-full transition-all"
          style={{
            width: `${pct}%`,
            backgroundColor: pct >= 100 ? 'var(--color-success)' : pct > 0 ? 'var(--color-accent)' : 'var(--color-bg-tertiary)',
          }}
        />
      </div>
      <span className="tabular-nums text-xs" style={{ color: 'var(--color-text-muted)' }}>{pct}%</span>
    </div>
  );
}

function ColumnPicker({ columns, visibleColumns, onToggle }) {
  return (
    <div
      className="absolute right-0 top-6 z-50 rounded-lg shadow-lg py-1 min-w-[160px]"
      style={{ backgroundColor: 'var(--color-bg-secondary)', border: '1px solid var(--color-border)' }}
    >
      <div className="px-3 py-1.5 text-[10px] font-semibold uppercase tracking-wider" style={{ color: 'var(--color-text-muted)' }}>
        Columns
      </div>
      {columns.map((col) => (
        <label
          key={col.key}
          className="flex items-center gap-2 px-3 py-1 cursor-pointer text-xs transition-colors"
          style={{ color: 'var(--color-text-secondary)' }}
          onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = 'var(--color-bg-hover)')}
          onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
        >
          <input type="checkbox" checked={visibleColumns.has(col.key)} onChange={() => onToggle(col.key)} className="accent-[var(--color-accent)]" />
          {col.label}
        </label>
      ))}
    </div>
  );
}

function AddRowBar({ onAddTask }) {
  return (
    <div className="flex-shrink-0 px-3 py-2" style={{ borderTop: '1px solid var(--color-border-subtle)' }}>
      <button
        onClick={onAddTask}
        className="flex items-center gap-1.5 px-2 py-1 rounded text-xs font-medium cursor-pointer transition-colors"
        style={{ color: 'var(--color-text-muted)', backgroundColor: 'transparent' }}
        onMouseEnter={(e) => { e.currentTarget.style.backgroundColor = 'var(--color-bg-hover)'; e.currentTarget.style.color = 'var(--color-accent)'; }}
        onMouseLeave={(e) => { e.currentTarget.style.backgroundColor = 'transparent'; e.currentTarget.style.color = 'var(--color-text-muted)'; }}
      >
        <Plus size={12} />
        Add task
      </button>
    </div>
  );
}
