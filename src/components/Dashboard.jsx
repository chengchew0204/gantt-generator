import { useMemo, useState, useRef, useEffect } from 'react';
import {
  FileDown,
  FileUp,
  Download,
  Image,
  Activity,
  CheckCircle2,
  Clock,
  AlertTriangle,
  Play,
  ArrowRight,
  Palette,
  SlidersHorizontal,
  CalendarRange,
  CircleDot,
  Pencil,
  BookOpen,
  Undo2,
  Redo2,
} from 'lucide-react';
import ViewOptions from './ViewOptions';

function formatShortDate(iso) {
  if (!iso) return '--';
  const [y, m, d] = iso.split('-');
  return `${Number(m)}/${Number(d)}/${y}`;
}

export default function Dashboard({
  tasks,
  projectName,
  onChangeProjectName,
  onImport,
  onDownloadTemplate,
  onExport,
  onExportPng,
  onOpenGuide,
  onOpenTheme,
  onUndo,
  onRedo,
  canUndo,
  canRedo,
  viewOptionsOpen,
  onToggleViewOptions,
  viewBtnRef,
  viewOptions,
  onToggleViewOption,
  columns,
  visibleColumns,
  onToggleColumn,
}) {
  const stats = useMemo(() => {
    const total = tasks.length;
    const completed = tasks.filter((t) => t.progress >= 100).length;
    const progress = total > 0
      ? Math.round(tasks.reduce((sum, t) => sum + (t.progress || 0), 0) / total)
      : 0;
    const inProgress = tasks.filter((t) => t.progress > 0 && t.progress < 100).length;
    const notStarted = tasks.filter((t) => !t.progress || t.progress === 0).length;
    const critical = tasks.filter((t) => t.isCritical).length;

    const today = new Date().toISOString().slice(0, 10);
    const overdue = tasks.filter(
      (t) => t.endDate && t.endDate < today && (t.progress || 0) < 100,
    ).length;

    let minStart = null;
    let maxEnd = null;
    for (const t of tasks) {
      if (t.startDate && (!minStart || t.startDate < minStart)) minStart = t.startDate;
      if (t.endDate && (!maxEnd || t.endDate > maxEnd)) maxEnd = t.endDate;
    }

    const sorted = [...tasks].sort((a, b) => {
      if (!a.startDate) return 1;
      if (!b.startDate) return -1;
      return a.startDate.localeCompare(b.startDate);
    });

    const currentTask = sorted.find((t) => t.progress > 0 && t.progress < 100);
    const nextTask = sorted.find(
      (t) => (t.progress === 0 || t.progress == null) && t.status !== 'Completed',
    );

    return { total, completed, progress, inProgress, notStarted, critical, overdue, minStart, maxEnd, currentTask, nextTask };
  }, [tasks]);

  const handleFileSelect = () => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx,.xls';
    input.onchange = (e) => {
      const file = e.target.files?.[0];
      if (file) onImport(file);
    };
    input.click();
  };

  return (
    <div className="flex flex-col gap-2 px-4 py-3 border-b" style={{ borderColor: 'var(--color-border)' }}>
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-1.5">
            <span className="text-xs font-semibold tracking-tight" style={{ color: 'var(--color-text-muted)' }}>
              GanttGen
            </span>
            <span className="text-[9px] px-1 py-px rounded font-medium"
              style={{ backgroundColor: 'var(--color-accent-muted)', color: 'var(--color-accent)' }}>
              v0.1
            </span>
          </div>
          <span style={{ color: 'var(--color-border)' }}>/</span>
          <EditableProjectName value={projectName} onChange={onChangeProjectName} />
          <div className="flex items-center gap-0.5">
            <IconButton icon={Undo2} title="Undo (Ctrl+Z)" onClick={onUndo} disabled={canUndo && !canUndo()} />
            <IconButton icon={Redo2} title="Redo (Ctrl+Shift+Z)" onClick={onRedo} disabled={canRedo && !canRedo()} />
          </div>
        </div>

        <div className="flex items-center gap-2">
          <div className="flex items-center gap-2" data-guide="toolbar-import">
            <ActionButton icon={FileUp} label="Import Excel" onClick={handleFileSelect} />
            <ActionButton icon={FileDown} label="Download Template" onClick={onDownloadTemplate} />
            <ActionButton icon={Download} label="Save to Excel" onClick={onExport} primary />
          </div>
          <ActionButton icon={Image} label="Export PNG" onClick={onExportPng} guideAttr="btn-export-png" />

          <Separator />

          <ActionButton icon={BookOpen} label="Guide" onClick={onOpenGuide} />
          <ActionButton icon={Palette} label="Theme" onClick={onOpenTheme} guideAttr="btn-theme" />
          <div className="relative" ref={viewBtnRef} data-guide="btn-view">
            <ActionButton icon={SlidersHorizontal} label="View" onClick={onToggleViewOptions} />
            <ViewOptions
              open={viewOptionsOpen}
              onClose={onToggleViewOptions}
              viewOptions={viewOptions}
              onToggleViewOption={onToggleViewOption}
              columns={columns}
              visibleColumns={visibleColumns}
              onToggleColumn={onToggleColumn}
              anchorRef={viewBtnRef}
            />
          </div>
        </div>
      </div>

      <div className="flex items-center gap-5 flex-wrap">
        <StatCard icon={Activity} label="Progress" value={`${stats.progress}%`} color="var(--color-accent)" />
        <StatCard icon={CheckCircle2} label="Completed" value={`${stats.completed} / ${stats.total}`} color="var(--color-success)" />
        <StatCard icon={Clock} label="In Progress" value={String(stats.inProgress)} color="var(--color-info)" />
        <StatCard icon={CircleDot} label="Not Started" value={String(stats.notStarted)} color="var(--color-text-muted)" />
        <StatCard icon={AlertTriangle} label="Critical" value={String(stats.critical)} color="var(--color-critical-path)" />
        {stats.overdue > 0 && (
          <StatCard icon={AlertTriangle} label="Overdue" value={String(stats.overdue)} color="var(--color-danger)" />
        )}

        {stats.total > 0 && (
          <>
            <Divider />
            <StatCard icon={CalendarRange} label="Date Range" value={`${formatShortDate(stats.minStart)} - ${formatShortDate(stats.maxEnd)}`} color="var(--color-text-secondary)" />
            <Divider />
            <TaskInfoCard icon={Play} label="Current Task" task={stats.currentTask} color="var(--color-info)" />
            <TaskInfoCard icon={ArrowRight} label="Next Task" task={stats.nextTask} color="var(--color-text-muted)" />
          </>
        )}
      </div>
    </div>
  );
}

function EditableProjectName({ value, onChange }) {
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(value);
  const inputRef = useRef(null);

  useEffect(() => {
    setDraft(value);
  }, [value]);

  useEffect(() => {
    if (editing && inputRef.current) {
      inputRef.current.focus();
      inputRef.current.select();
    }
  }, [editing]);

  const commit = () => {
    setEditing(false);
    if (onChange) onChange(draft.trim());
  };

  if (editing) {
    return (
      <input
        ref={inputRef}
        type="text"
        value={draft}
        onChange={(e) => setDraft(e.target.value)}
        onBlur={commit}
        onKeyDown={(e) => {
          if (e.key === 'Enter') commit();
          if (e.key === 'Escape') { setDraft(value); setEditing(false); }
        }}
        className="text-sm font-semibold tracking-tight bg-transparent outline-none border-b"
        style={{
          color: 'var(--color-text-primary)',
          borderColor: 'var(--color-accent)',
          minWidth: 120,
          maxWidth: 320,
        }}
        placeholder="Untitled Project"
      />
    );
  }

  return (
    <button
      onClick={() => setEditing(true)}
      className="flex items-center gap-1.5 group cursor-pointer"
      title="Click to rename project"
    >
      <span
        className="text-sm font-semibold tracking-tight"
        style={{ color: value ? 'var(--color-text-primary)' : 'var(--color-text-muted)' }}
      >
        {value || 'Untitled Project'}
      </span>
      <Pencil size={11} className="opacity-0 group-hover:opacity-60 transition-opacity" style={{ color: 'var(--color-text-muted)' }} />
    </button>
  );
}

function Separator() {
  return <div className="w-px h-6 self-center" style={{ backgroundColor: 'var(--color-border)' }} />;
}

function Divider() {
  return <div className="w-px h-8 self-center" style={{ backgroundColor: 'var(--color-border)' }} />;
}

function TaskInfoCard({ icon: Icon, label, task, color }) {
  return (
    <div className="flex items-center gap-2.5">
      <div className="flex items-center justify-center w-7 h-7 rounded-md" style={{ backgroundColor: `${color}18` }}>
        <Icon size={14} style={{ color }} />
      </div>
      <div>
        <div className="text-xs font-medium" style={{ color: 'var(--color-text-muted)' }}>{label}</div>
        <div
          className="text-xs font-semibold max-w-[140px] truncate"
          style={{ color: task ? 'var(--color-text-primary)' : 'var(--color-text-muted)' }}
          title={task?.name || ''}
        >
          {task ? task.name : '--'}
        </div>
      </div>
    </div>
  );
}

function ActionButton({ icon: Icon, label, onClick, primary, guideAttr, disabled }) {
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      title={label}
      data-guide={guideAttr || undefined}
      className="flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-medium transition-colors cursor-pointer disabled:opacity-30 disabled:cursor-default"
      style={{
        backgroundColor: primary ? 'var(--color-accent)' : 'var(--color-bg-tertiary)',
        color: primary ? '#fff' : 'var(--color-text-secondary)',
        border: primary ? 'none' : '1px solid var(--color-border)',
      }}
      onMouseEnter={(e) => { if (!primary && !disabled) e.currentTarget.style.backgroundColor = 'var(--color-bg-hover)'; }}
      onMouseLeave={(e) => { if (!primary) e.currentTarget.style.backgroundColor = 'var(--color-bg-tertiary)'; }}
    >
      <Icon size={14} />
      <span className="hidden sm:inline">{label}</span>
    </button>
  );
}

function IconButton({ icon: Icon, title, onClick, disabled }) {
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      title={title}
      className="p-1 rounded transition-colors cursor-pointer disabled:opacity-30 disabled:cursor-default"
      style={{ color: 'var(--color-text-muted)' }}
      onMouseEnter={(e) => { if (!disabled) e.currentTarget.style.color = 'var(--color-text-secondary)'; }}
      onMouseLeave={(e) => { e.currentTarget.style.color = 'var(--color-text-muted)'; }}
    >
      <Icon size={14} />
    </button>
  );
}

function StatCard({ icon: Icon, label, value, color }) {
  return (
    <div className="flex items-center gap-2.5">
      <div className="flex items-center justify-center w-7 h-7 rounded-md" style={{ backgroundColor: `${color}18` }}>
        <Icon size={14} style={{ color }} />
      </div>
      <div>
        <div className="text-xs font-medium" style={{ color: 'var(--color-text-muted)' }}>{label}</div>
        <div className="text-sm font-semibold tabular-nums" style={{ color: 'var(--color-text-primary)' }}>{value}</div>
      </div>
    </div>
  );
}
