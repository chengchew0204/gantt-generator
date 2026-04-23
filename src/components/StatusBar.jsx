import { useState, useEffect } from 'react';
import { Plus, Calendar, CalendarDays, ZoomIn, ZoomOut } from 'lucide-react';

const ZOOM_MIN = 5;
const ZOOM_MAX = 300;
const ZOOM_STEP = 10;

export default function StatusBar({
  onAddTask,
  scale,
  onScaleChange,
  zoomPct,
  onZoomChange,
  activeTab,
  tabs,
  onSelectTab,
  onAddTab,
  onRenameTab,
  onDeleteTab,
  onReorderTab,
}) {
  const [zoomInput, setZoomInput] = useState(String(zoomPct));
  useEffect(() => { setZoomInput(String(zoomPct)); }, [zoomPct]);

  const handleZoomStep = (direction) => {
    const next = Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, zoomPct + direction * ZOOM_STEP));
    onZoomChange(next);
  };

  const handleZoomInputChange = (e) => {
    const raw = e.target.value;
    setZoomInput(raw);
    const parsed = parseInt(raw, 10);
    if (!isNaN(parsed) && parsed >= ZOOM_MIN && parsed <= ZOOM_MAX) {
      onZoomChange(parsed);
    }
  };

  const handleZoomInputCommit = () => {
    const parsed = parseInt(zoomInput, 10);
    if (!isNaN(parsed)) {
      const clamped = Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, Math.round(parsed)));
      onZoomChange(clamped);
      setZoomInput(String(clamped));
    } else {
      setZoomInput(String(zoomPct));
    }
  };

  const isGantt = activeTab === 'gantt';

  return (
    <div
      className="flex items-center justify-between flex-shrink-0 px-2 gap-2"
      style={{
        height: 32,
        backgroundColor: 'var(--color-bg-secondary)',
        borderTop: '1px solid var(--color-border)',
      }}
    >
      {/* Left: Add Task */}
      <div className="flex items-center gap-1 flex-shrink-0">
        <button
          onClick={onAddTask}
          className="flex items-center gap-1 px-2 py-0.5 rounded text-[12px] font-medium cursor-pointer transition-colors"
          style={{ color: 'var(--color-text-muted)' }}
          onMouseEnter={(e) => { e.currentTarget.style.color = 'var(--color-accent)'; e.currentTarget.style.backgroundColor = 'var(--color-accent-muted)'; }}
          onMouseLeave={(e) => { e.currentTarget.style.color = 'var(--color-text-muted)'; e.currentTarget.style.backgroundColor = 'transparent'; }}
          title="Add a new task row"
        >
          <Plus size={12} />
          <span>Add Task</span>
        </button>
      </div>

      {/* Middle: Tab Bar */}
      <TabBar
        activeTab={activeTab}
        tabs={tabs}
        onSelectTab={onSelectTab}
        onAddTab={onAddTab}
        onRenameTab={onRenameTab}
        onDeleteTab={onDeleteTab}
        onReorderTab={onReorderTab}
      />

      {/* Right: Scale + Zoom (visible only on Gantt tab) */}
      {isGantt && (
        <div className="flex items-center gap-1.5 flex-shrink-0">
          <ScaleButton active={scale === 'day'} onClick={() => onScaleChange('day')} icon={Calendar} label="Day" />
          <ScaleButton active={scale === 'week'} onClick={() => onScaleChange('week')} icon={CalendarDays} label="Week" />
          <div className="w-px h-4 mx-0.5" style={{ backgroundColor: 'var(--color-border)' }} />
          <ZoomButton icon={ZoomOut} onClick={() => handleZoomStep(-1)} disabled={zoomPct <= ZOOM_MIN} />
          <input
            type="number"
            value={zoomInput}
            min={ZOOM_MIN}
            max={ZOOM_MAX}
            onChange={handleZoomInputChange}
            onBlur={handleZoomInputCommit}
            onKeyDown={(e) => { if (e.key === 'Enter') e.currentTarget.blur(); }}
            className="tabular-nums text-center font-medium bg-transparent outline-none border-b"
            style={{
              width: 34,
              fontSize: 11,
              color: 'var(--color-text-secondary)',
              borderColor: 'var(--color-border)',
              MozAppearance: 'textfield',
            }}
          />
          <span className="text-[11px]" style={{ color: 'var(--color-text-muted)' }}>%</span>
          <ZoomButton icon={ZoomIn} onClick={() => handleZoomStep(1)} disabled={zoomPct >= ZOOM_MAX} />
        </div>
      )}
      {!isGantt && <div />}
    </div>
  );
}

function TabBar({ activeTab, tabs = [], onSelectTab, onAddTab, onRenameTab, onDeleteTab, onReorderTab }) {
  const [renamingId, setRenamingId] = useState(null);
  const [renameValue, setRenameValue] = useState('');

  const startRename = (tabId, currentName) => {
    setRenamingId(tabId);
    setRenameValue(currentName);
  };

  const commitRename = () => {
    if (renamingId && onRenameTab && renameValue.trim()) {
      onRenameTab(renamingId, renameValue.trim());
    }
    setRenamingId(null);
    setRenameValue('');
  };

  return (
    <div className="flex items-center gap-0.5 flex-1 min-w-0 overflow-x-auto px-1" style={{ scrollbarWidth: 'none' }}>
      {/* Fixed Gantt tab */}
      <TabPill
        active={activeTab === 'gantt'}
        label="Gantt"
        onClick={() => onSelectTab('gantt')}
      />

      {/* User tabs */}
      {tabs.map((tab) => (
        renamingId === tab.id ? (
          <input
            key={tab.id}
            autoFocus
            type="text"
            value={renameValue}
            onChange={(e) => setRenameValue(e.target.value)}
            onBlur={commitRename}
            onKeyDown={(e) => {
              if (e.key === 'Enter') commitRename();
              if (e.key === 'Escape') { setRenamingId(null); setRenameValue(''); }
            }}
            className="text-[11px] font-medium outline-none px-1.5 py-0.5 rounded"
            style={{
              width: 80,
              backgroundColor: 'var(--color-bg-tertiary)',
              color: 'var(--color-text-primary)',
              border: '1px solid var(--color-accent)',
            }}
          />
        ) : (
          <TabPill
            key={tab.id}
            active={activeTab === tab.id}
            label={tab.name}
            onClick={() => onSelectTab(tab.id)}
            onDoubleClick={() => startRename(tab.id, tab.name)}
            onClose={onDeleteTab ? () => onDeleteTab(tab.id) : null}
          />
        )
      ))}

      {/* Add Tab button */}
      {onAddTab ? (
        <button
          onClick={onAddTab}
          className="flex items-center justify-center w-5 h-5 rounded cursor-pointer transition-colors flex-shrink-0"
          style={{ color: 'var(--color-text-muted)' }}
          onMouseEnter={(e) => { e.currentTarget.style.color = 'var(--color-accent)'; e.currentTarget.style.backgroundColor = 'var(--color-accent-muted)'; }}
          onMouseLeave={(e) => { e.currentTarget.style.color = 'var(--color-text-muted)'; e.currentTarget.style.backgroundColor = 'transparent'; }}
          title="Add a new tab"
        >
          <Plus size={11} />
        </button>
      ) : (
        <button
          disabled
          className="flex items-center justify-center w-5 h-5 rounded flex-shrink-0 opacity-30 cursor-default"
          style={{ color: 'var(--color-text-muted)' }}
          title="Tab management coming soon"
        >
          <Plus size={11} />
        </button>
      )}
    </div>
  );
}

function TabPill({ active, label, onClick, onDoubleClick, onClose }) {
  const [hovered, setHovered] = useState(false);

  return (
    <div
      className="flex items-center gap-1 px-2 py-0.5 rounded text-[11px] font-medium cursor-pointer transition-colors flex-shrink-0 select-none"
      style={{
        backgroundColor: active ? 'var(--color-accent-muted)' : hovered ? 'var(--color-bg-hover)' : 'transparent',
        color: active ? 'var(--color-accent)' : 'var(--color-text-muted)',
      }}
      onClick={onClick}
      onDoubleClick={onDoubleClick}
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
    >
      <span className="truncate max-w-[100px]" title={label}>{label}</span>
      {onClose && hovered && (
        <button
          onClick={(e) => { e.stopPropagation(); onClose(); }}
          className="p-0 ml-0.5 rounded-sm transition-colors flex-shrink-0"
          style={{ color: 'var(--color-text-muted)', lineHeight: 0 }}
          onMouseEnter={(e) => { e.currentTarget.style.color = 'var(--color-danger)'; }}
          onMouseLeave={(e) => { e.currentTarget.style.color = 'var(--color-text-muted)'; }}
          title={`Close ${label}`}
        >
          &times;
        </button>
      )}
    </div>
  );
}

function ScaleButton({ active, onClick, icon: Icon, label }) {
  return (
    <button
      onClick={onClick}
      className="flex items-center gap-1 px-2 py-0.5 rounded text-[11px] font-medium cursor-pointer transition-colors"
      style={{
        backgroundColor: active ? 'var(--color-accent-muted)' : 'transparent',
        color: active ? 'var(--color-accent)' : 'var(--color-text-muted)',
      }}
    >
      <Icon size={10} />{label}
    </button>
  );
}

function ZoomButton({ icon: Icon, onClick, disabled }) {
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      className="p-0.5 rounded cursor-pointer transition-colors disabled:opacity-30 disabled:cursor-default"
      style={{ color: 'var(--color-text-muted)' }}
    >
      <Icon size={11} />
    </button>
  );
}
