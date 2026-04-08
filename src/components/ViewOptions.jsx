import { useRef, useEffect } from 'react';

const CHART_TOGGLES = [
  { key: 'showCriticalPath', label: 'Critical Path Highlighting' },
  { key: 'showSlack', label: 'Slack Bars' },
  { key: 'showDependencies', label: 'Dependency Arrows' },
  { key: 'showTodayLine', label: 'Today Line' },
  { key: 'showBaseline', label: 'Baseline Bars' },
  { key: 'showTaskNames', label: 'Task Names on Chart' },
  { key: 'showProgressPercent', label: 'Progress % on Bars' },
  { key: 'skipWeekends', label: 'Skip Weekends (working days)' },
  { key: 'showScaleButtons', label: 'Day/Week Scale Buttons' },
  { key: 'showWeekLabels', label: 'Week Number Labels (W1, W2...)' },
  { key: 'showMonthLabels', label: 'Month Labels' },
  { key: 'showDayLabels', label: 'Day/Date Labels' },
];

export default function ViewOptions({
  open,
  onClose,
  viewOptions,
  onToggleViewOption,
  columns,
  visibleColumns,
  onToggleColumn,
  anchorRef,
}) {
  const panelRef = useRef(null);

  useEffect(() => {
    if (!open) return;
    const handler = (e) => {
      if (
        panelRef.current && !panelRef.current.contains(e.target) &&
        anchorRef?.current && !anchorRef.current.contains(e.target)
      ) {
        onClose();
      }
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [open, onClose, anchorRef]);

  if (!open) return null;

  return (
    <div
      ref={panelRef}
      className="absolute right-0 top-full mt-1 z-50 rounded-lg shadow-xl py-2 min-w-[220px] max-h-[70vh] overflow-y-auto"
      style={{
        backgroundColor: 'var(--color-bg-secondary)',
        border: '1px solid var(--color-border)',
      }}
    >
      <SectionLabel>Chart Elements</SectionLabel>
      {CHART_TOGGLES.map(({ key, label }) => (
        <ToggleRow
          key={key}
          label={label}
          checked={viewOptions[key] !== false && viewOptions[key] !== 'false'}
          onChange={() => onToggleViewOption(key)}
        />
      ))}

      <Divider />

      <SectionLabel>Table Columns</SectionLabel>
      {columns.map((col) => (
        <ToggleRow
          key={col.key}
          label={col.label}
          checked={visibleColumns.has(col.key)}
          onChange={() => onToggleColumn(col.key)}
        />
      ))}
    </div>
  );
}

function SectionLabel({ children }) {
  return (
    <div
      className="px-3 py-1.5 text-[11px] font-semibold uppercase tracking-wider"
      style={{ color: 'var(--color-text-muted)' }}
    >
      {children}
    </div>
  );
}

function ToggleRow({ label, checked, onChange }) {
  return (
    <label
      className="flex items-center gap-2 px-3 py-1.5 cursor-pointer text-[13px] transition-colors"
      style={{ color: 'var(--color-text-secondary)' }}
      onMouseEnter={(e) => (e.currentTarget.style.backgroundColor = 'var(--color-bg-hover)')}
      onMouseLeave={(e) => (e.currentTarget.style.backgroundColor = 'transparent')}
    >
      <input
        type="checkbox"
        checked={checked}
        onChange={onChange}
        className="accent-[var(--color-accent)]"
      />
      {label}
    </label>
  );
}

function Divider() {
  return (
    <div
      className="my-1 mx-3"
      style={{ borderBottom: '1px solid var(--color-border-subtle)' }}
    />
  );
}
