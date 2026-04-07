import { X, Check } from 'lucide-react';

const THEME_PRESETS = {
  'Linear Dark': {
    'bg-primary': '#0f0f12',
    'bg-secondary': '#1a1a23',
    'bg-tertiary': '#23232f',
    'bg-hover': '#2a2a38',
    'border': '#2e2e3a',
    'border-subtle': '#232330',
    'text-primary': '#e8e8ed',
    'text-secondary': '#9898a8',
    'text-muted': '#66667a',
    'accent': '#6366f1',
    'accent-hover': '#818cf8',
    'success': '#22c55e',
    'warning': '#f59e0b',
    'danger': '#ef4444',
    'info': '#3b82f6',
    'critical-path': '#f43f5e',
  },
  'Notion Light': {
    'bg-primary': '#ffffff',
    'bg-secondary': '#f7f7f5',
    'bg-tertiary': '#edece9',
    'bg-hover': '#e8e7e4',
    'border': '#e0dfdc',
    'border-subtle': '#ebebea',
    'text-primary': '#1f1f1f',
    'text-secondary': '#6b6b6b',
    'text-muted': '#9b9b9b',
    'accent': '#2383e2',
    'accent-hover': '#1a6dbe',
    'success': '#0f7b0f',
    'warning': '#d97706',
    'danger': '#e03e3e',
    'info': '#2383e2',
    'critical-path': '#e03e3e',
  },
  'Classic Tremor': {
    'bg-primary': '#0f172a',
    'bg-secondary': '#1e293b',
    'bg-tertiary': '#273548',
    'bg-hover': '#334155',
    'border': '#334155',
    'border-subtle': '#1e293b',
    'text-primary': '#f1f5f9',
    'text-secondary': '#94a3b8',
    'text-muted': '#64748b',
    'accent': '#10b981',
    'accent-hover': '#34d399',
    'success': '#22c55e',
    'warning': '#f59e0b',
    'danger': '#ef4444',
    'info': '#3b82f6',
    'critical-path': '#f43f5e',
  },
  Gray: {
    'bg-primary': '#2c3038',
    'bg-secondary': '#353a44',
    'bg-tertiary': '#3e4450',
    'bg-hover': '#4a5160',
    'border': '#5c6575',
    'border-subtle': '#323842',
    'text-primary': '#e9ebef',
    'text-secondary': '#a7b0bf',
    'text-muted': '#7c8696',
    'accent': '#9aa5b8',
    'accent-hover': '#b6c0d4',
    'success': '#7dcda3',
    'warning': '#e3b565',
    'danger': '#e89393',
    'info': '#8eb4dc',
    'critical-path': '#f0a0a0',
  },
  'Black & White': {
    'bg-primary': '#0c0c0c',
    'bg-secondary': '#141414',
    'bg-tertiary': '#1c1c1c',
    'bg-hover': '#262626',
    'border': '#3d3d3d',
    'border-subtle': '#181818',
    'text-primary': '#f5f5f5',
    'text-secondary': '#bdbdbd',
    'text-muted': '#858585',
    'accent': '#e8e8e8',
    'accent-hover': '#ffffff',
    'success': '#c6c6c6',
    'warning': '#9e9e9e',
    'danger': '#ebebeb',
    'info': '#adadad',
    'critical-path': '#ffffff',
  },
};

const EDITABLE_KEYS = [
  { key: 'bg-primary', label: 'Background' },
  { key: 'bg-secondary', label: 'Surface' },
  { key: 'bg-tertiary', label: 'Elevated' },
  { key: 'text-primary', label: 'Text' },
  { key: 'text-secondary', label: 'Text Secondary' },
  { key: 'accent', label: 'Accent' },
  { key: 'critical-path', label: 'Critical Path' },
  { key: 'success', label: 'Success' },
  { key: 'warning', label: 'Warning' },
  { key: 'danger', label: 'Danger' },
];

export { THEME_PRESETS };

export default function ThemePanel({ open, onClose, activeTheme, onApplyPreset, onApplyCustomColor, categoryColors = {}, onChangeCategoryColor }) {
  if (!open) return null;

  return (
    <div className="fixed inset-0 z-50 flex justify-end">
      <div
        className="absolute inset-0"
        style={{ backgroundColor: 'rgba(0,0,0,0.4)' }}
        onClick={onClose}
      />
      <div
        className="relative w-80 h-full overflow-y-auto shadow-xl"
        style={{
          backgroundColor: 'var(--color-bg-secondary)',
          borderLeft: '1px solid var(--color-border)',
        }}
      >
        <div
          className="flex items-center justify-between px-4 py-3 sticky top-0 z-10"
          style={{
            backgroundColor: 'var(--color-bg-secondary)',
            borderBottom: '1px solid var(--color-border)',
          }}
        >
          <h2 className="text-[15px] font-semibold" style={{ color: 'var(--color-text-primary)' }}>
            Theme
          </h2>
          <button
            onClick={onClose}
            className="p-1 rounded cursor-pointer transition-colors"
            style={{ color: 'var(--color-text-muted)' }}
          >
            <X size={16} />
          </button>
        </div>

        <div className="p-4">
          <SectionLabel>Presets</SectionLabel>
          <div className="flex flex-col gap-2 mb-6">
            {Object.entries(THEME_PRESETS).map(([name, colors]) => (
              <PresetCard
                key={name}
                name={name}
                colors={colors}
                active={activeTheme === name}
                onSelect={() => onApplyPreset(name, colors)}
              />
            ))}
          </div>

          <SectionLabel>Custom Colors</SectionLabel>
          <div className="flex flex-col gap-2">
            {EDITABLE_KEYS.map(({ key, label }) => {
              const cssVar = `--color-${key}`;
              const current = getComputedStyle(document.documentElement)
                .getPropertyValue(cssVar)
                .trim();
              return (
                <ColorRow
                  key={key}
                  label={label}
                  value={current}
                  onChange={(hex) => onApplyCustomColor(key, hex)}
                />
              );
            })}
          </div>

          {Object.keys(categoryColors).length > 0 && (
            <>
              <div className="mt-6" />
              <SectionLabel>Category Colors</SectionLabel>
              <div className="flex flex-col gap-2">
                {Object.entries(categoryColors).map(([cat, hex]) => (
                  <ColorRow
                    key={cat}
                    label={cat}
                    value={hex}
                    onChange={(newHex) => onChangeCategoryColor && onChangeCategoryColor(cat, newHex)}
                  />
                ))}
              </div>
            </>
          )}
        </div>
      </div>
    </div>
  );
}

function SectionLabel({ children }) {
  return (
    <div
      className="text-[11px] font-semibold uppercase tracking-wider mb-2"
      style={{ color: 'var(--color-text-muted)' }}
    >
      {children}
    </div>
  );
}

function PresetCard({ name, colors, active, onSelect }) {
  const swatches = [colors['bg-primary'], colors['accent'], colors['text-primary'], colors['success'], colors['critical-path']];

  return (
    <button
      onClick={onSelect}
      className="flex items-center gap-3 px-3 py-2.5 rounded-lg text-left cursor-pointer transition-colors w-full"
      style={{
        backgroundColor: active ? 'var(--color-accent-muted)' : 'var(--color-bg-tertiary)',
        border: active ? '1px solid var(--color-accent)' : '1px solid var(--color-border)',
      }}
    >
      <div className="flex gap-1">
        {swatches.map((c, i) => (
          <div key={i} className="w-4 h-4 rounded-sm" style={{ backgroundColor: c }} />
        ))}
      </div>
      <span
        className="text-[13px] font-medium flex-1"
        style={{ color: active ? 'var(--color-accent)' : 'var(--color-text-secondary)' }}
      >
        {name}
      </span>
      {active && <Check size={14} style={{ color: 'var(--color-accent)' }} />}
    </button>
  );
}

function ColorRow({ label, value, onChange }) {
  const normalizedValue = normalizeToHex(value);
  return (
    <label className="flex items-center justify-between gap-2">
      <span className="text-[13px]" style={{ color: 'var(--color-text-secondary)' }}>
        {label}
      </span>
      <div className="flex items-center gap-1.5">
        <span className="text-[11px] tabular-nums" style={{ color: 'var(--color-text-muted)' }}>
          {normalizedValue}
        </span>
        <input
          type="color"
          value={normalizedValue}
          onChange={(e) => onChange(e.target.value)}
          className="w-6 h-6 rounded cursor-pointer border-0 p-0"
          style={{ backgroundColor: 'transparent' }}
        />
      </div>
    </label>
  );
}

function normalizeToHex(value) {
  if (!value) return '#000000';
  const trimmed = value.trim();
  if (trimmed.startsWith('#') && (trimmed.length === 7 || trimmed.length === 4)) {
    return trimmed;
  }
  if (trimmed.startsWith('rgb')) {
    const nums = trimmed.match(/\d+/g);
    if (nums && nums.length >= 3) {
      const r = parseInt(nums[0], 10);
      const g = parseInt(nums[1], 10);
      const b = parseInt(nums[2], 10);
      return `#${((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1)}`;
    }
  }
  return '#000000';
}
