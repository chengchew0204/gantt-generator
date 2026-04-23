import { useState, useRef, useEffect, useCallback } from 'react';
import {
  Bold,
  Italic,
  Underline,
  Grid3x3,
  AlignLeft,
  AlignCenter,
  AlignRight,
  AlignVerticalJustifyStart,
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  ChevronDown,
  Shapes,
  Sparkles,
  BringToFront,
  SendToBack,
  MoveUp,
  MoveDown,
  Trash2,
} from 'lucide-react';
import { groupShapePresets, renderShapePath, isLineShape, getShapePreset } from '../utils/ShapePresets';

export default function SpreadsheetToolbar({
  cellRef,
  displayValue,
  editText,
  isEditing,
  onEditTextChange,
  onStartBarEdit,
  onBarKeyDown,
  barInputRef,
  cellStyle,
  onApplyStyle,
  onApplyBorderPreset,
  onToggleGridLines,
  showGridLines,
  onMerge,
  onUnmerge,
  isMergedSelection,
  shapeMode,
  onSetShapeMode,
  selectedShape,
  onApplyShapeStyle,
  onApplyTextStyle,
  onArrange,
  onDeleteShape,
}) {
  const style = cellStyle || {};
  const showShapeRow = !!selectedShape || !!shapeMode;

  return (
    <div className="flex flex-col flex-shrink-0" style={{ borderBottom: '1px solid var(--color-border)' }}>
      {/* Row 1: Format buttons */}
      <div
        className="flex items-center gap-1 px-2"
        style={{
          height: 28,
          borderBottom: '1px solid var(--color-border-subtle)',
          backgroundColor: 'var(--color-bg-secondary)',
        }}
      >
        <FormatButton icon={Bold} active={style.bold} onClick={() => onApplyStyle({ bold: !style.bold })} title="Bold (Ctrl+B)" />
        <FormatButton icon={Italic} active={style.italic} onClick={() => onApplyStyle({ italic: !style.italic })} title="Italic (Ctrl+I)" />
        <FormatButton icon={Underline} active={style.underline} onClick={() => onApplyStyle({ underline: !style.underline })} title="Underline (Ctrl+U)" />

        <div className="w-px h-4 mx-0.5 flex-shrink-0" style={{ backgroundColor: 'var(--color-border)' }} />

        <FontSizeInput value={style.fontSize || 12} onChange={(v) => onApplyStyle({ fontSize: v })} />

        <div className="w-px h-4 mx-0.5 flex-shrink-0" style={{ backgroundColor: 'var(--color-border)' }} />

        <ColorPicker value={style.color || ''} onChange={(hex) => onApplyStyle({ color: hex })} title="Font color" label="A" />
        <ColorPicker value={style.bg || ''} onChange={(hex) => onApplyStyle({ bg: hex })} title="Background color" label="Bg" isBg />

        <div className="w-px h-4 mx-0.5 flex-shrink-0" style={{ backgroundColor: 'var(--color-border)' }} />

        <BorderPicker onApply={onApplyBorderPreset} />

        <div className="w-px h-4 mx-0.5 flex-shrink-0" style={{ backgroundColor: 'var(--color-border)' }} />

        <FormatButton
          icon={AlignLeft}
          active={style.hAlign === 'left'}
          onClick={() => onApplyStyle({ hAlign: style.hAlign === 'left' ? undefined : 'left' })}
          title="Align left"
        />
        <FormatButton
          icon={AlignCenter}
          active={style.hAlign === 'center'}
          onClick={() => onApplyStyle({ hAlign: style.hAlign === 'center' ? undefined : 'center' })}
          title="Align center"
        />
        <FormatButton
          icon={AlignRight}
          active={style.hAlign === 'right'}
          onClick={() => onApplyStyle({ hAlign: style.hAlign === 'right' ? undefined : 'right' })}
          title="Align right"
        />

        <div className="w-px h-4 mx-0.5 flex-shrink-0" style={{ backgroundColor: 'var(--color-border)' }} />

        <FormatButton
          icon={AlignVerticalJustifyStart}
          active={style.vAlign === 'top'}
          onClick={() => onApplyStyle({ vAlign: style.vAlign === 'top' ? undefined : 'top' })}
          title="Align top"
        />
        <FormatButton
          icon={AlignVerticalJustifyCenter}
          active={style.vAlign === 'middle'}
          onClick={() => onApplyStyle({ vAlign: style.vAlign === 'middle' ? undefined : 'middle' })}
          title="Align middle"
        />
        <FormatButton
          icon={AlignVerticalJustifyEnd}
          active={style.vAlign === 'bottom'}
          onClick={() => onApplyStyle({ vAlign: style.vAlign === 'bottom' ? undefined : 'bottom' })}
          title="Align bottom"
        />

        <div className="w-px h-4 mx-0.5 flex-shrink-0" style={{ backgroundColor: 'var(--color-border)' }} />

        <MergeMenu
          onMerge={onMerge}
          onUnmerge={onUnmerge}
          isMergedSelection={!!isMergedSelection}
        />

        <div className="w-px h-4 mx-0.5 flex-shrink-0" style={{ backgroundColor: 'var(--color-border)' }} />

        <FormatButton icon={Grid3x3} active={showGridLines !== false} onClick={onToggleGridLines} title="Toggle grid lines" />

        <div className="w-px h-4 mx-0.5 flex-shrink-0" style={{ backgroundColor: 'var(--color-border)' }} />

        <InsertShapesMenu shapeMode={shapeMode} onSetShapeMode={onSetShapeMode} />
      </div>

      {showShapeRow && (
        <ShapeToolbarRow
          shape={selectedShape}
          onApplyShapeStyle={onApplyShapeStyle}
          onApplyTextStyle={onApplyTextStyle}
          onArrange={onArrange}
          onDeleteShape={onDeleteShape}
        />
      )}

      {/* Row 2: Formula bar */}
      <div
        className="flex items-center gap-0 px-0"
        style={{
          height: 26,
          backgroundColor: 'var(--color-bg-primary)',
        }}
      >
        {/* Name box */}
        <div
          className="flex items-center justify-center flex-shrink-0 text-[11px] font-semibold tabular-nums select-none"
          style={{
            width: 52,
            height: '100%',
            borderRight: '1px solid var(--color-border)',
            color: 'var(--color-text-muted)',
            backgroundColor: 'var(--color-bg-secondary)',
          }}
        >
          {cellRef || ''}
        </div>

        {/* fx label */}
        <div
          className="flex items-center justify-center flex-shrink-0 text-[11px] italic select-none"
          style={{
            width: 24,
            height: '100%',
            borderRight: '1px solid var(--color-border-subtle)',
            color: 'var(--color-text-muted)',
          }}
        >
          fx
        </div>

        {/* Formula input */}
        <input
          ref={barInputRef}
          type="text"
          value={isEditing ? editText : (displayValue ?? '')}
          readOnly={!isEditing}
          onChange={(e) => onEditTextChange(e.target.value)}
          onFocus={onStartBarEdit}
          onKeyDown={onBarKeyDown}
          onMouseDown={(e) => {
            if (isEditing) e.stopPropagation();
          }}
          className="flex-1 min-w-0 h-full px-2 text-[12px] outline-none"
          style={{
            backgroundColor: isEditing ? 'var(--color-bg-primary)' : 'var(--color-bg-primary)',
            color: 'var(--color-text-primary)',
            cursor: isEditing ? 'text' : 'default',
          }}
        />
      </div>
    </div>
  );
}

function FormatButton({ icon: Icon, active, onClick, title }) {
  return (
    <button
      onClick={onClick}
      onMouseDown={(e) => e.preventDefault()}
      title={title}
      className="flex items-center justify-center w-6 h-6 rounded cursor-pointer transition-colors flex-shrink-0"
      style={{
        backgroundColor: active ? 'var(--color-accent-muted)' : 'transparent',
        color: active ? 'var(--color-accent)' : 'var(--color-text-muted)',
      }}
      onMouseEnter={(e) => {
        if (!active) {
          e.currentTarget.style.backgroundColor = 'var(--color-bg-hover)';
          e.currentTarget.style.color = 'var(--color-text-secondary)';
        }
      }}
      onMouseLeave={(e) => {
        e.currentTarget.style.backgroundColor = active ? 'var(--color-accent-muted)' : 'transparent';
        e.currentTarget.style.color = active ? 'var(--color-accent)' : 'var(--color-text-muted)';
      }}
    >
      <Icon size={13} />
    </button>
  );
}

function FontSizeInput({ value, onChange }) {
  const [draft, setDraft] = useState(String(value));
  useEffect(() => { setDraft(String(value)); }, [value]);

  // Commit on every change when the draft is a valid in-range number so
  // that font size updates apply in real time (important feedback when
  // editing shape text). On blur, snap back to the prop value if the
  // draft is still invalid or out of range.
  const handleChange = (e) => {
    const next = e.target.value;
    setDraft(next);
    const n = parseInt(next, 10);
    if (Number.isFinite(n) && n >= 6 && n <= 72) {
      onChange(n);
    }
  };

  const handleBlur = () => {
    const n = parseInt(draft, 10);
    if (!Number.isFinite(n) || n < 6 || n > 72) {
      setDraft(String(value));
    }
  };

  return (
    <input
      type="number"
      value={draft}
      min={6}
      max={72}
      onChange={handleChange}
      onBlur={handleBlur}
      onMouseDown={(e) => e.stopPropagation()}
      onKeyDown={(e) => { if (e.key === 'Enter') e.currentTarget.blur(); }}
      className="tabular-nums text-center font-medium bg-transparent outline-none border-b flex-shrink-0"
      style={{
        width: 28,
        fontSize: 11,
        color: 'var(--color-text-secondary)',
        borderColor: 'var(--color-border)',
        MozAppearance: 'textfield',
      }}
      title="Font size"
    />
  );
}

const PRESET_COLORS = [
  '#000000', '#434343', '#666666', '#999999', '#b7b7b7', '#cccccc', '#d9d9d9', '#efefef', '#f3f3f3', '#ffffff',
  '#980000', '#ff0000', '#ff9900', '#ffff00', '#00ff00', '#00ffff', '#4a86e8', '#0000ff', '#9900ff', '#ff00ff',
  '#e6b8af', '#f4cccc', '#fce5cd', '#fff2cc', '#d9ead3', '#d0e0e3', '#c9daf8', '#cfe2f3', '#d9d2e9', '#ead1dc',
  '#dd7e6b', '#ea9999', '#f9cb9c', '#ffe599', '#b6d7a8', '#a2c4c9', '#a4c2f4', '#9fc5e8', '#b4a7d6', '#d5a6bd',
  '#cc4125', '#e06666', '#f6b26b', '#ffd966', '#93c47d', '#76a5af', '#6d9eeb', '#6fa8dc', '#8e7cc3', '#c27ba0',
  '#a61c00', '#cc0000', '#e69138', '#f1c232', '#6aa84f', '#45818e', '#3c78d8', '#3d85c6', '#674ea7', '#a64d79',
  '#85200c', '#990000', '#b45f06', '#bf9000', '#38761d', '#134f5c', '#1155cc', '#0b5394', '#351c75', '#741b47',
  '#5b0f00', '#660000', '#783f04', '#7f6000', '#274e13', '#0c343d', '#1c4587', '#073763', '#20124d', '#4c1130',
];

function ColorPicker({ value, onChange, title, label, isBg }) {
  const [open, setOpen] = useState(false);
  const panelRef = useRef(null);
  const customInputRef = useRef(null);

  const close = useCallback(() => setOpen(false), []);

  useEffect(() => {
    if (!open) return;
    const handleClickOutside = (e) => {
      if (panelRef.current && !panelRef.current.contains(e.target)) close();
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [open, close]);

  const handleSelect = (hex) => {
    onChange(hex);
    close();
  };

  const underlineColor = value || (isBg ? 'var(--color-text-muted)' : 'var(--color-text-primary)');

  return (
    <div ref={panelRef} className="relative flex items-center flex-shrink-0">
      <button
        onClick={() => setOpen((v) => !v)}
        onMouseDown={(e) => e.preventDefault()}
        className="flex flex-col items-center justify-center w-7 h-6 rounded cursor-pointer transition-colors"
        style={{
          color: isBg ? 'var(--color-text-muted)' : (value || 'var(--color-text-primary)'),
          backgroundColor: 'transparent',
        }}
        onMouseEnter={(e) => { e.currentTarget.style.backgroundColor = 'var(--color-bg-hover)'; }}
        onMouseLeave={(e) => { e.currentTarget.style.backgroundColor = 'transparent'; }}
        title={title}
      >
        <span className="text-[11px] font-bold leading-none" style={{ marginBottom: 1 }}>{label}</span>
        <div style={{ width: 14, height: 3, backgroundColor: underlineColor, borderRadius: 1 }} />
      </button>

      {open && (
        <div
          className="absolute top-full left-0 mt-1 p-2 rounded-md shadow-lg z-50"
          style={{
            backgroundColor: 'var(--color-bg-secondary)',
            border: '1px solid var(--color-border)',
            width: 208,
          }}
        >
          <div className="grid gap-[3px]" style={{ gridTemplateColumns: 'repeat(10, 1fr)' }}>
            {PRESET_COLORS.map((hex) => (
              <button
                key={hex}
                onClick={() => handleSelect(hex)}
                onMouseDown={(e) => e.preventDefault()}
                className="rounded-sm cursor-pointer transition-transform hover:scale-125"
                style={{
                  width: 17,
                  height: 17,
                  backgroundColor: hex,
                  border: hex === '#ffffff' ? '1px solid var(--color-border)' : (value === hex ? '2px solid var(--color-accent)' : '1px solid transparent'),
                }}
                title={hex}
              />
            ))}
          </div>

          {/* No color (reset) + custom */}
          <div className="flex items-center justify-between mt-2 pt-2" style={{ borderTop: '1px solid var(--color-border-subtle)' }}>
            <button
              onClick={() => handleSelect('')}
              onMouseDown={(e) => e.preventDefault()}
              className="text-[10px] px-2 py-0.5 rounded cursor-pointer"
              style={{ color: 'var(--color-text-muted)' }}
              onMouseEnter={(e) => { e.currentTarget.style.backgroundColor = 'var(--color-bg-hover)'; }}
              onMouseLeave={(e) => { e.currentTarget.style.backgroundColor = 'transparent'; }}
            >
              No color
            </button>
            <button
              onClick={() => customInputRef.current?.click()}
              onMouseDown={(e) => e.preventDefault()}
              className="text-[10px] px-2 py-0.5 rounded cursor-pointer"
              style={{ color: 'var(--color-accent)' }}
              onMouseEnter={(e) => { e.currentTarget.style.backgroundColor = 'var(--color-bg-hover)'; }}
              onMouseLeave={(e) => { e.currentTarget.style.backgroundColor = 'transparent'; }}
            >
              Custom...
            </button>
            <input
              ref={customInputRef}
              type="color"
              value={value || '#000000'}
              onChange={(e) => handleSelect(e.target.value)}
              className="absolute opacity-0 w-0 h-0"
              tabIndex={-1}
            />
          </div>
        </div>
      )}
    </div>
  );
}

const THIN = { style: 'thin' };
const THICK = { style: 'thick' };
const DOUBLE = { style: 'double' };

const BORDER_SECTIONS = [
  [
    { id: 'bottom',  label: 'Bottom Border',  borders: { bottom: THIN } },
    { id: 'top',     label: 'Top Border',     borders: { top: THIN } },
    { id: 'left',    label: 'Left Border',    borders: { left: THIN } },
    { id: 'right',   label: 'Right Border',   borders: { right: THIN } },
  ],
  [
    { id: 'none',    label: 'No Border',       borders: null },
    { id: 'all',     label: 'All Borders',     borders: { top: THIN, bottom: THIN, left: THIN, right: THIN }, showInner: true },
    { id: 'outside', label: 'Outside Borders', borders: { top: THIN, bottom: THIN, left: THIN, right: THIN } },
  ],
  [
    { id: 'thick-outside',   label: 'Thick Outside Borders',      borders: { top: THICK, bottom: THICK, left: THICK, right: THICK } },
    { id: 'bottom-double',   label: 'Bottom Double Border',       borders: { bottom: DOUBLE } },
    { id: 'thick-bottom',    label: 'Thick Bottom Border',        borders: { bottom: THICK } },
  ],
  [
    { id: 'top-bottom',        label: 'Top and Bottom Border',              borders: { top: THIN, bottom: THIN } },
    { id: 'top-thick-bottom',  label: 'Top and Thick Bottom Border',       borders: { top: THIN, bottom: THICK } },
    { id: 'top-double-bottom', label: 'Top and Double Bottom Border',      borders: { top: THIN, bottom: DOUBLE } },
  ],
];

function BorderOptionIcon({ borders, showInner, showInnerDashed, size = 16 }) {
  const p = 2;
  const s = size - p * 2;
  const faint = 'var(--color-border-subtle)';
  // Darker dashed lines used by the toolbar button (showInnerDashed=true)
  // so the 2x2 cell-grid reads clearly next to the solid bottom border.
  // Menu items keep using the very faint outer rect so their border
  // previews remain the visual focus.
  const gridColor = 'var(--color-text-muted)';
  const mx = p + s / 2;
  const my = p + s / 2;

  function sw(b) {
    if (!b) return 0;
    if (b.style === 'thick') return 2.5;
    if (b.style === 'medium') return 2;
    if (b.style === 'double') return 1;
    return 1.2;
  }

  function da(b) {
    if (!b) return undefined;
    if (b.style === 'dashed') return '2 1';
    if (b.style === 'dotted') return '0.5 1.5';
    return undefined;
  }

  const isNone = !borders && !showInner;
  const outerColor = showInnerDashed ? gridColor : faint;
  const outerStrokeWidth = showInnerDashed ? 0.8 : 0.5;

  return (
    <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`} style={{ flexShrink: 0 }}>
      <rect x={p} y={p} width={s} height={s} fill="none" stroke={outerColor} strokeWidth={outerStrokeWidth} strokeDasharray="2 1" />
      {showInnerDashed && (<>
        <line x1={p} y1={my} x2={p + s} y2={my} stroke={gridColor} strokeWidth={0.8} strokeDasharray="2 1" />
        <line x1={mx} y1={p} x2={mx} y2={p + s} stroke={gridColor} strokeWidth={0.8} strokeDasharray="2 1" />
      </>)}
      {isNone && (
        <line x1={p} y1={p + s} x2={p + s} y2={p} stroke="var(--color-danger)" strokeWidth={1} opacity={0.5} />
      )}
      {borders?.top && (
        <line x1={p} y1={p} x2={p + s} y2={p} stroke="currentColor" strokeWidth={sw(borders.top)} strokeDasharray={da(borders.top)} />
      )}
      {borders?.bottom && borders.bottom.style !== 'double' && (
        <line x1={p} y1={p + s} x2={p + s} y2={p + s} stroke="currentColor" strokeWidth={sw(borders.bottom)} strokeDasharray={da(borders.bottom)} />
      )}
      {borders?.bottom?.style === 'double' && (<>
        <line x1={p} y1={p + s - 1.5} x2={p + s} y2={p + s - 1.5} stroke="currentColor" strokeWidth={0.8} />
        <line x1={p} y1={p + s + 0.5} x2={p + s} y2={p + s + 0.5} stroke="currentColor" strokeWidth={0.8} />
      </>)}
      {borders?.left && (
        <line x1={p} y1={p} x2={p} y2={p + s} stroke="currentColor" strokeWidth={sw(borders.left)} strokeDasharray={da(borders.left)} />
      )}
      {borders?.right && (
        <line x1={p + s} y1={p} x2={p + s} y2={p + s} stroke="currentColor" strokeWidth={sw(borders.right)} strokeDasharray={da(borders.right)} />
      )}
      {showInner && (<>
        <line x1={p} y1={my} x2={p + s} y2={my} stroke="currentColor" strokeWidth={1.2} />
        <line x1={mx} y1={p} x2={mx} y2={p + s} stroke="currentColor" strokeWidth={1.2} />
      </>)}
    </svg>
  );
}

function BorderPicker({ onApply }) {
  const [open, setOpen] = useState(false);
  const panelRef = useRef(null);
  const close = useCallback(() => setOpen(false), []);

  useEffect(() => {
    if (!open) return;
    const handleClickOutside = (e) => {
      if (panelRef.current && !panelRef.current.contains(e.target)) close();
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [open, close]);

  const handleSelect = (borders) => {
    onApply(borders);
    close();
  };

  return (
    <div ref={panelRef} className="relative flex items-center flex-shrink-0">
      <button
        onClick={() => setOpen((v) => !v)}
        onMouseDown={(e) => e.preventDefault()}
        className="flex items-center justify-center gap-0.5 h-6 px-1 rounded cursor-pointer transition-colors"
        style={{ color: 'var(--color-text-muted)', backgroundColor: 'transparent' }}
        onMouseEnter={(e) => { e.currentTarget.style.backgroundColor = 'var(--color-bg-hover)'; }}
        onMouseLeave={(e) => { e.currentTarget.style.backgroundColor = 'transparent'; }}
        title="Borders"
      >
        <BorderOptionIcon borders={{ bottom: THIN }} showInnerDashed size={16} />
        <svg width={8} height={8} viewBox="0 0 8 8" style={{ opacity: 0.5 }}>
          <path d="M1 3 L4 6 L7 3" fill="none" stroke="currentColor" strokeWidth={1.2} />
        </svg>
      </button>

      {open && (
        <div
          className="absolute top-full left-0 mt-1 py-1 rounded-md shadow-lg z-50"
          style={{
            backgroundColor: 'var(--color-bg-secondary)',
            border: '1px solid var(--color-border)',
            width: 220,
          }}
        >
          {BORDER_SECTIONS.map((section, si) => (
            <div key={si}>
              {si > 0 && (
                <div className="my-1 mx-2" style={{ height: 1, backgroundColor: 'var(--color-border-subtle)' }} />
              )}
              {section.map((opt) => (
                <button
                  key={opt.id}
                  onClick={() => handleSelect(opt.id)}
                  onMouseDown={(e) => e.preventDefault()}
                  className="flex items-center gap-2 w-full px-3 py-1 text-left text-[11px] cursor-pointer transition-colors"
                  style={{ color: 'var(--color-text-secondary)' }}
                  onMouseEnter={(e) => { e.currentTarget.style.backgroundColor = 'var(--color-bg-hover)'; e.currentTarget.style.color = 'var(--color-text-primary)'; }}
                  onMouseLeave={(e) => { e.currentTarget.style.backgroundColor = 'transparent'; e.currentTarget.style.color = 'var(--color-text-secondary)'; }}
                >
                  <BorderOptionIcon borders={opt.borders} showInner={opt.showInner} size={16} />
                  <span>{opt.label}</span>
                </button>
              ))}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// Excel-style merge icons. Each draws a 2x2 (or related) cell grid outline
// to communicate the action visually, rather than Lucide's generic graph
// glyphs. Sized at 13px to match the surrounding FormatButton icons.
function MergeCenterIcon({ size = 13 }) {
  return (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.1" strokeLinecap="round" strokeLinejoin="round">
      <rect x="1.5" y="3.5" width="13" height="9" />
      <path d="M3 8 L6.5 8" />
      <path d="M13 8 L9.5 8" />
      <path d="M5 6.5 L3 8 L5 9.5" />
      <path d="M11 6.5 L13 8 L11 9.5" />
    </svg>
  );
}

function MergeAcrossIcon({ size = 13 }) {
  return (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.1" strokeLinecap="round" strokeLinejoin="round">
      <rect x="1.5" y="2" width="13" height="5" />
      <rect x="1.5" y="9" width="13" height="5" />
      <path d="M3 4.5 L6.5 4.5 M13 4.5 L9.5 4.5" />
      <path d="M4.5 3.5 L3 4.5 L4.5 5.5 M11.5 3.5 L13 4.5 L11.5 5.5" />
      <path d="M3 11.5 L6.5 11.5 M13 11.5 L9.5 11.5" />
      <path d="M4.5 10.5 L3 11.5 L4.5 12.5 M11.5 10.5 L13 11.5 L11.5 12.5" />
    </svg>
  );
}

function MergeCellsIcon({ size = 13 }) {
  return (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.1" strokeLinecap="round" strokeLinejoin="round">
      <rect x="1.5" y="2.5" width="13" height="11" />
      <path d="M1.5 8 L14.5 8" strokeDasharray="1.2 1.2" />
      <path d="M8 2.5 L8 13.5" strokeDasharray="1.2 1.2" />
    </svg>
  );
}

function UnmergeIcon({ size = 13 }) {
  return (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.1" strokeLinecap="round" strokeLinejoin="round">
      <rect x="1.5" y="2.5" width="13" height="11" />
      <path d="M8 2.5 L8 13.5" />
      <path d="M5 6 L8 8 L11 6" />
      <path d="M5 10 L8 8 L11 10" />
    </svg>
  );
}

const MERGE_OPTIONS = [
  { id: 'center',  label: 'Merge & Center', Icon: MergeCenterIcon },
  { id: 'across',  label: 'Merge Across',   Icon: MergeAcrossIcon },
  { id: 'cells',   label: 'Merge Cells',    Icon: MergeCellsIcon },
  { id: 'unmerge', label: 'Unmerge Cells',  Icon: UnmergeIcon },
];

function MergeMenu({ onMerge, onUnmerge, isMergedSelection }) {
  const [open, setOpen] = useState(false);
  const panelRef = useRef(null);
  const close = useCallback(() => setOpen(false), []);

  useEffect(() => {
    if (!open) return;
    const handleClickOutside = (e) => {
      if (panelRef.current && !panelRef.current.contains(e.target)) close();
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [open, close]);

  const handleDefaultClick = () => {
    if (isMergedSelection) {
      onUnmerge?.();
    } else {
      onMerge?.('center');
    }
  };

  const handleSelect = (id) => {
    close();
    if (id === 'unmerge') onUnmerge?.();
    else onMerge?.(id);
  };

  const MainIcon = isMergedSelection ? UnmergeIcon : MergeCenterIcon;
  const mainTitle = isMergedSelection ? 'Unmerge cells' : 'Merge & Center';

  return (
    <div ref={panelRef} className="relative flex items-center flex-shrink-0">
      <button
        onClick={handleDefaultClick}
        onMouseDown={(e) => e.preventDefault()}
        title={mainTitle}
        className="flex items-center justify-center w-6 h-6 rounded-l cursor-pointer transition-colors"
        style={{ color: 'var(--color-text-muted)', backgroundColor: 'transparent' }}
        onMouseEnter={(e) => { e.currentTarget.style.backgroundColor = 'var(--color-bg-hover)'; e.currentTarget.style.color = 'var(--color-text-secondary)'; }}
        onMouseLeave={(e) => { e.currentTarget.style.backgroundColor = 'transparent'; e.currentTarget.style.color = 'var(--color-text-muted)'; }}
      >
        <MainIcon size={13} />
      </button>
      <button
        onClick={() => setOpen((v) => !v)}
        onMouseDown={(e) => e.preventDefault()}
        title="Merge options"
        className="flex items-center justify-center h-6 rounded-r cursor-pointer transition-colors"
        style={{ width: 12, color: 'var(--color-text-muted)', backgroundColor: 'transparent' }}
        onMouseEnter={(e) => { e.currentTarget.style.backgroundColor = 'var(--color-bg-hover)'; e.currentTarget.style.color = 'var(--color-text-secondary)'; }}
        onMouseLeave={(e) => { e.currentTarget.style.backgroundColor = 'transparent'; e.currentTarget.style.color = 'var(--color-text-muted)'; }}
      >
        <ChevronDown size={10} />
      </button>

      {open && (
        <div
          className="absolute top-full left-0 mt-1 py-1 rounded-md shadow-lg z-50"
          style={{
            backgroundColor: 'var(--color-bg-secondary)',
            border: '1px solid var(--color-border)',
            width: 180,
          }}
        >
          {MERGE_OPTIONS.map((opt) => (
            <button
              key={opt.id}
              onClick={() => handleSelect(opt.id)}
              onMouseDown={(e) => e.preventDefault()}
              className="flex items-center gap-2 w-full px-3 py-1 text-left text-[11px] cursor-pointer transition-colors"
              style={{ color: 'var(--color-text-secondary)' }}
              onMouseEnter={(e) => { e.currentTarget.style.backgroundColor = 'var(--color-bg-hover)'; e.currentTarget.style.color = 'var(--color-text-primary)'; }}
              onMouseLeave={(e) => { e.currentTarget.style.backgroundColor = 'transparent'; e.currentTarget.style.color = 'var(--color-text-secondary)'; }}
            >
              <opt.Icon size={13} />
              <span>{opt.label}</span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

function InsertShapesMenu({ shapeMode, onSetShapeMode }) {
  const [open, setOpen] = useState(false);
  const panelRef = useRef(null);
  const close = useCallback(() => setOpen(false), []);

  useEffect(() => {
    if (!open) return;
    const handleClickOutside = (e) => {
      if (panelRef.current && !panelRef.current.contains(e.target)) close();
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [open, close]);

  const groups = groupShapePresets();
  const active = !!shapeMode;

  const handleSelect = (presetId) => {
    onSetShapeMode?.(presetId);
    close();
  };

  return (
    <div ref={panelRef} className="relative flex items-center flex-shrink-0">
      <button
        onClick={() => setOpen((v) => !v)}
        onMouseDown={(e) => e.preventDefault()}
        title="Insert shapes"
        className="flex items-center justify-center gap-0.5 h-6 px-1 rounded cursor-pointer transition-colors"
        style={{
          color: active ? 'var(--color-accent)' : 'var(--color-text-muted)',
          backgroundColor: active ? 'var(--color-accent-muted)' : 'transparent',
        }}
        onMouseEnter={(e) => { if (!active) e.currentTarget.style.backgroundColor = 'var(--color-bg-hover)'; }}
        onMouseLeave={(e) => { e.currentTarget.style.backgroundColor = active ? 'var(--color-accent-muted)' : 'transparent'; }}
      >
        <Shapes size={13} />
        <ChevronDown size={9} style={{ opacity: 0.6 }} />
      </button>
      {open && (
        <div
          className="absolute top-full left-0 mt-1 rounded-md shadow-lg z-50"
          style={{
            backgroundColor: 'var(--color-bg-secondary)',
            border: '1px solid var(--color-border)',
            width: 248,
            maxHeight: 420,
            overflowY: 'auto',
          }}
        >
          {Object.entries(groups).map(([groupName, presets]) => (
            <div key={groupName}>
              <div
                className="px-2 py-1 text-[10px] font-semibold uppercase tracking-wide select-none"
                style={{
                  color: 'var(--color-text-muted)',
                  backgroundColor: 'var(--color-bg-tertiary)',
                  borderBottom: '1px solid var(--color-border-subtle)',
                }}
              >
                {groupName}
              </div>
              <div className="grid gap-1 p-2" style={{ gridTemplateColumns: 'repeat(8, 1fr)' }}>
                {presets.map((p) => (
                  <button
                    key={p.id}
                    onClick={() => handleSelect(p.id)}
                    onMouseDown={(e) => e.preventDefault()}
                    title={p.label}
                    className="flex items-center justify-center rounded cursor-pointer transition-colors"
                    style={{
                      width: 24,
                      height: 24,
                      color: 'var(--color-text-secondary)',
                      backgroundColor: shapeMode === p.id ? 'var(--color-accent-muted)' : 'transparent',
                      border: '1px solid var(--color-border-subtle)',
                    }}
                    onMouseEnter={(e) => { e.currentTarget.style.backgroundColor = 'var(--color-bg-hover)'; }}
                    onMouseLeave={(e) => { e.currentTarget.style.backgroundColor = shapeMode === p.id ? 'var(--color-accent-muted)' : 'transparent'; }}
                  >
                    <ShapeGlyph presetId={p.id} size={16} />
                  </button>
                ))}
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function ShapeGlyph({ presetId, size = 16 }) {
  const preset = getShapePreset(presetId);
  const padding = 2;
  const inner = size - padding * 2;
  const d = renderShapePath(presetId, inner, inner);
  const isLine = isLineShape(presetId);
  const arrowEnd = preset?.arrowEnd;
  const arrowStart = preset?.arrowStart;
  const uid = `shape-glyph-${presetId}`;

  // Text-box variants render as just a "T" / "T|" glyph since the shape
  // itself has no fill and no outline and would be invisible otherwise.
  if (preset?.isTextBox) {
    return (
      <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`} style={{ flexShrink: 0 }}>
        <rect x={1.5} y={1.5} width={size - 3} height={size - 3} fill="none" stroke="currentColor" strokeWidth={0.6} strokeDasharray="1.5 1" />
        <text
          x={size / 2}
          y={size / 2}
          fontSize={size * 0.65}
          fontWeight={700}
          textAnchor="middle"
          dominantBaseline="central"
          fill="currentColor"
          transform={preset.vertical ? `rotate(90 ${size / 2} ${size / 2})` : undefined}
        >T</text>
      </svg>
    );
  }

  return (
    <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`} style={{ flexShrink: 0 }}>
      {(arrowEnd || arrowStart) && (
        <defs>
          {arrowEnd && (
            <marker id={`${uid}-end`} viewBox="0 0 10 10" refX="9" refY="5" markerWidth="4" markerHeight="4" orient="auto-start-reverse">
              <path d="M0,0 L10,5 L0,10 Z" fill="currentColor" />
            </marker>
          )}
          {arrowStart && (
            <marker id={`${uid}-start`} viewBox="0 0 10 10" refX="1" refY="5" markerWidth="4" markerHeight="4" orient="auto">
              <path d="M10,0 L0,5 L10,10 Z" fill="currentColor" />
            </marker>
          )}
        </defs>
      )}
      <g transform={`translate(${padding} ${padding})`}>
        <path
          d={d}
          fill={isLine ? 'none' : 'currentColor'}
          stroke="currentColor"
          strokeWidth={isLine ? 1.2 : 0.7}
          strokeLinejoin="round"
          strokeLinecap="round"
          markerEnd={arrowEnd ? `url(#${uid}-end)` : undefined}
          markerStart={arrowStart ? `url(#${uid}-start)` : undefined}
        />
      </g>
    </svg>
  );
}

const SHAPE_DASH_OPTIONS = [
  { id: 'solid', label: 'Solid' },
  { id: 'dash', label: 'Dash' },
  { id: 'dot', label: 'Dot' },
  { id: 'dashDot', label: 'Dash-Dot' },
  { id: 'longDash', label: 'Long Dash' },
];

const SHAPE_WIDTH_OPTIONS = [0.5, 1, 1.5, 2, 3, 4, 6];

function ShapeToolbarRow({ shape, onApplyShapeStyle, onApplyTextStyle, onArrange, onDeleteShape }) {
  const style = shape?.style || {};
  const textStyle = shape?.textStyle || {};
  const fill = style.fill || {};
  const outline = style.outline || {};
  const effects = style.effects || {};
  const tFill = textStyle.fill || {};

  const disabled = !shape;

  return (
    <div
      className="flex items-center gap-1 px-2"
      style={{
        height: 28,
        borderBottom: '1px solid var(--color-border-subtle)',
        backgroundColor: 'var(--color-bg-tertiary)',
      }}
    >
      <span
        className="text-[10px] font-semibold uppercase tracking-wide select-none pr-1"
        style={{ color: 'var(--color-text-muted)' }}
      >
        Shape
      </span>
      <ShapeColorPicker
        value={fill.color || ''}
        alpha={fill.alpha != null ? fill.alpha : 1}
        disabled={disabled}
        showAlpha
        onChange={(color, alpha) => onApplyShapeStyle?.({ fill: { type: color ? 'solid' : 'none', color, alpha } })}
        title="Shape Fill"
        label="Fill"
      />
      <ShapeColorPicker
        value={outline.color || ''}
        alpha={outline.alpha != null ? outline.alpha : 1}
        disabled={disabled}
        onChange={(color, alpha) => onApplyShapeStyle?.({ outline: { ...outline, color, alpha } })}
        title="Shape Outline"
        label="Outline"
      />
      <ShapeOutlineMenu
        outline={outline}
        disabled={disabled}
        onChange={(next) => onApplyShapeStyle?.({ outline: { ...outline, ...next } })}
      />
      <ShapeEffectsMenu
        effects={effects}
        disabled={disabled}
        onChange={(next) => onApplyShapeStyle?.({ effects: next })}
      />

      <div className="w-px h-4 mx-0.5 flex-shrink-0" style={{ backgroundColor: 'var(--color-border)' }} />

      <span
        className="text-[10px] font-semibold uppercase tracking-wide select-none pr-1"
        style={{ color: 'var(--color-text-muted)' }}
      >
        Text
      </span>
      <ShapeColorPicker
        value={tFill.color || ''}
        alpha={tFill.alpha != null ? tFill.alpha : 1}
        disabled={disabled}
        onChange={(color, alpha) => onApplyTextStyle?.({ fill: { color, alpha } })}
        title="Text Fill"
        label="A"
      />
      <ShapeColorPicker
        value={textStyle.outline?.color || ''}
        alpha={textStyle.outline?.alpha != null ? textStyle.outline.alpha : 1}
        disabled={disabled}
        onChange={(color, alpha) => onApplyTextStyle?.({ outline: { ...(textStyle.outline || {}), color, alpha, width: textStyle.outline?.width || 1 } })}
        title="Text Outline"
        label="Ao"
      />
      <TextEffectsMenu
        effects={textStyle.effects || {}}
        disabled={disabled}
        onChange={(next) => onApplyTextStyle?.({ effects: next })}
      />

      <div className="w-px h-4 mx-0.5 flex-shrink-0" style={{ backgroundColor: 'var(--color-border)' }} />

      <FormatButton
        icon={BringToFront}
        active={false}
        onClick={() => onArrange?.('front')}
        title="Bring to front"
      />
      <FormatButton
        icon={SendToBack}
        active={false}
        onClick={() => onArrange?.('back')}
        title="Send to back"
      />
      <FormatButton
        icon={MoveUp}
        active={false}
        onClick={() => onArrange?.('forward')}
        title="Bring forward"
      />
      <FormatButton
        icon={MoveDown}
        active={false}
        onClick={() => onArrange?.('backward')}
        title="Send backward"
      />

      <div className="w-px h-4 mx-0.5 flex-shrink-0" style={{ backgroundColor: 'var(--color-border)' }} />

      <FormatButton
        icon={Trash2}
        active={false}
        onClick={onDeleteShape}
        title="Delete shape"
      />
    </div>
  );
}

function ShapeColorPicker({ value, alpha, onChange, title, label, showAlpha, disabled }) {
  const [open, setOpen] = useState(false);
  const panelRef = useRef(null);
  const customInputRef = useRef(null);

  const close = useCallback(() => setOpen(false), []);

  useEffect(() => {
    if (!open) return;
    const handleClickOutside = (e) => {
      if (panelRef.current && !panelRef.current.contains(e.target)) close();
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [open, close]);

  const handleSelect = (hex) => {
    onChange?.(hex, alpha);
    close();
  };

  const underlineColor = value || 'var(--color-text-muted)';

  return (
    <div ref={panelRef} className="relative flex items-center flex-shrink-0">
      <button
        disabled={disabled}
        onClick={() => !disabled && setOpen((v) => !v)}
        onMouseDown={(e) => e.preventDefault()}
        className="flex flex-col items-center justify-center h-6 rounded cursor-pointer transition-colors"
        style={{
          width: label && label.length > 2 ? 38 : 28,
          color: value || 'var(--color-text-primary)',
          backgroundColor: 'transparent',
          opacity: disabled ? 0.4 : 1,
          cursor: disabled ? 'not-allowed' : 'pointer',
        }}
        onMouseEnter={(e) => { if (!disabled) e.currentTarget.style.backgroundColor = 'var(--color-bg-hover)'; }}
        onMouseLeave={(e) => { e.currentTarget.style.backgroundColor = 'transparent'; }}
        title={title}
      >
        <span className="text-[10px] font-semibold leading-none" style={{ marginBottom: 1 }}>{label}</span>
        <div style={{ width: 18, height: 3, backgroundColor: underlineColor, borderRadius: 1 }} />
      </button>

      {open && (
        <div
          className="absolute top-full left-0 mt-1 p-2 rounded-md shadow-lg z-50"
          style={{
            backgroundColor: 'var(--color-bg-secondary)',
            border: '1px solid var(--color-border)',
            width: 208,
          }}
        >
          <div className="grid gap-[3px]" style={{ gridTemplateColumns: 'repeat(10, 1fr)' }}>
            {PRESET_COLORS.map((hex) => (
              <button
                key={hex}
                onClick={() => handleSelect(hex)}
                onMouseDown={(e) => e.preventDefault()}
                className="rounded-sm cursor-pointer transition-transform hover:scale-125"
                style={{
                  width: 17,
                  height: 17,
                  backgroundColor: hex,
                  border: hex === '#ffffff' ? '1px solid var(--color-border)' : (value === hex ? '2px solid var(--color-accent)' : '1px solid transparent'),
                }}
                title={hex}
              />
            ))}
          </div>
          {showAlpha && (
            <div className="flex items-center gap-2 mt-2 pt-2" style={{ borderTop: '1px solid var(--color-border-subtle)' }}>
              <span className="text-[10px]" style={{ color: 'var(--color-text-muted)' }}>Opacity</span>
              <input
                type="range"
                min={0}
                max={100}
                value={Math.round((alpha == null ? 1 : alpha) * 100)}
                onChange={(e) => onChange?.(value, parseInt(e.target.value, 10) / 100)}
                className="flex-1"
              />
              <span className="text-[10px] tabular-nums" style={{ color: 'var(--color-text-muted)', width: 24, textAlign: 'right' }}>
                {Math.round((alpha == null ? 1 : alpha) * 100)}%
              </span>
            </div>
          )}
          <div className="flex items-center justify-between mt-2 pt-2" style={{ borderTop: '1px solid var(--color-border-subtle)' }}>
            <button
              onClick={() => handleSelect('')}
              onMouseDown={(e) => e.preventDefault()}
              className="text-[10px] px-2 py-0.5 rounded cursor-pointer"
              style={{ color: 'var(--color-text-muted)' }}
              onMouseEnter={(e) => { e.currentTarget.style.backgroundColor = 'var(--color-bg-hover)'; }}
              onMouseLeave={(e) => { e.currentTarget.style.backgroundColor = 'transparent'; }}
            >
              No color
            </button>
            <button
              onClick={() => customInputRef.current?.click()}
              onMouseDown={(e) => e.preventDefault()}
              className="text-[10px] px-2 py-0.5 rounded cursor-pointer"
              style={{ color: 'var(--color-accent)' }}
              onMouseEnter={(e) => { e.currentTarget.style.backgroundColor = 'var(--color-bg-hover)'; }}
              onMouseLeave={(e) => { e.currentTarget.style.backgroundColor = 'transparent'; }}
            >
              Custom...
            </button>
            <input
              ref={customInputRef}
              type="color"
              value={value || '#000000'}
              onChange={(e) => handleSelect(e.target.value)}
              className="absolute opacity-0 w-0 h-0"
              tabIndex={-1}
            />
          </div>
        </div>
      )}
    </div>
  );
}

function ShapeOutlineMenu({ outline, disabled, onChange }) {
  const [open, setOpen] = useState(false);
  const panelRef = useRef(null);
  const close = useCallback(() => setOpen(false), []);

  useEffect(() => {
    if (!open) return;
    const handleClickOutside = (e) => {
      if (panelRef.current && !panelRef.current.contains(e.target)) close();
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [open, close]);

  const currentWidth = outline?.width != null ? outline.width : 1;
  const currentDash = outline?.dash || 'solid';

  return (
    <div ref={panelRef} className="relative flex items-center flex-shrink-0">
      <button
        disabled={disabled}
        onClick={() => !disabled && setOpen((v) => !v)}
        onMouseDown={(e) => e.preventDefault()}
        className="flex items-center justify-center h-6 px-1 rounded cursor-pointer transition-colors gap-0.5"
        style={{ color: 'var(--color-text-muted)', opacity: disabled ? 0.4 : 1 }}
        onMouseEnter={(e) => { if (!disabled) e.currentTarget.style.backgroundColor = 'var(--color-bg-hover)'; }}
        onMouseLeave={(e) => { e.currentTarget.style.backgroundColor = 'transparent'; }}
        title="Outline weight & style"
      >
        <svg width={16} height={12} viewBox="0 0 16 12">
          <line x1={1} y1={6} x2={15} y2={6} stroke="currentColor" strokeWidth={currentWidth} strokeDasharray={currentDash === 'solid' ? undefined : '3 2'} />
        </svg>
        <ChevronDown size={9} style={{ opacity: 0.6 }} />
      </button>
      {open && (
        <div
          className="absolute top-full left-0 mt-1 py-1 rounded-md shadow-lg z-50"
          style={{
            backgroundColor: 'var(--color-bg-secondary)',
            border: '1px solid var(--color-border)',
            width: 180,
          }}
        >
          <div className="px-2 py-1 text-[10px] font-semibold uppercase tracking-wide" style={{ color: 'var(--color-text-muted)' }}>
            Weight
          </div>
          {SHAPE_WIDTH_OPTIONS.map((w) => (
            <button
              key={w}
              onClick={() => { onChange?.({ width: w }); close(); }}
              onMouseDown={(e) => e.preventDefault()}
              className="flex items-center gap-3 w-full px-3 py-1 text-left text-[11px] cursor-pointer transition-colors"
              style={{
                color: 'var(--color-text-secondary)',
                backgroundColor: currentWidth === w ? 'var(--color-accent-muted)' : 'transparent',
              }}
              onMouseEnter={(e) => { e.currentTarget.style.backgroundColor = 'var(--color-bg-hover)'; }}
              onMouseLeave={(e) => { e.currentTarget.style.backgroundColor = currentWidth === w ? 'var(--color-accent-muted)' : 'transparent'; }}
            >
              <svg width={36} height={8} viewBox="0 0 36 8">
                <line x1={0} y1={4} x2={36} y2={4} stroke="currentColor" strokeWidth={w} />
              </svg>
              <span className="tabular-nums">{w} pt</span>
            </button>
          ))}
          <div className="my-1 mx-2" style={{ height: 1, backgroundColor: 'var(--color-border-subtle)' }} />
          <div className="px-2 py-1 text-[10px] font-semibold uppercase tracking-wide" style={{ color: 'var(--color-text-muted)' }}>
            Dash
          </div>
          {SHAPE_DASH_OPTIONS.map((opt) => (
            <button
              key={opt.id}
              onClick={() => { onChange?.({ dash: opt.id }); close(); }}
              onMouseDown={(e) => e.preventDefault()}
              className="flex items-center gap-3 w-full px-3 py-1 text-left text-[11px] cursor-pointer transition-colors"
              style={{
                color: 'var(--color-text-secondary)',
                backgroundColor: currentDash === opt.id ? 'var(--color-accent-muted)' : 'transparent',
              }}
              onMouseEnter={(e) => { e.currentTarget.style.backgroundColor = 'var(--color-bg-hover)'; }}
              onMouseLeave={(e) => { e.currentTarget.style.backgroundColor = currentDash === opt.id ? 'var(--color-accent-muted)' : 'transparent'; }}
            >
              <svg width={36} height={8} viewBox="0 0 36 8">
                <line x1={0} y1={4} x2={36} y2={4} stroke="currentColor" strokeWidth={1.5} strokeDasharray={dashPreview(opt.id)} />
              </svg>
              <span>{opt.label}</span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

function dashPreview(id) {
  switch (id) {
    case 'dash': return '6 3';
    case 'dot': return '2 2';
    case 'dashDot': return '6 2 2 2';
    case 'longDash': return '10 3';
    default: return undefined;
  }
}

function ShapeEffectsMenu({ effects, disabled, onChange }) {
  const [open, setOpen] = useState(false);
  const panelRef = useRef(null);
  const close = useCallback(() => setOpen(false), []);

  useEffect(() => {
    if (!open) return;
    const handleClickOutside = (e) => {
      if (panelRef.current && !panelRef.current.contains(e.target)) close();
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [open, close]);

  const hasShadow = !!effects?.shadow;
  const hasGlow = !!effects?.glow;
  const hasSoft = !!effects?.softEdge;
  const any = hasShadow || hasGlow || hasSoft;

  const toggle = (key, defaultVal) => {
    const next = { ...(effects || {}) };
    if (next[key]) delete next[key];
    else next[key] = defaultVal;
    onChange?.(next);
  };

  return (
    <div ref={panelRef} className="relative flex items-center flex-shrink-0">
      <button
        disabled={disabled}
        onClick={() => !disabled && setOpen((v) => !v)}
        onMouseDown={(e) => e.preventDefault()}
        className="flex items-center justify-center h-6 px-1 rounded cursor-pointer transition-colors gap-0.5"
        style={{
          color: any ? 'var(--color-accent)' : 'var(--color-text-muted)',
          backgroundColor: any ? 'var(--color-accent-muted)' : 'transparent',
          opacity: disabled ? 0.4 : 1,
        }}
        onMouseEnter={(e) => { if (!disabled) e.currentTarget.style.backgroundColor = 'var(--color-bg-hover)'; }}
        onMouseLeave={(e) => { e.currentTarget.style.backgroundColor = any ? 'var(--color-accent-muted)' : 'transparent'; }}
        title="Shape effects"
      >
        <Sparkles size={13} />
        <ChevronDown size={9} style={{ opacity: 0.6 }} />
      </button>
      {open && (
        <div
          className="absolute top-full left-0 mt-1 py-2 rounded-md shadow-lg z-50"
          style={{
            backgroundColor: 'var(--color-bg-secondary)',
            border: '1px solid var(--color-border)',
            width: 240,
          }}
        >
          <EffectSection
            label="Shadow"
            enabled={hasShadow}
            onToggle={() => toggle('shadow', { color: '#000000', alpha: 0.35, offsetX: 3, offsetY: 3, blur: 4 })}
          >
            {hasShadow && (
              <EffectSliders
                rows={[
                  { label: 'Offset X', value: effects.shadow.offsetX ?? 3, min: -20, max: 20, onChange: (v) => onChange?.({ ...effects, shadow: { ...effects.shadow, offsetX: v } }) },
                  { label: 'Offset Y', value: effects.shadow.offsetY ?? 3, min: -20, max: 20, onChange: (v) => onChange?.({ ...effects, shadow: { ...effects.shadow, offsetY: v } }) },
                  { label: 'Blur', value: effects.shadow.blur ?? 4, min: 0, max: 30, onChange: (v) => onChange?.({ ...effects, shadow: { ...effects.shadow, blur: v } }) },
                  { label: 'Opacity', value: Math.round((effects.shadow.alpha ?? 0.35) * 100), min: 0, max: 100, onChange: (v) => onChange?.({ ...effects, shadow: { ...effects.shadow, alpha: v / 100 } }) },
                ]}
              />
            )}
          </EffectSection>
          <EffectSection
            label="Glow"
            enabled={hasGlow}
            onToggle={() => toggle('glow', { color: '#ffea00', alpha: 0.6, radius: 5 })}
          >
            {hasGlow && (
              <EffectSliders
                rows={[
                  { label: 'Radius', value: effects.glow.radius ?? 5, min: 0, max: 30, onChange: (v) => onChange?.({ ...effects, glow: { ...effects.glow, radius: v } }) },
                  { label: 'Opacity', value: Math.round((effects.glow.alpha ?? 0.6) * 100), min: 0, max: 100, onChange: (v) => onChange?.({ ...effects, glow: { ...effects.glow, alpha: v / 100 } }) },
                ]}
              />
            )}
          </EffectSection>
          <EffectSection
            label="Soft Edges"
            enabled={hasSoft}
            onToggle={() => toggle('softEdge', { radius: 3 })}
          >
            {hasSoft && (
              <EffectSliders
                rows={[
                  { label: 'Radius', value: effects.softEdge.radius ?? 3, min: 0, max: 20, onChange: (v) => onChange?.({ ...effects, softEdge: { ...effects.softEdge, radius: v } }) },
                ]}
              />
            )}
          </EffectSection>
        </div>
      )}
    </div>
  );
}

function TextEffectsMenu({ effects, disabled, onChange }) {
  const [open, setOpen] = useState(false);
  const panelRef = useRef(null);
  const close = useCallback(() => setOpen(false), []);

  useEffect(() => {
    if (!open) return;
    const handleClickOutside = (e) => {
      if (panelRef.current && !panelRef.current.contains(e.target)) close();
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [open, close]);

  const hasShadow = !!effects?.shadow;
  const hasGlow = !!effects?.glow;
  const any = hasShadow || hasGlow;

  const toggle = (key, defaultVal) => {
    const next = { ...(effects || {}) };
    if (next[key]) delete next[key];
    else next[key] = defaultVal;
    onChange?.(next);
  };

  return (
    <div ref={panelRef} className="relative flex items-center flex-shrink-0">
      <button
        disabled={disabled}
        onClick={() => !disabled && setOpen((v) => !v)}
        onMouseDown={(e) => e.preventDefault()}
        className="flex items-center justify-center h-6 px-1 rounded cursor-pointer transition-colors gap-0.5"
        style={{
          color: any ? 'var(--color-accent)' : 'var(--color-text-muted)',
          backgroundColor: any ? 'var(--color-accent-muted)' : 'transparent',
          opacity: disabled ? 0.4 : 1,
        }}
        onMouseEnter={(e) => { if (!disabled) e.currentTarget.style.backgroundColor = 'var(--color-bg-hover)'; }}
        onMouseLeave={(e) => { e.currentTarget.style.backgroundColor = any ? 'var(--color-accent-muted)' : 'transparent'; }}
        title="Text effects"
      >
        <span className="text-[10px] font-semibold">Afx</span>
        <ChevronDown size={9} style={{ opacity: 0.6 }} />
      </button>
      {open && (
        <div
          className="absolute top-full left-0 mt-1 py-2 rounded-md shadow-lg z-50"
          style={{
            backgroundColor: 'var(--color-bg-secondary)',
            border: '1px solid var(--color-border)',
            width: 240,
          }}
        >
          <EffectSection
            label="Shadow"
            enabled={hasShadow}
            onToggle={() => toggle('shadow', { color: '#000000', alpha: 0.6, offsetX: 1, offsetY: 1, blur: 2 })}
          >
            {hasShadow && (
              <EffectSliders
                rows={[
                  { label: 'Offset X', value: effects.shadow.offsetX ?? 1, min: -10, max: 10, onChange: (v) => onChange?.({ ...effects, shadow: { ...effects.shadow, offsetX: v } }) },
                  { label: 'Offset Y', value: effects.shadow.offsetY ?? 1, min: -10, max: 10, onChange: (v) => onChange?.({ ...effects, shadow: { ...effects.shadow, offsetY: v } }) },
                  { label: 'Blur', value: effects.shadow.blur ?? 2, min: 0, max: 20, onChange: (v) => onChange?.({ ...effects, shadow: { ...effects.shadow, blur: v } }) },
                  { label: 'Opacity', value: Math.round((effects.shadow.alpha ?? 0.6) * 100), min: 0, max: 100, onChange: (v) => onChange?.({ ...effects, shadow: { ...effects.shadow, alpha: v / 100 } }) },
                ]}
              />
            )}
          </EffectSection>
          <EffectSection
            label="Glow"
            enabled={hasGlow}
            onToggle={() => toggle('glow', { color: '#ffea00', alpha: 0.8, radius: 3 })}
          >
            {hasGlow && (
              <EffectSliders
                rows={[
                  { label: 'Radius', value: effects.glow.radius ?? 3, min: 0, max: 20, onChange: (v) => onChange?.({ ...effects, glow: { ...effects.glow, radius: v } }) },
                  { label: 'Opacity', value: Math.round((effects.glow.alpha ?? 0.8) * 100), min: 0, max: 100, onChange: (v) => onChange?.({ ...effects, glow: { ...effects.glow, alpha: v / 100 } }) },
                ]}
              />
            )}
          </EffectSection>
        </div>
      )}
    </div>
  );
}

function EffectSection({ label, enabled, onToggle, children }) {
  return (
    <div className="px-3 py-1.5">
      <label className="flex items-center gap-2 text-[11px] cursor-pointer" style={{ color: 'var(--color-text-primary)' }}>
        <input type="checkbox" checked={enabled} onChange={onToggle} />
        <span className="font-medium">{label}</span>
      </label>
      {children}
    </div>
  );
}

function EffectSliders({ rows }) {
  return (
    <div className="pl-5 pt-1 flex flex-col gap-1">
      {rows.map((r) => (
        <div key={r.label} className="flex items-center gap-2">
          <span className="text-[10px]" style={{ color: 'var(--color-text-muted)', width: 54 }}>{r.label}</span>
          <input
            type="range"
            min={r.min}
            max={r.max}
            value={r.value}
            onChange={(e) => r.onChange(parseInt(e.target.value, 10))}
            className="flex-1"
          />
          <span className="text-[10px] tabular-nums" style={{ color: 'var(--color-text-muted)', width: 28, textAlign: 'right' }}>
            {r.value}
          </span>
        </div>
      ))}
    </div>
  );
}
