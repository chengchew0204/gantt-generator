import { useRef, useState, useEffect, useCallback, useMemo } from 'react';
import { getShapePreset, renderShapePath, isLineShape } from '../utils/ShapePresets';

const HANDLE_SIZE = 8;
const ROTATE_HANDLE_OFFSET = 22;
const MIN_SHAPE_SIZE = 4;
const DEFAULT_FILL = '#2383e2';
const DEFAULT_OUTLINE = '#0b5394';

/**
 * SVG-based floating drawing layer that scrolls with the DataGrid's inner
 * positioning container. Handles shape drawing, selection, movement,
 * resize, rotation, and text editing; commits committed state to the
 * parent via onChange(nextShapes) at the end of each gesture so undo
 * entries coalesce per-action.
 */
export default function ShapeLayer({
  shapes,
  selectedIds,
  shapeMode,
  colOffsets,
  rowOffsets,
  rowHeaderWidth,
  headerHeight,
  totalWidth,
  totalHeight,
  onChange,
  onSelect,
  onExitShapeMode,
  editingTextId,
  onBeginTextEdit,
  onCommitText,
  onToggleTextFormat,
}) {
  const svgRef = useRef(null);
  const [gesture, setGesture] = useState(null);

  const contentW = Math.max(0, totalWidth - rowHeaderWidth);
  const contentH = Math.max(0, totalHeight - headerHeight);

  const shapeMap = useMemo(() => {
    const m = new Map();
    for (const s of shapes) m.set(s.id, s);
    return m;
  }, [shapes]);

  const selectedSet = useMemo(() => new Set(selectedIds || []), [selectedIds]);

  // Screen pointer -> layer-local (shape) coordinates.
  const getLocal = useCallback((clientX, clientY) => {
    const svg = svgRef.current;
    if (!svg) return { x: 0, y: 0 };
    const rect = svg.getBoundingClientRect();
    return { x: clientX - rect.left, y: clientY - rect.top };
  }, []);

  // Apply the active gesture to a shape and return the effective visible shape.
  const effectiveShape = useCallback(
    (shape) => {
      if (!gesture) return shape;
      if (gesture.kind === 'move' && gesture.ids.includes(shape.id)) {
        const original = gesture.originals[shape.id];
        return { ...shape, x: original.x + gesture.dx, y: original.y + gesture.dy };
      }
      if (gesture.kind === 'resize' && gesture.id === shape.id) {
        return { ...shape, ...gesture.bounds };
      }
      if (gesture.kind === 'rotate' && gesture.id === shape.id) {
        return { ...shape, rot: gesture.rot };
      }
      return shape;
    },
    [gesture],
  );

  // Cancel current gesture & shape-insert mode on Escape.
  useEffect(() => {
    const handler = (e) => {
      if (e.key !== 'Escape') return;
      if (gesture) {
        setGesture(null);
      } else if (shapeMode) {
        onExitShapeMode?.();
      }
    };
    window.addEventListener('keydown', handler);
    return () => window.removeEventListener('keydown', handler);
  }, [gesture, shapeMode, onExitShapeMode]);

  // SVG root mousedown. In shape mode, start drawing a new shape. Outside
  // shape mode, the SVG has pointer-events: none, so this only fires when
  // the user is inserting.
  const onSvgMouseDown = useCallback(
    (e) => {
      if (!shapeMode) return;
      e.preventDefault();
      e.stopPropagation();
      const { x, y } = getLocal(e.clientX, e.clientY);
      setGesture({ kind: 'draw', startX: x, startY: y, currentX: x, currentY: y, presetId: shapeMode });
    },
    [shapeMode, getLocal],
  );

  // Global mousemove / mouseup while a gesture is in flight.
  useEffect(() => {
    if (!gesture) return;

    const onMove = (e) => {
      const { x, y } = getLocal(e.clientX, e.clientY);
      if (gesture.kind === 'draw') {
        setGesture({ ...gesture, currentX: x, currentY: y });
      } else if (gesture.kind === 'move') {
        setGesture({ ...gesture, dx: x - gesture.startX, dy: y - gesture.startY });
      } else if (gesture.kind === 'resize') {
        const next = computeResizeBounds(gesture, x, y, e.shiftKey);
        setGesture({ ...gesture, bounds: next });
      } else if (gesture.kind === 'rotate') {
        const dx = x - gesture.cx;
        const dy = y - gesture.cy;
        let deg = Math.atan2(dy, dx) * (180 / Math.PI) + 90;
        if (e.shiftKey) deg = Math.round(deg / 15) * 15;
        deg = ((deg % 360) + 360) % 360;
        setGesture({ ...gesture, rot: deg });
      }
    };

    const onUp = () => {
      if (gesture.kind === 'draw') {
        const x0 = Math.min(gesture.startX, gesture.currentX);
        const y0 = Math.min(gesture.startY, gesture.currentY);
        const w = Math.abs(gesture.currentX - gesture.startX);
        const h = Math.abs(gesture.currentY - gesture.startY);
        if (w >= MIN_SHAPE_SIZE && h >= MIN_SHAPE_SIZE) {
          const newShape = createShape(gesture.presetId, x0, y0, w, h, shapes);
          onChange([...shapes, newShape]);
          onSelect([newShape.id]);
        }
        onExitShapeMode?.();
      } else if (gesture.kind === 'move') {
        if (gesture.dx !== 0 || gesture.dy !== 0) {
          const next = shapes.map((s) => {
            if (!gesture.ids.includes(s.id)) return s;
            const orig = gesture.originals[s.id];
            return { ...s, x: orig.x + gesture.dx, y: orig.y + gesture.dy };
          });
          onChange(next);
        }
      } else if (gesture.kind === 'resize') {
        const next = shapes.map((s) =>
          s.id === gesture.id ? { ...s, ...gesture.bounds } : s,
        );
        onChange(next);
      } else if (gesture.kind === 'rotate') {
        const next = shapes.map((s) =>
          s.id === gesture.id ? { ...s, rot: gesture.rot } : s,
        );
        onChange(next);
      }
      setGesture(null);
    };

    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
    return () => {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
    };
  }, [gesture, getLocal, shapes, onChange, onSelect, onExitShapeMode]);

  const beginMove = useCallback(
    (e, shapeId) => {
      if (shapeMode) return; // In insert mode, clicks on shapes start a new draw instead.
      e.preventDefault();
      e.stopPropagation();
      if (editingTextId) return;
      let moveIds;
      if (selectedSet.has(shapeId)) {
        moveIds = selectedIds;
      } else if (e.shiftKey) {
        moveIds = [...selectedIds, shapeId];
        onSelect(moveIds);
      } else {
        moveIds = [shapeId];
        onSelect(moveIds);
      }
      const originals = {};
      for (const id of moveIds) {
        const s = shapeMap.get(id);
        if (s) originals[id] = { x: s.x, y: s.y };
      }
      const { x, y } = getLocal(e.clientX, e.clientY);
      setGesture({ kind: 'move', ids: moveIds, startX: x, startY: y, dx: 0, dy: 0, originals });
    },
    [shapeMode, editingTextId, selectedSet, selectedIds, onSelect, shapeMap, getLocal],
  );

  const beginResize = useCallback(
    (e, shapeId, handle) => {
      e.preventDefault();
      e.stopPropagation();
      if (editingTextId) return;
      const shape = shapeMap.get(shapeId);
      if (!shape) return;
      const { x, y } = getLocal(e.clientX, e.clientY);
      setGesture({
        kind: 'resize',
        id: shapeId,
        handle,
        startX: x,
        startY: y,
        initial: { x: shape.x, y: shape.y, w: shape.w, h: shape.h, rot: shape.rot || 0 },
        bounds: { x: shape.x, y: shape.y, w: shape.w, h: shape.h },
      });
    },
    [editingTextId, shapeMap, getLocal],
  );

  const beginRotate = useCallback(
    (e, shapeId) => {
      e.preventDefault();
      e.stopPropagation();
      if (editingTextId) return;
      const shape = shapeMap.get(shapeId);
      if (!shape) return;
      setGesture({
        kind: 'rotate',
        id: shapeId,
        cx: shape.x + shape.w / 2,
        cy: shape.y + shape.h / 2,
        rot: shape.rot || 0,
      });
    },
    [editingTextId, shapeMap],
  );

  const handleShapeDoubleClick = useCallback(
    (e, shapeId) => {
      e.stopPropagation();
      onSelect([shapeId]);
      onBeginTextEdit?.(shapeId);
    },
    [onSelect, onBeginTextEdit],
  );

  const svgPointerEvents = shapeMode ? 'auto' : 'none';
  const svgCursor = shapeMode ? 'crosshair' : 'default';

  // Preview rectangle while drawing a new shape.
  const drawPreview =
    gesture && gesture.kind === 'draw'
      ? {
          x: Math.min(gesture.startX, gesture.currentX),
          y: Math.min(gesture.startY, gesture.currentY),
          w: Math.abs(gesture.currentX - gesture.startX),
          h: Math.abs(gesture.currentY - gesture.startY),
          presetId: gesture.presetId,
        }
      : null;

  if (contentW <= 0 || contentH <= 0) return null;

  return (
    <svg
      ref={svgRef}
      width={contentW}
      height={contentH}
      style={{
        position: 'absolute',
        left: rowHeaderWidth,
        top: headerHeight,
        zIndex: 2,
        overflow: 'visible',
        pointerEvents: svgPointerEvents,
        cursor: svgCursor,
      }}
      onMouseDown={onSvgMouseDown}
    >
      <defs>
        <marker
          id="ganttgen-shape-arrow-end"
          viewBox="0 0 10 10"
          refX="9"
          refY="5"
          markerWidth="6"
          markerHeight="6"
          orient="auto-start-reverse"
        >
          <path d="M0,0 L10,5 L0,10 Z" fill="context-stroke" />
        </marker>
        <marker
          id="ganttgen-shape-arrow-start"
          viewBox="0 0 10 10"
          refX="1"
          refY="5"
          markerWidth="6"
          markerHeight="6"
          orient="auto"
        >
          <path d="M10,0 L0,5 L10,10 Z" fill="context-stroke" />
        </marker>
      </defs>

      {shapes.map((s) => {
        const eff = effectiveShape(s);
        const isSelected = selectedSet.has(s.id);
        const isEditingText = editingTextId === s.id;
        return (
          <ShapeNode
            key={s.id}
            shape={eff}
            isSelected={isSelected}
            isEditingText={isEditingText}
            onMouseDown={(e) => beginMove(e, s.id)}
            onDoubleClick={(e) => handleShapeDoubleClick(e, s.id)}
            onCommitText={(text) => onCommitText?.(s.id, text)}
            onToggleTextFormat={onToggleTextFormat}
          />
        );
      })}

      {/* Selection decorations drawn on top of all shapes */}
      {shapes.map((s) => {
        if (!selectedSet.has(s.id)) return null;
        if (editingTextId === s.id) return null;
        const eff = effectiveShape(s);
        return (
          <SelectionDecorations
            key={`sel-${s.id}`}
            shape={eff}
            onResizeStart={(e, handle) => beginResize(e, s.id, handle)}
            onRotateStart={(e) => beginRotate(e, s.id)}
          />
        );
      })}

      {drawPreview && drawPreview.w > 0 && drawPreview.h > 0 && (
        <g
          transform={`translate(${drawPreview.x} ${drawPreview.y})`}
          style={{ pointerEvents: 'none' }}
        >
          <path
            d={renderShapePath(drawPreview.presetId, drawPreview.w, drawPreview.h)}
            fill={isLineShape(drawPreview.presetId) ? 'none' : 'var(--color-accent)'}
            fillOpacity={0.2}
            stroke="var(--color-accent)"
            strokeWidth={1.5}
            strokeDasharray="4 3"
          />
        </g>
      )}
    </svg>
  );
}

function ShapeNode({ shape, isSelected, isEditingText, onMouseDown, onDoubleClick, onCommitText, onToggleTextFormat }) {
  const preset = getShapePreset(shape.type);
  if (!preset) return null;

  const w = Math.max(1, shape.w);
  const h = Math.max(1, shape.h);
  const rot = shape.rot || 0;
  const style = shape.style || {};
  const fill = style.fill;
  const outline = style.outline || {};
  const effects = style.effects || {};
  const isLine = preset.kind === 'line';
  const isTextBox = !!preset.isTextBox;

  const fillValue = isLine
    ? 'none'
    : !fill || fill.type === 'none'
      ? 'none'
      : fill.color || DEFAULT_FILL;
  const fillOpacity =
    isLine || !fill || fill.type === 'none' || fill.alpha == null ? 1 : fill.alpha;

  const hasVisibleOutline = !!outline.color && (outline.width == null || outline.width > 0);
  const strokeValue = hasVisibleOutline ? outline.color : 'none';
  const strokeWidth = hasVisibleOutline ? (outline.width != null ? outline.width : 1) : 0;
  const strokeOpacity = outline.alpha != null ? outline.alpha : 1;
  const strokeDash = dashToSvg(outline.dash);

  const transform = `translate(${shape.x} ${shape.y}) rotate(${rot} ${w / 2} ${h / 2})`;

  const hasEffects = !!(effects.shadow || effects.glow || effects.softEdge);
  const filterId = hasEffects ? `shape-fx-${shape.id}` : null;

  const d = renderShapePath(shape.type, w, h);

  const markers = {};
  if (isLine) {
    if (preset.arrowEnd) markers.markerEnd = 'url(#ganttgen-shape-arrow-end)';
    if (preset.arrowStart) markers.markerStart = 'url(#ganttgen-shape-arrow-start)';
  }

  const textStyle = shape.textStyle || {};
  const textValue = shape.text?.value ?? '';

  // Hit-test transparent/no-fill shapes (e.g. text boxes) via the full
  // bounding box; lines stick to stroke-only so clicks on empty interior
  // don't grab them.
  const hitPointerEvents = isLine ? 'visiblePainted' : 'all';

  return (
    <g transform={transform} onMouseDown={onMouseDown} onDoubleClick={onDoubleClick}>
      {hasEffects && (
        <defs>
          <ShapeFilter id={filterId} effects={effects} />
        </defs>
      )}
      {isTextBox && !isEditingText && !isSelected && (
        <rect
          x={0}
          y={0}
          width={w}
          height={h}
          fill="none"
          stroke="var(--color-text-muted)"
          strokeWidth={0.5}
          strokeDasharray="2 2"
          style={{ pointerEvents: 'none', opacity: 0.5 }}
        />
      )}
      <path
        d={d}
        fill={fillValue}
        fillOpacity={fillOpacity}
        stroke={strokeValue}
        strokeOpacity={strokeOpacity}
        strokeWidth={strokeWidth}
        strokeDasharray={strokeDash}
        strokeLinejoin="round"
        strokeLinecap="round"
        fillRule={preset.compound ? 'evenodd' : 'nonzero'}
        filter={filterId ? `url(#${filterId})` : undefined}
        style={{ cursor: isEditingText ? 'text' : 'move', pointerEvents: hitPointerEvents }}
        {...markers}
      />
      {(textValue || isEditingText) && !isLine && (
        <ShapeTextBody
          w={w}
          h={h}
          textStyle={textStyle}
          textValue={textValue}
          isEditing={isEditingText}
          vertical={!!preset.vertical}
          onCommit={onCommitText}
          onToggleTextFormat={onToggleTextFormat}
        />
      )}
    </g>
  );
}

function ShapeTextBody({ w, h, textStyle, textValue, isEditing, vertical, onCommit, onToggleTextFormat }) {
  const ts = textStyle || {};
  const fontSize = ts.fontSize || 12;
  const color = ts.fill?.color || '#1f1f1f';
  const hAlign = ts.hAlign || 'center';
  const vAlign = ts.vAlign || 'middle';
  const justifyContent = hAlign === 'left' ? 'flex-start' : hAlign === 'right' ? 'flex-end' : 'center';
  const alignItems = vAlign === 'top' ? 'flex-start' : vAlign === 'bottom' ? 'flex-end' : 'center';
  const textOutline = ts.outline;
  const textShadow = buildTextShadowCss(ts.effects?.shadow, ts.effects?.glow);
  const webkitStroke = textOutline && textOutline.color
    ? `${textOutline.width || 1}px ${textOutline.color}`
    : undefined;

  const writingMode = vertical ? 'vertical-rl' : undefined;

  const [draft, setDraft] = useState(textValue);
  const inputRef = useRef(null);

  useEffect(() => {
    if (isEditing) {
      setDraft(textValue);
      requestAnimationFrame(() => {
        inputRef.current?.focus();
        const node = inputRef.current;
        if (node) {
          const range = document.createRange();
          range.selectNodeContents(node);
          const sel = window.getSelection();
          sel.removeAllRanges();
          sel.addRange(range);
        }
      });
    }
  }, [isEditing, textValue]);

  return (
    <foreignObject
      x={4}
      y={4}
      width={Math.max(0, w - 8)}
      height={Math.max(0, h - 8)}
      style={{ overflow: 'visible', pointerEvents: isEditing ? 'auto' : 'none' }}
    >
      <div
        style={{
          width: '100%',
          height: '100%',
          display: 'flex',
          justifyContent,
          alignItems,
          textAlign: hAlign,
          fontSize,
          fontWeight: ts.bold ? 700 : 400,
          fontStyle: ts.italic ? 'italic' : 'normal',
          textDecoration: ts.underline ? 'underline' : 'none',
          fontFamily: ts.fontFamily || 'inherit',
          color,
          textShadow,
          WebkitTextStroke: webkitStroke,
          writingMode,
          whiteSpace: 'pre-wrap',
          wordBreak: 'break-word',
          padding: '2px 4px',
          boxSizing: 'border-box',
          pointerEvents: isEditing ? 'auto' : 'none',
          userSelect: isEditing ? 'text' : 'none',
        }}
      >
        {isEditing ? (
          <div
            ref={inputRef}
            contentEditable
            suppressContentEditableWarning
            onInput={(e) => setDraft(e.currentTarget.innerText)}
            onBlur={(e) => onCommit?.(e.currentTarget.innerText)}
            onKeyDown={(e) => {
              e.stopPropagation();
              if (e.key === 'Escape') {
                e.preventDefault();
                onCommit?.(textValue);
                return;
              }
              if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                onCommit?.(e.currentTarget.innerText);
                return;
              }
              // Intercept Ctrl/Cmd + B / I / U and route to the parent so
              // the whole-shape textStyle updates instead of the browser
              // injecting <b>/<i>/<u> tags that innerText later strips.
              if ((e.ctrlKey || e.metaKey) && !e.altKey) {
                const k = e.key.toLowerCase();
                if (k === 'b' || k === 'i' || k === 'u') {
                  e.preventDefault();
                  const key = k === 'b' ? 'bold' : k === 'i' ? 'italic' : 'underline';
                  onToggleTextFormat?.(key);
                }
              }
            }}
            onPaste={(e) => {
              // Keep the stored text plain. Rich-text paste would inject
              // HTML that innerText strips on commit, producing surprising
              // results.
              e.preventDefault();
              const text = e.clipboardData?.getData('text/plain') || '';
              if (text) document.execCommand('insertText', false, text);
            }}
            style={{ outline: 'none', minWidth: 4, minHeight: 4 }}
          >
            {textValue}
          </div>
        ) : (
          textValue
        )}
      </div>
    </foreignObject>
  );
}

function SelectionDecorations({ shape, onResizeStart, onRotateStart }) {
  const w = Math.max(1, shape.w);
  const h = Math.max(1, shape.h);
  const rot = shape.rot || 0;
  const transform = `translate(${shape.x} ${shape.y}) rotate(${rot} ${w / 2} ${h / 2})`;
  const handles = [
    { id: 'tl', x: 0, y: 0, cursor: 'nwse-resize' },
    { id: 't', x: w / 2, y: 0, cursor: 'ns-resize' },
    { id: 'tr', x: w, y: 0, cursor: 'nesw-resize' },
    { id: 'r', x: w, y: h / 2, cursor: 'ew-resize' },
    { id: 'br', x: w, y: h, cursor: 'nwse-resize' },
    { id: 'b', x: w / 2, y: h, cursor: 'ns-resize' },
    { id: 'bl', x: 0, y: h, cursor: 'nesw-resize' },
    { id: 'l', x: 0, y: h / 2, cursor: 'ew-resize' },
  ];
  return (
    <g transform={transform} style={{ pointerEvents: 'none' }}>
      <rect
        x={0}
        y={0}
        width={w}
        height={h}
        fill="none"
        stroke="var(--color-accent)"
        strokeWidth={1}
        strokeDasharray="3 2"
        vectorEffect="non-scaling-stroke"
      />
      <line
        x1={w / 2}
        y1={0}
        x2={w / 2}
        y2={-ROTATE_HANDLE_OFFSET}
        stroke="var(--color-accent)"
        strokeWidth={1}
      />
      <circle
        cx={w / 2}
        cy={-ROTATE_HANDLE_OFFSET}
        r={HANDLE_SIZE / 2 + 1}
        fill="#ffffff"
        stroke="var(--color-accent)"
        strokeWidth={1}
        style={{ pointerEvents: 'all', cursor: 'grab' }}
        onMouseDown={onRotateStart}
      />
      {handles.map((hd) => (
        <rect
          key={hd.id}
          x={hd.x - HANDLE_SIZE / 2}
          y={hd.y - HANDLE_SIZE / 2}
          width={HANDLE_SIZE}
          height={HANDLE_SIZE}
          fill="#ffffff"
          stroke="var(--color-accent)"
          strokeWidth={1}
          style={{ pointerEvents: 'all', cursor: hd.cursor }}
          onMouseDown={(e) => onResizeStart(e, hd.id)}
        />
      ))}
    </g>
  );
}

function dashToSvg(dash) {
  switch (dash) {
    case 'dash': return '6 3';
    case 'dot': return '2 2';
    case 'dashDot': return '6 2 2 2';
    case 'longDash': return '10 3';
    default: return undefined;
  }
}

function ShapeFilter({ id, effects }) {
  const primitives = [];
  if (effects.shadow) {
    const s = effects.shadow;
    const dx = s.offsetX != null ? s.offsetX : 3;
    const dy = s.offsetY != null ? s.offsetY : 3;
    const blur = s.blur != null ? s.blur : 4;
    const color = s.color || '#000000';
    const alpha = s.alpha != null ? s.alpha : 0.35;
    primitives.push(
      <feGaussianBlur key="sh-blur" in="SourceAlpha" stdDeviation={blur} result="fx-sh-blur" />,
      <feOffset key="sh-off" in="fx-sh-blur" dx={dx} dy={dy} result="fx-sh-off" />,
      <feFlood key="sh-flood" floodColor={color} floodOpacity={alpha} result="fx-sh-color" />,
      <feComposite key="sh-comp" in="fx-sh-color" in2="fx-sh-off" operator="in" result="fx-shadow" />,
    );
  }
  if (effects.glow) {
    const g = effects.glow;
    const radius = g.radius != null ? g.radius : 4;
    const color = g.color || '#ffea00';
    const alpha = g.alpha != null ? g.alpha : 0.6;
    primitives.push(
      <feGaussianBlur key="gl-blur" in="SourceAlpha" stdDeviation={radius} result="fx-gl-blur" />,
      <feFlood key="gl-flood" floodColor={color} floodOpacity={alpha} result="fx-gl-color" />,
      <feComposite key="gl-comp" in="fx-gl-color" in2="fx-gl-blur" operator="in" result="fx-glow" />,
    );
  }
  if (effects.softEdge) {
    const radius = effects.softEdge.radius != null ? effects.softEdge.radius : 2;
    primitives.push(
      <feGaussianBlur key="se-blur" in="SourceGraphic" stdDeviation={radius} result="fx-softedge" />,
    );
  }
  // Merge all effects back with the original source graphic on top.
  const mergeNodes = [];
  if (effects.shadow) mergeNodes.push(<feMergeNode key="mg-shadow" in="fx-shadow" />);
  if (effects.glow) mergeNodes.push(<feMergeNode key="mg-glow" in="fx-glow" />);
  if (effects.softEdge) {
    mergeNodes.push(<feMergeNode key="mg-soft" in="fx-softedge" />);
  } else {
    mergeNodes.push(<feMergeNode key="mg-src" in="SourceGraphic" />);
  }
  return (
    <filter id={id} x="-50%" y="-50%" width="200%" height="200%">
      {primitives}
      <feMerge>{mergeNodes}</feMerge>
    </filter>
  );
}

function buildTextShadowCss(shadow, glow) {
  const parts = [];
  if (shadow) {
    const dx = shadow.offsetX != null ? shadow.offsetX : 1;
    const dy = shadow.offsetY != null ? shadow.offsetY : 1;
    const blur = shadow.blur != null ? shadow.blur : 2;
    const color = applyAlpha(shadow.color || '#000000', shadow.alpha != null ? shadow.alpha : 0.6);
    parts.push(`${dx}px ${dy}px ${blur}px ${color}`);
  }
  if (glow) {
    const radius = glow.radius != null ? glow.radius : 3;
    const color = applyAlpha(glow.color || '#ffea00', glow.alpha != null ? glow.alpha : 0.8);
    parts.push(`0 0 ${radius}px ${color}`);
  }
  return parts.length > 0 ? parts.join(', ') : undefined;
}

function applyAlpha(hex, alpha) {
  const m = /^#([0-9a-f]{6})$/i.exec(hex);
  if (!m) return hex;
  const r = parseInt(m[1].slice(0, 2), 16);
  const g = parseInt(m[1].slice(2, 4), 16);
  const b = parseInt(m[1].slice(4, 6), 16);
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

function computeResizeBounds(gesture, curX, curY, uniform) {
  const { initial, handle, startX, startY } = gesture;
  const dx = curX - startX;
  const dy = curY - startY;
  let { x, y, w, h } = initial;

  if (handle.includes('l')) { x = initial.x + dx; w = initial.w - dx; }
  if (handle.includes('r')) { w = initial.w + dx; }
  if (handle.includes('t')) { y = initial.y + dy; h = initial.h - dy; }
  if (handle.includes('b')) { h = initial.h + dy; }

  if (w < MIN_SHAPE_SIZE) {
    if (handle.includes('l')) x = initial.x + initial.w - MIN_SHAPE_SIZE;
    w = MIN_SHAPE_SIZE;
  }
  if (h < MIN_SHAPE_SIZE) {
    if (handle.includes('t')) y = initial.y + initial.h - MIN_SHAPE_SIZE;
    h = MIN_SHAPE_SIZE;
  }

  if (uniform && handle.length === 2) {
    const aspect = initial.w / initial.h;
    if (w / h > aspect) {
      w = h * aspect;
      if (handle.includes('l')) x = initial.x + initial.w - w;
    } else {
      h = w / aspect;
      if (handle.includes('t')) y = initial.y + initial.h - h;
    }
  }

  return { x, y, w, h };
}

function createShape(presetId, x, y, w, h, existingShapes) {
  const preset = getShapePreset(presetId);
  const id = `shp_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`;
  const maxZ = existingShapes.reduce((m, s) => Math.max(m, s.z || 0), 0);
  const isLine = preset?.kind === 'line';
  const isTextBox = !!preset?.isTextBox;
  return {
    id,
    type: presetId,
    x,
    y,
    w,
    h,
    rot: 0,
    z: maxZ + 1,
    text: isLine ? undefined : { value: '' },
    style: {
      fill: isLine || isTextBox ? { type: 'none' } : { type: 'solid', color: DEFAULT_FILL, alpha: 1 },
      outline: isTextBox
        ? { color: undefined, alpha: 1, width: 0, dash: 'solid' }
        : { color: DEFAULT_OUTLINE, alpha: 1, width: isLine ? 2 : 1, dash: 'solid' },
      effects: {},
    },
    textStyle: {
      fontFamily: 'Calibri',
      fontSize: 12,
      fill: isTextBox ? { color: '#1f1f1f', alpha: 1 } : { color: '#ffffff', alpha: 1 },
      hAlign: isTextBox ? 'left' : 'center',
      vAlign: isTextBox ? 'top' : 'middle',
    },
  };
}

export function nudgeShapes(shapes, ids, dx, dy) {
  if (!ids || ids.length === 0) return shapes;
  const idSet = new Set(ids);
  return shapes.map((s) => (idSet.has(s.id) ? { ...s, x: s.x + dx, y: s.y + dy } : s));
}

export function deleteShapes(shapes, ids) {
  if (!ids || ids.length === 0) return shapes;
  const idSet = new Set(ids);
  return shapes.filter((s) => !idSet.has(s.id));
}

export function duplicateShapes(shapes, ids) {
  if (!ids || ids.length === 0) return { shapes, newIds: [] };
  const idSet = new Set(ids);
  const originals = shapes.filter((s) => idSet.has(s.id));
  const maxZ = shapes.reduce((m, s) => Math.max(m, s.z || 0), 0);
  const newShapes = originals.map((s, i) => ({
    ...JSON.parse(JSON.stringify(s)),
    id: `shp_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}_${i}`,
    x: s.x + 16,
    y: s.y + 16,
    z: maxZ + i + 1,
  }));
  const newIds = newShapes.map((s) => s.id);
  return { shapes: [...shapes, ...newShapes], newIds };
}

export function reorderShape(shapes, id, direction) {
  const idx = shapes.findIndex((s) => s.id === id);
  if (idx < 0) return shapes;
  const sorted = [...shapes].sort((a, b) => (a.z || 0) - (b.z || 0));
  const sIdx = sorted.findIndex((s) => s.id === id);
  if (direction === 'front') {
    sorted.splice(sIdx, 1);
    sorted.push(shapes[idx]);
  } else if (direction === 'back') {
    sorted.splice(sIdx, 1);
    sorted.unshift(shapes[idx]);
  } else if (direction === 'forward' && sIdx < sorted.length - 1) {
    [sorted[sIdx], sorted[sIdx + 1]] = [sorted[sIdx + 1], sorted[sIdx]];
  } else if (direction === 'backward' && sIdx > 0) {
    [sorted[sIdx], sorted[sIdx - 1]] = [sorted[sIdx - 1], sorted[sIdx]];
  }
  return sorted.map((s, i) => ({ ...s, z: i + 1 }));
}
