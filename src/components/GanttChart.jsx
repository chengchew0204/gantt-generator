import { useState, useMemo, useRef, useCallback, useEffect } from 'react';
import { BarChart3 } from 'lucide-react';

const ROW_HEIGHT = 32;
const MONTH_LABEL_HEIGHT = 16;
const WEEK_LABEL_HEIGHT = 18;
const DAY_LABEL_HEIGHT = 18;
const BAR_HEIGHT = 16;
const BASELINE_HEIGHT = 4;
const SUMMARY_HEIGHT = 10;

const ZOOM_MIN = 5;
const ZOOM_MAX = 300;
const BASE_UNIT_AT_100 = 32;
// Extra right-side pixels reserved so task name labels that appear after
// the last bar are not clipped by the SVG boundary.
const LABEL_RIGHT_PADDING = 300;

function toBool(v) {
  if (typeof v === 'boolean') return v;
  return v !== 'false' && v !== false;
}

export default function GanttChart({ tasks, allTasks, viewOptions = {}, scrollTop, onScroll, onUpdateTask, onUpdateTaskFields, selectedTaskId, onSelectTask, categoryColors = {}, datePickField, onDatePickField, onBeginDrag, onEndDrag, scale, zoomPct }) {
  const effectiveScale = scale ?? 'day';
  const effectiveZoom = Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, Number.isFinite(zoomPct) ? zoomPct : 100));
  const [hScrollLeft, setHScrollLeft] = useState(0);
  const [hoveredDate, setHoveredDate] = useState(null);
  const [dragPickStart, setDragPickStart] = useState(null);
  const dragPickStartRef = useRef(null);
  const scrollRef = useRef(null);
  const suppressScroll = useRef(false);

  const show = {
    criticalPath: toBool(viewOptions.showCriticalPath ?? true),
    slack: toBool(viewOptions.showSlack ?? true),
    dependencies: toBool(viewOptions.showDependencies ?? true),
    todayLine: toBool(viewOptions.showTodayLine ?? true),
    baseline: toBool(viewOptions.showBaseline ?? true),
    weekLabels: toBool(viewOptions.showWeekLabels ?? false),
    monthLabels: toBool(viewOptions.showMonthLabels ?? true),
    dayLabels: toBool(viewOptions.showDayLabels ?? true),
    taskNames: toBool(viewOptions.showTaskNames ?? true),
    progressPercent: toBool(viewOptions.showProgressPercent ?? true),
  };

  // Keep baseUnit as a float so zoom scales continuously at every integer percentage.
  // Per-element minimums (TaskBar/SummaryBar/BaselineBar cap widths at >=4 px, MilestoneDiamond
  // has a fixed 7 px radius, TodayLine has a fixed stroke) keep each rendered element readable
  // even when unitWidth drops below 1 px at very low zoom.
  const baseUnit = BASE_UNIT_AT_100 * effectiveZoom / 100;
  const unitWidth = effectiveScale === 'day' ? baseUnit : baseUnit * 0.56;

  const { minDate, totalDays } = useMemo(() => {
    const src = allTasks?.length > 0 ? allTasks : tasks;
    if (src.length === 0) {
      return { minDate: todayIso(), totalDays: 1 };
    }
    const min = getMinDate(src);
    const max = getMaxDate(src);
    return { minDate: min, totalDays: daysBetween(min, max) + 1 };
  }, [tasks, allTasks]);

  const isPicking = !!datePickField && !!selectedTaskId;

  const xToDate = useCallback((clientX, svgEl) => {
    if (!svgEl) return null;
    const rect = svgEl.getBoundingClientRect();
    const x = clientX - rect.left;
    const dayOffset = Math.floor(x / unitWidth);
    return addDaysIso(minDate, dayOffset);
  }, [unitWidth, minDate]);

  const handleGridMouseDown = useCallback((e) => {
    if (!isPicking || !selectedTaskId) {
      if (onSelectTask) onSelectTask(null);
      return;
    }
    e.preventDefault();
    const svgEl = e.currentTarget;
    const date = xToDate(e.clientX, svgEl);
    if (!date) return;

    setDragPickStart(date);
    dragPickStartRef.current = date;
    if (onBeginDrag) onBeginDrag();

    const handleDocMove = (ev) => {
      const moveDate = xToDate(ev.clientX, svgEl);
      if (moveDate) setHoveredDate(moveDate);
    };

    const handleDocUp = (ev) => {
      document.removeEventListener('mousemove', handleDocMove);
      document.removeEventListener('mouseup', handleDocUp);

      const anchor = dragPickStartRef.current;
      const releaseDate = xToDate(ev.clientX, svgEl) || anchor;
      setDragPickStart(null);
      dragPickStartRef.current = null;
      setHoveredDate(null);

      if (!anchor) {
        if (onEndDrag) onEndDrag();
        return;
      }

      const dragged = releaseDate !== anchor;
      if (dragged && onUpdateTaskFields) {
        const startDate = anchor < releaseDate ? anchor : releaseDate;
        const endDate = anchor < releaseDate ? releaseDate : anchor;
        onUpdateTaskFields(selectedTaskId, { startDate, endDate });
      } else if (onUpdateTask && datePickField) {
        onUpdateTask(selectedTaskId, datePickField, anchor);
      }

      if (onEndDrag) onEndDrag();
      if (onDatePickField) onDatePickField(null);
    };

    document.addEventListener('mousemove', handleDocMove);
    document.addEventListener('mouseup', handleDocUp);
  }, [isPicking, selectedTaskId, datePickField, xToDate, onUpdateTask, onUpdateTaskFields, onSelectTask, onDatePickField, onBeginDrag, onEndDrag]);

  const handleGridMouseMove = useCallback((e) => {
    if (!isPicking) { setHoveredDate(null); return; }
    const date = xToDate(e.clientX, e.currentTarget);
    setHoveredDate(date);
  }, [isPicking, xToDate]);

  const handleGridMouseLeave = useCallback(() => {
    if (!dragPickStart) setHoveredDate(null);
  }, [dragPickStart]);

  const isDragPicking = !!dragPickStart;

  const hoveredDateLabel = useMemo(() => {
    if (!hoveredDate) return null;
    if (isDragPicking && dragPickStart) {
      const s = dragPickStart < hoveredDate ? dragPickStart : hoveredDate;
      const e = dragPickStart < hoveredDate ? hoveredDate : dragPickStart;
      if (s === e) return `Start: ${s}`;
      return `${s}  --  ${e}`;
    }
    if (!datePickField) return null;
    const label = datePickField === 'startDate' ? 'Start' : 'End';
    return `${label}: ${hoveredDate}`;
  }, [hoveredDate, datePickField, isDragPicking, dragPickStart]);

  const taskIndex = useMemo(() => {
    const map = new Map();
    tasks.forEach((t, i) => map.set(String(t.id), i));
    return map;
  }, [tasks]);

  const deps = useMemo(() => {
    if (!show.dependencies || tasks.length === 0) return [];
    const edges = [];
    for (const task of tasks) {
      if (!task.dependency) continue;
      const preds = String(task.dependency).split(',').map((s) => s.trim()).filter(Boolean);
      for (const predId of preds) {
        if (taskIndex.has(predId) && taskIndex.has(String(task.id))) {
          edges.push({ from: predId, to: String(task.id) });
        }
      }
    }
    return edges;
  }, [tasks, taskIndex, show.dependencies]);

  useEffect(() => {
    if (scrollRef.current && scrollTop != null) {
      if (Math.abs(scrollRef.current.scrollTop - scrollTop) > 1) {
        suppressScroll.current = true;
        scrollRef.current.scrollTop = scrollTop;
      }
    }
  }, [scrollTop]);

  const handleScroll = useCallback(() => {
    if (!scrollRef.current) return;
    setHScrollLeft(scrollRef.current.scrollLeft);
    if (suppressScroll.current) {
      suppressScroll.current = false;
      return;
    }
    if (onScroll) {
      onScroll(scrollRef.current.scrollTop);
    }
  }, [onScroll]);

  if (tasks.length === 0) {
    return (
      <div className="flex flex-col items-center justify-center h-full gap-3 px-6">
        <div className="flex items-center justify-center w-12 h-12 rounded-xl" style={{ backgroundColor: 'var(--color-accent-muted)' }}>
          <BarChart3 size={22} style={{ color: 'var(--color-accent)' }} />
        </div>
        <p className="text-[16px] font-medium" style={{ color: 'var(--color-text-secondary)' }}>No chart data</p>
        <p className="text-[14px] text-center max-w-[200px]" style={{ color: 'var(--color-text-muted)' }}>Import tasks to render the Gantt chart.</p>
      </div>
    );
  }

  const dataWidth = totalDays * unitWidth;
  const bodyHeight = tasks.length * ROW_HEIGHT;
  const containerWidth = scrollRef.current?.clientWidth || 0;
  const chartWidth = Math.max(dataWidth + LABEL_RIGHT_PADDING, containerWidth, 600);
  return (
    <div className="flex flex-col h-full">
      {/* Month Labels row */}
      {show.monthLabels && (
        <div className="flex-shrink-0 overflow-hidden" style={{ height: MONTH_LABEL_HEIGHT, backgroundColor: 'var(--color-bg-secondary)', borderBottom: show.weekLabels || show.dayLabels ? '1px solid var(--color-border-subtle)' : '1px solid var(--color-border)' }}>
          <svg
            width={Math.max(chartWidth, dataWidth)}
            height={MONTH_LABEL_HEIGHT}
            className="block"
            style={{ transform: `translateX(${-hScrollLeft}px)`, cursor: isPicking ? 'crosshair' : 'default' }}
            onMouseDown={handleGridMouseDown}
            onMouseMove={handleGridMouseMove}
            onMouseLeave={handleGridMouseLeave}
          >
            <MonthLabels minDate={minDate} totalDays={totalDays} unitWidth={unitWidth} chartWidth={chartWidth} height={MONTH_LABEL_HEIGHT} scale={effectiveScale} />
          </svg>
        </div>
      )}

      {/* Week Number Labels row */}
      {show.weekLabels && (
        <div className="flex-shrink-0 overflow-hidden" style={{ height: WEEK_LABEL_HEIGHT, backgroundColor: 'var(--color-bg-secondary)', borderBottom: show.dayLabels ? '1px solid var(--color-border-subtle)' : '1px solid var(--color-border)' }}>
          <svg
            width={Math.max(chartWidth, dataWidth)}
            height={WEEK_LABEL_HEIGHT}
            className="block"
            style={{ transform: `translateX(${-hScrollLeft}px)` }}
          >
            <WeekLabels minDate={minDate} totalDays={totalDays} unitWidth={unitWidth} chartWidth={chartWidth} height={WEEK_LABEL_HEIGHT} />
          </svg>
        </div>
      )}

      {/* Day/Date Labels row */}
      {show.dayLabels && (
        <div className="flex-shrink-0 overflow-hidden" style={{ height: DAY_LABEL_HEIGHT, backgroundColor: 'var(--color-bg-secondary)', borderBottom: '1px solid var(--color-border)' }}>
          <svg
            width={Math.max(chartWidth, dataWidth)}
            height={DAY_LABEL_HEIGHT}
            className="block"
            style={{ transform: `translateX(${-hScrollLeft}px)`, cursor: isPicking ? 'crosshair' : 'default' }}
            onMouseDown={handleGridMouseDown}
            onMouseMove={handleGridMouseMove}
            onMouseLeave={handleGridMouseLeave}
          >
            <DayLabels minDate={minDate} totalDays={totalDays} unitWidth={unitWidth} chartWidth={chartWidth} height={DAY_LABEL_HEIGHT} scale={effectiveScale} />
            {isPicking && hoveredDate && (
              <HoverDateHighlight date={hoveredDate} minDate={minDate} unitWidth={unitWidth} height={DAY_LABEL_HEIGHT} label={hoveredDateLabel} />
            )}
          </svg>
        </div>
      )}

      {/* Scrollable body with grid background + task bars */}
      <div ref={scrollRef} className="flex-1 overflow-auto" onScroll={handleScroll}
        onClick={(e) => { if (e.target === e.currentTarget && onSelectTask) onSelectTask(null); }}>
        <svg width="100%" height={bodyHeight} className="block"
          style={{ minWidth: chartWidth, cursor: isPicking ? 'crosshair' : 'default' }}
          onMouseDown={handleGridMouseDown}
          onMouseMove={handleGridMouseMove}
          onMouseLeave={handleGridMouseLeave}
        >
          <defs>
            <marker id="arrowhead" markerWidth="6" markerHeight="5" refX="6" refY="2.5" orient="auto">
              <path d="M0,0 L6,2.5 L0,5 Z" fill="var(--color-text-muted)" opacity="0.6" />
            </marker>
            <pattern id="slack-hatch" width="6" height="6" patternUnits="userSpaceOnUse" patternTransform="rotate(45)">
              <line x1="0" y1="0" x2="0" y2="6" stroke="var(--color-warning)" strokeWidth="2" />
            </pattern>
          </defs>

          <BodyGrid minDate={minDate} totalDays={totalDays} unitWidth={unitWidth} bodyHeight={bodyHeight} chartWidth={chartWidth} scale={effectiveScale} />
          {isPicking && hoveredDate && (
            <HoverDateHighlight date={hoveredDate} minDate={minDate} unitWidth={unitWidth} height={bodyHeight} />
          )}
          {isDragPicking && hoveredDate && selectedTaskId && (
            <DragPreviewBar
              dragStart={dragPickStart}
              dragEnd={hoveredDate}
              minDate={minDate}
              unitWidth={unitWidth}
              rowIndex={taskIndex.get(String(selectedTaskId))}
            />
          )}

          {tasks.map((task, i) => {
            const y = i * ROW_HEIGHT;
            return (
              <g key={task.id}>
                <line x1={0} y1={y + ROW_HEIGHT} x2={chartWidth} y2={y + ROW_HEIGHT} stroke="var(--color-border-subtle)" strokeWidth={1} />

                {show.baseline && task.baselineStart && task.baselineEnd && (
                  <BaselineBar task={task} y={y} minDate={minDate} unitWidth={unitWidth} />
                )}

                {task.isParent ? (
                  <SummaryBar
                    task={task}
                    y={y}
                    minDate={minDate}
                    unitWidth={unitWidth}
                    showTaskNames={show.taskNames}
                    showProgressPercent={show.progressPercent}
                  />
                ) : task.duration === 0 ? (
                  <MilestoneDiamond task={task} y={y} minDate={minDate} unitWidth={unitWidth} showCritical={show.criticalPath} onUpdateTaskFields={onUpdateTaskFields} selected={String(task.id) === String(selectedTaskId)} onSelect={onSelectTask} categoryColors={categoryColors} onBeginDrag={onBeginDrag} onEndDrag={onEndDrag} showTaskNames={show.taskNames} />
                ) : (
                  <TaskBar
                    task={task}
                    y={y}
                    minDate={minDate}
                    unitWidth={unitWidth}
                    showCritical={show.criticalPath}
                    showSlack={show.slack}
                    onUpdateTaskFields={onUpdateTaskFields}
                    selected={String(task.id) === String(selectedTaskId)}
                    onSelect={onSelectTask}
                    categoryColors={categoryColors}
                    onBeginDrag={onBeginDrag}
                    onEndDrag={onEndDrag}
                    showTaskNames={show.taskNames}
                    showProgressPercent={show.progressPercent}
                  />
                )}
              </g>
            );
          })}

          {deps.map((dep, i) => (
            <DependencyArrow key={`${dep.from}-${dep.to}-${i}`} dep={dep} tasks={tasks} taskIndex={taskIndex} minDate={minDate} unitWidth={unitWidth} />
          ))}

          {show.todayLine && <TodayLine minDate={minDate} unitWidth={unitWidth} chartHeight={bodyHeight} />}

          {isPicking && (
            <rect x={0} y={0} width={chartWidth} height={bodyHeight}
              fill="transparent" style={{ cursor: 'crosshair' }} />
          )}
        </svg>
      </div>
    </div>
  );
}

function MonthLabels({ minDate, totalDays, unitWidth, chartWidth, height, scale }) {
  const labels = [];
  const base = parseLocal(minDate);
  const visibleDays = Math.ceil((chartWidth || totalDays * unitWidth) / unitWidth) + 1;
  const drawDays = Math.max(totalDays, visibleDays);
  const fullWidth = chartWidth || drawDays * unitWidth;

  if (scale === 'day') {
    let lastMonth = -1;
    for (let i = 0; i < drawDays; i++) {
      const d = new Date(base);
      d.setDate(d.getDate() + i);
      const month = d.getMonth();
      if (month !== lastMonth) {
        lastMonth = month;
        const monthName = d.toLocaleDateString('en-US', { month: 'short' });
        const showYear = d.getMonth() === 0 || i === 0;
        const monthText = showYear ? `${monthName} ${d.getFullYear()}` : monthName;
        labels.push(
          <text key={i} x={i * unitWidth + 3} y={height - 3} fill="var(--color-text-muted)" fontSize={11} fontWeight={600}>{monthText}</text>,
        );
      }
    }
  } else {
    const startDate = new Date(base);
    let weekStart = new Date(startDate);
    weekStart.setDate(weekStart.getDate() - weekStart.getDay() + 1);
    let weekIndex = 0;
    const totalW = chartWidth || drawDays * unitWidth;
    let lastMonth = -1;
    while (daysBetween(toLocalIso(startDate), toLocalIso(weekStart)) < drawDays + 7) {
      const dayOffset = daysBetween(minDate, toLocalIso(weekStart));
      const x = dayOffset * unitWidth;
      if (x >= -unitWidth * 7 && x < totalW + unitWidth * 7) {
        const month = weekStart.getMonth();
        if (month !== lastMonth) {
          lastMonth = month;
          const monthName = weekStart.toLocaleDateString('en-US', { month: 'short' });
          const showYear = weekStart.getMonth() === 0 || weekIndex === 0;
          const monthText = showYear ? `${monthName} ${weekStart.getFullYear()}` : monthName;
          labels.push(
            <text key={weekIndex} x={Math.max(x, 0) + 3} y={height - 3} fill="var(--color-text-muted)" fontSize={11} fontWeight={600}>{monthText}</text>,
          );
        }
      }
      weekStart.setDate(weekStart.getDate() + 7);
      weekIndex++;
      if (weekIndex > 200) break;
    }
  }

  return (
    <g>
      <rect x={0} y={0} width={fullWidth} height={height} fill="var(--color-bg-secondary)" />
      {labels}
    </g>
  );
}

function DayLabels({ minDate, totalDays, unitWidth, chartWidth, height, scale }) {
  const labels = [];
  const base = parseLocal(minDate);
  const visibleDays = Math.ceil((chartWidth || totalDays * unitWidth) / unitWidth) + 1;
  const drawDays = Math.max(totalDays, visibleDays);
  const fullWidth = chartWidth || drawDays * unitWidth;

  if (scale === 'day') {
    const getDayLabelStep = () => {
      if (unitWidth >= 16) return 1;
      if (unitWidth >= 12) return 2;
      if (unitWidth >= 9) return 3;
      if (unitWidth >= 7) return 5;
      if (unitWidth >= 5) return 7;
      return 14;
    };
    const labelStep = getDayLabelStep();
    const compactLabel = labelStep >= 7;
    for (let i = 0; i < drawDays; i++) {
      const d = new Date(base);
      d.setDate(d.getDate() + i);
      const x = i * unitWidth;
      const isWeekend = d.getDay() === 0 || d.getDay() === 6;
      const isMonthStart = d.getDate() === 1;
      const shouldShowDay = i === 0 || isMonthStart || i % labelStep === 0;
      labels.push(
        <g key={i}>
          {isWeekend && <rect x={x} y={0} width={unitWidth} height={height} fill="var(--color-bg-tertiary)" opacity={0.4} />}
          {shouldShowDay && (
            <text
              x={x + unitWidth / 2}
              y={height - 3}
              fill={isWeekend ? 'var(--color-text-muted)' : 'var(--color-text-secondary)'}
              fontSize={compactLabel ? 9 : 10}
              fontWeight={isMonthStart ? 600 : 500}
              textAnchor="middle"
              opacity={isWeekend ? 0.5 : 0.8}
            >
              {compactLabel ? d.toLocaleDateString('en-US', { month: 'numeric', day: 'numeric' }) : d.getDate()}
            </text>
          )}
        </g>,
      );
    }
  } else {
    const startDate = new Date(base);
    let weekStart = new Date(startDate);
    weekStart.setDate(weekStart.getDate() - weekStart.getDay() + 1);
    let weekIndex = 0;
    const totalW = chartWidth || drawDays * unitWidth;
    while (daysBetween(toLocalIso(startDate), toLocalIso(weekStart)) < drawDays + 7) {
      const dayOffset = daysBetween(minDate, toLocalIso(weekStart));
      const x = dayOffset * unitWidth;
      if (x >= -unitWidth * 7 && x < totalW + unitWidth * 7) {
        labels.push(
          <text key={weekIndex} x={x + 3} y={height - 3} fill="var(--color-text-secondary)" fontSize={9} opacity={0.8}>
            {weekStart.toLocaleDateString('en-US', { month: 'numeric', day: 'numeric' })}
          </text>,
        );
      }
      weekStart.setDate(weekStart.getDate() + 7);
      weekIndex++;
      if (weekIndex > 200) break;
    }
  }

  return (
    <g>
      <rect x={0} y={0} width={fullWidth} height={height} fill="var(--color-bg-secondary)" />
      {labels}
    </g>
  );
}

function WeekLabels({ minDate, totalDays, unitWidth, chartWidth, height }) {
  const labels = [];
  const base = parseLocal(minDate);
  const visibleDays = Math.ceil((chartWidth || totalDays * unitWidth) / unitWidth) + 1;
  const drawDays = Math.max(totalDays, visibleDays);
  const weekWidth = 7 * unitWidth;

  // Find the Monday on or before the project start
  let weekStart = new Date(base);
  weekStart.setDate(weekStart.getDate() - ((weekStart.getDay() + 6) % 7));

  let weekNum = 1;
  const limit = Math.ceil(drawDays / 7) + 2;

  for (let i = 0; i < limit; i++) {
    const dayOffset = daysBetween(minDate, toLocalIso(weekStart));
    const x = dayOffset * unitWidth;
    const fullWidth = chartWidth || drawDays * unitWidth;

    if (x + weekWidth >= 0 && x < fullWidth) {
      labels.push(
        <g key={i}>
          <rect x={x} y={0} width={weekWidth} height={height} fill="transparent" />
          <line x1={x} y1={0} x2={x} y2={height} stroke="var(--color-border-subtle)" strokeWidth={1} opacity={0.5} />
          <text x={x + 4} y={height - 4} fill="var(--color-text-muted)" fontSize={10} fontWeight={500}>
            W{weekNum}
          </text>
        </g>,
      );
    }

    weekStart.setDate(weekStart.getDate() + 7);
    weekNum++;
  }

  const fullWidth = chartWidth || drawDays * unitWidth;
  return (
    <g>
      <rect x={0} y={0} width={fullWidth} height={height} fill="var(--color-bg-secondary)" />
      {labels}
    </g>
  );
}

function BodyGrid({ minDate, totalDays, unitWidth, bodyHeight, chartWidth, scale }) {
  const lines = [];
  const base = parseLocal(minDate);
  const visibleDays = Math.ceil(chartWidth / unitWidth) + 1;
  const drawDays = Math.max(totalDays, visibleDays);

  if (scale === 'day') {
    for (let i = 0; i < drawDays; i++) {
      const d = new Date(base);
      d.setDate(d.getDate() + i);
      const x = i * unitWidth;
      const isWeekend = d.getDay() === 0 || d.getDay() === 6;

      if (isWeekend) {
        lines.push(<rect key={`w${i}`} x={x} y={0} width={unitWidth} height={bodyHeight} fill="var(--color-bg-tertiary)" opacity={0.4} />);
      }
      lines.push(<line key={`l${i}`} x1={x} y1={0} x2={x} y2={bodyHeight} stroke="var(--color-border-subtle)" strokeWidth={1} opacity={0.3} />);
    }
  } else {
    const startDate = new Date(base);
    let weekStart = new Date(startDate);
    weekStart.setDate(weekStart.getDate() - weekStart.getDay() + 1);
    let weekIndex = 0;
    while (daysBetween(toLocalIso(startDate), toLocalIso(weekStart)) < drawDays + 7) {
      const dayOffset = daysBetween(minDate, toLocalIso(weekStart));
      const x = dayOffset * unitWidth;
      if (x >= 0 && x < chartWidth) {
        lines.push(<line key={`wk${weekIndex}`} x1={x} y1={0} x2={x} y2={bodyHeight} stroke="var(--color-border-subtle)" strokeWidth={1} opacity={0.4} />);
      }
      weekStart.setDate(weekStart.getDate() + 7);
      weekIndex++;
      if (weekIndex > 200) break;
    }
  }

  return <g>{lines}</g>;
}

function HoverDateHighlight({ date, minDate, unitWidth, height, label }) {
  if (!date) return null;
  const offset = daysBetween(minDate, date);
  const x = offset * unitWidth;
  const pillW = label && label.length > 16 ? 160 : 80;
  return (
    <g style={{ pointerEvents: 'none' }}>
      <rect
        x={x} y={0} width={unitWidth} height={height}
        fill="var(--color-accent)" opacity={0.10}
      />
      {label && (
        <>
          <rect x={x + unitWidth / 2 - pillW / 2} y={2} width={pillW} height={16} rx={3}
            fill="var(--color-accent)" opacity={0.85} />
          <text x={x + unitWidth / 2} y={13} textAnchor="middle"
            fill="var(--color-on-accent)" fontSize={11} fontWeight={500}>
            {label}
          </text>
        </>
      )}
    </g>
  );
}

function DragPreviewBar({ dragStart, dragEnd, minDate, unitWidth, rowIndex }) {
  if (rowIndex == null || !dragStart || !dragEnd) return null;
  const s = dragStart < dragEnd ? dragStart : dragEnd;
  const e = dragStart < dragEnd ? dragEnd : dragStart;
  const startOffset = daysBetween(minDate, s);
  const endOffset = daysBetween(minDate, e);
  const x = startOffset * unitWidth;
  const width = Math.max((endOffset - startOffset + 1) * unitWidth - 2, 4);
  const y = rowIndex * ROW_HEIGHT + (ROW_HEIGHT - BAR_HEIGHT) / 2;
  return (
    <g style={{ pointerEvents: 'none' }}>
      <rect x={x} y={y} width={width} height={BAR_HEIGHT} rx={3}
        fill="var(--color-accent)" opacity={0.3} />
      <rect x={x} y={y} width={width} height={BAR_HEIGHT} rx={3}
        fill="none" stroke="var(--color-accent)" strokeWidth={1.5} strokeDasharray="4 2" opacity={0.7} />
    </g>
  );
}

function clampProgress(progress) {
  const numeric = Number(progress);
  if (!Number.isFinite(numeric)) return 0;
  return Math.max(0, Math.min(100, Math.round(numeric)));
}

function TaskBar({ task, y, minDate, unitWidth, showCritical, showSlack, onUpdateTaskFields, selected, onSelect, categoryColors = {}, onBeginDrag, onEndDrag, showTaskNames, showProgressPercent }) {
  if (!task.startDate || !task.endDate) return null;
  const startOffset = daysBetween(minDate, task.startDate);
  const duration = daysBetween(task.startDate, task.endDate) + 1;
  const x = startOffset * unitWidth;
  const width = Math.max(duration * unitWidth - 2, 4);
  const barY = y + (ROW_HEIGHT - BAR_HEIGHT) / 2;
  const isCritical = showCritical && task.isCritical;
  const catColor = !showCritical && task.category ? categoryColors[task.category.trim()] : null;
  const barColor = isCritical ? 'var(--color-critical-path)' : (catColor || 'var(--color-accent)');
  const progress = clampProgress(task.progress);
  const renderedProgress = showProgressPercent ? progress : 100;
  const progressWidth = width * (renderedProgress / 100);
  const progressLabel = `${progress}%`;
  const progressTextX = progress > 0 ? x + progressWidth / 2 : x + width / 2;
  const progressTextY = barY + BAR_HEIGHT / 2 + 3.5;
  const progressClipId = `task-progress-${String(task.id)}`;
  const slackWidth = showSlack && task.totalFloat > 0 ? task.totalFloat * unitWidth : 0;
  const resizeWidth = 6;

  const handleMoveStart = (e) => {
    e.stopPropagation();
    if (onSelect) onSelect(String(task.id));
    if (!onUpdateTaskFields) return;
    if (onBeginDrag) onBeginDrag();
    const startX = e.clientX;
    const origStart = task.startDate;
    const origEnd = task.endDate;

    const onMove = (ev) => {
      const deltaDays = Math.round((ev.clientX - startX) / unitWidth);
      if (deltaDays === 0) return;
      onUpdateTaskFields(task.id, {
        startDate: addDaysIso(origStart, deltaDays),
        endDate: addDaysIso(origEnd, deltaDays),
      });
    };
    const onUp = () => {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
      if (onEndDrag) onEndDrag();
    };
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  };

  const handleResizeEndStart = (e) => {
    e.stopPropagation();
    if (onSelect) onSelect(String(task.id));
    if (!onUpdateTaskFields) return;
    if (onBeginDrag) onBeginDrag();
    const startX = e.clientX;
    const origEnd = task.endDate;

    const onMove = (ev) => {
      const deltaDays = Math.round((ev.clientX - startX) / unitWidth);
      if (deltaDays === 0) return;
      const newEnd = addDaysIso(origEnd, deltaDays);
      if (newEnd >= task.startDate) {
        onUpdateTaskFields(task.id, { endDate: newEnd });
      }
    };
    const onUp = () => {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
      if (onEndDrag) onEndDrag();
    };
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  };

  const handleResizeStartStart = (e) => {
    e.stopPropagation();
    if (onSelect) onSelect(String(task.id));
    if (!onUpdateTaskFields) return;
    if (onBeginDrag) onBeginDrag();
    const startX = e.clientX;
    const origStart = task.startDate;

    const onMove = (ev) => {
      const deltaDays = Math.round((ev.clientX - startX) / unitWidth);
      if (deltaDays === 0) return;
      const newStart = addDaysIso(origStart, deltaDays);
      if (newStart <= task.endDate) {
        onUpdateTaskFields(task.id, { startDate: newStart });
      }
    };
    const onUp = () => {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
      if (onEndDrag) onEndDrag();
    };
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  };

  const handleZone = Math.min(resizeWidth, width / 3);
  const middleX = x + handleZone;
  const middleWidth = Math.max(width - handleZone * 2, 2);
  const clampedSlackW = Math.min(slackWidth, 400);
  const nameX = slackWidth > 0
    ? x + width + clampedSlackW + 4
    : x + width + 4;

  return (
    <g onClick={(e) => e.stopPropagation()}>
      <defs>
        <clipPath id={progressClipId}>
          <rect x={x} y={barY} width={width} height={BAR_HEIGHT} rx={3} />
        </clipPath>
      </defs>
      {selected && (
        <rect x={0} y={y} width={9999} height={ROW_HEIGHT} fill={barColor} opacity={0.06} />
      )}
      {slackWidth > 0 && (
        <g>
          <rect x={x + width + 1} y={barY + 1} width={clampedSlackW} height={BAR_HEIGHT - 2} rx={2}
            fill="var(--color-warning)" opacity={0.13} />
          <rect x={x + width + 1} y={barY + 1} width={clampedSlackW} height={BAR_HEIGHT - 2} rx={2}
            fill="url(#slack-hatch)" opacity={0.45} />
        </g>
      )}
      <rect x={x} y={barY} width={width} height={BAR_HEIGHT} rx={3} fill={barColor} opacity={0.25} />
      {progressWidth > 0 && <rect x={x} y={barY} width={progressWidth} height={BAR_HEIGHT} rx={3} fill={barColor} opacity={0.85} />}
      {showProgressPercent && (
        <text
          x={progressTextX}
          y={progressTextY}
          textAnchor="middle"
          fill="#ffffff"
          fontSize={11}
          fontWeight={700}
          clipPath={`url(#${progressClipId})`}
          style={{ pointerEvents: 'none' }}
        >
          {progressLabel}
        </text>
      )}
      {selected && (
        <rect x={x - 2} y={barY - 2} width={width + 4} height={BAR_HEIGHT + 4}
          rx={4} fill="none" stroke={barColor} strokeWidth={2} opacity={0.8} />
      )}
      <rect x={x} y={barY + 3} width={2} height={BAR_HEIGHT - 6} rx={1} fill={barColor} opacity={selected ? 0.9 : 0.5} />
      <rect x={x + width - 2} y={barY + 3} width={2} height={BAR_HEIGHT - 6} rx={1} fill={barColor} opacity={selected ? 0.9 : 0.5} />
      <rect x={x} y={barY} width={handleZone} height={BAR_HEIGHT} fill="transparent" style={{ cursor: 'w-resize' }} onMouseDown={handleResizeStartStart} />
      <rect x={middleX} y={barY} width={middleWidth} height={BAR_HEIGHT} fill="transparent" style={{ cursor: 'grab' }} onMouseDown={handleMoveStart} />
      <rect x={x + width - handleZone} y={barY} width={handleZone} height={BAR_HEIGHT} fill="transparent" style={{ cursor: 'e-resize' }} onMouseDown={handleResizeEndStart} />
      {slackWidth > 0 && (
        <text x={x + width + clampedSlackW / 2} y={barY + BAR_HEIGHT / 2 + 3}
          textAnchor="middle" fill="var(--color-warning)" fontSize={9} fontWeight={600} opacity={0.8}>
          {task.totalFloat}d
        </text>
      )}
      {showTaskNames && (
        <text x={nameX} y={barY + BAR_HEIGHT / 2 + 3.5}
          fill="var(--color-text-muted)" fontSize={12} fontWeight={500}>
          {task.name}
        </text>
      )}
    </g>
  );
}

function MilestoneDiamond({ task, y, minDate, unitWidth, showCritical, onUpdateTaskFields, selected, onSelect, categoryColors = {}, onBeginDrag, onEndDrag, showTaskNames }) {
  if (!task.startDate) return null;
  const offset = daysBetween(minDate, task.startDate);
  const cx = offset * unitWidth + unitWidth / 2;
  const cy = y + ROW_HEIGHT / 2;
  const size = 7;
  const isCritical = showCritical && task.isCritical;
  const catColor = !showCritical && task.category ? categoryColors[task.category.trim()] : null;
  const color = isCritical ? 'var(--color-critical-path)' : (catColor || 'var(--color-warning)');

  const handleMoveStart = (e) => {
    e.stopPropagation();
    if (onSelect) onSelect(String(task.id));
    if (!onUpdateTaskFields) return;
    if (onBeginDrag) onBeginDrag();
    const startX = e.clientX;
    const origStart = task.startDate;

    const onMove = (ev) => {
      const deltaDays = Math.round((ev.clientX - startX) / unitWidth);
      if (deltaDays === 0) return;
      onUpdateTaskFields(task.id, { startDate: addDaysIso(origStart, deltaDays) });
    };
    const onUp = () => {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
      if (onEndDrag) onEndDrag();
    };
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  };

  return (
    <g style={{ cursor: 'grab' }} onMouseDown={handleMoveStart} onClick={(e) => e.stopPropagation()}>
      {selected && <rect x={0} y={y} width={9999} height={ROW_HEIGHT} fill={color} opacity={0.06} />}
      {selected && <circle cx={cx} cy={cy} r={size + 3} fill="none" stroke={color} strokeWidth={2} opacity={0.7} />}
      <polygon points={`${cx},${cy - size} ${cx + size},${cy} ${cx},${cy + size} ${cx - size},${cy}`} fill={color} opacity={0.9} />
      <rect x={cx - size} y={cy - size} width={size * 2} height={size * 2} fill="transparent" />
      {showTaskNames && <text x={cx + size + 4} y={cy + 3.5} fill="var(--color-text-muted)" fontSize={12} fontWeight={500}>{task.name}</text>}
    </g>
  );
}

function SummaryBar({ task, y, minDate, unitWidth, showTaskNames, showProgressPercent }) {
  if (!task.startDate || !task.endDate) return null;
  const startOffset = daysBetween(minDate, task.startDate);
  const duration = daysBetween(task.startDate, task.endDate) + 1;
  const x = startOffset * unitWidth;
  const width = Math.max(duration * unitWidth - 2, 4);
  const barY = y + (ROW_HEIGHT - SUMMARY_HEIGHT) / 2;
  const capWidth = 3;
  const progress = clampProgress(task.progress);
  const progressWidth = width * (progress / 100);
  const progressLabel = `${progress}%`;
  const progressTextX = progress > 0 ? x + progressWidth / 2 : x + width / 2;
  const progressTextY = barY + SUMMARY_HEIGHT / 2 + 3.5;
  const progressClipId = `summary-progress-${String(task.id)}`;
  const summaryBaseOpacity = showProgressPercent ? 0.5 : 0.72;
  return (
    <g>
      <defs>
        <clipPath id={progressClipId}>
          <rect x={x} y={barY} width={width} height={SUMMARY_HEIGHT} />
        </clipPath>
      </defs>
      <rect x={x} y={barY} width={width} height={SUMMARY_HEIGHT} fill="var(--color-text-muted)" opacity={summaryBaseOpacity} />
      {showProgressPercent && progressWidth > 0 && (
        <rect x={x} y={barY} width={progressWidth} height={SUMMARY_HEIGHT} fill="var(--color-accent)" opacity={0.88} />
      )}
      {showProgressPercent && (
        <text
          x={progressTextX}
          y={progressTextY}
          textAnchor="middle"
          fill="#ffffff"
          fontSize={10}
          fontWeight={700}
          clipPath={`url(#${progressClipId})`}
          style={{ pointerEvents: 'none' }}
        >
          {progressLabel}
        </text>
      )}
      <rect x={x} y={barY} width={capWidth} height={SUMMARY_HEIGHT + 4} fill="var(--color-text-muted)" opacity={0.7} />
      <rect x={x + width - capWidth} y={barY} width={capWidth} height={SUMMARY_HEIGHT + 4} fill="var(--color-text-muted)" opacity={0.7} />
      {showTaskNames && <text x={x + width + 4} y={barY + SUMMARY_HEIGHT / 2 + 3.5} fill="var(--color-text-secondary)" fontSize={12} fontWeight={700}>{task.name}</text>}
    </g>
  );
}

function BaselineBar({ task, y, minDate, unitWidth }) {
  const startOffset = daysBetween(minDate, task.baselineStart);
  const duration = daysBetween(task.baselineStart, task.baselineEnd) + 1;
  const x = startOffset * unitWidth;
  const width = Math.max(duration * unitWidth - 2, 4);
  const barY = y + ROW_HEIGHT / 2 + BAR_HEIGHT / 2 + 1;
  return <rect x={x} y={barY} width={width} height={BASELINE_HEIGHT} rx={1} fill="var(--color-text-muted)" opacity={0.3} />;
}

function DependencyArrow({ dep, tasks, taskIndex, minDate, unitWidth }) {
  const fromIdx = taskIndex.get(dep.from);
  const toIdx = taskIndex.get(dep.to);
  if (fromIdx == null || toIdx == null) return null;
  const fromTask = tasks[fromIdx];
  const toTask = tasks[toIdx];
  if (!fromTask.endDate || !toTask.startDate) return null;
  const fromEndDay = daysBetween(minDate, fromTask.endDate) + 1;
  const toStartDay = daysBetween(minDate, toTask.startDate);
  const x1 = fromEndDay * unitWidth;
  const y1 = fromIdx * ROW_HEIGHT + ROW_HEIGHT / 2;
  const x2 = toStartDay * unitWidth;
  const y2 = toIdx * ROW_HEIGHT + ROW_HEIGHT / 2;
  const midX = x1 + 8;
  return <path d={`M${x1},${y1} L${midX},${y1} L${midX},${y2} L${x2},${y2}`} fill="none" stroke="var(--color-text-muted)" strokeWidth={1} opacity={0.5} markerEnd="url(#arrowhead)" />;
}

function TodayLine({ minDate, unitWidth, chartHeight }) {
  const today = todayIso();
  const offset = daysBetween(minDate, today);
  if (offset < 0) return null;
  const x = offset * unitWidth + unitWidth / 2;
  return <line x1={x} y1={0} x2={x} y2={chartHeight} stroke="var(--color-danger)" strokeWidth={1.5} strokeDasharray="4 3" opacity={0.7} />;
}

// Parse "YYYY-MM-DD" as local time (noon) to avoid timezone day-shift.
function parseLocal(iso) {
  const [y, m, d] = iso.split('-').map(Number);
  return new Date(y, m - 1, d, 12);
}

function toLocalIso(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function todayIso() {
  return toLocalIso(new Date());
}

function addDaysIso(isoDate, days) {
  const d = parseLocal(isoDate);
  d.setDate(d.getDate() + days);
  return toLocalIso(d);
}

function getMinDate(tasks) {
  let min = null;
  for (const t of tasks) {
    if (t.startDate && (!min || t.startDate < min)) min = t.startDate;
    if (t.baselineStart && (!min || t.baselineStart < min)) min = t.baselineStart;
  }
  return min || todayIso();
}

function getMaxDate(tasks) {
  let max = null;
  for (const t of tasks) {
    if (t.endDate && (!max || t.endDate > max)) max = t.endDate;
    if (t.baselineEnd && (!max || t.baselineEnd > max)) max = t.baselineEnd;
  }
  return max || todayIso();
}

function daysBetween(a, b) {
  return Math.round((parseLocal(b) - parseLocal(a)) / (1000 * 60 * 60 * 24));
}
