import { useState, useEffect, useCallback, useRef } from 'react';
import {
  X,
  ChevronLeft,
  ChevronRight,
  Rocket,
  FileSpreadsheet,
  Table2,
  BarChart3,
  GitBranch,
  Diamond,
  FolderTree,
  Palette,
  SlidersHorizontal,
  Image,
} from 'lucide-react';

const TOUR_STEPS = [
  {
    target: null,
    icon: Rocket,
    title: 'Welcome to GanttGen',
    description:
      'GanttGen is a zero-installation, single-file project management tool that runs entirely in your browser. No backend, no login, no setup required.',
    details: [
      'Works offline and on the file:// protocol -- just open index.html',
      'All project data lives in Excel files (.xlsx) that you import and export',
      'No data leaves your machine; everything runs client-side',
    ],
    placement: 'center',
  },
  {
    target: '[data-guide="toolbar-import"]',
    icon: FileSpreadsheet,
    title: 'Excel Import / Export',
    description:
      'These buttons handle all data persistence. Download a template, fill it in, import it, and save your work back to Excel.',
    details: [
      '"Import Excel" reads your .xlsx and populates the table and chart',
      '"Download Template" gives you a pre-formatted .xlsx with correct column headers',
      '"Save to Excel" exports tasks plus theme, view options, and column settings',
      'Round-trip safe: import, edit, save, re-import -- everything is restored',
    ],
    placement: 'bottom',
  },
  {
    target: '[data-guide="data-table"]',
    icon: Table2,
    title: 'Data Table',
    description:
      'The left pane is an editable task grid. Click any cell to edit inline. Add or delete rows with the controls at the bottom.',
    details: [
      'Click a cell to edit text, numbers, dates, or status',
      'Toggle visible columns via the View dropdown or column picker in the header',
      'Resize columns by dragging header borders',
      'Editing Start Date + Duration auto-calculates End Date (and vice versa)',
    ],
    placement: 'over',
  },
  {
    target: '[data-guide="gantt-chart"]',
    icon: BarChart3,
    title: 'Gantt Chart',
    description:
      'The right pane renders an interactive SVG Gantt chart synchronized with the data table.',
    details: [
      'Drag a bar horizontally to move its start/end dates',
      'Drag the right edge of a bar to resize (change duration)',
      'Day / Week scale toggle and +/- zoom buttons in the chart header',
      'Click a bar to select the corresponding row in the data table',
      'A vertical "Today" line marks the current date',
    ],
    placement: 'over',
  },
  {
    target: '[data-guide="gantt-chart"]',
    icon: GitBranch,
    title: 'Dependencies & Critical Path',
    description:
      'Define task dependencies using predecessor IDs. GanttGen automatically computes the Critical Path (CPM) and draws arrows on the chart.',
    details: [
      'Enter predecessor task IDs in the Dependency column (comma-separated, e.g. "2,3")',
      'Dependency arrows are drawn on the chart between linked tasks',
      'Critical path tasks (zero total float) are highlighted in red',
      'Slack/float bars show how much a non-critical task can slip',
    ],
    placement: 'over',
  },
  {
    target: '[data-guide="gantt-chart"]',
    icon: Diamond,
    title: 'Milestones & Baselines',
    description:
      'Milestones are zero-duration tasks shown as diamonds on the chart. Baseline dates let you track schedule variance visually.',
    details: [
      'Set Duration to 0 to create a milestone (renders as a diamond)',
      'Fill in Baseline Start / End columns to record the original schedule',
      'Baseline bars appear as a thin strip behind the actual bar',
      'Toggle baseline visibility in View Options',
    ],
    placement: 'over',
  },
  {
    target: '[data-guide="data-table"]',
    icon: FolderTree,
    title: 'WBS (Work Breakdown Structure)',
    description:
      'Organize tasks hierarchically using the Parent ID column. Parent tasks become summary bars that aggregate child dates and progress.',
    details: [
      "Set a task's Parent ID to another task's ID to make it a child",
      'Parent rows auto-compute start/end dates and average progress',
      'Click the expand/collapse arrow on a parent to hide or show children',
      'Summary bars on the chart span the full duration of child tasks',
    ],
    placement: 'over',
  },
  {
    target: '[data-guide="btn-theme"]',
    icon: Palette,
    title: 'Theming',
    description:
      'Customize the look and feel with built-in presets or your own colors. Category-specific bar colors are also configurable.',
    details: [
      'Three presets: Linear Dark, Notion Light, Classic Tremor',
      'Custom color builder for background, text, accent, and more',
      'Category colors: assign distinct bar colors per task category',
      'Theme settings are saved into your Excel file on export',
    ],
    placement: 'bottom',
  },
  {
    target: '[data-guide="btn-view"]',
    icon: SlidersHorizontal,
    title: 'View Options',
    description:
      'Fine-tune what the chart displays. All preferences are persisted when you save to Excel.',
    details: [
      'Toggle: Critical Path, Slack, Dependency Arrows, Today Line, Baseline Bars',
      'Toggle: Day/Week/Month labels on the timeline header',
      'Skip Weekends: auto-calculates durations using working days only',
      'Column visibility: show or hide data table columns',
    ],
    placement: 'bottom',
  },
  {
    target: '[data-guide="btn-export-png"]',
    icon: Image,
    title: 'PNG Export',
    description:
      'Export the Gantt chart area as a PNG image for presentations, reports, or documentation.',
    details: [
      'Click "Export PNG" to download a snapshot of the chart',
      'The exported image uses the current theme colors and background',
      'Only the chart pane is captured (not the data table)',
    ],
    placement: 'bottom',
  },
];

export default function GuideOverlay({ open, onClose }) {
  const [step, setStep] = useState(0);
  const [targetRect, setTargetRect] = useState(null);
  const tooltipRef = useRef(null);

  const currentStep = TOUR_STEPS[step];
  const totalSteps = TOUR_STEPS.length;

  const measureTarget = useCallback(() => {
    if (!open) return;
    const s = TOUR_STEPS[step];
    if (!s.target) {
      setTargetRect(null);
      return;
    }
    const el = document.querySelector(s.target);
    if (!el) {
      setTargetRect(null);
      return;
    }
    const rect = el.getBoundingClientRect();
    setTargetRect({
      top: rect.top,
      left: rect.left,
      width: rect.width,
      height: rect.height,
    });
  }, [open, step]);

  useEffect(() => {
    if (!open) {
      setStep(0);
      setTargetRect(null);
      return;
    }
    measureTarget();
  }, [open, step, measureTarget]);

  useEffect(() => {
    if (!open) return;
    const onResize = () => measureTarget();
    window.addEventListener('resize', onResize);
    window.addEventListener('scroll', onResize, true);
    return () => {
      window.removeEventListener('resize', onResize);
      window.removeEventListener('scroll', onResize, true);
    };
  }, [open, measureTarget]);

  const goNext = () => {
    if (step < totalSteps - 1) setStep(step + 1);
    else onClose();
  };
  const goPrev = () => {
    if (step > 0) setStep(step - 1);
  };

  useEffect(() => {
    if (!open) return;
    const onKey = (e) => {
      if (e.key === 'Escape') onClose();
      if (e.key === 'ArrowRight' || e.key === 'Enter') goNext();
      if (e.key === 'ArrowLeft') goPrev();
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  });

  if (!open) return null;

  const Icon = currentStep.icon;
  const isCenter = !targetRect || currentStep.placement === 'center';
  const PADDING = 8;

  const tooltipStyle = {};
  if (!isCenter && targetRect) {
    const p = currentStep.placement;
    const TOOLTIP_MAX_W = 380;
    const GAP = 12;

    if (p === 'bottom') {
      tooltipStyle.top = targetRect.top + targetRect.height + GAP;
      tooltipStyle.left = Math.max(
        16,
        Math.min(
          targetRect.left + targetRect.width / 2 - TOOLTIP_MAX_W / 2,
          window.innerWidth - TOOLTIP_MAX_W - 16,
        ),
      );
    } else if (p === 'right') {
      tooltipStyle.top = Math.max(16, targetRect.top);
      tooltipStyle.left = targetRect.left + targetRect.width + GAP;
    } else if (p === 'left') {
      tooltipStyle.top = Math.max(16, targetRect.top);
      tooltipStyle.left = Math.max(16, targetRect.left - TOOLTIP_MAX_W - GAP);
    } else if (p === 'over') {
      tooltipStyle.top = Math.max(
        targetRect.top + 40,
        targetRect.top + (targetRect.height - 300) / 2,
      );
      tooltipStyle.left = Math.max(
        16,
        Math.min(
          targetRect.left + (targetRect.width - TOOLTIP_MAX_W) / 2,
          window.innerWidth - TOOLTIP_MAX_W - 16,
        ),
      );
    }
    tooltipStyle.maxWidth = TOOLTIP_MAX_W;
  }

  return (
    <div className="fixed inset-0 z-[9999]">
      {isCenter ? (
        <div
          className="absolute inset-0"
          style={{ backgroundColor: 'rgba(0,0,0,0.6)' }}
          onClick={onClose}
        />
      ) : (
        <svg className="absolute inset-0 w-full h-full" style={{ pointerEvents: 'none' }}>
          <defs>
            <mask id="guide-spotlight-mask">
              <rect x="0" y="0" width="100%" height="100%" fill="white" />
              {targetRect && (
                <rect
                  x={targetRect.left - PADDING}
                  y={targetRect.top - PADDING}
                  width={targetRect.width + PADDING * 2}
                  height={targetRect.height + PADDING * 2}
                  rx="8"
                  fill="black"
                />
              )}
            </mask>
          </defs>
          <rect
            x="0" y="0" width="100%" height="100%"
            fill="rgba(0,0,0,0.55)"
            mask="url(#guide-spotlight-mask)"
            style={{ pointerEvents: 'auto' }}
            onClick={onClose}
          />
          {targetRect && (
            <rect
              x={targetRect.left - PADDING}
              y={targetRect.top - PADDING}
              width={targetRect.width + PADDING * 2}
              height={targetRect.height + PADDING * 2}
              rx="8"
              fill="none"
              stroke="var(--color-accent)"
              strokeWidth="2"
              className="animate-pulse"
            />
          )}
        </svg>
      )}

      <div
        ref={tooltipRef}
        className={
          isCenter
            ? 'absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 w-full max-w-md mx-4'
            : 'absolute'
        }
        style={{
          ...(!isCenter ? tooltipStyle : {}),
          zIndex: 10000,
        }}
      >
        <div
          className="rounded-xl shadow-2xl overflow-hidden"
          style={{
            backgroundColor: 'var(--color-bg-secondary)',
            border: '1px solid var(--color-border)',
          }}
        >
          <div
            className="flex items-center justify-between px-5 py-3"
            style={{ borderBottom: '1px solid var(--color-border)' }}
          >
            <div className="flex items-center gap-2.5">
              <div
                className="flex items-center justify-center w-7 h-7 rounded-md"
                style={{ backgroundColor: 'var(--color-accent-muted)' }}
              >
                <Icon size={15} style={{ color: 'var(--color-accent)' }} />
              </div>
              <h3
                className="text-sm font-semibold"
                style={{ color: 'var(--color-text-primary)' }}
              >
                {currentStep.title}
              </h3>
            </div>
            <button
              onClick={onClose}
              className="p-1 rounded-md cursor-pointer transition-colors"
              style={{ color: 'var(--color-text-muted)' }}
              onMouseEnter={(e) =>
                (e.currentTarget.style.backgroundColor = 'var(--color-bg-hover)')
              }
              onMouseLeave={(e) =>
                (e.currentTarget.style.backgroundColor = 'transparent')
              }
            >
              <X size={16} />
            </button>
          </div>

          <div className="px-5 py-4">
            <p
              className="text-xs leading-relaxed mb-3"
              style={{ color: 'var(--color-text-secondary)' }}
            >
              {currentStep.description}
            </p>

            {currentStep.details && currentStep.details.length > 0 && (
              <ul className="flex flex-col gap-1.5 mb-3">
                {currentStep.details.map((item, i) => (
                  <li key={i} className="flex items-start gap-2">
                    <ChevronRight
                      size={11}
                      className="mt-0.5 flex-shrink-0"
                      style={{ color: 'var(--color-accent)' }}
                    />
                    <span
                      className="text-xs leading-relaxed"
                      style={{ color: 'var(--color-text-secondary)' }}
                    >
                      {item}
                    </span>
                  </li>
                ))}
              </ul>
            )}
          </div>

          <div
            className="flex items-center justify-between px-5 py-3"
            style={{ borderTop: '1px solid var(--color-border)' }}
          >
            <div className="flex items-center gap-1.5">
              {Array.from({ length: totalSteps }).map((_, i) => (
                <button
                  key={i}
                  onClick={() => setStep(i)}
                  className="w-2 h-2 rounded-full cursor-pointer transition-all"
                  style={{
                    backgroundColor:
                      i === step ? 'var(--color-accent)' : 'var(--color-border)',
                    transform: i === step ? 'scale(1.3)' : 'scale(1)',
                  }}
                />
              ))}
            </div>

            <div className="flex items-center gap-3">
              <span
                className="text-[10px] tabular-nums"
                style={{ color: 'var(--color-text-muted)' }}
              >
                {step + 1} / {totalSteps}
              </span>

              <div className="flex items-center gap-1.5">
                <button
                  onClick={goPrev}
                  disabled={step === 0}
                  className="p-1.5 rounded-md cursor-pointer transition-colors disabled:opacity-30 disabled:cursor-default"
                  style={{
                    backgroundColor: 'var(--color-bg-tertiary)',
                    color: 'var(--color-text-secondary)',
                    border: '1px solid var(--color-border)',
                  }}
                  onMouseEnter={(e) => {
                    if (step > 0)
                      e.currentTarget.style.backgroundColor = 'var(--color-bg-hover)';
                  }}
                  onMouseLeave={(e) => {
                    e.currentTarget.style.backgroundColor = 'var(--color-bg-tertiary)';
                  }}
                >
                  <ChevronLeft size={14} />
                </button>

                <button
                  onClick={goNext}
                  className="flex items-center gap-1 px-3 py-1.5 rounded-md text-xs font-medium cursor-pointer transition-colors"
                  style={{
                    backgroundColor: 'var(--color-accent)',
                    color: '#fff',
                  }}
                  onMouseEnter={(e) =>
                    (e.currentTarget.style.backgroundColor =
                      'var(--color-accent-hover)')
                  }
                  onMouseLeave={(e) =>
                    (e.currentTarget.style.backgroundColor = 'var(--color-accent)')
                  }
                >
                  {step === totalSteps - 1 ? 'Finish' : 'Next'}
                  {step < totalSteps - 1 && <ChevronRight size={14} />}
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
