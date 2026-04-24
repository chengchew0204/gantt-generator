import { useState, useEffect, useLayoutEffect, useCallback, useRef } from 'react';
import {
  X,
  ChevronLeft,
  ChevronRight,
  Rocket,
  FileSpreadsheet,
  FolderTree,
  BarChart3,
  GitBranch,
  Diamond,
  LayoutGrid,
  Palette,
  SlidersHorizontal,
  Share2,
} from 'lucide-react';

const TOUR_STEPS = [
  {
    target: null,
    icon: Rocket,
    title: 'Welcome to GanttGen',
    description:
      'A zero-installation, single-file project tool that runs entirely in your browser. No backend, no login, no setup -- and your data never leaves your machine.',
    details: [
      'Works offline; opens directly from index.html on file://',
      'Excel (.xlsx) is the only persistence layer -- import, edit, save, repeat',
      'Use Arrow keys, Enter to advance, or Esc to exit this tour',
    ],
    placement: 'center',
  },
  {
    target: '[data-guide="toolbar-import"]',
    icon: FileSpreadsheet,
    title: 'Top Toolbar',
    description:
      'Everything you need to load, save, and share lives along the top. Click the project name to rename it; the small arrows beside it are Undo / Redo (Ctrl+Z, Ctrl+Shift+Z).',
    details: [
      'Import Excel / Download Template -- start from your data or a blank file',
      'Save to Excel -- exports tasks, theme, view options, columns, and grid sheets',
      'Export PNG -- chart-only or full layout. Share -- a single-file HTML download',
      'Below 1280 px the labels collapse to icons; hover to see the name',
    ],
    placement: 'bottom',
  },
  {
    target: '[data-guide="data-table"]',
    icon: FolderTree,
    title: 'Data Table & WBS',
    description:
      'The left pane is an editable task grid. Click any cell to edit. Use the Parent ID column to nest tasks -- multi-level hierarchies render automatic indent guides.',
    details: [
      'Editing Start Date + Duration auto-fills End Date (Skip Weekends honored)',
      'Progress (%) updates the bar live as you type',
      'Collapse a parent to hide every descendant; deleting a parent cascades',
    ],
    placement: 'over',
  },
  {
    target: '[data-guide="gantt-chart"]',
    icon: BarChart3,
    title: 'Gantt Chart',
    description:
      'The right pane is an interactive SVG chart synced to the data table. Drag, resize, and zoom to shape your schedule.',
    details: [
      'Drag a bar to move it; drag its right edge to resize the duration',
      'Click a bar to select the matching row',
      'Day / Week scale and 5%-300% zoom live in the bottom status bar',
    ],
    placement: 'over',
  },
  {
    target: '[data-guide="gantt-chart"]',
    icon: GitBranch,
    title: 'Dependencies & Critical Path',
    description:
      'Type predecessor IDs into the Dependency column (e.g. "2,3"). GanttGen runs CPM live and can highlight the critical path and slack.',
    details: [
      'Critical (zero-slack) tasks render in red; slack bars show how far a task can slip',
      'Critical Path and Slack are off by default -- enable them in View options',
      'Dependency arrows and the Today line are toggled there too',
    ],
    placement: 'over',
  },
  {
    target: '[data-guide="gantt-chart"]',
    icon: Diamond,
    title: 'Milestones & Baselines',
    description:
      'Use zero-duration tasks for milestones. Fill the Baseline columns to track schedule variance against your original plan.',
    details: [
      'Duration = 0 renders as a diamond on the chart',
      'Baseline Start / End create a thin strip behind the actual bar',
      'Toggle baselines on or off in View options',
    ],
    placement: 'over',
  },
  {
    target: '[data-guide="status-bar"]',
    icon: LayoutGrid,
    title: 'Spreadsheet Tabs',
    description:
      'The bottom tab strip lets you add full Excel-like data sheets next to your Gantt -- handy for budgets, references, or supporting calculations. Click + to add a tab.',
    details: [
      '400+ formula functions, plus bold / italic / underline, colors, borders, and alignment',
      'Excel-parity number formats (Currency, Accounting, Percentage...) with Ctrl+Shift+1..6',
      'Merge cells, draw shapes (rectangles, arrows, lines, text boxes), copy/paste with Excel',
      'Each tab saves as its own sheet inside your .xlsx',
    ],
    placement: 'top',
  },
  {
    target: '[data-guide="btn-theme"]',
    icon: Palette,
    title: 'Theming',
    description:
      'Switch presets or build your own. Per-category bar colors keep similar tasks visually grouped.',
    details: [
      'Presets: Linear Dark, Notion Light, Classic Tremor, Gray, Black & White',
      'Custom color builder for background, text, accent, and more',
      'Theme + category colors are saved into your Excel file on export',
    ],
    placement: 'bottom',
  },
  {
    target: '[data-guide="btn-view"]',
    icon: SlidersHorizontal,
    title: 'View Options',
    description:
      'Fine-tune what the chart shows and which columns appear in the data table. Every preference round-trips through Excel.',
    details: [
      'Chart toggles: Critical Path, Slack, Dependency Arrows, Today Line, Baseline',
      'Header labels: Day / Week / Month',
      'Skip Weekends -- compute durations in working days',
      'Column visibility -- show or hide any data table column',
    ],
    placement: 'bottom',
  },
  {
    target: '[data-guide="btn-export-png"]',
    icon: Share2,
    title: 'Export & Share',
    description:
      'Take your project anywhere. Export PNG snapshots for slides, or share the entire app as a single HTML file.',
    details: [
      'Export PNG -- pick "Gantt Chart Only" or "Include Headers" (full dashboard)',
      'Share -- generates a self-contained HTML file you can open offline anywhere',
      'Save to Excel -- the canonical format; everything round-trips losslessly',
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

  // useLayoutEffect (not useEffect) so the new target's rect is measured and
  // `targetRect` updated synchronously after DOM commit but BEFORE the browser
  // paints. Otherwise on every step change the tooltip would render once with
  // (newStep, oldRect) -- placing it at the wrong screen position -- and then
  // snap to the correct spot on the next render, producing a visible flash.
  useLayoutEffect(() => {
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
    } else if (p === 'top') {
      // Anchor the tooltip's bottom edge GAP px above the target's top edge.
      // Using `bottom` (rather than `top` with an estimated tooltip height)
      // lets the browser size the tooltip from its content without us having
      // to guess the height. Clamp to >= 16 so it never disappears off-screen
      // when the target sits near the viewport top.
      tooltipStyle.bottom = Math.max(16, window.innerHeight - targetRect.top + GAP);
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
                className="text-[15px] font-semibold"
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
              className="text-[13px] leading-relaxed mb-3"
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
                      className="text-[13px] leading-relaxed"
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
                className="text-[11px] tabular-nums"
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
                  className="flex items-center gap-1 px-3 py-1.5 rounded-md text-[13px] font-medium cursor-pointer transition-colors"
                  style={{
                    backgroundColor: 'var(--color-accent)',
                    color: 'var(--color-on-accent)',
                  }}
                  onMouseEnter={(e) => {
                    e.currentTarget.style.backgroundColor = 'var(--color-accent-hover)';
                    e.currentTarget.style.color = 'var(--color-on-accent-hover)';
                  }}
                  onMouseLeave={(e) => {
                    e.currentTarget.style.backgroundColor = 'var(--color-accent)';
                    e.currentTarget.style.color = 'var(--color-on-accent)';
                  }}
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
