# Backlog

> Update this file when tasks are added, re-prioritized, or completed.

## Active

| Priority | Task | Notes |
|---|---|---|
| Medium | Manual browser E2E testing | Real project file, interactive walkthrough of all features |
| Low | ThemePanel getComputedStyle optimization | Reads DOM during render; should use React state instead |

## Icebox

<!-- Tasks that are identified but not yet prioritized. -->

- Keyboard shortcuts for common actions (import, toggle view)
- Performance spike: SVG vs Canvas for 100+ task projects

## Completed

<!-- Move tasks here when done, with a brief note on outcome. -->

- Project scaffolding: Cursor rules and Obsidian vault set up (2026-03-30)
- Scaffold Vite + React + Tailwind project -- `vite-plugin-singlefile` configured, single-file build verified (2026-03-30)
- Build `App.jsx` root layout -- Dashboard top bar + horizontal split-pane with draggable resizer (2026-03-30)
- Implement `ExcelUtils`: downloadTemplate, importExcel, exportExcel with Settings sheet support (2026-03-30)
- Implement `CpmEngine`: forward/backward pass, total float, critical path detection (2026-03-30)
- Implement `DataTable` component -- fully inline-editable cells, dynamic column visibility, add/delete rows, WBS expand/collapse (2026-04-01)
- Implement `GanttChart` v1 -- SVG bars, milestone diamonds, baseline bars, dependency arrows, WBS summary bars, daily/weekly toggle, today line (2026-04-01)
- Add `Dashboard` summary stats -- live reactive progress %, completed/in-progress/critical counts, current/next task (2026-04-01)
- Implement milestone rendering -- diamond shape for Duration=0 tasks (2026-04-01)
- Implement baseline bars -- thin grey bar below main task bar (2026-04-01)
- Implement WBS parent rows -- expandable/collapsible, auto-computed start/end/progress, summary bar (2026-04-01)
- Implement dependency arrows -- SVG path arrows with arrowhead markers (2026-04-01)
- Multi-dependency support -- comma-separated predecessor IDs (2026-04-01)
- Implement `ThemePanel` -- 3 presets (Linear Dark, Notion Light, Classic Tremor) + custom color builder, persisted to Excel (2026-04-02)
- Implement `ViewOptions` dropdown -- chart/column toggles, skip weekends, all persisted to Excel (2026-04-02)
- Working-days auto-calculation -- `DateUtils.js`, skip weekends toggle, auto-compute End Date from Start+Duration (2026-04-02)
- Add zoom/pan to GanttChart -- CSS transform zoom 50%-300% with +/- controls (2026-04-02)
- PNG export via `html-to-image` -- verified working (2026-04-02)
- Final build verification -- single `dist/index.html` (756 KB), zero external network calls (2026-04-02)
- Code audit and bug fixes -- fixed 7 bugs: hooks rule violation in GanttChart, DateUtils weekend off-by-one, duration string coercion, DataTable undefined guard, ViewOptions effect churn, unused imports (2026-04-02)
- Undo/redo stack for table edits with batch support for drag operations; keyboard shortcuts Ctrl+S (save), Ctrl+Z (undo), Ctrl+Shift+Z (redo); toolbar buttons (2026-04-05)
