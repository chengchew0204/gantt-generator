# Architecture

> Update this file whenever components, data flows, or system boundaries change.

## Overview

GanttGen is a single-page React application bundled into one self-contained `index.html` by `vite-plugin-singlefile`. There is no server, no API, and no database. All state lives in React memory during the session; persistence is achieved entirely through Excel file import/export.

## High-level layout

```
+-----------------------------------------------------------+
|  Top Section -- Project Dashboard                          |
|  (Overall Progress, Status, Current/Next Tasks,            |
|   Action Bar: Import / Template / Save / PNG / Theme /     |
|   View Options)                                            |
+---------------------------+-------------------------------+
|  Left Pane                |  Right Pane                    |
|  Data Table               |  Gantt Chart                   |
|  (editable task rows,     |  (SVG, scrollable,             |
|   toggleable columns)     |   zoomable, day/week toggle)   |
+---------------------------+-------------------------------+
```

## Components

| Component | Responsibility |
|---|---|
| `App.jsx` | Root state, layout wiring, theme CSS variable injection, date auto-calc, WBS tree builder |
| `Dashboard` | Summary stats (progress, completed, in-progress, critical, current/next task), action buttons, hosts ViewOptions dropdown |
| `DataTable` | Editable left-pane task grid with dynamic column visibility, inline editing (text/number/date/status), add/delete rows, WBS expand/collapse |
| `GanttChart` | SVG Gantt renderer: task bars, baseline bars, milestones, today line, dependency arrows, slack bars, WBS summary bars, day/week toggle, zoom controls |
| `ThemePanel` | Slide-over panel with 3 theme presets and custom `<input type="color">` builder; applies CSS variables on `:root` |
| `ViewOptions` | Dropdown with chart element toggles (critical path, slack, dependencies, today line, baseline, skip weekends) and column visibility checkboxes |
| `ExcelUtils` | Import/export logic using `xlsx`; reads and writes the hidden "Settings" sheet with 24 fields (theme, toggles, columns) |
| `CpmEngine` | Critical Path Method: computes Early Start, Early Finish, Late Start, Late Finish, Total Float via Kahn's topological sort |
| `DateUtils` | Working-days arithmetic: `addWorkingDays`, `workingDaysBetween`; skip-weekends support |

## Data flow

```
User edits DataTable
        |
        v
  React state (tasks[])
        |
   +----+----+
   |         |
   v         v
CpmEngine  GanttChart
(slack,    (renders bars,
 critical   critical path,
 path)      baseline, etc.)
   |
   v
Dashboard
(progress %, status, current/next tasks)
        |
        v (on "Save to Excel")
   ExcelUtils.export()
   -> tasks sheet + Settings sheet -> .xlsx download
        |
        v (on "Import Excel")
   ExcelUtils.import()
   -> populates tasks[] + restores theme + view options + columns

User changes theme (ThemePanel)
        |
        v
  CSS variables on :root updated
  Settings state updated
        |
        v (on "Save to Excel")
   Persisted to Settings sheet

User toggles view option (ViewOptions)
        |
        v
  viewOptions state updated -> GanttChart re-renders conditionally
  Settings state updated -> persisted on next export
```

## Key design decisions

- **Single-file output**: `vite-plugin-singlefile` inlines all JS, CSS, and assets into one HTML file. No runtime network calls.
- **Excel as database**: The `.xlsx` file is the only persistence mechanism. A hidden "Settings" sheet stores 24 fields: theme colors, chart toggles, column visibility, and skip-weekends preference.
- **CSS variables for theming**: All colors are defined as CSS custom properties on `:root`. Theme changes update variables only; no component re-renders for color switches. `accent-muted` is computed from accent hex at runtime.
- **CPM on every state change**: The CPM engine runs as a derived computation on the task list whenever tasks change (via `useMemo`).
- **Working-days auto-calc**: When a user edits Start Date + Duration, End Date is auto-computed via `DateUtils.addWorkingDays`. When End Date is edited, Duration is recomputed. Controlled by the "Skip Weekends" toggle.
- **Zoom via CSS transform**: The SVG is rendered at natural size inside a wrapper div with `transform: scale(zoom)` and `transform-origin: top left`. The wrapper sits inside a scrollable container. This avoids modifying SVG coordinate space.

## Open questions

- SVG vs Canvas for the Gantt renderer: SVG is working well for moderate task counts. Canvas may be needed for performance with 100+ tasks. Decision deferred pending performance spike.
