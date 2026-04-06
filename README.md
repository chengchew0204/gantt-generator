# GanttGen

A zero-installation, single-file Gantt chart and project management tool that runs entirely in the browser. No backend, no login, no setup -- just open the HTML file and start planning.

GanttGen is built for project managers and engineers who need a capable scheduling tool that works offline, stores data in portable Excel files, and can be distributed as a single `index.html`.

## Features

**Interactive Gantt Chart**
- Drag task bars to move start/end dates; drag edges to resize duration
- Click the timeline grid to fill date fields directly
- Day / Week scale toggle with zoom controls (50%--300%)
- Today line, dependency arrows, slack bars, and critical path highlighting

**Critical Path Method (CPM)**
- Automatic forward/backward pass calculation via topological sort
- Early Start, Early Finish, Late Start, Late Finish, and Total Float for every task
- Zero-float tasks highlighted as the critical path

**Excel-Based Persistence**
- Import and export `.xlsx` files -- Excel is the database
- A hidden "Settings" sheet preserves theme, view options, and column visibility across save/load cycles
- Download a blank template to get started quickly

**Work Breakdown Structure (WBS)**
- Hierarchical parent/child task grouping via Parent ID
- Parent rows auto-aggregate child start/end dates and average progress
- Expandable/collapsible groups in both the table and the chart

**Milestones & Baselines**
- Zero-duration tasks render as diamond milestones
- Baseline Start/End columns for tracking schedule variance
- Baseline bars displayed as a thin strip behind the actual bar

**Theming**
- Three built-in presets: Linear Dark, Notion Light, Classic Tremor
- Custom color builder for background, text, accent, and status colors
- Per-category bar colors
- All theme settings saved to the Excel file

**Table Editing**
- Inline-editable cells for text, numbers, dates, and status dropdowns
- Arrow key, Tab, and Enter navigation between cells
- Drag-to-reorder rows
- Undo/Redo with full history (Ctrl+Z / Ctrl+Shift+Z)
- Working-days auto-calculation (skip weekends toggle)

**Export**
- Save to Excel (Ctrl+S)
- PNG export of the Gantt chart or the full table+chart view

## Quick Start

### Use the pre-built file

Download `dist/index.html` and open it in any modern browser -- including from the `file://` protocol. No server needed.

### Build from source

```bash
git clone https://github.com/<your-username>/GanttGen.git
cd GanttGen
npm install
npm run build
```

The output is a single self-contained file at `dist/index.html`.

### Development

```bash
npm run dev
```

Opens a Vite dev server with hot module replacement at `http://localhost:5173`.

## Usage

1. Click **Download Template** to get a blank `.xlsx` with the correct column headers.
2. Fill in your tasks in Excel (or use one of the sample templates in `excel-template/`).
3. Click **Import Excel** to load your file.
4. Edit tasks inline in the data table or drag bars on the Gantt chart.
5. Click **Save to Excel** (or press Ctrl+S) to export your project back to `.xlsx`.

All settings -- theme, visible columns, view toggles, category colors -- are saved into the Excel file and restored on the next import.

## Excel Columns

| Column | Description |
|---|---|
| Task ID | Unique numeric identifier |
| Task Name | Display name of the task |
| Dependency | Comma-separated predecessor Task IDs (e.g. `2,3`) |
| Task Category | Grouping label; used for category-specific bar colors |
| Start Date | ISO date (`YYYY-MM-DD`) |
| End Date | Auto-calculated from Start Date + Duration when editing |
| Duration | Number of working days (or calendar days if weekends are included) |
| Progress (%) | 0--100; drives bar fill and dashboard stats |
| Status | `Not Started`, `In Progress`, `Completed`, `On Hold` |
| Owner | Person responsible |
| Remarks | Free-text notes |
| Baseline Start | Original planned start date for variance tracking |
| Baseline End | Original planned end date for variance tracking |
| Parent ID | Task ID of the WBS parent (leave blank for top-level tasks) |

## Keyboard Shortcuts

| Shortcut | Action |
|---|---|
| Ctrl+S | Save to Excel |
| Ctrl+Z | Undo |
| Ctrl+Shift+Z / Ctrl+Y | Redo |
| Arrow keys | Navigate between cells while editing |
| Tab / Shift+Tab | Move to next/previous column |
| Enter | Commit cell and move down |
| Escape | Cancel cell edit |

## Tech Stack

| Library | Purpose |
|---|---|
| [React](https://react.dev/) | UI framework |
| [Tailwind CSS](https://tailwindcss.com/) | Utility-first styling |
| [SheetJS (xlsx)](https://sheetjs.com/) | In-browser Excel read/write |
| [lucide-react](https://lucide.dev/) | Icon set |
| [html-to-image](https://github.com/bubkoo/html-to-image) | Client-side PNG export |
| [Vite](https://vite.dev/) | Build tool |
| [vite-plugin-singlefile](https://github.com/nickreid/vite-plugin-singlefile) | Bundle everything into one HTML file |

## Project Structure

```
GanttGen/
  src/
    main.jsx              Entry point
    App.jsx               Root state, layout, date auto-calc, WBS tree
    index.css             CSS variables, Tailwind import, global styles
    components/
      Dashboard.jsx       Toolbar, stats bar, project name editor
      DataTable.jsx       Editable task grid with column toggles
      GanttChart.jsx      SVG Gantt renderer (bars, arrows, milestones)
      ThemePanel.jsx      Slide-over theme configurator
      ViewOptions.jsx     Chart/column toggle dropdown
      GuideOverlay.jsx    Interactive guided tour
    utils/
      CpmEngine.js        Critical Path Method computation
      DateUtils.js        Working-days arithmetic
      ExcelUtils.js       Import/export logic with Settings sheet
    hooks/
      useUndoRedo.js      Undo/redo stack with batch support
  scripts/
    generate-templates.mjs  Generates sample Excel templates
  dist/
    index.html            Production build (single file, ~800 KB)
```

## Constraints

- **Single-file output** -- the production build is one self-contained `index.html` with all JS, CSS, and assets inlined. No CDN calls at runtime.
- **`file://` protocol** -- must work when opened directly from the filesystem. No `localStorage` (unreliable on `file://`); Excel is the only persistence layer.
- **No backend** -- no server, database, API, or cloud sync of any kind.
- **Desktop-first** -- optimized for desktop browsers; mobile is not a primary target.

## License

This project is not yet licensed. All rights reserved.
