# Project Brief — GanttGen

> This is the authoritative description of the project. Update it whenever the scope, goals, or target users change.

## What is GanttGen?

GanttGen is a zero-installation, single-file React web application for professional project management. It runs entirely in the browser with no backend, no server, and no setup required. It produces a fully interactive Gantt chart with Critical Path Method (CPM) analysis, Excel import/export, and a highly customizable UI.

## Problem statement

Project managers and engineers frequently need a capable Gantt chart tool that:
- Works offline and without installation
- Stores project data in a portable, shareable format (Excel)
- Supports CPM/critical path analysis without expensive desktop software
- Can be distributed as a single HTML file and opened anywhere
- I need to run it on cooperative computers with strict IT compliances that prevent me from accessing internet, so it needs to be completely offline, and executable without installing additional softwares.
Existing tools are either cloud-dependent, expensive, or lack CPM/baseline/WBS capabilities.

## Goals

- Deliver a single `index.html` file (bundled via `vite-plugin-singlefile`) that opens in any modern browser with no installation.
- Support full Excel-based data persistence: import project data from `.xlsx`, save edits back to `.xlsx`, including UI theme and column visibility settings in a hidden "Settings" sheet.
- Implement a complete CPM engine: calculate Total Float (Slack), highlight the Critical Path (zero-slack tasks), and display slack time visually on the chart.
- Render professional Gantt bars with status colors, baseline tracking, milestone diamonds, WBS summary bars, dependency arrows, and a "Today" line.
- Support multiple dependencies (comma-separated predecessor IDs) and working-days calculation (skip weekends, with a toggle).
- Provide a dynamic theming system with CSS variables, 3 premium presets, and a custom color builder — all persisted in the Excel file.
- Allow toggling chart elements (critical path, slack, dependency arrows, today line) and data table columns, with those preferences saved to Excel.

## Non-goals

- No backend server, database, or cloud sync of any kind.
- No user authentication or multi-user collaboration.
- No mobile-first design (desktop browser is the primary target).
- No support for resource leveling or cost tracking (out of scope for v1).
- No Gantt chart printing / PDF export beyond PNG via `html-to-image`.

## Target users

- Project managers and engineers who manage small-to-medium projects.
- Teams that already use Excel as their project data format.
- Users who need a portable tool that can be sent as a single file and opened without setup.

## Success criteria

- The app opens as a single `index.html` on `file://` protocol with no errors.
- A blank Excel template can be downloaded, filled in, and re-imported to fully populate the chart.
- CPM (critical path) is correctly calculated and highlighted for a multi-task project with dependencies.
- Theme and column visibility preferences survive a save → re-import cycle.
- Milestones, WBS parent rows, and baseline bars all render correctly on the chart.

## Constraints

- Must output a single self-contained `index.html` — no external CDN calls at runtime.
- Must run on `file://` protocol (no `localStorage` assumption; persistence is Excel-only).
- Libraries: React, Tailwind CSS, `xlsx`, `lucide-react`, `html-to-image`, `vite-plugin-singlefile`.
- UI style target: Linear / Notion / Tremor.so — modern, clean, professional dashboard.
