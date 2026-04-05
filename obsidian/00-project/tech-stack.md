# Tech Stack

> Update this file whenever a new tool, library, language, or framework is added or removed.

## Languages

| Language | Usage | Rationale |
|---|---|---|
| JavaScript (JSX) | All application logic and UI | React ecosystem; no TypeScript required for this scope |
| CSS (via Tailwind) | Styling | Utility-first, compatible with single-file output |

## Frameworks & Libraries

| Package | Purpose | Rationale |
|---|---|---|
| React | UI framework and state management | Component model suits split-pane interactive layout |
| Tailwind CSS | Utility-first styling | Rapid UI iteration; purged CSS keeps the single-file small |
| `vite-plugin-singlefile` | Bundle everything into one `index.html` | Core requirement — zero-installation delivery |
| `xlsx` (SheetJS) | Read and write `.xlsx` files in-browser | In-memory Excel I/O without any server; supports custom sheets |
| `lucide-react` | Icon set | Lightweight, tree-shakeable, consistent with modern dashboard style |
| `html-to-image` | Export chart area as PNG | Client-side image capture with no server dependency |

## Tooling

| Tool | Purpose |
|---|---|
| Vite | Build tool and dev server |
| `vite-plugin-singlefile` | Inline all assets into a single HTML file at build time |
| Node.js / npm | Dependency management and build scripts |

## UI style reference

Target aesthetic: **Linear / Notion / Tremor.so** — clean, dark-mode-capable, professional dashboard.
- CSS Custom Properties (variables) drive all theme colors.
- Three built-in theme presets: "Linear Dark", "Notion Light", "Classic Tremor".
- Custom color builder via native `<input type="color">` elements in a settings modal.

## Rejected alternatives

| Alternative | Reason not chosen |
|---|---|
| TypeScript | Adds build complexity; not required for this project's scope |
| Chart.js / D3 | Gantt requires custom SVG/Canvas rendering; these libraries don't offer first-class Gantt support |
| `jspdf` for export | PNG via `html-to-image` is simpler and sufficient; PDF not required |
| `localStorage` for persistence | Does not work on `file://` protocol reliably; Excel file is the persistence layer |
| Backend / Electron | Core requirement is zero-installation browser-only delivery |
