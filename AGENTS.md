# AGENTS.md

## Cursor Cloud specific instructions

GanttGen is a single-package, client-side-only React SPA with zero backend dependencies. There is no database, no API server, and no environment variables or secrets required.

### Quick reference

| Action       | Command          |
|------------- |------------------|
| Install deps | `npm install`    |
| Dev server   | `npm run dev`    |
| Lint         | `npm run lint`   |
| Build        | `npm run build`  |
| Preview      | `npm run preview`|

### Notes

- The dev server (Vite) runs at `http://localhost:5173` by default.
- The production build outputs a single self-contained `dist/index.html` (~1.5 MB) via `vite-plugin-singlefile`. All JS/CSS/assets are inlined.
- The `xlsx` (SheetJS) dependency is fetched from `https://cdn.sheetjs.com`, not the npm registry. If network issues occur during `npm install`, this is the most likely cause.
- ESLint has pre-existing warnings/errors in the codebase (mostly `no-unused-vars` and React hooks warnings). These are not blocking and are part of the current codebase state.
- There are no automated tests (no test framework configured). Validation is done through lint and build.
- Node.js 22 LTS is required. The environment comes with Node.js pre-installed via nodesource.
