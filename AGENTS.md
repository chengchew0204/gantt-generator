# AGENTS.md

## Cursor Cloud specific instructions

GanttGen is a zero-installation, single-file Gantt chart tool that runs entirely client-side in the browser. There is no backend, database, or API.

### Services

| Service | Command | Notes |
|---|---|---|
| Vite dev server | `npm run dev` | Serves at `http://localhost:5173` with HMR |

### Key commands

- **Lint**: `npm run lint` -- ESLint 9. The codebase has pre-existing lint errors (unused variables, React hooks warnings) that are not regressions.
- **Build**: `npm run build` -- produces a single self-contained `dist/index.html` (~800 KB) via `vite-plugin-singlefile`.
- **Dev**: `npm run dev` -- starts Vite dev server with HMR. Use `--host 0.0.0.0` if you need external access.
- **Preview**: `npm run preview` -- serves the production build locally.

### Caveats

- There are no automated tests configured in the project (no test script in `package.json`, no test framework).
- The `xlsx` dependency is fetched from a CDN tarball (`https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz`), not npm registry. This may cause `npm install` to be slower or fail if that CDN is unreachable.
- The project uses Tailwind CSS v4 with `@tailwindcss/vite` plugin (not the older PostCSS-based setup).
- Excel files in `excel-template/` can be used as sample data for manual testing of import/export features.
