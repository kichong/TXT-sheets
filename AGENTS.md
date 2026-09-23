# Repository workflow

- When asked to fix an open GitHub issue, complete the delivery flow unless the user says otherwise: implement and test the fix, update the patch version, commit with a GitHub closing keyword, push the default branch, publish the normal release artifacts, and verify that the issue closed.
- Keep compatibility reports privacy-safe. Never include filenames, paths, workbook or worksheet names, cell values, formulas, comments, authors, workbook metadata, or other user content in issue URLs or reports.

## Start here

- This repository is **TXT Sheets**. The sibling `TXT-docs` checkout is a separate app; check the working directory before editing or releasing.
- Read `README.md` for supported formats and user behavior, then inspect the owner files below. Use `rg` to locate a feature before opening the large renderer or grid files.
- Keep this component map current: whenever a file or folder is added, moved, renamed, or given a new responsibility, update the relevant route here in the same change. Remove stale routes.

## Component map and edit routes

| Task | Start here | Also inspect |
| --- | --- | --- |
| Workbook UI, toolbar, sheet tabs, save/close | `src/renderer/App.tsx`, `src/renderer/styles.css` | `src/renderer/SheetTabMenu.tsx`, `src/renderer/sheet-actions.ts`, `tests/sheet-actions.test.ts`, `scripts/sheet-reorder-smoke.mjs` |
| Grid editing, selection, resize, fill | `src/renderer/SpreadsheetGrid.tsx`, `src/renderer/workbook-model.ts` | `src/renderer/App.tsx`, `tests/dimensions.test.ts` |
| Formulas and named functions | `src/renderer/formulas.ts`, `src/renderer/formula-functions.ts` | `tests/formulas.test.ts`, `src/shared/types.ts` |
| XLSX/CSV/TSV import and export | `src/main/workbook-io.ts` | `src/shared/types.ts`, `tests/workbook-roundtrip.test.ts`, `tests/fixtures/` |
| File lifecycle, recent files, recovery, launch files | `src/main.ts`, `src/main/storage.ts`, `src/main/launch-files.ts` | `src/preload.ts`, `tests/launch-files.test.ts` |
| Electron API / IPC contract | `src/main.ts`, `src/preload.ts`, `src/shared/types.ts` | `src/global.d.ts`, `src/renderer/browser-mock.ts` |
| Updates and releases | `src/main/updater.ts`, `src/shared/updates.ts`, `package.json` | `.github/workflows/release.yml`, `tests/updates.test.ts` |
| Compatibility reporting | `src/shared/compatibility-report.ts` | `src/main.ts`, `tests/compatibility-report.test.ts` |

`src/renderer.tsx` mounts the React app. `vite.main.config.ts`, `vite.preload.config.ts`, and `vite.renderer.config.ts` build its three processes. `tests/` holds focused checks; `scripts/` includes Electron smoke checks. Treat `dist/`, `release/`, and `node_modules/` as generated output.

## Change workflow

1. Trace the full path of the requested behavior through renderer, workbook model, preload, main process, and file I/O as applicable. Preserve workbook order, active sheet, formulas, undo, recovery, and existing XLSX/CSV/TSV behavior.
2. When fixing, updating, or optimizing a feature, inspect its related and adjacent functions for the same defect, duplicated logic, or a safe improvement that would make the app work better. Make small, clearly connected improvements when they can be validated in the same change; report larger ideas separately. Do not expand into unrelated redesigns.
3. Add focused tests at the relevant boundary, especially for workbook round trips and formulas. Verify existing workbooks before changing the persisted schema; keep temporary UI state out of workbook data.
4. Run `pnpm check` and `pnpm build`; for interaction changes, exercise the actual Electron UI and keyboard/focus flow. For releases, validate the packaged app and published installer/update assets as well.
5. Before a release, check `package.json`, `.github/workflows/release.yml`, and the current Git state. Follow the issue delivery rule above when it applies. Do not claim a release is complete from a local build alone.
