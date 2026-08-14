# TXT Sheets

<p align="center">
  <img src="src/public/txt-sheets-logo.svg" width="112" height="112" alt="TXT Sheets logo">
</p>

<p align="center">
  A beautiful, minimal, local-first spreadsheet application for Windows.
</p>

<p align="center">
  <a href="https://github.com/kichong/TXT-sheets/actions/workflows/ci.yml"><img alt="CI" src="https://github.com/kichong/TXT-sheets/actions/workflows/ci.yml/badge.svg"></a>
  <a href="LICENSE"><img alt="Apache 2.0 license" src="https://img.shields.io/badge/license-Apache--2.0-23805d"></a>
  <a href="https://github.com/kichong/TXT-sheets/releases/latest"><img alt="Latest release" src="https://img.shields.io/github/v/release/kichong/TXT-sheets?color=23805d"></a>
</p>

TXT Sheets follows the focused desktop design and release architecture of [TXT Docs](https://github.com/kichong/TXT-Docs), while keeping the grid—not application chrome—as the primary workspace. Files remain on your computer unless you choose to move or share them.

## Essential features

- Create, open, edit, and save `.xlsx`, `.csv`, and `.tsv` files
- Multiple worksheets with add, rename, delete, and keyboard-friendly navigation
- Direct cell editing, a formula bar, copy/cut/paste, undo/redo, and find
- Drag-resizable rows and columns with dimensions preserved in Excel workbooks
- Row and column insertion/deletion
- Bold, italic, underline, text/fill color, number formats, and alignment
- Excel fidelity for values, formulas, merged cells, sizing, common styling, number formats, and frozen panes
- Autosave recovery, recent files, Windows **Open with**, light/dark themes, and atomic local saves
- Visible compatibility notes for advanced Excel features outside the essential editing model
- User-confirmed updates through GitHub Releases, with unsaved-work protection before restart

## Formulas

TXT Sheets starts with Excel-compatible arithmetic (`+`, `-`, `*`, `/`, `^`), parentheses, cell/range references, cross-sheet references, and these functions:

- `SUM`
- `AVERAGE`
- `MIN`
- `MAX`
- `COUNT`

Formula tokenization, parsing, reference resolution, evaluation, and named-function registration are separate layers. New named formulas can be added to [`src/renderer/formula-functions.ts`](src/renderer/formula-functions.ts) without changing the grid or workbook file layer.

## Development

TXT Sheets uses Electron, React, TypeScript, ExcelJS, and pnpm.

```powershell
pnpm install
pnpm icons
pnpm start
```

Run validation and create a Windows installer:

```powershell
pnpm check
pnpm build
pnpm make
```

The installer is written to `release/`.

## Releases and updates

A tagged release is validated and built on a clean Windows GitHub Actions runner. The workflow publishes the NSIS installer, block map, and `latest.yml` update metadata consumed by `electron-updater`.

1. Update the version in `package.json` and commit it.
2. Tag the commit with the matching version, such as `v0.2.0`.
3. Push the tag. The **Release TXT Sheets** workflow validates, builds, and publishes the release.

Installed builds check shortly after launch and every six hours. A user chooses **Download**, then **Restart to update**; an unsaved workbook blocks restart. The current Windows installer is not code-signed, so Windows SmartScreen may warn on the initial download.

## Contributing and security

Contributions are welcome. Read [CONTRIBUTING.md](CONTRIBUTING.md) before opening a pull request. Please report security concerns privately using the process in [SECURITY.md](SECURITY.md).

## License

TXT Sheets is available under the [Apache License 2.0](LICENSE).
