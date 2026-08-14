# Contributing to TXT Sheets

Thank you for helping improve TXT Sheets. Keep contributions focused on a clear spreadsheet task and preserve the app's minimal, local-first product direction.

## Development setup

1. Install Node.js 22 and pnpm 11.
2. Run `pnpm install`.
3. Run `pnpm start` for local development.

Before opening a pull request, run:

```powershell
pnpm check
pnpm build
```

Changes to workbook import/export should include a round-trip test. Formula additions should be registered in `src/renderer/formula-functions.ts` and include focused tests. UI changes should remain keyboard accessible and work in both light and dark themes.

## Pull requests

- Keep each pull request scoped to one feature or fix.
- Explain the user-facing behavior and any Excel compatibility impact.
- Include screenshots for visible interface changes.
- Do not include private workbooks, generated installers, or secrets.

By contributing, you agree that your contribution is licensed under Apache License 2.0.
