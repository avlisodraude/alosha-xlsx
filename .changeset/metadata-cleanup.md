---
"@alosha/xlsx": patch
---

chore: metadata & packaging cleanup

A batch of small correctness/packaging fixes with no runtime behaviour change beyond the two API
tweaks noted below:

- **`version` now reports the real package version.** `export const version` was hardcoded to
  `"0.0.0"`; it's now injected from `package.json` at build time via tsup's `define` (mirrored in the
  vitest config so it resolves under test too).
- **`getWorksheet(name)` falls back to a case-insensitive match.** Exact-case matches still win, but a
  miss now retries case-insensitively, lining up with `addWorksheet`'s case-insensitive duplicate rule
  (so `getWorksheet("report")` finds a `"Report"` sheet).
- **`cell.note = undefined` and `cell.dataValidation = undefined` now clear** the comment / validation
  (routing to the store's `remove`), instead of only accepting a value.
- **Smaller published package.** Sourcemaps are no longer shipped (emitted only for `pnpm dev` watch
  builds), dropping the tarball from ~585KB to ~195KB (unpacked ~2.55MB → ~800KB).
- **`sideEffects` now covers the hashed ESM chunk** (`./dist/*.js`, `./dist/*.cjs`) that actually holds
  the `workbook.xlsx` accessor's `Object.defineProperty` side effect, so aggressive tree-shakers can't
  drop it.
- **Docs:** README now states Node.js `>= 22` (matching `engines`) and documents the case-insensitive
  `getWorksheet` fallback.
