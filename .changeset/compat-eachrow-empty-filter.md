---
"@alosha/xlsx": patch
---

compat: `eachRow`/`eachCell` now skip empty rows/cells by default, matching ExcelJS.

Without `includeEmpty`, the `@alosha/xlsx/compat` facade's `worksheet.eachRow` and
`row`/`column.eachCell` now skip rows/cells that are empty or were explicitly set to `null` (as
ExcelJS does), rather than yielding every row/cell ever touched via `getRow`/`getCell`. Pass
`{ includeEmpty: true }` for full enumeration. Callback rows/cells remain proxied (so `cell.type`
still reports ExcelJS `ValueType` numbers).

Docs: the README `BENCHMARKS.md`/`CHANGELOG.md` links are now absolute so they resolve on npmjs.com,
and the corresponding note is removed from `docs/MIGRATION.md`.
