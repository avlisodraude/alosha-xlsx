# Migrating from ExcelJS

`@alosha/xlsx/compat` is a drop-in for [`exceljs`](https://github.com/exceljs/exceljs): the same
`new Workbook()` / `addWorksheet` / `getCell` / `addRow` / `mergeCells` /
`workbook.xlsx.writeBuffer`/`writeFile`/`load`/`readFile` surface, backed by `@alosha/xlsx`'s model and
OOXML writer/reader instead of ExcelJS's. For the common case — build a workbook, set values, styles,
number formats, merges, column defs, then serialize/reload it — **the only change is the import.**

```diff
- import ExcelJS from "exceljs";
+ import ExcelJS from "@alosha/xlsx/compat";

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Report");

  sheet.getCell("A1").value = "Item";
  sheet.getCell("B1").value = "Qty";
  sheet.getColumn(2).width = 12;
  sheet.getCell("B1").font = { bold: true };

  await workbook.xlsx.writeFile("report.xlsx");
```

Everything below the diff line is unmodified ExcelJS code.

If you're migrating *away* from the ExcelJS surface entirely (not just swapping the import), see the
main [`README.md`](../README.md#migrating-from-exceljs) instead — that's a deeper rewrite (0-based dense
`worksheets`, an object-shaped `inheritFrom` instead of the `'i'|'o'|'i+'|'o+'` DSL, etc.) with more
freedom to diverge, since it doesn't have to stay drop-in compatible.

## Two behavioural notes — read these before porting

**D1 — `cell.type` / `cell.effectiveType` report ExcelJS's numbers, not `@alosha/xlsx`'s own.**
`@alosha/xlsx`'s core `CellType` enum shares names with ExcelJS's `ValueType` but **not** the same
numbers (e.g. `Number` is `1` in the core model, `2` in ExcelJS). In compat mode this is fully emulated:
every cell reached through a `@alosha/xlsx/compat` workbook reports the ExcelJS number, so `if (cell.type
=== ExcelJS.ValueType.Formula)`-style checks port unchanged. This *only* applies to workbooks built via
`@alosha/xlsx/compat`'s `Workbook` — a plain `@alosha/xlsx` `Workbook` still reports its own numbers.

**D2 — style setters replace the whole category, not just the field you touched.**
`@alosha/xlsx`'s core model does a **deep field-level merge** when resolving a cell's style: setting
`cell.font = { bold: true }` on a cell that inherited `{ name: "Arial", size: 12 }` from its row leaves
you with `{ name: "Arial", size: 12, bold: true }`. ExcelJS instead **replaces the category outright** —
the same assignment discards the inherited `name`/`size`, leaving just `{ bold: true }`. A file migrated
from ExcelJS relies on the latter to render identically, so compat-mode style setters
(`font`/`fill`/`border`/`alignment`/`protection`/`numFmt`) always apply ExcelJS's whole-object precedence,
never the core's field merge. Again, this is scoped to compat-constructed workbooks; the plain
`@alosha/xlsx` API is unaffected.

## Supported / deferred matrix

| Feature | Status | Notes |
| --- | :---: | --- |
| `Workbook`/`Worksheet`/`Row`/`Column`/`Cell` core API, `getCell`/`getRow`/`addRow`/`addRows`/`mergeCells`/`eachRow`/`eachCell` | ✅ | |
| `addRow(array \| object)`, `row.values` (1-based, empty `[0]`), `worksheet.columns = [{header,key,width}]` | ✅ | |
| Values — string/number/boolean/date/null, formula (`{formula,result}`), rich text, error | ✅ | |
| `cell.type` / `cell.effectiveType` | ✅ | ExcelJS `ValueType` numbers — see D1 above |
| Style setters (`font`/`fill`/`border`/`alignment`/`protection`/`numFmt`) | ✅ | ExcelJS whole-object precedence — see D2 above |
| `Enums.ValueType`, `Enums.FormulaType` | ✅ | |
| `workbook.xlsx.writeBuffer`/`writeFile`/`load`/`readFile`/`read` | ✅ | |
| Hyperlink cells (`{ text, hyperlink }`, external URL/mailto) | ✅ | persists as a real OOXML relationship; round-trips to `{text, hyperlink, tooltip?}` with `cell.type` `Hyperlink`. Internal `location=` (in-workbook) links are still ⚠️ deferred. |
| `worksheet.dataValidations` (`add`/`find`/`remove`/`model`), `cell.dataValidation` | ✅ | list dropdowns (inline or same-sheet range), whole/decimal/date/time/textLength constraints, custom formulas. Note: OOXML's `showDropDown` attribute is inverted — it *hides* the dropdown when present — which the model corrects (`showDropDown: false` means hidden); see DATA-VALIDATION-SPEC §4. A list referencing a range on **another sheet** (the x14/`<extLst>` form) is still ⚠️ deferred. |
| `stream.xlsx.WorkbookWriter` / `WorkbookReader` | ❌ deferred | throws `UnsupportedFeatureError` |
| `workbook.csv.*` | ❌ deferred | throws `UnsupportedFeatureError` |
| `addImage`/`getImage`/`workbook.media`/`worksheet.addImage` | ❌ deferred | throws `UnsupportedFeatureError` |
| `worksheet.addTable`/`getTable(s)` (tables, pivot tables) | ❌ deferred | throws `UnsupportedFeatureError` |
| `worksheet.addConditionalFormatting`/`removeConditionalFormatting` | ❌ deferred | throws `UnsupportedFeatureError` |
| `cell.note` (comments) | ❌ deferred | throws `UnsupportedFeatureError` |
| `workbook.definedNames.*` | ❌ deferred | throws `UnsupportedFeatureError` |
| `worksheet.protect`/`unprotect` | ❌ deferred | throws `UnsupportedFeatureError` |

Every deferred method fails loudly and precisely at the exact call — `UnsupportedFeatureError` names the
feature — rather than silently no-op'ing or producing a subtly-wrong file. If your workbook only touches
supported features, migration genuinely is a one-line import swap; if it hits anything in the deferred
list, that call throws immediately so you find out at the call site, not by diffing a corrupt `.xlsx`.

## Other differences worth knowing about

These aren't compat-specific bugs — they're pre-existing `@alosha/xlsx` core-model behaviours that a
migrated file can still trip over, surfaced while building the parity corpus (below):

- **An unstyled cell's `fill` is `undefined`, not `{ pattern: "none" }`.** ExcelJS always resolves a
  cell's fill from the styles pool, so an unstyled cell still reports `{ type: "pattern", pattern:
  "none" }`. `@alosha/xlsx`'s reader treats the baseline "no fill" style id as "nothing to report" and
  leaves `cell.fill` as `undefined`. Code that does `cell.fill.pattern` unconditionally (rather than
  `cell.fill?.pattern`) will need a null check either way — but the difference is worth knowing before
  you rely on `cell.fill` always being an object.

## Verifying the claim: the parity corpus

`tests/compat/parity.test.ts` is the mechanical proof behind this document. It takes a corpus of ~10
representative ExcelJS snippets (every value type, a styled header row, number formats, a merge, multiple
worksheets, column keys + `addRow(object)`, a formula with a cached result, a `date1904` workbook, and a
row/cell style-precedence case), runs each snippet **unmodified** against both real `exceljs` and
`@alosha/xlsx/compat`, and diffs the two outputs after reloading both with real `exceljs` — plus reloads
the compat-produced buffer with `@alosha/xlsx`'s own reader, to confirm the shim's output is valid input
for our reader too. It's guarded on the optional `exceljs` dev dependency and skips cleanly if that
package isn't installed.
