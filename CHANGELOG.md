# Changelog

## 0.6.0

### Minor Changes

- Freeze/split panes and sheet auto-filter (write + read + compat).

  Worksheets can now carry a single `<sheetView>` with a `<pane>` — `views = [{ state: "frozen" |
"split", xSplit?, ySplit?, topLeftCell?, activePane?, showGridLines? }]` — and a sheet-level
  `autoFilter` range (`"A1:D1"`).

  - **Model.** `WorksheetView` and `worksheet.autoFilter` round-trip through `WorksheetModel` like
    merges and column widths.
  - **Write.** `<pane>` is emitted inside `<sheetView>`, and `<autoFilter ref="…">` after
    `<sheetData>`/before `<mergeCells>`, matching CT_Worksheet element order.
  - **Read.** Both invert back to an equal model; a malformed `<pane>` (e.g. a non-numeric `xSplit`)
    raises a typed `AloshaXlsxError` instead of `NaN`.
  - **Compat.** `@alosha/xlsx/compat`'s `worksheet.views` and `worksheet.autoFilter` are a real
    pass-through — the core field names already match ExcelJS's. The one bridge: setting
    `worksheet.autoFilter` to ExcelJS's `{ from, to }` object form (each an address string or
    `{ row, column }`) normalizes to our `"A1:D1"` range string before it reaches the core.

  **Deferred:** multiple views per sheet, `<selection>`, per-pane active-cell tracking, split-window
  zoom, and auto-filter criteria/sort state (`<filterColumn>`, `<sortState>`) — only the `ref` range
  persists, matching where ExcelJS's basic `autoFilter` sits.

## 0.5.0

### Minor Changes

- Cell data validations — dropdowns and value constraints (write + read + compat).

  Cells and ranges can now carry a data validation: a **list dropdown** (an inline
  `'"Yes,No,Maybe"'` or a same-sheet range `"$E$1:$E$5"`), a numeric/date/length constraint
  (`whole` / `decimal` / `date` / `time` / `textLength`) with an operator
  (`between` / `notBetween` / `equal` / `notEqual` / `greaterThan` / `lessThan` /
  `greaterThanOrEqual` / `lessThanOrEqual`) and one or two formulae, or a `custom` boolean formula —
  plus the usual flags and messages (`allowBlank`, `showInputMessage`, `showErrorMessage`,
  `errorStyle`, `errorTitle`, `error`, `promptTitle`, `prompt`, `showDropDown`). Applied per cell via
  `cell.dataValidation` or per range via `worksheet.dataValidations.add(sqref, dv)`, where `sqref` may
  be a single address, a range, or space-separated ranges.

  - **Model.** New `DataValidation` type and a per-worksheet `DataValidations` manager
    (`add` / `find` / `remove` / `model`, keyed by `sqref`); `cell.dataValidation` get/set resolves
    through it, and validations round-trip through `WorksheetModel` like merges.
  - **Write.** A `<dataValidations>` block is emitted in the sheet, in CT_Worksheet order (after
    `<mergeCells>`, before `<hyperlinks>`), so the validation re-opens in Excel/LibreOffice as a
    working dropdown/constraint. No extra parts or relationships.
  - **Read.** The `<dataValidations>` block parses back to an equal model; malformed entries are
    skipped rather than throwing.
  - **Compat.** `@alosha/xlsx/compat`'s `cell.dataValidation` and `worksheet.dataValidations` are now
    a real pass-through (previously threw `UnsupportedFeatureError`) — the core shape already matches
    ExcelJS's `DataValidations`.

  **OOXML `showDropDown` inversion:** the OOXML attribute is inverted — `showDropDown="1"` _hides_ the
  dropdown. The model uses the intuitive sense, so `showDropDown: false` means "hide it" and maps to
  the attribute; unset/`true` shows the dropdown (the default) and emits nothing.

  **Deferred:** a list that references a range on **another sheet** (the x14 / `<extLst>` form Excel
  uses for cross-sheet list refs) is not yet supported — only same-sheet ranges and inline lists.

## 0.4.0

### Minor Changes

- Persist external cell hyperlinks as real OOXML relationships (write + read + compat).

  Setting a `{ text, hyperlink, tooltip? }` cell value now writes a proper worksheet
  `<hyperlinks>` record plus an external relationship in
  `xl/worksheets/_rels/sheetN.xml.rels`, so the link is clickable in Excel/LibreOffice and
  survives a round-trip — instead of degrading to its plain display text. External
  `http(s)` and `mailto:` targets are supported; the display text still renders as an
  ordinary string cell.

  - **Write.** Hyperlink cells emit a `<hyperlinks>` block (after `<mergeCells>`) with
    worksheet-local `r:id`s and a matching `.rels` part.
  - **Read.** The `<hyperlinks>` block is resolved against the worksheet's own rels, so
    hyperlink cells parse back to `{ text, hyperlink, tooltip? }` with `CellType.Hyperlink`.
  - **Compat.** No API change — `@alosha/xlsx/compat` hyperlink cells now round-trip, and
    `cell.type` is `ValueType.Hyperlink` (5).

  **Behaviour change (why this is a minor):** reading a workbook that contains hyperlinks now
  yields a `{ text, hyperlink }` cell value with `CellType.Hyperlink` (compat:
  `ValueType.Hyperlink`), where 0.3.x returned a plain **string**. Code that read hyperlink
  cells expecting a string will now receive an object.

  Internal in-workbook (`location="Sheet2!A1"`) hyperlinks remain deferred; the previous
  "hyperlink cells round-trip as their display text" known limitation is removed from the
  README and `docs/MIGRATION.md` in favour of that narrower note.

## 0.3.1

### Patch Changes

- 36706d6: compat: `eachRow`/`eachCell` now skip empty rows/cells by default, matching ExcelJS.

  Without `includeEmpty`, the `@alosha/xlsx/compat` facade's `worksheet.eachRow` and
  `row`/`column.eachCell` now skip rows/cells that are empty or were explicitly set to `null` (as
  ExcelJS does), rather than yielding every row/cell ever touched via `getRow`/`getCell`. Pass
  `{ includeEmpty: true }` for full enumeration. Callback rows/cells remain proxied (so `cell.type`
  still reports ExcelJS `ValueType` numbers).

  Docs: the README `BENCHMARKS.md`/`CHANGELOG.md` links are now absolute so they resolve on npmjs.com,
  and the corresponding note is removed from `docs/MIGRATION.md`.

All notable changes to `@alosha/xlsx` are documented here. This file is maintained by
[Changesets](https://github.com/changesets/changesets) from the release `0.2.0` onward; entries
are grouped by version, newest first.

## 0.3.0

### Minor Changes

- Add the `@alosha/xlsx/compat` ExcelJS drop-in subpath.

  Import `ExcelJS from "@alosha/xlsx/compat"` to run existing ExcelJS code against the
  `@alosha/xlsx` core with read + write parity. The facade mirrors ExcelJS's shape:

  - **Numeric `cell.type` emulation (D1).** `cell.type` returns ExcelJS's `ValueType`
    numbers (and the `ValueType` enum is re-exported), rather than the core's string tags.
  - **Whole-object style precedence (D2).** Style setters (`cell.font`, `cell.fill`,
    `cell.border`, `cell.alignment`, `cell.numFmt`, and `cell.style`) follow ExcelJS's
    whole-object assignment and precedence semantics.
  - Familiar `addRow` / `row.values` / `worksheet.columns` behaviour.

  Two documented gaps remain for this release (both in
  [`docs/MIGRATION.md`](https://github.com/avlisodraude/alosha-xlsx/blob/main/docs/MIGRATION.md)):
  an unstyled-fill cosmetic difference, and `eachRow` / `eachCell` not filtering nulls the
  way ExcelJS does (a workaround is documented; a faithfulness fix is deferred to 0.3.1).

## 0.2.0

### Minor Changes

- First public release (read + write).

  - Read and write real `.xlsx` (OOXML) workbooks through an ESM-first document model.
  - ExcelJS-style accessor: `workbook.xlsx.writeBuffer` / `writeFile` / `load` / `readFile`.
  - Environment-agnostic core: `writeWorkbookBuffer` / `readWorkbookBuffer` (sync) plus the Node
    `writeWorkbookFile` / `readWorkbookFile` variants.
  - Cell values (string, number, boolean, date, error, rich text), formulas with cached results,
    shared strings, styles (fonts, fills, borders, alignment, number formats, protection), merged
    cells, column/row metadata, multiple worksheets, and workbook metadata.
  - See the [known limitations](./README.md#known-limitations) for what is deferred (hyperlink
    relationships, streaming, drawings/images, data validation).
