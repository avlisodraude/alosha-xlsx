# Changelog

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
