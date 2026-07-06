# Changelog

All notable changes to `@alosha/xlsx` are documented here. This file is maintained by
[Changesets](https://github.com/changesets/changesets) from the release `0.2.0` onward; entries
are grouped by version, newest first.

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
