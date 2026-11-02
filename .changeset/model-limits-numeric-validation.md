---
"@alosha/xlsx": minor
---

Bound imported worksheet-model expansion and validate grid coordinates and numeric cell values consistently.

- Add workbook-wide read budgets for cells (1,000,000), allocated rows (1,048,576), merged-range areas (100,000 cells), and expanded column definitions (100,000). Overrides and Infinity are supported. Merge and column-range work is checked before expansion; cells created by comments and hyperlinks also count.
- Keep row numbering intact after a streamed addRow() rejects an invalid payload.
- Enforce Excel's row and column limits across numeric/string coordinates, ranges, constructors, imported references, and row/column edits. Preflight edit bounds and numeric inserts before moving existing data.
- Reject NaN, infinities, invalid Dates, and invalid cached formula results with InvalidCellValueError. Revalidate mutable values at serialization time and reject invalid date conversions.
- Preserve the workbook.xlsx type in packed declarations so downstream ESM/CommonJS consumers can use the documented accessor and new read options.

Import budgets reset for each read and are removed before returning the workbook, so later user edits are unaffected. These limits count model objects and expansion work, not exact heap bytes.
