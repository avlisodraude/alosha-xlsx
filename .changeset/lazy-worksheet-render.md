---
"@alosha/xlsx": patch
---

Buffered writer: render worksheet rows lazily instead of snapshotting the whole sheet.

`renderWorksheet` previously read `ws.model`, whose getter deep-snapshots every cell into a
`WorksheetModel` before serializing — a full redundant copy of the sheet held alongside the live
model and the growing output. It now reads `ws.sheetMeta` (sheet-level model without rows) and pulls
rows one at a time through a lazy `ws.rowModelsInOrder()` generator, so only the current row's cells
are materialized during serialization. Output is byte-identical; `ws.model` / `workbook.model` are
unchanged. Peak memory on large buffered exports drops (~6% at 500k rows × 5 cols) and the write is
modestly faster.
