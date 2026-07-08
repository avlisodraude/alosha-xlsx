---
"@alosha/xlsx": minor
---

feat: round-trip shared and array formulas

The README's feature matrix claimed shared/array formulas round-trip, but they were silently dropped.
Three gaps are now closed:

- `Cell.model` (getter and setter) preserves `sharedFormula`, `ref`, and `shareType` on a
  `CellFormula`, so the writer actually sees them.
- The worksheet writer emits `<f t="shared" …>` / `<f t="array" …>` with the right `si` and `ref`. It
  plans each sheet up front the way ExcelJS does: one `si` per shared master, each `sharedFormula`
  slave matched to its master's `si`, and a plain-formula master promoted to a shared group (with a
  computed bounding `ref`) the first time a slave references it.
- The reader parses `<f t= si= ref=>`: it records shared masters by `si` and materializes slaves as
  `{ sharedFormula: masterAddress, result }` (instead of the previous `{ formula: "" }`, which made us
  write an empty `<f></f>` that Excel flagged for repair).

A write→read→write cycle is now stable and Excel/ExcelJS opens the file without repair. The streaming
writer, which flushes one row at a time and so has no whole-sheet view, continues to render each
formula stand-alone.
