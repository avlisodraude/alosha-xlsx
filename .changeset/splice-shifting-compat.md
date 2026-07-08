---
"@alosha/xlsx": minor
---

fix: re-anchor comments/validations/conditional formatting on splice, and bridge spliceColumns' 0-based inserts in compat

`spliceRows`/`spliceColumns` re-anchored cell values and merges but left the ref-keyed annotations —
cell comments (`ref`), data validations (`sqref`), and conditional-formatting `ref`s — at their old
addresses, so a note on `A5` stayed on `A5` after inserting a row pushed its cell to `A6`. Both splice
methods now shift those annotations with the same keep/shift/drop rules the merge re-anchoring already
used (a ref entirely past the removed region shifts, one entirely before is kept, one straddling it is
dropped — space-separated `sqref`/`ref` parts are handled independently). Images and `autoFilter` are
deliberately left in place, matching ExcelJS, which moves neither on a splice (documented in
docs/MIGRATION.md D7).

Separately, `spliceColumns` insert arrays were ambiguous across the core/compat boundary: the core
reads them **1-based** (`inserts[k][n]` fills row `n`, consistent with `Column.values`), while real
ExcelJS reads them **0-based** (`inserts[k][i]` fills row `i + 1`) — so ported ExcelJS code silently
landed inserted-column data one row too high. The core convention is now documented, and the compat
facade bridges `spliceColumns` to accept ExcelJS's 0-based arrays, so migrated code keeps working while
the plain API stays internally consistent.
