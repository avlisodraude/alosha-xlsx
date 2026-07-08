---
"@alosha/xlsx": patch
---

fix: preserve per-cell styles inside merged ranges on read

Reading a workbook destroyed the styles of non-master (slave) cells inside a merged range. Repro:
merge `A1:B2`, set a thick bottom border on `B2`, `writeWorkbookBuffer` → `readWorkbookBuffer` →
`B2.border` came back `undefined`. Cause: the reader (and `Worksheet` `set model`) applied merges via
the public `mergeCells()`, whose `Cell.merge(master)` defaults to copying the master's style over each
slave — correct at authoring time, but it clobbered the slave's just-parsed style when re-hydrating a
saved workbook. The read / `set model` / merge re-anchoring paths now link slaves with `ignoreStyle`,
so each slave keeps its own style; authoring-time `mergeCells()` still propagates the master's style
(unchanged ExcelJS-compatible behavior).
