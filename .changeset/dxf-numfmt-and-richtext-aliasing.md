---
"@alosha/xlsx": patch
---

Two correctness fixes:

- **Reader no longer aliases rich text across cells.** A `t="s"` cell resolved its value to the shared
  RichText object by reference, so multiple cells pointing at the same rich shared string shared one
  mutable object — editing one cell's runs silently mutated the others (and the shared-strings pool).
  Rich-text values are now cloned per cell on read.

- **Distinct custom dxf number-formats get distinct ids.** Every custom number-format code in a
  conditional-formatting differential style (`<dxf>`) was written with `numFmtId="164"`, so two rules
  with different custom formats collided on one id. Each distinct custom code now gets its own id
  (164, 165, …); built-in codes keep their reserved ids. The dxf `formatCode` (which our reader inverts
  by) is unchanged, so this is transparent to round-trips but no longer confuses stricter consumers.
