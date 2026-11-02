---
"@alosha/xlsx": minor
---

Protect address caches and improve OOXML namespace interoperability.

- Return frozen cached address results with readonly coordinate fields; preserve cached object identity and bounded eviction.
- Decode doubled apostrophes in quoted sheet references, including ranges and absolute addresses.
- Resolve namespaces by scoped URI bindings across package wiring, worksheets, styles, shared strings, comments, metadata, and drawings. Normalize relationship attributes and skip foreign namespace subtrees.
- Retain the namespace-free fragment behavior of the low-level part-reader helpers.
- Add deterministic address property tests, an independently handwritten package, and an ExcelJS-produced workbook with systematically renamed namespace prefixes, including styles, links, notes, and an image.

Use a spread copy of decodeAddress() results if mutable coordinates are needed.
