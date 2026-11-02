---
"@alosha/xlsx": minor
---

Harden exports and imports and provide conditional browser builds.

- Reject destination errors and premature closes without hanging flushes or background rejections.
- Ship browser builds without filesystem imports; verify packed ESM, CommonJS, downstream types, and browser consumers in CI.
- Preserve inline rich text, literal ST_Xstring escape tokens, worksheet defaults, outline levels, tab colors, and custom descenders.
- Support explicitly declared shared/array formula groups in streamed workbooks; reject ambiguous shared references.
- Validate streaming worksheet names, row order, duplicate commits, output options, and flush intervals.
- Track the highest buffered row incrementally so appending rows no longer scans every existing row.
- Add configurable ZIP allocation limits and typed ReadLimitError failures. Reads default to 128 MiB of input, 10,000 entries, 256 MiB per entry, and 1 GiB total; overrides and Infinity are available for trusted large inputs.

Streaming rows must be committed once, in creation order, before committing the worksheet.
Update the streaming examples to yield to their destination through writeRows() or periodic flush().
