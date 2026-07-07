# Changelog

## 0.10.0

### Minor Changes

- Streaming `WorkbookWriter` (write-only, ExcelJS `stream.xlsx.WorkbookWriter`-shaped): add a sheet,
  commit rows one at a time, and each row's XML is deflated straight into the output archive —
  bounded memory for large workbooks. At 500k rows × 5 cols it holds ~15.6× less peak memory than
  the buffered writer at ~1.9× the throughput (`pnpm bench:stream`, BENCHMARKS.md).

  ```ts
  const wb = new WorkbookWriter({ filename }); // or { stream }, or neither → AsyncIterable<Uint8Array>
  const ws = wb.addWorksheet("Data");
  ws.addRow([1, "a", new Date()]).commit(); // flushed now
  ws.commit();
  await wb.commit();
  ```

  - **Sinks.** A Node `Writable`/`filename` (backpressure-honouring, `node:fs` imported lazily), or —
    with no destination — the writer itself is an `AsyncIterable<Uint8Array>` of archive chunks
    (browser/edge-safe; no `node:*` on that path).
  - **Memory model.** Defaults `useSharedStrings: false` (inline `<is>` strings) so a committed row
    retains nothing — the one thing OOXML would force into memory is the shared-string table, written
    after all sheets. Opt in to `useSharedStrings: true` and rows still stream; only the string table
    is held.
  - **Parity.** Full buffered-writer feature coverage — values (string/number/boolean/date/error/
    formula/rich text/hyperlink), styles, merges, panes + auto-filter, data validation, conditional
    formatting, comments, images — asserted structurally equal to `writeWorkbookBuffer` per feature;
    `<dimension>` is intentionally omitted (extent unknown until the last row).
  - **Docs.** New `docs/STREAMING.md` usage guide + README section.

  **Known limitations:** write-only (no streaming `WorkbookReader` yet); column widths / `<cols>` are
  not yet supported in streaming mode (`getColumn`/`columns` intentionally not exposed) — a planned
  follow-up; `ExcelJS.stream.xlsx.WorkbookWriter` in `@alosha/xlsx/compat` still throws — import
  `WorkbookWriter` from the main entry instead.

## 0.9.0

### Minor Changes

- Embedded images (`workbook.addImage` + `worksheet.addImage`, write + read + compat), the second real
  consumer of the F0 multi-part worksheet seam.

  Registering an image once at the workbook — `workbook.addImage({ buffer, extension: "png"|"jpeg"|"gif" })`
  — returns a numeric `imageId`; placing it on a sheet — `worksheet.addImage(id, "B2:D6")` (two-cell anchor)
  or `worksheet.addImage(id, { tl: { col, row }, ext: { width, height } })` (one-cell anchor, pixel extents)
  — is written as an `xl/drawings/drawing{n}.xml` DrawingML anchor part, the raw bytes in
  `xl/media/image{k}.{ext}`, the drawing's own `.rels`, and a `<drawing r:id>` in the worksheet (CT_Worksheet
  order: after `<hyperlinks>`, before any `<legacyDrawing>`).

  - **Model.** A workbook-level `Image` registry (dedup-free, insertion-order `imageId`s) plus a per-worksheet
    `WorksheetImage[]` of `{ imageId, range }`; the `"B2:D6"` string anchor decodes through the existing
    `Range` parser into 0-based `tl`/`br`.
  - **Write.** EMU conversion (1 pixel = 9525 EMU) for one-cell/`ext` anchors; two-cell anchors keep corner
    offsets at 0 unless `nativeColOff`/`nativeRowOff` are given. Media bytes push through the F0 `ExtraPart`
    seam with the right content type; a workbook with no images is byte-identical to a pre-images write.
  - **Read.** Follows the worksheet's `<drawing>` rel to `drawing{n}.xml`, each anchor's `<a:blip r:embed>`
    through the drawing's `.rels` to the media path, and reconstructs the registry + anchors (EMU back to
    pixels). Missing/unresolvable drawing or media parts leave a sheet without that image rather than
    throwing.
  - **Compat.** `workbook.addImage`/`getImage`/`media` and `worksheet.addImage`/`getImages` are un-deferred:
    ExcelJS's `{ buffer }` image input and its three anchor forms (`"B2:D6"` / `{ tl, br }` / `{ tl, ext }`)
    already match the core shapes, so they pass straight through. `{ base64 }`/`{ filename }` image inputs
    throw `UnsupportedFeatureError` instead of silently mishandling the image — see MIGRATION.md's new D5
    note, which also covers real ExcelJS's `Anchor`-class `getImages()` shape. `worksheet.addBackgroundImage`
    stays a permanent, documented stub.

  **Deferred:** image cropping/rotation/effects, hyperlinked images, header/footer (background) images, SVG,
  and EMF/WMF.

## 0.8.0

### Minor Changes

- Cell comments (`cell.note`, write + read + compat), the first real consumer of the F0 multi-part
  worksheet seam.

  A comment on a cell — `worksheet.getCell("B2").note = "text"` or the object form `{ texts: RichText |
string, author? }` — is written as an `xl/comments{n}.xml` note/author table plus a legacy
  `xl/drawings/vmlDrawing{n}.vml` anchor part, wired through a `<legacyDrawing r:id>` (CT_Worksheet's last
  child, after `<drawing>`).

  - **Model.** A worksheet-level `Comments` manager keyed by cell ref (single-cell, unlike range-keyed data
    validations); `cell.note` get/set routes through it, mirroring `cell.dataValidation`. Default author
    "Author" when unset (ExcelJS parity).
  - **Write.** `<authors>` (distinct authors, stable first-seen order) + `<commentList>` with a rich
    `<text>`, reusing the shared-string run renderer. A plain-string comment still renders as a single
    `<r><t>` run rather than a lone `<t>`, since exceljs's own comment parser only reads `<r>` children —
    writing a lone `<t>` would make the text invisible to exceljs (and to any file it re-opens). The VML
    part carries one `<v:shape>` per comment with a 0-based row/column anchor; geometry is a fixed default.
  - **Read.** Follows `<legacyDrawing>`/rels to `comments{n}.xml`, resolving `authorId` back to an author
    string; a lone `<t>` or `<r>` runs both invert correctly (real Excel and other tools may still emit the
    lone-`<t>` form). Malformed/unresolvable parts leave a cell without a note rather than throwing.
  - **Compat.** `cell.note` is un-deferred: a bare string passes straight through, and ExcelJS's flat
    `texts: [{font, text}]` run array is bridged to/from the core's `RichText` wrapper. `author` is a
    compat/native-API extension — real ExcelJS itself never round-trips one (its own writer hardcodes a
    single author, and its reader never resolves one back onto `note.author`) — see MIGRATION.md's new D4
    note, which also covers the plain-string round-trip caveat and exceljs's valued-cell requirement for
    a note to survive being reopened.

  **Deferred:** threaded/modern comments (`xl/threadedComments/` + persons), comment box
  geometry/margins/`editAs`/visibility.

## 0.7.0

### Minor Changes

- Conditional formatting (write + read + compat), backed by a new differential-style (`<dxfs>`) pool.

  Worksheets can now carry `<conditionalFormatting>` blocks — `addConditionalFormatting({ ref, rules })`,
  `removeConditionalFormatting(ref)`, and a `conditionalFormattings` getter — covering `cellIs`,
  `expression`, the text-contains family (`containsText`/`notContainsText`/`beginsWith`/`endsWith`),
  `top10`, `aboveAverage`/`belowAverage`, `duplicateValues`/`uniqueValues`, and the self-visualising
  `colorScale`/`dataBar`/`iconSet` rules.

  - **Model.** `ConditionalRule` is a discriminated union keyed by `type`, mirroring ExcelJS's
    `addConditionalFormatting` shape so the compat layer stays a pass-through; `ref` is validated
    through the existing `Range` decoder.
  - **Write.** A real `<dxfs>` pool in `xl/styles.xml` (dxf ids start at 0, unlike the seeded
    font/fill/border pools) backs every dxf-bearing rule's differential `style`; `<conditionalFormatting>`
    is emitted between `<mergeCells>` and `<dataValidations>`, matching CT_Worksheet element order.
    Rules with no explicit `priority` draw a workbook-monotonic one.
  - **Read.** Inverts every write shape, including the dxf pool (**footgun**: a dxf highlight's colour
    lives in `<patternFill><bgColor>`, not `<fgColor>` like a cellXf solid fill). A malformed
    `priority`/`rank` raises a typed `AloshaXlsxError` instead of `NaN`; an unrecognized `cfRule` type,
    or a block missing `sqref`, is skipped rather than thrown.
  - **Compat.** `worksheet.addConditionalFormatting`/`removeConditionalFormatting`/
    `conditionalFormattings` are un-deferred. `colorScale`/`dataBar` colours are bridged bidirectionally
    between ExcelJS's `{ argb }` objects and the core model's bare-ARGB-string `CfColor`. Two shapes are
    _not_ bridged, by design — see MIGRATION.md's D3 note: the text-contains family's ExcelJS
    `type:"containsText"` + `operator` split (construct these with the core's literal `type` per variant
    instead), and `removeConditionalFormatting`'s ExcelJS index/predicate/clear-all argument (the core
    removes by exact `ref` string).

  **Deferred:** the x14/`<extLst>` dataBar extensions (gradient fill, axis, negative-value fill) — only
  the ECMA-376 baseline `<dataBar>` is emitted — `timePeriod` rules, and custom `iconSet` icons.

## 0.6.0

### Minor Changes

- Freeze/split panes and sheet auto-filter (write + read + compat).

  Worksheets can now carry a single `<sheetView>` with a `<pane>` — `views = [{ state: "frozen" |
"split", xSplit?, ySplit?, topLeftCell?, activePane?, showGridLines? }]` — and a sheet-level
  `autoFilter` range (`"A1:D1"`).

  - **Model.** `WorksheetView` and `worksheet.autoFilter` round-trip through `WorksheetModel` like
    merges and column widths.
  - **Write.** `<pane>` is emitted inside `<sheetView>`, and `<autoFilter ref="…">` after
    `<sheetData>`/before `<mergeCells>`, matching CT_Worksheet element order.
  - **Read.** Both invert back to an equal model; a malformed `<pane>` (e.g. a non-numeric `xSplit`)
    raises a typed `AloshaXlsxError` instead of `NaN`.
  - **Compat.** `@alosha/xlsx/compat`'s `worksheet.views` and `worksheet.autoFilter` are a real
    pass-through — the core field names already match ExcelJS's. The one bridge: setting
    `worksheet.autoFilter` to ExcelJS's `{ from, to }` object form (each an address string or
    `{ row, column }`) normalizes to our `"A1:D1"` range string before it reaches the core.

  **Deferred:** multiple views per sheet, `<selection>`, per-pane active-cell tracking, split-window
  zoom, and auto-filter criteria/sort state (`<filterColumn>`, `<sortState>`) — only the `ref` range
  persists, matching where ExcelJS's basic `autoFilter` sits.

## 0.5.0

### Minor Changes

- Cell data validations — dropdowns and value constraints (write + read + compat).

  Cells and ranges can now carry a data validation: a **list dropdown** (an inline
  `'"Yes,No,Maybe"'` or a same-sheet range `"$E$1:$E$5"`), a numeric/date/length constraint
  (`whole` / `decimal` / `date` / `time` / `textLength`) with an operator
  (`between` / `notBetween` / `equal` / `notEqual` / `greaterThan` / `lessThan` /
  `greaterThanOrEqual` / `lessThanOrEqual`) and one or two formulae, or a `custom` boolean formula —
  plus the usual flags and messages (`allowBlank`, `showInputMessage`, `showErrorMessage`,
  `errorStyle`, `errorTitle`, `error`, `promptTitle`, `prompt`, `showDropDown`). Applied per cell via
  `cell.dataValidation` or per range via `worksheet.dataValidations.add(sqref, dv)`, where `sqref` may
  be a single address, a range, or space-separated ranges.

  - **Model.** New `DataValidation` type and a per-worksheet `DataValidations` manager
    (`add` / `find` / `remove` / `model`, keyed by `sqref`); `cell.dataValidation` get/set resolves
    through it, and validations round-trip through `WorksheetModel` like merges.
  - **Write.** A `<dataValidations>` block is emitted in the sheet, in CT_Worksheet order (after
    `<mergeCells>`, before `<hyperlinks>`), so the validation re-opens in Excel/LibreOffice as a
    working dropdown/constraint. No extra parts or relationships.
  - **Read.** The `<dataValidations>` block parses back to an equal model; malformed entries are
    skipped rather than throwing.
  - **Compat.** `@alosha/xlsx/compat`'s `cell.dataValidation` and `worksheet.dataValidations` are now
    a real pass-through (previously threw `UnsupportedFeatureError`) — the core shape already matches
    ExcelJS's `DataValidations`.

  **OOXML `showDropDown` inversion:** the OOXML attribute is inverted — `showDropDown="1"` _hides_ the
  dropdown. The model uses the intuitive sense, so `showDropDown: false` means "hide it" and maps to
  the attribute; unset/`true` shows the dropdown (the default) and emits nothing.

  **Deferred:** a list that references a range on **another sheet** (the x14 / `<extLst>` form Excel
  uses for cross-sheet list refs) is not yet supported — only same-sheet ranges and inline lists.

## 0.4.0

### Minor Changes

- Persist external cell hyperlinks as real OOXML relationships (write + read + compat).

  Setting a `{ text, hyperlink, tooltip? }` cell value now writes a proper worksheet
  `<hyperlinks>` record plus an external relationship in
  `xl/worksheets/_rels/sheetN.xml.rels`, so the link is clickable in Excel/LibreOffice and
  survives a round-trip — instead of degrading to its plain display text. External
  `http(s)` and `mailto:` targets are supported; the display text still renders as an
  ordinary string cell.

  - **Write.** Hyperlink cells emit a `<hyperlinks>` block (after `<mergeCells>`) with
    worksheet-local `r:id`s and a matching `.rels` part.
  - **Read.** The `<hyperlinks>` block is resolved against the worksheet's own rels, so
    hyperlink cells parse back to `{ text, hyperlink, tooltip? }` with `CellType.Hyperlink`.
  - **Compat.** No API change — `@alosha/xlsx/compat` hyperlink cells now round-trip, and
    `cell.type` is `ValueType.Hyperlink` (5).

  **Behaviour change (why this is a minor):** reading a workbook that contains hyperlinks now
  yields a `{ text, hyperlink }` cell value with `CellType.Hyperlink` (compat:
  `ValueType.Hyperlink`), where 0.3.x returned a plain **string**. Code that read hyperlink
  cells expecting a string will now receive an object.

  Internal in-workbook (`location="Sheet2!A1"`) hyperlinks remain deferred; the previous
  "hyperlink cells round-trip as their display text" known limitation is removed from the
  README and `docs/MIGRATION.md` in favour of that narrower note.

## 0.3.1

### Patch Changes

- 36706d6: compat: `eachRow`/`eachCell` now skip empty rows/cells by default, matching ExcelJS.

  Without `includeEmpty`, the `@alosha/xlsx/compat` facade's `worksheet.eachRow` and
  `row`/`column.eachCell` now skip rows/cells that are empty or were explicitly set to `null` (as
  ExcelJS does), rather than yielding every row/cell ever touched via `getRow`/`getCell`. Pass
  `{ includeEmpty: true }` for full enumeration. Callback rows/cells remain proxied (so `cell.type`
  still reports ExcelJS `ValueType` numbers).

  Docs: the README `BENCHMARKS.md`/`CHANGELOG.md` links are now absolute so they resolve on npmjs.com,
  and the corresponding note is removed from `docs/MIGRATION.md`.

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
