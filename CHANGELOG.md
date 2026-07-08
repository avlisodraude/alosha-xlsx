# Changelog

## 0.12.0

### Minor Changes

- 57b6f96: feat: round-trip shared and array formulas

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

- 9a31506: fix: re-anchor comments/validations/conditional formatting on splice, and bridge spliceColumns' 0-based inserts in compat

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

### Patch Changes

- c506ec2: fix: escape embedded double quotes in conditional-formatting text-rule formulae

  `containsText`/`notContainsText`/`beginsWith`/`endsWith` conditional-formatting rules generate an
  Excel `<formula>` that interpolates the rule's `text` into a string literal (e.g.
  `SEARCH("...",A1)`). A `text` value containing a `"` was interpolated unescaped, producing an
  invalid formula that made Excel prompt to repair the file on open. Embedded quotes are now doubled
  (`"` → `""`) per Excel string-literal syntax. The `text=` XML attribute was already escaped
  correctly and is unaffected.

- cc493ba: fix: guard numeric character reference decoding against RangeError on out-of-range code points

  `decodeEntities` in `src/xml/xml-reader.ts` guarded `&#nn;`/`&#xhh;` numeric character references
  with `Number.isFinite` before calling `String.fromCodePoint`, but that function throws `RangeError`
  for code points beyond `0x10FFFF` or lone surrogate halves (`0xD800`-`0xDFFF`) — neither of which
  `Number.isFinite` rejects. A corrupt or crafted `.xlsx` part (e.g. `&#x110000;` in
  `sharedStrings.xml`) crashed `readWorkbookBuffer` with an untyped `RangeError`, bypassing the typed
  error contract in `src/errors.ts`. `decodeEntities` now validates the code point range before
  converting and falls back to the original match text when it's invalid, the same fallback already
  used for unknown named entities.

- e62e5f3: fix: preserve per-cell styles inside merged ranges on read

  Reading a workbook destroyed the styles of non-master (slave) cells inside a merged range. Repro:
  merge `A1:B2`, set a thick bottom border on `B2`, `writeWorkbookBuffer` → `readWorkbookBuffer` →
  `B2.border` came back `undefined`. Cause: the reader (and `Worksheet` `set model`) applied merges via
  the public `mergeCells()`, whose `Cell.merge(master)` defaults to copying the master's style over each
  slave — correct at authoring time, but it clobbered the slave's just-parsed style when re-hydrating a
  saved workbook. The read / `set model` / merge re-anchoring paths now link slaves with `ignoreStyle`,
  so each slave keeps its own style; authoring-time `mergeCells()` still propagates the master's style
  (unchanged ExcelJS-compatible behavior).

- 1b4057b: chore: metadata & packaging cleanup

  A batch of small correctness/packaging fixes with no runtime behaviour change beyond the two API
  tweaks noted below:

  - **`version` now reports the real package version.** `export const version` was hardcoded to
    `"0.0.0"`; it's now injected from `package.json` at build time via tsup's `define` (mirrored in the
    vitest config so it resolves under test too).
  - **`getWorksheet(name)` falls back to a case-insensitive match.** Exact-case matches still win, but a
    miss now retries case-insensitively, lining up with `addWorksheet`'s case-insensitive duplicate rule
    (so `getWorksheet("report")` finds a `"Report"` sheet).
  - **`cell.note = undefined` and `cell.dataValidation = undefined` now clear** the comment / validation
    (routing to the store's `remove`), instead of only accepting a value.
  - **Smaller published package.** Sourcemaps are no longer shipped (emitted only for `pnpm dev` watch
    builds), dropping the tarball from ~585KB to ~195KB (unpacked ~2.55MB → ~800KB).
  - **`sideEffects` now covers the hashed ESM chunk** (`./dist/*.js`, `./dist/*.cjs`) that actually holds
    the `workbook.xlsx` accessor's `Object.defineProperty` side effect, so aggressive tree-shakers can't
    drop it.
  - **Docs:** README now states Node.js `>= 22` (matching `engines`) and documents the case-insensitive
    `getWorksheet` fallback.

- da8eb7a: fix: round-trip sheet view zoom scale

  `renderSheetViews` (`src/xlsx/parts/worksheet.ts`) only emitted `showGridLines` and `<pane>` on the
  `<sheetView>` element, silently dropping `zoomScale`; `parseSheetViews`
  (`src/xlsx/read/worksheet.ts`) only read those two back. A view set to `{ zoomScale: 150 }`
  round-tripped to an empty `views` array, contradicting the README's claim that views (including
  zoom) round-trip. `renderSheetViews` now emits `zoomScale`/`zoomScaleNormal` (mirrored, since we
  don't track a separate normal-view zoom) whenever the zoom isn't the default 100, in both the
  frozen/split and plain branches; `parseSheetViews` reads `zoomScale` back, coercing to a number and
  dropping it (along with the rest of the view, if otherwise empty) when it's the default 100. The
  streaming writer's prologue calls the same `renderSheetViews`, so it inherits the fix without
  changes.

- 47346e7: fix: dedupe repeated workbook media in the streaming writer

  `WorkbookWriter` wrote duplicate ZIP entries when the same workbook-registered image was placed on
  multiple sheets. Repro: `wb.addImage` once, `addImage` the returned id on two sheets, each
  `ws.commit()` independently pushes the same `xl/media/imageN.ext` `ExtraPart` — `WorkbookWriter.addExtraPart`
  wrote it into the archive unconditionally every time, producing two ZIP local-file-headers (and two
  central-directory records) for one part name, an OPC violation. The buffered writer never had this bug:
  `collectWorksheetParts` gathers parts into a `Map` keyed by path, so a repeat silently overwrites the
  same key. `addExtraPart` now tracks written part paths in a `Set` and skips a repeat — its content is
  identical by construction (same `imageId` resolves to the same registry buffer).

- d27a942: fix: sanitize characters illegal in XML 1.0 when writing cell text and attributes

  `escapeXmlText`/`escapeXmlAttr` only escaped `& < > " '`. Cell text containing control characters
  (e.g. `\x0B` vertical tab, common when pasting from PDFs), lone surrogate halves, or the
  U+FFFE/U+FFFF noncharacters was written verbatim into `sharedStrings.xml`/inline strings and
  attribute values, producing XML that parsers and Excel reject. These characters are now re-encoded
  using Excel's own `_xHHHH_` convention (rather than stripped), and the reader decodes the same
  token back, so values round-trip through our own writer/reader intact.

## 0.11.1

### Patch Changes

- 75a10e6: docs: homepage → xlsx.alosha.dev, live demo link.

  `package.json` `homepage` now points at `https://xlsx.alosha.dev` (was the GitHub `#readme`
  anchor), so the npm package page's "Homepage" link resolves to the project site. The README
  gains a top-of-file live demo link and badge (`https://xlsx.alosha.dev/demo`) so it surfaces on
  the npm page too.

## 0.11.0

### Minor Changes

- Follow-up batch: streaming column widths, a real compat streaming writer, and two conditional-formatting parity fixes.

  - **Streaming column-width support (`<cols>`).** The streaming `WorksheetWriter` now exposes
    `columns` / `getColumn` (widths, styles, hidden/outline flags), backed by its real model
    worksheet and emitted in the sheet prologue after `<sheetFormatPr>` (CT_Worksheet order).
    Same rule as `views`/`properties`: configure columns **before the first row commits** — a
    column set afterward is silently ignored. This closes the 0.10.0 known limitation; streaming
    now has full feature parity with the buffered writer (write-side).

  - **`ExcelJS.stream.xlsx.WorkbookWriter` in `@alosha/xlsx/compat` is a real bridge** onto the
    streaming writer instead of a throwing stub (MIGRATION.md D6). Constructor options
    `filename`/`stream`/`useSharedStrings`/`useStyles` (plus `creator`/`lastModifiedBy`/`created`/
    `modified`) map through; the full ExcelJS lifecycle works (`addWorksheet()` →
    `ws.addRow(...).commit()` → `ws.commit()` → `await workbook.commit()`), with rows/cells proxied
    like the buffered compat facade (ExcelJS `cell.type` numbers, whole-object style setters,
    `cell.note`). Options with no streaming equivalent — `zip`, `lastPrinted`, and `addWorksheet`'s
    `pageSetup`/`headerFooter`/per-sheet `useSharedStrings` — throw `UnsupportedFeatureError`
    rather than being silently dropped. `stream.xlsx.WorkbookReader` still throws (no streaming
    reader yet).

  - **Conditional-formatting compat parity (closes both MIGRATION.md D3 divergences).**
    `removeConditionalFormatting` now accepts ExcelJS's index, predicate, and no-argument
    (clear-all) forms, resolved onto the core's by-ref removal. The text-contains family is
    bridged bidirectionally: ExcelJS's `{ type: "containsText", operator: "beginsWith" | "endsWith"
| "notContains" | "containsText" }` shape normalizes to the core's per-variant `type` on
    `addConditionalFormatting`, and `conditionalFormattings` always reads back the ExcelJS shape.

  - **Lint cleanup:** removed the last `noNonNullAssertion` warning in the buffered write
    orchestrator; `pnpm lint` is warning-free.

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
