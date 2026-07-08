# Migrating from ExcelJS

`@alosha/xlsx/compat` is a drop-in for [`exceljs`](https://github.com/exceljs/exceljs): the same
`new Workbook()` / `addWorksheet` / `getCell` / `addRow` / `mergeCells` /
`workbook.xlsx.writeBuffer`/`writeFile`/`load`/`readFile` surface, backed by `@alosha/xlsx`'s model and
OOXML writer/reader instead of ExcelJS's. For the common case — build a workbook, set values, styles,
number formats, merges, column defs, then serialize/reload it — **the only change is the import.**

```diff
- import ExcelJS from "exceljs";
+ import ExcelJS from "@alosha/xlsx/compat";

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Report");

  sheet.getCell("A1").value = "Item";
  sheet.getCell("B1").value = "Qty";
  sheet.getColumn(2).width = 12;
  sheet.getCell("B1").font = { bold: true };

  await workbook.xlsx.writeFile("report.xlsx");
```

Everything below the diff line is unmodified ExcelJS code.

If you're migrating *away* from the ExcelJS surface entirely (not just swapping the import), see the
main [`README.md`](../README.md#migrating-from-exceljs) instead — that's a deeper rewrite (0-based dense
`worksheets`, an object-shaped `inheritFrom` instead of the `'i'|'o'|'i+'|'o+'` DSL, etc.) with more
freedom to diverge, since it doesn't have to stay drop-in compatible.

## Seven behavioural notes — read these before porting

**D1 — `cell.type` / `cell.effectiveType` report ExcelJS's numbers, not `@alosha/xlsx`'s own.**
`@alosha/xlsx`'s core `CellType` enum shares names with ExcelJS's `ValueType` but **not** the same
numbers (e.g. `Number` is `1` in the core model, `2` in ExcelJS). In compat mode this is fully emulated:
every cell reached through a `@alosha/xlsx/compat` workbook reports the ExcelJS number, so `if (cell.type
=== ExcelJS.ValueType.Formula)`-style checks port unchanged. This *only* applies to workbooks built via
`@alosha/xlsx/compat`'s `Workbook` — a plain `@alosha/xlsx` `Workbook` still reports its own numbers.

**D2 — style setters replace the whole category, not just the field you touched.**
`@alosha/xlsx`'s core model does a **deep field-level merge** when resolving a cell's style: setting
`cell.font = { bold: true }` on a cell that inherited `{ name: "Arial", size: 12 }` from its row leaves
you with `{ name: "Arial", size: 12, bold: true }`. ExcelJS instead **replaces the category outright** —
the same assignment discards the inherited `name`/`size`, leaving just `{ bold: true }`. A file migrated
from ExcelJS relies on the latter to render identically, so compat-mode style setters
(`font`/`fill`/`border`/`alignment`/`protection`/`numFmt`) always apply ExcelJS's whole-object precedence,
never the core's field merge. Again, this is scoped to compat-constructed workbooks; the plain
`@alosha/xlsx` API is unaffected.

**D3 — conditional formatting: the colour bridge is bidirectional; the text-family and
`removeConditionalFormatting` shapes are too, as of 0.10.1.**
`worksheet.addConditionalFormatting`/`conditionalFormattings` normalize a `colorScale`/`dataBar` rule's
colour between ExcelJS's `{ argb }` object and the core model's bare ARGB string automatically — pass
either shape in, and `conditionalFormattings` always reads back ExcelJS's `{ argb }` form. Two more
shapes are bridged the same way:
- **The text-contains family's `type`/`operator` split.** Real ExcelJS represents `beginsWith`,
  `endsWith`, `notContainsText`, etc. as `{ type: "containsText", operator: "beginsWith", … }` — a
  single `type` with the real variant carried on `operator`. `@alosha/xlsx`'s core model instead uses
  a literal `type` per variant (`{ type: "beginsWith", … }`, no `operator` field). Either shape works
  through the compat facade now: an ExcelJS-idiomatic `{ type: "containsText", operator: "beginsWith" }`
  call normalizes to the core's `{ type: "beginsWith" }` before it reaches the writer, and
  `conditionalFormattings` always reads back ExcelJS's `{ type: "containsText", operator }` form
  (mirroring the colour bridge above), whether or not the block was constructed with `operator`.
- **`removeConditionalFormatting`'s argument.** ExcelJS accepts an index, a predicate function, or no
  argument (clear all); the core model's `removeConditionalFormatting(ref)` removes by exact `ref`
  string match instead. All three ExcelJS forms now work through the compat facade — an index or
  predicate is resolved against a snapshot of `conditionalFormattings` (the predicate sees the same
  denormalized, ExcelJS-shaped blocks that getter returns) and mapped onto the matching block(s)' core
  `removeConditionalFormatting(ref)` call; a call with no argument (or any other value) clears every
  block, matching ExcelJS's `else` branch. One edge case doesn't fully match real ExcelJS: because the
  bridge ultimately removes **by `ref`**, an index/predicate call that's meant to keep one block but
  drop another sharing the *exact same* `ref` string will drop both — real ExcelJS's index-based splice
  has no such collision. Give blocks distinct `ref`s (the common case) to avoid it.

**D4 — `cell.note` (comments): author isn't part of ExcelJS's real contract, and a note needs a valued
cell.** Real ExcelJS's own comments writer hardcodes a single author name ("Author") for every comment
and never resolves one back onto `note.author` when reading a file — `author` simply isn't part of
ExcelJS's actual round trip, even though its object form doesn't forbid the key. `@alosha/xlsx/compat`
accepts and returns an optional `author` on `cell.note` regardless (useful when interoperating with the
plain `@alosha/xlsx` API, whose `Comments` manager is genuinely per-author), but a file that has passed
through real ExcelJS will only ever show `"Author"`. Separately, a plain-string note (`cell.note =
"text"`) always serializes as a single `<r><t>` run rather than a lone `<t>`, because ExcelJS's own
comment parser only reads `<r>` children and would otherwise see no text at all; reading that file back
through `@alosha/xlsx`'s own reader recovers it as a one-run rich text (`{ richText: [{ text: "text"
}] }`) rather than the original bare string — expect either shape rather than relying on the string
staying a string across a full write→read cycle. Finally, ExcelJS only recognises a comment on a cell
that also carries a value in the sheet XML, so give a noted cell a value if you want the note to survive
being opened elsewhere.

**D5 — embedded images: only the `{ buffer }` input, and `worksheet.addBackgroundImage` is a
permanent stub.** `workbook.addImage(image)` accepts ExcelJS's `{ extension, buffer }` form and routes
straight to the core image registry; the `{ base64 }` and `{ filename }` input forms throw
`UnsupportedFeatureError` instead of silently decoding (or worse, silently dropping) the image —
re-encode to a `Uint8Array` and pass it as `buffer`. `worksheet.addImage`'s three ExcelJS anchor forms
(`"B2:D6"`, `{ tl, br }`, `{ tl, ext }`) already match the core `ImageAnchor` type, so they pass
straight through with no bridging. `worksheet.addBackgroundImage` (header/footer images) stays a
documented stub — out of scope, not just "not yet" — and always throws. Note for anyone diffing
against real ExcelJS directly: real ExcelJS's own `worksheet.getImages()` returns each anchor's
`tl`/`br` as `Anchor` **class instances** (position under `nativeCol`/`nativeRow`, plus a
`worksheet` back-reference), not the plain `{ col, row }` object `@alosha/xlsx/compat`'s
`getImages()` returns — reload through real `exceljs` if you need that exact shape.

**D6 — `stream.xlsx.WorkbookWriter` is a real bridge onto our own streaming writer; a few
constructor/`addWorksheet` options with no equivalent throw instead of silently no-op'ing.**
`new ExcelJS.stream.xlsx.WorkbookWriter(options)` maps `filename`/`stream`/`useSharedStrings`/
`useStyles` straight onto our own `xlsx/stream` `WorkbookWriter` (docs/STREAMING.md) and applies
`creator`/`lastModifiedBy`/`created`/`modified` as a post-construction assignment (our streaming
writer has matching setters, just no constructor form for them). The lifecycle is the same as real
ExcelJS: `addWorksheet()` → `ws.addRow(...).commit()` → `ws.commit()` → `await workbook.commit()`,
and the workbook is itself an `AsyncIterable<Uint8Array>` when neither `filename` nor `stream` is
given. Rows/cells crossing the boundary are real model objects, so they get the same proxying as the
buffered facade for free (ExcelJS `cell.type` numbers, whole-object style setters, `cell.note`
bridging, the §4 conditional-formatting colour/text-family bridge, ExcelJS's `{from,to}` `autoFilter`
form). Two constructor options and three `addWorksheet` options have **no equivalent** on the
streaming writer and throw `UnsupportedFeatureError` naming the option, rather than being silently
dropped: the constructor's `zip` (archiver-specific tuning) and `lastPrinted` (untracked by the
streaming writer's doc-metadata surface, unlike the buffered `Workbook`); `addWorksheet`'s `pageSetup`
and `headerFooter` (printing config, unimplemented for streaming) and a per-sheet `useSharedStrings`
override (only a workbook-wide setting exists). `stream.xlsx.WorkbookReader` has no streaming
counterpart yet and still throws unconditionally.

**D7 — `spliceColumns` insert arrays are 1-based in the core, 0-based through compat; splice re-anchors
comments/validations/conditional formatting but not images/autoFilter.** `@alosha/xlsx`'s core
`worksheet.spliceColumns(start, count, ...inserts)` reads each insert array **1-based** — `inserts[k][n]`
fills row `n`, index `0` unused — consistent with `Column.values` and every other 1-based array in the
core API. Real ExcelJS's `spliceColumns` instead reads them **0-based** (`inserts[k][i]` fills row
`i + 1`), so ported ExcelJS code would silently land data one row too high against the raw core. The
compat facade bridges this: through a `@alosha/xlsx/compat` workbook, `spliceColumns` accepts ExcelJS's
0-based insert arrays unchanged, so migrated code keeps working — only the plain `@alosha/xlsx` API uses
the 1-based convention. (`spliceRows` inserts are 0-based in both — its `inserts[i]` is a whole row, the
same as ExcelJS — so it needs no bridge.) Separately, both splice methods now **re-anchor** the
ref-keyed annotations that don't live on cells — comment refs, data-validation `sqref`s, and
conditional-formatting `ref`s — with the same rules the merge re-anchoring uses: an annotation entirely
past the removed region shifts, one entirely before it is untouched, and one straddling the removed
region is dropped. **Images and `autoFilter` are deliberately left at their original addresses** on a
splice — real ExcelJS doesn't move them either (its `spliceRows`/`spliceColumns` only re-merge cells and
update defined names), so matching that keeps behaviour predictable for anyone porting ExcelJS code. If
you splice around an image or filter range, re-place or reset it yourself afterwards.

## Supported / deferred matrix

| Feature | Status | Notes |
| --- | :---: | --- |
| `Workbook`/`Worksheet`/`Row`/`Column`/`Cell` core API, `getCell`/`getRow`/`addRow`/`addRows`/`mergeCells`/`eachRow`/`eachCell` | ✅ | |
| `addRow(array \| object)`, `row.values` (1-based, empty `[0]`), `worksheet.columns = [{header,key,width}]` | ✅ | |
| Values — string/number/boolean/date/null, formula (`{formula,result}`), rich text, error | ✅ | |
| `cell.type` / `cell.effectiveType` | ✅ | ExcelJS `ValueType` numbers — see D1 above |
| Style setters (`font`/`fill`/`border`/`alignment`/`protection`/`numFmt`) | ✅ | ExcelJS whole-object precedence — see D2 above |
| `Enums.ValueType`, `Enums.FormulaType` | ✅ | |
| `workbook.xlsx.writeBuffer`/`writeFile`/`load`/`readFile`/`read` | ✅ | |
| Hyperlink cells (`{ text, hyperlink }`, external URL/mailto) | ✅ | persists as a real OOXML relationship; round-trips to `{text, hyperlink, tooltip?}` with `cell.type` `Hyperlink`. Internal `location=` (in-workbook) links are still ⚠️ deferred. |
| `worksheet.dataValidations` (`add`/`find`/`remove`/`model`), `cell.dataValidation` | ✅ | list dropdowns (inline or same-sheet range), whole/decimal/date/time/textLength constraints, custom formulas. Note: OOXML's `showDropDown` attribute is inverted — it *hides* the dropdown when present — which the model corrects (`showDropDown: false` means hidden); see DATA-VALIDATION-SPEC §4. A list referencing a range on **another sheet** (the x14/`<extLst>` form) is still ⚠️ deferred. |
| `worksheet.views` (freeze/split pane), `worksheet.autoFilter` | ✅ | pass-through — the core field names already match ExcelJS's. `autoFilter` also accepts ExcelJS's `{ from, to }` object form (each an address string or `{ row, column }`), normalized to an `"A1:D1"` range string. Multiple views per sheet, `<selection>`, and auto-filter criteria/sort state are still ⚠️ deferred. |
| `stream.xlsx.WorkbookWriter` | ✅ | bridges onto our own streaming writer — see D6 above. `zip`/`lastPrinted` ctor options and `addWorksheet`'s `pageSetup`/`headerFooter`/per-sheet `useSharedStrings` throw `UnsupportedFeatureError` (no streaming equivalent). |
| `stream.xlsx.WorkbookReader` | ❌ deferred | throws `UnsupportedFeatureError`; no streaming reader yet |
| `workbook.csv.*` | ❌ deferred | throws `UnsupportedFeatureError` |
| `workbook.addImage`/`getImage`/`media`, `worksheet.addImage`/`getImages` | ✅ | `{ buffer }` image input; `"B2:D6"` / `{tl,br}` / `{tl,ext}` anchors — see D5 above. `{ base64 }`/`{ filename }` image inputs throw `UnsupportedFeatureError`. Cropping/rotation/effects, hyperlinked images, SVG, and EMF/WMF are deferred. |
| `worksheet.addBackgroundImage` (header/footer images) | ❌ deferred | throws `UnsupportedFeatureError`; permanent stub, not planned |
| `worksheet.addTable`/`getTable(s)` (tables, pivot tables) | ❌ deferred | throws `UnsupportedFeatureError` |
| `worksheet.addConditionalFormatting`/`removeConditionalFormatting`, `worksheet.conditionalFormattings` | ✅ | `cellIs`, `expression`, the text-contains family, `top10`, `aboveAverage`/`belowAverage`, `duplicateValues`/`uniqueValues`, `colorScale`, `dataBar`, `iconSet`. `colorScale`/`dataBar` colours are bridged bidirectionally between ExcelJS's `{argb}` objects and the core's bare-ARGB-string `CfColor` — see D3 below. Deferred: x14 dataBar gradient/axis/negative-fill extensions, `timePeriod` rules, custom `iconSet` icons. |
| `cell.note` (comments) | ✅ | plain-string and rich-text notes, multi-author sheets. `author` is a compat/native-API extension ExcelJS itself doesn't round-trip — see D4 above. Deferred: threaded/modern comments (`xl/threadedComments/`), comment box geometry/margins/`editAs`/visibility. |
| `workbook.definedNames.*` | ❌ deferred | throws `UnsupportedFeatureError` |
| `worksheet.protect`/`unprotect` | ❌ deferred | throws `UnsupportedFeatureError` |

Every deferred method fails loudly and precisely at the exact call — `UnsupportedFeatureError` names the
feature — rather than silently no-op'ing or producing a subtly-wrong file. If your workbook only touches
supported features, migration genuinely is a one-line import swap; if it hits anything in the deferred
list, that call throws immediately so you find out at the call site, not by diffing a corrupt `.xlsx`.

## Other differences worth knowing about

These aren't compat-specific bugs — they're pre-existing `@alosha/xlsx` core-model behaviours that a
migrated file can still trip over, surfaced while building the parity corpus (below):

- **An unstyled cell's `fill` is `undefined`, not `{ pattern: "none" }`.** ExcelJS always resolves a
  cell's fill from the styles pool, so an unstyled cell still reports `{ type: "pattern", pattern:
  "none" }`. `@alosha/xlsx`'s reader treats the baseline "no fill" style id as "nothing to report" and
  leaves `cell.fill` as `undefined`. Code that does `cell.fill.pattern` unconditionally (rather than
  `cell.fill?.pattern`) will need a null check either way — but the difference is worth knowing before
  you rely on `cell.fill` always being an object.

## Verifying the claim: the parity corpus

`tests/compat/parity.test.ts` is the mechanical proof behind this document. It takes a corpus of ~10
representative ExcelJS snippets (every value type, a styled header row, number formats, a merge, multiple
worksheets, column keys + `addRow(object)`, a formula with a cached result, a `date1904` workbook, and a
row/cell style-precedence case), runs each snippet **unmodified** against both real `exceljs` and
`@alosha/xlsx/compat`, and diffs the two outputs after reloading both with real `exceljs` — plus reloads
the compat-produced buffer with `@alosha/xlsx`'s own reader, to confirm the shim's output is valid input
for our reader too. It's guarded on the optional `exceljs` dev dependency and skips cleanly if that
package isn't installed.
