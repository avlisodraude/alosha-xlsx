# Stage 2 Build Spec — The Reader (`.xlsx` → model)

> **Input for the build sessions.** The inverse of [`STAGE-1.3-SPEC.md`](./STAGE-1.3-SPEC.md): turn a
> valid `.xlsx` back into the Stage 1.2 in-memory `Workbook`. This is the roadmap's **"Stage 3 — add
> READ path"** and PLAN.md phase 3 ("Reader"), implemented as the Stage 2 code increment after the
> Stage 1.3 writer. Completing it makes `@alosha/xlsx` a genuine read+write ExcelJS replacement and
> unblocks the `@alosha/xlsx/compat` shim.
>
> Prereqs (all merged): the Stage 1.2 model (`src/model/*`, `src/address/*`), the Stage 1.3 writer
> (`src/xlsx/*` — we reuse `date.ts`'s `serialToDate`, the `zip-layer`, and the styles/shared-strings
> *shapes* to invert), `src/errors.ts`. Conventions unchanged: TS strict, ESM, `Uint8Array`/`TextDecoder`
> over Node `Buffer`, no `console.*`, pure functions where practical, **fflate stays the only runtime dep**.

## 0. Scope

**In:** parse a single buffered `.xlsx` into a `Workbook` — sheets (name/order/state), rows, cells (every
value type the writer emits: string via shared or inline, number, boolean, date, error, formula+cached
result, richtext), cell/row/column styles and number formats, merged cells, column widths, `date1904`,
and workbook/core/app metadata. Provide the ExcelJS-style `workbook.xlsx.load/readFile/read` surface plus
free functions.

**Out (defer to v0.3+):** the streaming `WorkbookReader`; drawings/images, tables, pivot tables,
conditional formatting, data validation, comments, hyperlink relationships (a hyperlink cell read back as
its display string is acceptable — it round-trips the Stage 1.3 writer's output), protection, calcChain,
formula *recalculation* (we preserve the cached `<v>` only), and the CSV reader.

**Definition of done:** `readWorkbookBuffer(bytes)` returns a `Workbook` such that (a) a
**write→read round-trip** of the Stage 1.3 fixture reproduces every cell value, style, number format,
merge, sheet name/order/state and `date1904` flag; (b) a **foreign file** produced by ExcelJS (generated
in-test via the existing `exceljs` devDep) parses with its sampled values, a merge, and a numFmt intact;
(c) malformed/kind-mismatched input throws a typed `AloshaXlsxError`, never a bare `TypeError`. Detail in §9.

## 1. Module layout

```
src/
  xml/
    xml-reader.ts        // tiny non-validating pull parser: string -> Iterable<XmlEvent> (+ cursor helpers)
  xlsx/
    read/
      package.ts         // unzip + discover parts: [Content_Types].xml, _rels/.rels, workbook.xml.rels
      shared-strings.ts  // parse sharedStrings.xml -> (string | RichText)[]
      styles.ts          // parse styles.xml -> StyleTable { styleFor(index): Style; isDateNumFmt(id) }
      worksheet.ts       // parse sheetN.xml -> populate a Worksheet (cells, cols, merges)
      workbook.ts        // parse workbook.xml -> sheet order/name/state/id, date1904, activeTab
      props.ts           // parse docProps/core.xml + app.xml -> workbook metadata
    read-workbook.ts     // orchestrator: bytes -> Map<path,bytes> -> parsers -> Workbook
    index.ts             // add the read exports to the xlsx barrel
  index.ts               // extend the xlsx accessor with load/readFile/read
tests/
  xml/xml-reader.test.ts
  xlsx/read/*.test.ts
  xlsx/read-roundtrip.test.ts    // write -> read equality (the real proof)
  xlsx/read-foreign.test.ts      // parse an exceljs-produced buffer (guarded on the devDep)
```

Build order = session order (§ appendix): parser + package discovery → sharedStrings + props → styles
(inverse pools) → worksheet grid → orchestrator + public API + round-trip/foreign verification.

## 2. `src/xml/xml-reader.ts` — the pull parser

OOXML is machine-generated and regular, so a compact **non-validating** pull parser keeps the single-dep
(fflate) story rather than pulling in `saxes`/`fast-xml-parser`. It is not a general XML engine — it
targets the well-formed subset Excel/LibreOffice/our writer emit.

```ts
export type XmlEvent =
  | { kind: "open"; name: string; attrs: Record<string, string>; selfClosing: boolean }
  | { kind: "text"; text: string }
  | { kind: "close"; name: string };

/** Stream events in document order. Prefixed names are kept verbatim ("r:id", "cp:coreProperties"). */
export function* parseXml(xml: string): Generator<XmlEvent>;

/** Convenience cursor over the event stream for the recursive-descent readers. */
export interface XmlCursor {
  next(): XmlEvent | undefined;
  peek(): XmlEvent | undefined;
  /** Collect the raw text content until the matching close of `name` (assumes `open` already consumed). */
  textUntilClose(name: string): string;
  /** Invoke `onOpen` for each direct child element of the current open element, then consume its close. */
  eachChild(name: string, onOpen: (ev: Extract<XmlEvent, { kind: "open" }>) => void): void;
}
export function cursor(xml: string): XmlCursor;
```

Requirements: skip the `<?xml …?>` declaration, `<!-- comments -->`, and `<![CDATA[…]]>` (emit its literal
text); decode the five named entities **and** numeric `&#nn;`/`&#xhh;`; parse single- and double-quoted
attributes; handle self-closing `<x/>`; do **not** collapse significant whitespace (shared-string `<t>`
may carry it, honoured via `xml:space="preserve"` on write — on read we keep text verbatim). Namespace
prefixes are matched literally because OOXML fixes them; a `localName(name)` helper strips the prefix for
the few places we compare loosely.

## 3. `src/xlsx/read/package.ts` — unzip & part discovery

```ts
export interface PackageParts {
  entry(path: string): Uint8Array | undefined;
  text(path: string): string | undefined;          // TextDecoder over entry()
  /** rId -> target resolved relative to xl/ (e.g. rId1 -> "xl/worksheets/sheet1.xml"). */
  workbookRels(): Map<string, { type: string; target: string }>;
  /** Absolute path of the workbook part from _rels/.rels (normally xl/workbook.xml). */
  workbookPath(): string;
}
export function openPackage(bytes: Uint8Array): PackageParts;
```

Read `_rels/.rels` for the officeDocument target, then `xl/_rels/workbook.xml.rels` for the styles /
sharedStrings / theme / per-worksheet targets. Resolve relative targets against the workbook part's
directory. `[Content_Types].xml` is parsed only to sanity-check the package is a spreadsheet (fail fast
with `UnsupportedFeatureError` otherwise); part *locations* come from the rels, not fixed paths, so
foreign producers with different file names still resolve.

## 4. `src/xlsx/read/shared-strings.ts` — the string table

```ts
export function parseSharedStrings(xml: string): (string | RichText)[];
```

Invert the Stage 1.3 builder: each `<si>` is a plain string when it holds a lone `<t>`, else a `RichText`
whose runs come from `<r>` (`<rPr>` → `Font` via the inverse of `renderRunProps`, `<t>` → text). Missing
part → empty array. A `t="s"` cell's `<v>` is an index into this array.

## 5. `src/xlsx/read/styles.ts` — inverse dedup pools (the fiddly one)

```ts
export interface StyleTable {
  /** Reconstructed Style for a cell's `s` index; index 0 (or out of range) -> {} (the default). */
  styleFor(index: number): Style;
  /** Whether a cellXf's resolved numFmt is a date/time format (drives number-vs-date on read). */
  isDateStyle(index: number): boolean;
}
export function parseStyles(xml: string | undefined): StyleTable;
```

Rebuild each pool into arrays, then compose per `cellXfs/<xf>`:
- `numFmts`: custom `<numFmt numFmtId formatCode>` (id ≥ 164) merged over the reserved built-in id→code
  map (the reverse of `BUILTIN_NUMFMT_BY_CODE`). `xf.numFmtId` → the code string on `Style.numFmt`
  (omit when id 0 / "General").
- `fonts` → `Font[]` (invert `renderFont`: `<b>/<i>/<strike>/<u val?>/<sz>/<color>/<name>/<family>/
  <scheme>/<vertAlign>`), `fills` → `Fill[]` (patternFill fg/bg, gradientFill angle/path),
  `borders` → `Borders[]` (four sides + diagonal, `diagonalUp/Down`).
- Each `<xf numFmtId fontId fillId borderId>` with optional `<alignment>`/`<protection>` children →
  `Style { numFmt?, font?, fill?, border?, alignment?, protection? }`, pulling each sub-object from its
  pool **only when the matching `applyX` flag is set or the id is non-default** (mirror the writer: fontId
  0 = default font, so emit no `font` unless it differs; this keeps round-tripped styles structurally
  equal to what was written). Alignment `textRotation` 255 → `"vertical"`, 91–180 → `90 - r`; reading
  order 1/2 → `"ltr"`/`"rtl"`.
- `isDateStyle(index)`: resolve the xf's numFmt code and test it with a shared `isDateFormatCode(code)`
  helper — true when the code contains date/time tokens (`y m d h s`, and `[h]`/`AM/PM`) outside quoted
  literals and colour/condition brackets. Built-in ids 14–22 and 45–47 are date/time; treat those as date
  without string-sniffing. This is exactly the ambiguity ExcelJS resolves the same way.

> Note: because Stage 1.2 baked row/column/cell precedence into each cell's `style`, the reader hangs the
> composed `Style` straight on the cell — it does **not** try to re-derive row/column defaults from the
> file. Column/row `s` attributes are applied to `Column.style`/`Row.style` respectively so the model and
> a subsequent re-write stay consistent.

## 6. `src/xlsx/read/worksheet.ts` — the grid

```ts
export interface WorksheetReadContext {
  sharedStrings: (string | RichText)[];
  styles: StyleTable;
  date1904: boolean;
}
/** Populate an already-created Worksheet from its sheetN.xml. */
export function parseWorksheet(ws: Worksheet, xml: string, ctx: WorksheetReadContext): void;
```

- `<cols><col min max width customWidth style hidden outlineLevel>` → for `n` in `[min,max]`,
  `ws.getColumn(n)` set width/hidden/outlineLevel and `.style = styles.styleFor(style)`.
- `<sheetData><row r ht customHeight hidden outlineLevel s customFormat>` → `ws.getRow(r)` height/hidden/
  outlineLevel and `.style` from `s`.
- `<c r s t>` value decoding by `t` (inverse of the writer's mapping):

  | `t` | source | becomes |
  |-----|--------|---------|
  | `s` | `<v>` = index | `sharedStrings[index]` (string or RichText) |
  | `inlineStr` | `<is><t>` (or `<r>` runs) | string / RichText |
  | `str` | `<v>` text | string (a formula string result) |
  | `b` | `<v>` 0/1 | boolean |
  | `e` | `<v>` code | `{ error }` |
  | `n` / absent | `<v>` number | **number**, or **Date** via `serialToDate` when `styles.isDateStyle(s)` |

  A `<f>` present → `CellFormula { formula: "=" + expr (or bare — match the writer, which stripped the
  leading `=`), result: the decoded `<v>` typed as above }`. Set `cell.style = styles.styleFor(s)` for
  every non-default `s`. Skip cells that are empty and unstyled.
- `<mergeCells><mergeCell ref>` → `ws.mergeCells(ref)` (after cells are placed, so masters keep values).

## 7. `src/xlsx/read/workbook.ts` + `props.ts`

```ts
export interface WorkbookHeader {
  sheets: { name: string; sheetId: number; state: "visible" | "hidden" | "veryHidden"; rId: string }[];
  date1904: boolean;
  activeTab?: number;
}
export function parseWorkbook(xml: string): WorkbookHeader;      // xl/workbook.xml
export function applyCoreProps(wb: Workbook, xml: string): void; // docProps/core.xml (W3CDTF -> Date)
export function applyAppProps(wb: Workbook, xml: string): void;  // Company/Manager
```

Sheet order follows `<sheet>` document order; each `r:id` resolves through `workbookRels()` to the
worksheet part. `<workbookPr date1904>` and `<bookViews><workbookView activeTab>` are read here. Core
dates parse from W3CDTF back into `Date`; unset elements leave the model default.

## 8. `src/xlsx/read-workbook.ts` + public API

```ts
export interface ReadOptions { /* reserved (e.g. ignoreStyles?) — none required for v1 */ }
export function readWorkbookBuffer(bytes: Uint8Array, options?: ReadOptions): Workbook;
export async function readWorkbookFile(path: string, options?: ReadOptions): Promise<Workbook>;
```

Orchestration (mirrors §2 write order in reverse):
1. `openPackage(bytes)`; resolve workbook + rels.
2. `parseSharedStrings` (if present), `parseStyles`.
3. `parseWorkbook` for the sheet list + `date1904`; for each sheet, create it via `wb.addWorksheet(name,
   { state })`, then `parseWorksheet` with the shared context.
4. `applyCoreProps` / `applyAppProps`.
5. Return the assembled `Workbook`.

Public surface — extend the `workbook.xlsx` accessor (src/index.ts) so it reads *into* the receiver,
matching ExcelJS (`load`/`read`/`readFile` mutate and return the workbook):

```ts
workbook.xlsx.load(buffer, options?): Promise<Workbook>   // parse then copy via `workbook.model = parsed.model`
workbook.xlsx.readFile(path, options?): Promise<Workbook> // Node-only wrapper (fs/promises.readFile)
workbook.xlsx.read(stream, options?): Promise<Workbook>   // Node Readable -> buffer -> load
```

Keep the buffer path environment-agnostic; `readFile`/`read` are the only Node-coupled pieces — isolate
them beside `writeWorkbookFile` so browser bundles stay `fs`-free.

## 9. Test & verification plan

Unit:
- **xml-reader:** entity decode (named + numeric), self-closing, CDATA, nested elements, attribute quoting,
  whitespace preservation, the declaration/comment skips.
- **shared-strings / styles / workbook / props:** golden fragments that are the *inverse* of the Stage 1.3
  writer's golden strings (assert the reconstructed model objects). Styles: built-in date-id detection,
  custom numFmt ≥ 164, identical-font dedup restored, alignment/rotation inversion.
- **worksheet:** each `t` mapping, date-vs-number via numFmt, formula + cached result, merges, cols.

Integration — the real proof:
1. **Round-trip** (`read-roundtrip.test.ts`): the Stage 1.3 fixture → `writeWorkbookBuffer` →
   `readWorkbookBuffer` → deep-assert values, styles, numFmts, merges, sheet name/order/state, `date1904`,
   creator/created/modified. Add an inline-strings variant (`useSharedStrings: false`) and a styles-off
   variant.
2. **Foreign file** (`read-foreign.test.ts`, guarded on the `exceljs` devDep — `ctx.skip()` when absent):
   build a workbook with ExcelJS, `writeBuffer`, then `readWorkbookBuffer` ours and assert values + a merge
   + a numFmt survive. This catches producer quirks our own writer never exercises (shared `<f>` groups,
   `spans`, differing part names, `t="str"`, whitespace runs).
3. **Errors:** truncated zip, missing workbook part, `t="b"` with a non-numeric `<v>` → typed
   `AloshaXlsxError` subclasses, asserted.

## 10. Deferred to v0.3 / later

Streaming `WorkbookReader`, drawings/images, tables, pivot tables, conditional formatting, data validation,
comments, hyperlink relationships (rebuild `Hyperlink` from the rels rather than reading the display
string), protection, calcChain, formula recalculation, CSV read, and the `@alosha/xlsx/compat` shim (which
this stage unblocks but does not implement).

## Appendix — session breakdown & model recommendations

1. **Parser + package discovery** — `xml/xml-reader.ts` + `read/package.ts` + tests. Mechanical, high
   reuse downstream; get entity/whitespace handling right. **Sonnet 5.**
2. **sharedStrings + props** — `read/shared-strings.ts`, `read/props.ts` + tests. Straightforward inverses.
   **Sonnet 5.**
3. **Styles reader** — `read/styles.ts` (inverse pools, numFmt reverse map, date-format detection) + tests.
   Fiddly, high blast-radius; must stay structurally symmetric with the Stage 1.3 styles builder. **Opus 4.8.**
4. **Worksheet reader** — `read/worksheet.ts` (t-mapping, date-vs-number, formulas, merges, cols) + tests.
   Some judgment on date detection. **Sonnet 5 (Opus if desired).**
5. **Orchestrator + public API + round-trip/foreign verification** — `read-workbook.ts`, wire
   `workbook.xlsx.load/readFile/read`, integration tests. Where subtle wiring bugs surface. **Opus 4.8.**

---

*Grounded in `OOXML-NOTES.md` §3–§6 and the completed Stage 1.2 model + Stage 1.3 writer. Clean-room; no
ExcelJS code copied — only the OOXML shapes it reads/writes.*
