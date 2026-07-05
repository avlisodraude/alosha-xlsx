# Stage 1.3 Build Spec — The XLSX Writer (model → `.xlsx`)

> **Input for the build sessions.** Turns [`OOXML-NOTES.md`](./OOXML-NOTES.md) §2–§8 into concrete module
> boundaries, XML shapes, and a test plan for serializing the Stage 1.2 in-memory model into a valid,
> Excel-openable `.xlsx`. This is the WRITE-first payoff: `Workbook` → OOXML parts → `zip-layer` →
> `Uint8Array`. The **read** path and streaming writer are out of scope.
>
> Prereqs (all merged): `src/model/*` (HEAD 174eb01), `src/address/*`, `src/zip-layer.ts`
> (`writeZipArchive(entries: Map<string, Uint8Array>): Uint8Array`), `src/errors.ts`.
> Conventions: TS strict, ESM, `Uint8Array`/`TextEncoder` over Node `Buffer`, no `console.*`, pure
> functions where practical.

## 0. Scope

**In:** produce a valid single-`.xlsx` for a workbook containing sheets, rows, cells (all value types
from `cell-value.ts`), cell/row/column styles, number formats, merged cells, column widths, and basic
sheet/workbook metadata. Emit shared strings and a deduplicated styles part. Public API to get a
`Uint8Array` (and a thin Node file writer).

**Out (defer):** reading `.xlsx`; the streaming `WorkbookWriter`; drawings/images, tables, pivot tables,
conditional formatting, data validation, comments, hyperlinks-as-relationships, protection, calcChain.
(Hyperlink *values* still render as inline cell text now; the external-relationship wiring is Stage 2.)

**Definition of done:** `writeBuffer(wb)` returns bytes that (a) unzip to the exact part set in §5,
(b) re-parse as well-formed XML, (c) round-trip every cell value/style through a generate→unzip→assert
test, and (d) **open in Excel / LibreOffice without a repair prompt**. Verification detail in §8.

## 1. Module layout

```
src/
  xml/
    xml-writer.ts        // tiny, allocation-light XML string builder + attribute/text escaping
    index.ts
  xlsx/
    shared-strings.ts    // SharedStrings builder: intern string/richtext -> index; render sst xml
    styles.ts            // Styles builder: dedup fonts/fills/borders/numFmts -> cellXfs; render styles.xml
    date.ts              // Date <-> Excel serial (1900 w/ leap bug, and 1904)
    parts/
      content-types.ts   // [Content_Types].xml
      root-rels.ts       // _rels/.rels
      workbook.ts        // xl/workbook.xml
      workbook-rels.ts   // xl/_rels/workbook.xml.rels
      worksheet.ts       // xl/worksheets/sheetN.xml  (the grid)
      core-props.ts      // docProps/core.xml
      app-props.ts       // docProps/app.xml
      theme.ts           // xl/theme/theme1.xml  (fixed default, a constant)
    write-workbook.ts    // orchestrator: model -> Map<path,bytes> -> zip-layer -> Uint8Array
    index.ts
  index.ts               // add the public write entry (see §7)
tests/
  xml/xml-writer.test.ts
  xlsx/shared-strings.test.ts
  xlsx/styles.test.ts
  xlsx/date.test.ts
  xlsx/parts/*.test.ts
  xlsx/write-workbook.test.ts     // integration + round-trip
```

Build order = session order (§ appendix): xml/date/shared-strings/styles first, then the simple parts,
then the worksheet writer, then the orchestrator + public API + verification.

## 2. `src/xml/xml-writer.ts` — the XML builder

Hand-rolled and lean (avoid a heavy XML dep; fflate is our only runtime dep and stays that way). A string
buffer with escaping is enough; no DOM.

```ts
export class XmlWriter {
  constructor(opts?: { declaration?: boolean }); // default emits <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  open(tag: string, attrs?: Record<string, string | number | boolean | undefined>): this; // omits undefined attrs
  /** self-closing element with attrs only */
  leaf(tag: string, attrs?: Record<string, string | number | boolean | undefined>): this;
  /** element with escaped text content: <t>...</t> */
  el(tag: string, text: string, attrs?: Record<string, ...>): this;
  text(value: string): this;   // escapes & < > and, in attributes, " '
  close(tag: string): this;
  toString(): string;
  toBytes(): Uint8Array;       // TextEncoder
}

export function escapeXmlText(s: string): string;   // & < >  (and \r\n normalization if needed)
export function escapeXmlAttr(s: string): string;   // adds " '
```

Requirements: escape `& < >` in text and additionally `" '` in attribute values; drop `undefined`/`null`
attributes (so optional props don't emit empty attrs); preserve significant whitespace in shared strings
by emitting `xml:space="preserve"` on `<t>` when the string has leading/trailing spaces.

## 3. `src/xlsx/shared-strings.ts` — the string pool

```ts
export interface SharedStrings {
  /** Intern a plain string or rich text; returns its 0-based index. Repeated values return the same index. */
  add(value: string | RichText): number;
  readonly count: number;        // total references added
  readonly uniqueCount: number;  // distinct entries
  /** Render xl/sharedStrings.xml, or undefined if empty. */
  toXml(): string | undefined;
}
export function createSharedStrings(): SharedStrings;
```

- Dedup by key: plain strings key on the string; rich text keys on a stable serialization of its runs.
- `<si>` renders `<t>` for plain, `<r><rPr/><t/></r>` runs for rich text (font props via the shared
  font renderer from §4 where practical).
- Only used when `useSharedStrings` is on (default **true** for the buffered writer — see §7). When off,
  string cells render inline (`t="inlineStr"`) and this module isn't invoked.

## 4. `src/xlsx/styles.ts` — the dedup pools (the fiddly one)

This mirrors OOXML-NOTES §5. Cells reference a **cellXf index**; the builder interns the sub-objects.

```ts
export interface Styles {
  /** Intern a resolved cell Style; returns its cellXf index (the value for a cell's s="" attr). */
  addStyle(style: Style | undefined): number;
  toXml(): string;   // xl/styles.xml
}
export function createStyles(): Styles;
```

Internal pools, each with the mandatory Excel baseline seeds pre-inserted (getting these wrong triggers a
repair prompt):
- `numFmts`: built-in codes map to reserved ids; **custom codes allocate from id 164 upward**.
- `fonts[0]` = `{ sz 11, color theme=1, name 'Calibri', family 2, scheme 'minor' }`.
- `fills[0]` = `patternFill none`, `fills[1]` = `patternFill gray125` (both always present, in this order).
- `borders[0]` = empty border (`<border><left/><right/><top/><bottom/><diagonal/></border>`).
- `cellStyleXfs` = one default `<xf>` (id 0); `cellXfs` starts with a default and grows as styles intern.

`addStyle` decomposes a resolved `Style` into `{ numFmtId, fontId, fillId, borderId, alignment?, applyX flags }`,
interns each piece (structural dedup — identical font → same fontId), then interns the composed `<xf>` and
returns its `cellXfs` index. Cache by a stable string key per pool; optionally a `WeakMap<Style, number>`
fast path for reused style objects.

Renderers to implement (pure): `renderFont(f)`, `renderFill(f)`, `renderBorder(b)`, `renderNumFmt(id,code)`,
`renderXf(xf)`. Colors: `<color rgb="FFRRGGBB"/>` for argb, `<color theme="n"/>` for theme, `indexed` for
legacy. Alignment/protection ride inside `<xf>` with the `applyAlignment`/`applyProtection` flags.

> Note the model already resolved row/column/cell precedence into each cell's `style` (Stage 1.2), so the
> styles builder only dedups final styles — it does **not** re-run precedence.

## 5. `src/xlsx/parts/*` — the part writers & exact archive

Target archive (must match for a clean open — paths are the `Map` keys handed to `writeZipArchive`):

```
[Content_Types].xml
_rels/.rels
docProps/core.xml
docProps/app.xml
xl/workbook.xml
xl/_rels/workbook.xml.rels
xl/styles.xml
xl/sharedStrings.xml           (only if useSharedStrings and non-empty)
xl/theme/theme1.xml
xl/worksheets/sheet1.xml
xl/worksheets/sheet2.xml ...
```

**`[Content_Types].xml`** — `<Default>` for `rels` and `xml` extensions, `<Override>` for each part
(workbook, each worksheet, styles, sharedStrings, theme, core, app). Add the sharedStrings override only
when that part is emitted.

**`_rels/.rels`** — three relationships: officeDocument → `xl/workbook.xml`, core-properties →
`docProps/core.xml`, extended-properties → `docProps/app.xml`. Use the rel-type URIs from
OOXML-NOTES (`rel-type` table); rIds are local per `.rels` file.

**`xl/workbook.xml`** — `<sheets>` with one `<sheet name= sheetId= state= r:id=/>` per worksheet
(the `r:id` resolves via workbook rels), `<workbookPr date1904=/>` when set, and `<bookViews>` if present.
Emit the `xmlns:r` relationship namespace on the root.

**`xl/_rels/workbook.xml.rels`** — rId → each `worksheets/sheetN.xml`, plus styles, sharedStrings (if
present), theme. **Worksheet rIds here must match the `r:id`s in workbook.xml** (a common bug — assign
ids in one pass and share them).

**`xl/worksheets/sheetN.xml`** — the grid (OOXML-NOTES §3):
```
<worksheet xmlns=... xmlns:r=...>
  <dimension ref="A1:F12"/>
  <sheetViews><sheetView .../></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>   <col min max width customWidth style hidden/>   </cols>   (from Column.toModel; omit if none)
  <sheetData>
    <row r="1" ht customHeight s hidden outlineLevel>
      <c r="A1" t="s" s="3"><v>0</v></c>          <!-- shared string: v = index -->
      <c r="B1" t="inlineStr"><is><t>hi</t></is></c>  <!-- when useSharedStrings=false -->
      <c r="C1" t="n" s="2"><v>42</v></c>          <!-- number; also dates as serial + numFmt -->
      <c r="D1" t="b"><v>1</v></c>                 <!-- boolean -->
      <c r="E1" t="e"><v>#REF!</v></c>            <!-- error -->
      <c r="F1"><f>A1+C1</f><v>42</v></c>          <!-- formula + cached result -->
    </row>
  </sheetData>
  <mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>
</worksheet>
```
Cell `t` mapping from `StoredValue`/`effectiveType`: string→`s` (or `inlineStr`), number→`n` (default,
attr may be omitted), boolean→`b`, error→`e`, date→`n` with a date `numFmt` (serial via §6),
formula→`<f>`+`<v>`, richtext→shared string, merge slave→emit an empty `<c>` (no `<v>`) carrying only its
style, null→skip. Rows/cells sorted ascending; `<c r>` explicit.

**`docProps/core.xml`** — Dublin Core: `dc:creator`, `cp:lastModifiedBy`, `dcterms:created`/`modified`
(ISO 8601, `xsi:type="dcterms:W3CDTF"`), `dc:title`/`subject`, `cp:keywords`, `dc:description`,
`cp:category`.

**`docProps/app.xml`** — `Application`, `Company`, `Manager`, and `TitlesOfParts`/`HeadingPairs` listing
the sheet names.

**`xl/theme/theme1.xml`** — ship a fixed, known-good default theme as a string constant (ExcelJS does the
same). No need to synthesize.

## 6. `src/xlsx/date.ts` — serial conversion

```ts
export function dateToSerial(d: Date, date1904?: boolean): number;
export function serialToDate(serial: number, date1904?: boolean): Date;
```
1900 system epoch is 1899-12-31 with Excel's **intentional 1900-leap-year bug** (serial 60 = the
non-existent 1900-02-29); account for it. 1904 system epoch is 1904-01-01. Prefer a small self-contained
implementation (no `dayjs`); test against known anchors (1 = 1900-01-01, 60 boundary, 2 = 1900-01-02,
a modern date, and the 1904 offset of 1462 days).

## 7. `src/xlsx/write-workbook.ts` + public API

```ts
export interface WriteOptions { useSharedStrings?: boolean; useStyles?: boolean; } // defaults: true, true
export function writeWorkbookBuffer(wb: Workbook, options?: WriteOptions): Uint8Array;
```
Orchestration (single pass, mirrors OOXML-NOTES §2 write order):
1. Assign sheet ids/rIds once; create `SharedStrings` + `Styles` builders.
2. Render each `worksheets/sheetN.xml` (this populates the shared-strings and styles pools as cells emit).
3. Render `sharedStrings.xml` (after worksheets), `styles.xml`, `theme1.xml`, `workbook.xml` + rels,
   content-types, root rels, core + app props.
4. Collect all into `Map<string, Uint8Array>` and hand to `writeZipArchive` → `Uint8Array`.

Public surface (mirror ExcelJS ergonomics for the future compat shim, but thin):

```ts
// src/index.ts — attach to Workbook via a lazy `xlsx` accessor, OR expose free functions. Recommend
// BOTH: free functions as the core, and a convenience Workbook method.
workbook.xlsx.writeBuffer(options?): Promise<Uint8Array>   // async to match exceljs; wraps the sync core
workbook.xlsx.writeFile(path, options?): Promise<void>     // Node-only thin wrapper (fs/promises.writeFile)
```
Keep the buffer path environment-agnostic (no `fs`); `writeFile` is the only Node-coupled piece — isolate
it so browser bundles don't pull `fs`.

## 8. Test & verification plan

Unit:
- **xml-writer:** escaping (`& < > " '`), undefined-attr omission, `xml:space="preserve"` trigger, nesting.
- **date:** the anchors in §6, both epochs, round-trip `serialToDate(dateToSerial(d))`.
- **shared-strings:** dedup (same string → same index), count vs uniqueCount, rich-text keying, empty → undefined.
- **styles:** baseline seeds present at the right indices; custom numFmt starts at 164; identical fonts
  dedup to one id; `addStyle` returns stable indices; alignment/fill/border render correctly.
- **each part writer:** golden-string assertions on a small fixture workbook.

Integration / round-trip (`write-workbook.test.ts`) — the real proof:
1. Build a fixture workbook exercising every value type, a merge, a number format, a styled cell, two sheets.
2. `writeWorkbookBuffer` → unzip with `unzipSync` (fflate) → assert the exact §5 part set exists.
3. Parse each XML (a lightweight parser or regex assertions) → assert cells/merges/styles are present.
4. **Strongest check:** add `exceljs` as a **devDependency** and re-open the buffer with
   `new ExcelJS.Workbook().xlsx.load(buf)`; assert sheet names, a sampling of cell values, a merged range,
   and a number format survive. (exceljs is not currently installed — `pnpm add -D exceljs`.) This catches
   structural mistakes that pure XML assertions miss.
5. Optional CI nicety: shell out to headless LibreOffice to confirm no-repair open (skip if unavailable).

Keep the fixture small and deterministic (fixed `created`/`modified` dates) so golden strings are stable.

## 9. Deferred to Stage 2 / v0.2

Read path (`load`/`read`/`readFile`), streaming `WorkbookWriter`/`WorkbookReader`, drawings/images,
tables, pivot tables, conditional formatting, data validation, comments, hyperlink relationships,
protection, calcChain, and the `@alosha/xlsx/compat` shim (Stage 3).

## Appendix — session breakdown & model recommendations

1. **Foundations** — `xml/xml-writer.ts`, `xlsx/date.ts`, `xlsx/shared-strings.ts` + tests. Mechanical,
   well-specified. **Sonnet 5.**
2. **Styles builder** — `xlsx/styles.ts` (dedup pools, baseline seeds, numFmt 164, renderers) + tests.
   Fiddly, high blast-radius if seeds are wrong. **Opus 4.8.**
3. **Simple parts** — `parts/content-types|root-rels|workbook|workbook-rels|core-props|app-props|theme`
   + tests (shared rId assignment is the one gotcha). **Sonnet 5.**
4. **Worksheet writer** — `parts/worksheet.ts` (cell `t` mapping, cols, mergeCells, dates) + tests.
   Some judgment on the type mapping. **Sonnet 5 (Opus if desired).**
5. **Orchestrator + public API + round-trip/exceljs verification** — `xlsx/write-workbook.ts`, wire
   `workbook.xlsx.writeBuffer/writeFile`, integration tests, add exceljs devDep. **Opus 4.8** (integration
   + verification is where subtle wiring bugs surface).

---

*Grounded in `OOXML-NOTES.md` and the completed Stage 1.2 model (HEAD 174eb01). Clean-room; no ExcelJS code copied.*
