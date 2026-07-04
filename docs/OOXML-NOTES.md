# OOXML & ExcelJS I/O Notes — Stage 1.1

> **Purpose.** Everything `@alosha/xlsx` needs to know about the `.xlsx` file format (Office Open XML,
> SpreadsheetML) and how ExcelJS reads/writes it — so Stage 1.2+ can build the serializer without
> re-reading upstream. Companion doc: [`API-AUDIT.md`](./API-AUDIT.md) covers the in-memory model.
>
> Audited against exceljs `4.4.0`.

## 1. What a `.xlsx` actually is

A `.xlsx` file is a **ZIP archive** (the OPC — Open Packaging Conventions container) of XML parts.
Rename any `.xlsx` to `.zip` and unzip to inspect. The spec is ECMA-376 / ISO/IEC 29500; the relevant
dialect is **SpreadsheetML**.

Two cross-cutting mechanisms hold the package together:

- **Content types** — `[Content_Types].xml` at the root declares the MIME type of every part (by
  extension default and by explicit override). A part not declared here is invalid.
- **Relationships** — `.rels` files (in `_rels/` folders) wire parts together by opaque `rId` tokens
  rather than by path. The package root has `_rels/.rels`; the workbook has `xl/_rels/workbook.xml.rels`;
  each worksheet that references external things has `xl/worksheets/_rels/sheetN.xml.rels`.

Everything is referenced indirectly: `workbook.xml` lists sheets with an `r:id`, and
`workbook.xml.rels` maps that `rId` → `worksheets/sheet1.xml`. **You never hardcode paths across parts —
you resolve through rels.** This indirection is the single most important thing to get right on both
read and write.

## 2. Canonical ZIP layout (as ExcelJS writes it)

Exact archive entries, taken from the `zip.append(..., {name})` calls in `lib/xlsx/xlsx.js`:

```
mybook.xlsx  (ZIP root)
├─ [Content_Types].xml                     ← part registry (MIME types)
├─ _rels/
│  └─ .rels                                ← root rels → officeDocument (xl/workbook.xml), core, app
├─ docProps/
│  ├─ core.xml                             ← Dublin-core metadata (creator, created, title…)
│  └─ app.xml                              ← app metadata (company, manager, sheet names…)
└─ xl/
   ├─ workbook.xml                         ← sheet list, defined names, views, calcPr, date1904
   ├─ _rels/
   │  └─ workbook.xml.rels                 ← rId → worksheets, styles, sharedStrings, theme, pivotCache…
   ├─ styles.xml                           ← the dedup'd style pools (see §5)
   ├─ sharedStrings.xml                    ← the string pool (optional, see §4)
   ├─ theme/theme1.xml                     ← color/font theme (ExcelJS ships a fixed default)
   ├─ calcChain.xml                        ← formula calc order (ExcelJS largely ignores on read)
   ├─ worksheets/
   │  ├─ sheet1.xml                        ← the grid: <cols>, <sheetData><row><c>, <mergeCells>…
   │  ├─ sheet2.xml
   │  └─ _rels/
   │     └─ sheet1.xml.rels                ← hyperlinks, drawings, tables, comments for that sheet
   ├─ drawings/
   │  ├─ drawing1.xml                      ← image/shape anchors
   │  ├─ vmlDrawing1.vml                   ← legacy VML (comments/notes markers)
   │  └─ _rels/drawing1.xml.rels
   ├─ media/
   │  └─ image1.png|jpeg|gif…              ← embedded binaries
   ├─ tables/
   │  └─ table1.xml                        ← structured table defs
   ├─ comments1.xml                        ← threaded/legacy comments per sheet
   └─ pivotCache/ , pivotTables/           ← pivot parts (experimental in ExcelJS)
```

Write **order matters** for streaming and for some readers: ExcelJS writes content-types → rels →
worksheets → sharedStrings (always after worksheets, because writing cells populates the string pool) →
drawings → tables → pivots → themes+styles → media → app+core → workbook, then finalizes the ZIP.

## 3. The key XML parts, briefly

### `xl/workbook.xml`
Lists `<sheets><sheet name= sheetId= state= r:id=/></sheets>`, plus `<definedNames>`, `<bookViews>`,
`<calcPr>`, and `<workbookPr date1904=/>`. The `r:id` resolves via `workbook.xml.rels`. `sheetId` is
stable identity; tab order is the element order.

### `xl/worksheets/sheetN.xml`
The grid. Shape:
```xml
<worksheet>
  <sheetPr/> <dimension ref="A1:F12"/>
  <sheetViews>…freeze/split, zoom…</sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols><col min="1" max="1" width="20" customWidth="1" style="3"/></cols>
  <sheetData>
    <row r="1" ht="18" customHeight="1" s="2">
      <c r="A1" t="s" s="1"><v>0</v></c>       <!-- t="s": v is index into sharedStrings -->
      <c r="B1" t="n"><v>42</v></c>            <!-- number -->
      <c r="C1"><f>A1+B1</f><v>42</v></c>      <!-- formula + cached result -->
      <c r="D1" t="str"><f>…</f><v>text</v></c><!-- t="b" boolean, t="e" error, t="inlineStr" -->
    </row>
  </sheetData>
  <mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>
  <hyperlinks/> <dataValidations/> <conditionalFormatting/> <pageSetup/> <drawing r:id=/>
</worksheet>
```
Cell `t` attribute encodes type: `s`=shared string, `str`=formula string, `n`=number (default),
`b`=boolean, `e`=error, `inlineStr`=inline string, `d`=ISO date (rare). Dates are normally **numbers**
with a date `numFmt`, not a distinct type — the "1900/1904 epoch serial" convention. `s` attribute on
`<c>`/`<row>`/`<col>` is the **cellXf index** into `styles.xml`.

### `xl/sharedStrings.xml`
`<sst count= uniqueCount=>` of `<si>` (string items), each plain `<t>` or rich `<r><rPr/><t/></r>`
runs. Cells with `t="s"` store the **index** here, deduplicating repeated strings across the workbook.

### `xl/styles.xml`
The pooled style tables — see §5.

### `docProps/core.xml` & `app.xml`
Metadata. `core.xml` is Dublin Core (`dc:creator`, `dcterms:created`…); `app.xml` is extended
(`Company`, `Manager`, `TitlesOfParts` = sheet names).

## 4. Shared strings vs inline strings

`useSharedStrings` (write option, default **true** in the buffered writer, **false** in the streaming
writer). When on, every string cell becomes an index into `sharedStrings.xml`; the pool is built while
worksheets serialize, which is why shared strings are written *after* worksheets. When off, strings are
inlined as `t="inlineStr"` — larger files but no second pass and lower memory for write-once streams.
The rewrite should keep this a per-write toggle with the same semantics.

## 5. Styles: the dedup pools (`styles.xml`)

`styles.xml` is a set of indexed pools, and cells/rows/columns reference entries by integer id:

```xml
<styleSheet>
  <numFmts>   <numFmt numFmtId="164" formatCode="0.00%"/> </numFmts>   <!-- custom fmts, id ≥ 164 -->
  <fonts>     <font><sz val="11"/><name val="Calibri"/>…</font> </fonts>
  <fills>     <fill><patternFill patternType="none"/></fill> … </fills>
  <borders>   <border><left/><right/>…</border> </borders>
  <cellXfs>   <xf numFmtId="164" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1">
                <alignment horizontal="center"/></xf> </cellXfs>
</styleSheet>
```

ExcelJS's `StylesXform` (`lib/xlsx/xform/style/styles-xform.js`) builds these on write:
- Pre-seeds `borders[0]={}`, `fills[0]=none`, `fills[1]=gray125`, `fonts[0]=Calibri 11` (Excel requires
  these baseline entries).
- Built-in number formats map to their reserved IDs; custom codes allocate from **164** upward.
- `addStyleModel(model, cellType)` returns a `cellXf` index. It caches by `WeakMap<styleObject, id>`
  (fast path for the common "same style object reused across many cells") **and** by serialized-XML
  string per sub-pool (structural dedup). This is what keeps files small — many cells → one `xf`.
- On read, the reverse: `<xf>` index → reconstructed `Style` object hung on the cell model.

**Reproduce faithfully:** the baseline seed entries, the 164 custom-numfmt base, and the row-over-column
style precedence (resolved in the model, see API audit §4). Getting the seed indices wrong makes Excel
repair the file on open.

## 6. How ExcelJS parses & generates — architecture

### The "xform" pattern
Every XML part has a transformer extending `BaseXform` (`lib/xlsx/xform/base-xform.js`, 118 xform files
total) implementing a small virtual interface:

```
prepare(model, options)   // mutate model to be write-ready (assign ids, filenames, style indices)
render(xmlStream, model)  // model → XML (push-based, writes into an XmlStream buffer)
parseOpen(node)           // SAX: element opened
parseText(text)           // SAX: text chunk
parseClose(name)          // SAX: element closed → returns true when this xform's subtree is done
reconcile(model, options) // post-parse fix-up (resolve rels, merge drawings/hyperlinks, styles → cells)
reset()                   // clear state between reuse (xforms are reused across many nodes)
```

Xforms compose: a `WorksheetXform` owns child xforms for `<cols>`, `<row>`, `<c>`, `<mergeCells>`, etc.,
each a `BaseXform`. This mirrors the XML tree and makes each part independently testable.

### Read pipeline (`workbook.xlsx.read/readFile/load`)
1. Open the ZIP. Buffered path uses **JSZip** (`load(buffer)` for browser/whole-file); file/stream path
   uses **unzipper** streaming entries.
2. Route each entry by path to its xform (`_processWorksheetEntry`, `_processStyles`,
   `_parseSharedStrings`, core/app, drawings, tables, comments…).
3. Parse XML via **saxes** (SAX, event-based) through each xform's `parseOpen/Text/Close`, producing a
   plain **model**.
4. `reconcile(model)` — second pass: resolve rels (hyperlink `rId`→URL, drawing→image, table→styles),
   attach shared strings and styles to cell models.
5. `workbook.model = model` rebuilds the live doc object graph.

### Write pipeline (`workbook.xlsx.write/writeFile/writeBuffer`)
1. `prepareModel(model, options)` — seed defaults (`creator='ExcelJS'`…), create a fresh
   `SharedStringsXform` and `StylesXform`, then `prepare` the workbook + each worksheet (assign table
   filenames/ids, drawing ids, and — as cells serialize — style indices and shared-string indices).
2. Render each part in the fixed order (§2) via `zip.append(xml, {name})`.
3. Finalize the ZIP.

### ZIP backends
- **Write:** `lib/utils/zip-stream.js` wraps **archiver** (`ZipStream.ZipWriter`) and pipes into any
  Node writable; `writeBuffer` pipes into an in-memory `StreamBuf`.
- **Read:** JSZip (whole-buffer) or unzipper (streaming entries) + `tmp` for spill files.
- **Browser:** `package.json` `browser` field points at a prebuilt `dist/exceljs.min.js`; JSZip handles
  the archive in-browser.

Runtime deps worth noting for the rewrite: `jszip`, `unzipper`, `archiver`, `saxes`, `readable-stream`,
`dayjs` (date serial math), `fast-csv`, `uuid`, `tmp`. That's a lot of surface — see §8.

## 7. Streaming API (large files)

ExcelJS ships a separate streaming path under `lib/stream/xlsx/` — this is the answer to "big files
without holding everything in memory":

- **`stream.xlsx.WorkbookWriter({ filename|stream, useSharedStrings, useStyles })`** — `addWorksheet()`
  returns a `WorksheetWriter`; you `row.commit()` / `worksheet.commit()` to flush rows to disk as you go,
  then `workbook.commit()` to finalize. Rows already committed are **not** retained. `useSharedStrings`
  defaults to **false** here (inline strings) to avoid the second pass.
- **`stream.xlsx.WorkbookReader(input, { worksheets:'emit', entries:'emit', sharedStrings:'cache' })`** —
  an `EventEmitter` / async-iterator (`for await (const worksheet of workbook)` → `for await (const row
  of worksheet)`), parsing entries incrementally so you never materialize the whole grid.

The doc `Worksheet.commit()` and `Row.commit()` are no-ops precisely so the same calling code works
against both the buffered and streaming worksheet.

## 8. Areas for improvement (design targets for `@alosha/xlsx`)

### Memory & large files
- **Per-cell Value-object allocation** (12 strategy classes, each with its own `model`) is the dominant
  cost on large sheets. Fold into a discriminated union + pure functions (see API audit §5). Target:
  a flat, columnar or typed-array-backed cell store for the hot numeric case.
- **Buffered write builds the whole ZIP in memory** before flush. Make **true streaming write** a
  first-class path (not a separate half-parallel `lib/stream` API with subtly different defaults). One
  worksheet abstraction, two drains: buffer or stream.
- **Shared-strings second pass** forces holding all worksheets before writing strings. Consider a
  single-pass inline-string default for write-heavy workloads, shared strings opt-in.

### Dependency footprint
- ExcelJS pulls **three** archive libs (`jszip`, `unzipper`, `archiver`) plus `saxes`, `readable-stream`,
  `dayjs`, `tmp`, `uuid`. A modern rewrite can target **one** small zip lib for both directions
  (roadmap picks `fflate` — tiny, streaming, works in Node + browser) and a lean SAX/pull parser, cutting
  install size and browser bundle substantially.
- Browser support via a prebuilt UMD bundle is dated. Ship **ESM + native `Uint8Array`/Web Streams**;
  let the bundler tree-shake.

### Correctness / robustness
- **`calcChain.xml`** is mostly ignored — fine for write (Excel recomputes with `fullCalcOnLoad`), but be
  deliberate about it.
- **Formula translation** (shared → per-cell) happens lazily on `cell.formula` read and mutates cached
  state — move to explicit, side-effect-free expansion.
- **Style seed/precedence** (baseline fills/fonts, numFmt id 164, row-over-column) must be pinned by
  golden-file tests against real Excel output, or files trigger "repair" on open.
- **Number/date epochs**: honor `workbookPr@date1904`; centralize serial↔Date conversion (ExcelJS uses
  `dayjs`) behind one utility with tests for the 1900 leap-year bug that Excel deliberately preserves.

### Scope for v0.1 (WRITE-first, per roadmap)
Minimum viable writer parts: `[Content_Types].xml`, `_rels/.rels`, `xl/workbook.xml` (+ rels),
`xl/worksheets/sheetN.xml`, `xl/styles.xml`, optional `xl/sharedStrings.xml`, `docProps/core.xml` +
`app.xml`, a fixed `theme1.xml`, packed with `fflate`. That alone produces a valid, Excel-openable file.
Defer drawings/media, tables, pivots, comments, conditional formatting, data validation, and the whole
**read** path to later stages.

---

*Sources: `lib/xlsx/xlsx.js`, `lib/xlsx/xform/**`, `lib/stream/xlsx/**`, `lib/xlsx/rel-type.js`,
`package.json` in exceljs 4.4.0. Format facts cross-checked against ECMA-376 SpreadsheetML. Clean-room
reimplementation — informed by, not derived from, ExcelJS source.*
