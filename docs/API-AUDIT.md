# ExcelJS API Audit — Stage 1.1

> **Purpose.** This document maps the public API surface of [ExcelJS](https://github.com/exceljs/exceljs)
> (MIT, v4.4.0 — the version audited) so that `@alosha/xlsx` can be designed as a modern TypeScript
> rewrite without re-reading the upstream source. It is a *reference of what exists*, plus opinionated
> notes on *what to keep, drop, or redesign*.
>
> Companion doc: [`OOXML-NOTES.md`](./OOXML-NOTES.md) covers the file format and I/O layer.

## How ExcelJS is structured

ExcelJS separates two concerns cleanly, and this separation is worth preserving:

1. **The in-memory document model** (`lib/doc/*`) — `Workbook`, `Worksheet`, `Row`, `Column`, `Cell`,
   `Range`, plus value/style helpers. This is a plain-ish object graph with getters/setters. No XML,
   no ZIP, no I/O.
2. **The serialization layer** (`lib/xlsx/*`, `lib/csv/*`, `lib/stream/*`) — "xform" transformers that
   convert the model to/from OOXML XML, and ZIP glue. Accessed only through `workbook.xlsx` / `workbook.csv`.

Every model object exposes a `get model()` / `set model(v)` pair. `model` is the serializable
snapshot; the doc classes are the ergonomic wrapper around it. The rewrite should adopt this same
"live object ⟷ plain model" split — it is the single cleanest idea in the codebase.

Cell values are represented by an internal **strategy/polymorphic Value class** per type
(`NullValue`, `NumberValue`, `StringValue`, `DateValue`, `HyperlinkValue`, `FormulaValue`,
`MergeValue`, `RichTextValue`, `BooleanValue`, `ErrorValue`, `SharedStringValue`, `JSONValue`).
`Cell` delegates `value`, `type`, `text`, `toString()` etc. to the current strategy. Type detection is
duck-typed in `Value.getType()`.

### The `ValueType` enum (public, on `Enums.ValueType` / `Cell.Types`)

| Name | # | Notes |
|------|---|-------|
| `Null` | 0 | empty cell |
| `Merge` | 1 | slave cell of a merge; delegates to master |
| `Number` | 2 | |
| `String` | 3 | plain text |
| `Date` | 4 | JS `Date`, stored as OOXML serial |
| `Hyperlink` | 5 | `{ text, hyperlink, tooltip? }` |
| `Formula` | 6 | `{ formula, result, ref?, shareType?, sharedFormula? }` |
| `SharedString` | 7 | index into shared-string table |
| `RichText` | 8 | `{ richText: [{ text, font? }] }` |
| `Boolean` | 9 | |
| `Error` | 10 | `{ error: '#REF!' … }` |

`FormulaType`: `None=0`, `Master=1`, `Shared=2`.

---

## Section 1 — Workbook public interface

`Workbook` (`lib/doc/workbook.js`, typed in `index.d.ts`) is the root aggregate. Construct with
`new ExcelJS.Workbook()` — the constructor takes **no arguments**; all metadata is set via properties.

### Metadata properties (plain read/write fields)

| Property | Type | Default | Notes |
|----------|------|---------|-------|
| `creator` | `string` | `'Unknown'` on serialize | author |
| `lastModifiedBy` | `string` | `'Unknown'` | |
| `created` | `Date` | `new Date()` | |
| `modified` | `Date` | `= created` | |
| `lastPrinted` | `Date` | — | |
| `title` | `string` | `''` | docProps/core.xml |
| `subject` | `string` | `''` | |
| `keywords` | `string` | `''` | |
| `category` | `string` | `''` | |
| `description` | `string` | `''` | |
| `company` | `string` | `''` | docProps/app.xml |
| `manager` | `string` | `''` | |
| `language` | `string` | — | |
| `revision` | — | — | |
| `contentStatus` | — | — | |
| `properties` | `WorkbookProperties` | `{}` | notably `date1904: boolean` |
| `calcProperties` | `CalculationProperties` | `{}` | e.g. `fullCalcOnLoad` |
| `views` | `WorkbookView[]` | `[]` | window layout, active tab |

### Worksheet management

| Member | Signature | Notes |
|--------|-----------|-------|
| `worksheets` | `get → Worksheet[]` | **clone**, sorted by `orderNo`, filtered of holes |
| `addWorksheet` | `(name?: string, options?: Partial<AddWorksheetOptions>) → Worksheet` | auto-names `sheetN`; assigns `id` from `nextId` |
| `getWorksheet` | `(idOrName?: number \| string) → Worksheet \| undefined` | no arg → first sheet |
| `removeWorksheet` | `(idOrName: number \| string) → void` | calls `worksheet.destroy()` |
| `removeWorksheetEx` | `(ws: Worksheet) → void` | low-level, deletes slot |
| `eachSheet` | `(cb: (ws, id) => void) → void` | preferred over `worksheets.forEach` |
| `nextId` | `get → number` | next free slot in sparse `_worksheets` |

> **Gotcha to fix:** `_worksheets` is a **1-based sparse array** (index 0 is always empty). `worksheets`
> getter does `.slice(1).sort().filter(Boolean)` on every access — O(n log n) each call. `addWorksheet`
> still carries a deprecated `tabColor`-as-string branch that emits `console.trace`. Drop both in the rewrite.

### I/O accessors (lazy singletons)

| Member | Type | Notes |
|--------|------|-------|
| `xlsx` | `get → Xlsx` | the OOXML reader/writer (see §I/O below) |
| `csv` | `get → Csv` | fast-csv based |

### Media & defined names

| Member | Signature | Notes |
|--------|-----------|-------|
| `addImage` | `(img: {base64\|filename\|buffer, extension}) → number` | returns media id |
| `getImage` | `(id: number) → Image` | |
| `definedNames` | `get → DefinedNames` | `.add(address, name)`, `.getRanges(name)`, `.remove(...)` |
| `clearThemes` | `() → void` | internal theme reset (undocumented) |
| `model` | `get/set WorkbookModel` | full serializable snapshot |

### Workbook I/O — `workbook.xlsx.*` (the `Xlsx` interface)

All async, all `Promise`-returning:

| Method | Signature |
|--------|-----------|
| `readFile` | `(path, options?) → Promise<Workbook>` |
| `read` | `(stream, options?) → Promise<Workbook>` (Node Readable) |
| `load` | `(buffer, options?) → Promise<Workbook>` (Buffer/ArrayBuffer — browser path) |
| `writeFile` | `(path, options?) → Promise<void>` |
| `write` | `(stream, options?) → Promise<void>` |
| `writeBuffer` | `(options?) → Promise<Buffer>` |

`XlsxReadOptions`: `{ ignoreNodes?, worksheets?, sheetRows?, styles?, hyperlinks?, ... }`.
`XlsxWriteOptions`: `{ useSharedStrings?, useStyles?, zip?: JSZipGeneratorOptions }`.
CSV mirror: `workbook.csv.readFile/read/writeFile/write/writeBuffer` with `{ dateFormats, map, parserOptions, formatterOptions, sheetName }`.

**~30 top-level Workbook members total.** The surface is small; the complexity lives in Worksheet/Cell/Style.

---

## Section 2 — Worksheet interface

`Worksheet` (`lib/doc/worksheet.js`) is the largest doc class (~950 LOC). Never constructed directly —
always via `workbook.addWorksheet()`.

### Identity & state

| Member | Type | Notes |
|--------|------|-------|
| `id` | `number` (readonly) | slot in workbook |
| `name` | `string` | **validated setter**: no `* ? : / \ [ ]`, no leading/trailing `'`, ≤31 chars (warns+truncates), `"History"` reserved, must be unique (case-insensitive) |
| `orderNo` | `number` | tab order |
| `state` | `'visible' \| 'hidden' \| 'veryHidden'` | |
| `workbook` | `Workbook` (readonly) | back-reference |
| `properties` | `WorksheetProperties` | `tabColor`, `defaultRowHeight=15`, `defaultColWidth?`, `outlineLevelCol/Row`, `dyDescent=55`, `showGridLines`, `outlineProperties` |
| `pageSetup` | `Partial<PageSetup>` | margins, orientation, scale, fitToWidth/Height, paperSize, printArea… (~30 fields) |
| `headerFooter` | `Partial<HeaderFooter>` | odd/even/first header & footer strings |
| `views` | `WorksheetView[]` | freeze/split panes, zoom, gridlines, RTL |
| `autoFilter` | `string \| {from,to}` | |
| `state`, `sheetProtection`, `conditionalFormattings`, `tables`, `pivotTables` | | feature stores |

### Dimensions & counts (all getters)

`dimensions → Range`, `rowCount`, `actualRowCount`, `columnCount`, `actualColumnCount`,
`lastRow`, `lastColumn`, `hasMerges`.

> `actualColumnCount` and `actualRowCount` are explicitly commented as "performance nightmare" — they
> scan every cell. Worth precomputing/caching in the rewrite.

### Columns

| Member | Signature | Notes |
|--------|-----------|-------|
| `columns` | `get/set Array<Partial<Column>>` | set from defns; headers written into row cells |
| `getColumn` | `(idxOrKey: number \| string) → Column` | creates gaps as needed; accepts letter or key |
| `lastColumn` | `get → Column` | |
| `spliceColumns` | `(start, count, ...inserts: any[][]) → void` | insert/remove columns; shifts cells + defined names |
| `getColumnKey` / `setColumnKey` / `deleteColumnKey` / `eachColumnKey` | key↔Column map for `addRow({obj})` convenience |

### Rows

| Member | Signature | Notes |
|--------|-----------|-------|
| `getRow` | `(r: number) → Row` | creates if absent |
| `getRows` | `(start, length) → Row[] \| undefined` | |
| `findRow` / `findRows` | non-creating lookup | |
| `addRow` | `(data: any[] \| object, style?: 'i'\|'o'\|'i+'\|'o+'\|'n') → Row` | appends at `_nextRow`; style copies from adjacent row |
| `addRows` | `(rows: any[], style?) → Row[]` | |
| `insertRow` | `(pos, value, style?) → Row` | |
| `insertRows` | `(pos, values, style?) → Row[]` | |
| `duplicateRow` | `(rowNum, count, insert=false) → void` | copies values+styles |
| `spliceRows` | `(start, count, ...inserts) → void` | the heavy one; also remerges cells & shifts defined names |
| `eachRow` | `(opts?{includeEmpty}, cb(row, rowNumber)) → void` | |
| `getSheetValues` | `() → RowValues[]` | sparse 1-based array of row values |
| `lastRow`, `rowCount`, `actualRowCount` | getters | |

> The `style` mini-DSL for `addRow`/`insertRow` (`'i'`=copy from row above, `'o'`=below, `+`=include
> empty cells) is cryptic and under-documented. In the rewrite prefer an explicit options object:
> `{ inheritStyleFrom: 'above' \| 'below', includeEmpty?: boolean }`.

### Cells & merges

| Member | Signature | Notes |
|--------|-----------|-------|
| `getCell` | `(r: number\|string, c?: number\|string) → Cell` | `getCell('B2')` or `getCell(2,3)`; creates |
| `findCell` | `(r, c) → Cell \| undefined` | non-creating |
| `mergeCells` | many overloads: `('A1:B2')`, `('A1','B2')`, `(t,l,b,r)`, `([...])` | throws if already merged |
| `mergeCellsWithoutStyle` | same overloads | skips style copy (used on import) |
| `unMergeCells` | same overloads | atomic per merge block |
| `fillFormula` | `(range, formula, results?, shareType='shared') → void` | shared/array formula fill |

### Features

Images: `addImage(imageId, range)`, `getImages()`, `addBackgroundImage(id)`, `getBackgroundImageId()`.
Tables: `addTable(props) → Table`, `getTable(name)`, `getTables()`, `removeTable(name)`.
Pivot (experimental, warns): `addPivotTable(model)`.
Conditional formatting: `addConditionalFormatting(cf)`, `removeConditionalFormatting(filterOrIndexOrFn)`.
Protection: `protect(password, options) → Promise<void>` (SHA-512, configurable `spinCount`), `unprotect()`.
Data validations: per-cell via `cell.dataValidation` (the sheet holds a `DataValidations` store).
Streaming hook: `commit()` (no-op on doc worksheet; real on stream writer).
`model` get/set — serializable `WorksheetModel`.

**~60+ Worksheet members.** This is the class most worth trimming and re-organizing (see §5).

---

## Section 3 — Row & Cell interface

### Row (`lib/doc/row.js`)

| Member | Type / Signature | Notes |
|--------|------------------|-------|
| `number` | `number` (readonly) | 1-based |
| `worksheet` | `Worksheet` (readonly) | |
| `values` | `get/set` | **sparse array**: contiguous array starts at col 1 (offset by presence of key `'0'`); or key'd object literal mapped via column keys. Setter is destructive (clears prior cells). |
| `getCell` | `(col: number \| string \| key) → Cell` | letter/number/column-key |
| `findCell` | `(col: number) → Cell \| undefined` | |
| `getCellEx` | `({address,row,col}) → Cell` | internal creating helper |
| `eachCell` | `(opts?{includeEmpty}, cb(cell, colNumber)) → void` | |
| `splice` | `(start, count, ...inserts) → void` | shifts cells, carries value+style+comment |
| `addPageBreak` | `(left?, right?) → void` | |
| `hasValues` | `get → boolean` | |
| `cellCount` / `actualCellCount` | getters | |
| `dimensions` | `get → {min,max} \| null` | |
| `height` | `number` | row height in points |
| `hidden` | `boolean` | |
| `outlineLevel` | `number` | + derived `collapsed` getter |
| `style` | `Style` | row-level default; applied to cells lazily |
| `numFmt` / `font` / `alignment` / `protection` / `border` / `fill` | style shortcuts | setter fans out to every existing cell via `_applyStyle` |
| `commit` | `() → void` | streaming-writer signal |
| `model` | get/set `RowModel` | |

### Column (`lib/doc/column.js`)

Column is a **projection over cells**, not a container. Key members:
`number`, `letter`, `key`, `width`, `hidden`, `outlineLevel`, `collapsed`, `style` + the six style
shortcuts (fan out to the whole column via `eachCell`), `header`/`headers` (writing a header sets the
top cells' values), `values` (get/set down the column), `eachCell(opts, cb)`, `defn` get/set, plus
statics `Column.toModel(columns)` / `Column.fromModel(ws, cols)` that **compress** runs of identical
columns into `<col min max>` ranges (and expand them back). `isDefault`/`equivalentTo`/`isCustomWidth`
drive that compression. Default width constant is `9`.

### Cell (`lib/doc/cell.js`, ~1100 LOC incl. value classes)

| Member | Type / Signature | Notes |
|--------|------------------|-------|
| `value` | `get/set CellValue` | polymorphic; setter picks a Value strategy via `Value.getType(v)`. Setting on a merge slave writes the master. |
| `type` | `ValueType` (readonly) | current strategy's type |
| `effectiveType` | `ValueType` (readonly) | for formulas, the type of `result` |
| `text` | `string` (readonly) | display text |
| `html` | `string` (readonly) | HTML-escaped text |
| `toString()` | `→ string` | |
| `toCsvString()` | `→ string` | per-type CSV rendering |
| `formula` | `string` (readonly) | translates shared formulas on read |
| `result` | any (readonly) | cached formula result |
| `formulaType` | `FormulaType` (readonly) | |
| `address` | `string` (readonly) | e.g. `'B2'` |
| `row` / `col` | `number` (readonly) | |
| `$col$row` | `string` (readonly) | absolute ref `'$B$2'` |
| `fullAddress` | `{sheetName,address,row,col}` (readonly) | |
| `master` | `Cell` (readonly) | self if unmerged |
| `isMerged` / `isMergedTo(m)` / `merge(master, ignoreStyle?)` / `unmerge()` | merge control | |
| `isHyperlink` / `hyperlink` | | |
| `note` | `get/set` (Comment) | wraps `Note` |
| `dataValidation` | `get/set DataValidation` | |
| `name` / `names` / `addName` / `removeName` / `removeAllNames` | defined-name membership | |
| `numFmt`,`font`,`alignment`,`border`,`fill`,`protection` | style shortcuts | read/write into `cell.style` |
| `style` | `Style` | merged from row+column defaults at construction |
| `model` | get/set `CellModel` | `{address, type, value/formula/..., style, comment?}` |

> **Design note.** The one-class-per-value-type strategy pattern (12 classes) is *elegant conceptually*
> but heavy: each cell allocates a Value wrapper object holding its own `model`. For 100k+ cells this is
> real memory. The rewrite can likely fold this into a **tagged union / discriminated `CellValue`** plus
> pure functions (`getEffectiveType`, `toText`, `toCsv`) — same behavior, far fewer allocations, better TS
> ergonomics.

---

## Section 4 — Style & Format objects

Styles are plain objects (`Style`), shared by reference across cells until serialization, where the
`StylesXform` **deduplicates** them into indexed pools (`<fonts>`, `<fills>`, `<borders>`, `<numFmts>`,
`<cellXfs>`) using a `WeakMap<model, styleId>` cache + string-keyed indices. That dedup is the crux of
small file size and is worth reproducing carefully.

### `Style`

```ts
interface Style {
  numFmt: string;                 // e.g. '0.00%', 'dd/mm/yyyy'
  font: Partial<Font>;
  alignment: Partial<Alignment>;
  protection: Partial<Protection>;
  border: Partial<Borders>;
  fill: Fill;                     // union — NOT Partial
}
```

### `Font`

```ts
interface Font {
  name: string; size: number; family: number;
  scheme: 'minor' | 'major' | 'none';
  charset: number;
  color: Partial<Color>;
  bold: boolean; italic: boolean;
  underline: boolean | 'none' | 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';
  vertAlign: 'superscript' | 'subscript';
  strike: boolean; outline: boolean;
}
```
Default font on serialize: `{ size: 11, color: {theme:1}, name: 'Calibri', family: 2, scheme: 'minor' }`.

### `Color`

```ts
interface Color { argb: string; /* 'FFRRGGBB' */  theme: number; /* theme index */ }
```
Also supports `indexed` (legacy palette) in practice. Colors appear as `Partial<Color>` almost everywhere.

### `Fill` (discriminated union)

```ts
type Fill = FillPattern | FillGradientAngle | FillGradientPath;

interface FillPattern { type: 'pattern'; pattern: FillPatterns; fgColor?: Partial<Color>; bgColor?: Partial<Color>; }
interface FillGradientAngle { type: 'gradient'; gradient: 'angle'; degree: number; stops: GradientStop[]; }
interface FillGradientPath  { type: 'gradient'; gradient: 'path';  center: {left:number;top:number}; stops: GradientStop[]; }
interface GradientStop { position: number; color: Partial<Color>; }
```
`FillPatterns` = `'none' | 'solid' | 'darkVertical' | … | 'gray125' | 'gray0625'` (18 values).
Two fills are always pre-seeded in the pool: pattern `none` (id 0) and `gray125` (id 1).

### `Border` / `Borders`

```ts
type BorderStyle = 'thin'|'dotted'|'hair'|'medium'|'double'|'thick'|'dashed'
                 | 'dashDot'|'dashDotDot'|'slantDashDot'|'mediumDashed'|'mediumDashDotDot'|'mediumDashDot';
interface Border  { style: BorderStyle; color: Partial<Color>; }
interface BorderDiagonal extends Border { up: boolean; down: boolean; }
interface Borders {
  top: Partial<Border>; left: Partial<Border>; bottom: Partial<Border>; right: Partial<Border>;
  diagonal: Partial<BorderDiagonal>;
}
```

### `Alignment`

```ts
interface Alignment {
  horizontal: 'left'|'center'|'right'|'fill'|'justify'|'centerContinuous'|'distributed';
  vertical: 'top'|'middle'|'bottom'|'distributed'|'justify';
  wrapText: boolean; shrinkToFit: boolean; indent: number;
  readingOrder: 'rtl'|'ltr';
  textRotation: number | 'vertical';
}
```

### `Protection`

```ts
interface Protection { locked: boolean; hidden: boolean; }
```

### Number formats

`numFmt` is just a format-code string. `StylesXform` maps built-in codes to their reserved numeric IDs
(via `NumFmtXform.getDefaultFmtId`) and allocates custom codes from `NUMFMT_BASE` (164) upward.
`lib/xlsx/defaultnumformats.js` holds the built-in table.

> **Style-precedence rule (important, and subtle).** A cell's effective style is resolved at
> `Cell` construction by `_mergeStyle(rowStyle, colStyle, {})` — **row style wins over column style**,
> field by field (`numFmt`, `font`, `alignment`, `border`, `fill`, `protection`), and an explicit cell
> style overrides both. Reproduce this precedence exactly or existing files will render differently.

---

## Section 5 — Observations (keep / drop / redesign)

### What's elegant — keep
- **Model ⟷ document split** (`get model` / `set model` on every class). Clean serialization boundary;
  makes streaming and reconciliation possible. Keep verbatim (in spirit).
- **Column run-length compression** (`Column.toModel/fromModel`) — small, correct, and directly mirrors
  the OOXML `<col min max>` shape.
- **Style deduplication pool** with `WeakMap` cache — this is why ExcelJS files are compact. Keep the
  idea; the implementation can be simpler with TS `Map` + structural hashing.
- **Address utilities** (`col-cache`: `l2n`, `n2l`, `encodeAddress`, `decodeAddress`, `validateAddress`).
  Heavily cached, hot path. Port with tests.
- **xform base interface** (`prepare / render / parseOpen / parseText / parseClose / reconcile`) — a tidy
  contract for each XML part. Worth keeping as the serializer's shape (see OOXML notes).

### What's complex / risky — simplify
- **12 Value strategy classes** in `cell.js` → collapse to a discriminated union + pure helpers. Biggest
  memory + readability win. (~700 of cell.js's 1100 lines are these classes.)
- **`addRow`/`insertRow` style mini-DSL** (`'i' | 'o' | 'i+' | 'o+' | 'n'`) → explicit options object.
- **`spliceRows`/`spliceColumns`** are the trickiest code in the repo (cell shifting + merge re-anchoring
  + defined-name adjustment intertwined). Isolate each concern; property-test against ExcelJS output.
- **1-based sparse `_worksheets` array** (index 0 unused, repeated `slice/sort/filter` in the `worksheets`
  getter) → back with a `Map<number, Worksheet>` + maintained order list; make `worksheets` O(n).
- **`actualRowCount` / `actualColumnCount`** full scans → maintain counts incrementally or memoize.

### What to drop
- Deprecated `tabColor`-as-string branches in `addWorksheet` and the `worksheet.tabColor` getter/setter
  (both emit `console.trace`). Ship the modern `properties.tabColor` only.
- `console.warn`/`console.trace` side effects sprinkled through the model (name truncation, pivot warning,
  deprecations) → surface via return values / typed errors / an opt-in logger, not stdout.
- `JSONValue` fallback (stringify-anything) — too magic; require callers to pass a supported type or a
  RichText/string explicitly.

### API-shape decisions for `@alosha/xlsx`
- **TS-first, `Partial<>` everywhere** the current `.d.ts` already does — but make required-vs-optional
  honest (ExcelJS types many always-present fields as required when they're often undefined).
- **Async-only I/O**, returning real `Promise`s (ExcelJS already does; `Worksheet.protect` fakes async —
  make it genuinely async or sync). 
- **Immutable-ish value reads, explicit mutation** — avoid the getter-with-side-effects pattern
  (e.g. `cell.formula` lazily translating shared formulas and caching on read).
- **Separate `Workbook` (build) from `WorkbookReader`/`WorkbookWriter` (stream)** as ExcelJS does, but
  unify their worksheet APIs behind one interface so `addRow`/`getCell` behave identically.
- **Scope v0.1 to WRITE** (per roadmap): Workbook + Worksheet + Row + Cell + Style + shared-strings +
  styles dedup + ZIP out. Defer pivot tables, conditional formatting, drawings, and CSV until later stages.

### Priority reimplementation order (feeds Stage 1.2 core model)
1. `col-cache` address math (+ `Range`) — foundation, pure, testable.
2. `Cell` value union + `Style` types.
3. `Row` (values sparse array, style fan-out) and `Column` (compression).
4. `Worksheet` (rows/cells/merges/dimensions) — minus features.
5. `Workbook` (sheet management, metadata, model snapshot).
6. Then the OOXML writer (see `OOXML-NOTES.md`).

---

*Audited against exceljs `4.4.0`. Source of truth: `lib/doc/*` and `index.d.ts`. MIT-licensed;
reimplementation is a clean-room rewrite informed by — not copied from — this reading.*
