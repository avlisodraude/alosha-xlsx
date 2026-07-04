# Stage 1.2 Build Spec — Core Document Model

> **Input for the build session.** This turns [`API-AUDIT.md`](./API-AUDIT.md) §5 into concrete module
> boundaries, TypeScript signatures, and a test plan so the core model can be built without re-reading
> the ExcelJS source. Serialization (model → OOXML → zip) is **Stage 1.3** and out of scope here.
>
> Conventions: TS strict, ESM, `Uint8Array` over Node `Buffer`, no `console.*` side effects (typed
> errors / return values instead), pure functions where possible. Depends on the shipped
> [`src/zip-layer.ts`](../src/zip-layer.ts) only indirectly (the writer will).

## 0. Scope

**In scope (this stage):** the in-memory document graph — addresses, `Range`, cell value model, style
types, `Cell`, `Row`, `Column`, `Worksheet`, `Workbook` — plus each class's `model` snapshot get/set.

**Out of scope (defer):** XML serialization/parsing, the write pipeline, reading `.xlsx`, shared-strings
and style dedup pools (Stage 1.3), and all features — pivot tables, conditional formatting, drawings /
images, tables, data validation, comments, hyperlinks-as-cells beyond the value type, protection.

**Definition of done:** you can build a workbook in memory, set typed cell values and styles, merge
cells, iterate rows/cells, and round-trip every class through `get model` → `set model` with deep
equality. 100% of the unit tests in §8 pass. No `console` output on any happy or validation path.

## 1. Module layout

```
src/
  address/
    col-cache.ts       // A1 <-> {row,col}, letter<->number, address validation (pure, cached)
    range.ts           // Range class (bounding box + set ops)
    index.ts
  model/
    enums.ts           // CellType, FormulaType, ErrorCode
    types.ts           // shared primitives: Address, RowValues, etc.
    style.ts           // Style, Font, Fill, Border(s), Alignment, Protection, Color
    cell-value.ts      // CellValue union + StoredValue + pure helpers (detect/effective/toText/toCsv)
    cell.ts            // Cell class
    row.ts             // Row class
    column.ts          // Column class
    worksheet.ts       // Worksheet class
    workbook.ts        // Workbook class
    index.ts
  errors.ts            // typed error classes
  index.ts             // public barrel export
tests/
  address/col-cache.test.ts
  address/range.test.ts
  model/cell-value.test.ts
  model/style.test.ts
  model/cell.test.ts
  model/row.test.ts
  model/column.test.ts
  model/worksheet.test.ts
  model/workbook.test.ts
```

Build (and merge) in the order below — each layer is fully tested before the next depends on it.

## 2. `address/` — addresses & Range (build first)

Pure, hot-path, zero dependencies. Port ExcelJS's `col-cache` behavior but as typed free functions.

```ts
// col-cache.ts
export interface Decoded { row: number; col: number; address: string; }
export interface DecodedEx extends Decoded { $col$row: string; sheetName?: string; }

/** 'A'->1, 'Z'->26, 'AA'->27. Throws InvalidAddressError on non-letters. */
export function colLetterToNumber(letter: string): number;
/** 1->'A', 27->'AA'. */
export function colNumberToLetter(col: number): string;
/** (row=1,col=2) -> 'B1'. */
export function encodeAddress(row: number, col: number): string;
/** 'B1' -> { row:1, col:2, address:'B1' }. Cached. */
export function decodeAddress(address: string): Decoded;
/** Handles '$B$1', "Sheet 1"!B1, absolute/relative. */
export function decodeEx(address: string): DecodedEx;
/** Throws InvalidAddressError if not a well-formed A1 cell address. */
export function validateAddress(address: string): void;
```

Requirements: memoize `decodeAddress`/`colLetterToNumber` (Map cache) — this is called per cell.
Support columns past `Z` (multi-letter) and rows into the 1,048,576 range. 1-based rows and cols
everywhere in the public surface; document it once.

```ts
// range.ts
export type RangeSource =
  | string                                   // 'A1:B2' or 'A1'
  | [string, string]                         // ['A1','B2']
  | [number, number, number, number]         // [top,left,bottom,right]
  | Range | Decoded;

export class Range {
  top: number; left: number; bottom: number; right: number;
  sheetName?: string;
  constructor(source?: RangeSource /* | (t,l,b,r) */);
  get tl(): string;               // top-left address 'A1'
  get br(): string;               // bottom-right address 'B2'
  get range(): string;            // 'A1:B2'
  get count(): number;            // cell count
  expand(top: number, left: number, bottom: number, right: number): void;
  contains(address: string): boolean;
  intersects(other: Range): boolean;
  forEachAddress(cb: (address: string, row: number, col: number) => void): void;
}
```

`Range` normalizes so `top<=bottom`, `left<=right` regardless of input order. Used by `Worksheet.dimensions`
and merge bookkeeping.

## 3. `model/` enums & shared types

```ts
// enums.ts
export enum CellType {
  Null, Number, String, Boolean, Date, Error, RichText, Hyperlink, Formula, Merge, SharedString,
}
export enum FormulaType { None, Master, Shared }
export type ErrorCode = '#N/A' | '#REF!' | '#NAME?' | '#DIV/0!' | '#NULL!' | '#VALUE!' | '#NUM!';
```

> Keep the *set* of types from ExcelJS (`API-AUDIT §ValueType`) but drop the numeric-wire coupling —
> the OOXML `t=` mapping lives in the Stage 1.3 serializer, not the model enum.

## 4. `model/style.ts` — style & format types

Port the audited shapes verbatim in spirit; make optionality honest (all fields optional on a partial
style, since a cell may set only `font.bold`).

```ts
export interface Color { argb?: string; theme?: number; indexed?: number; }

export interface Font {
  name?: string; size?: number; family?: number;
  scheme?: 'minor' | 'major' | 'none'; charset?: number;
  color?: Color; bold?: boolean; italic?: boolean;
  underline?: boolean | 'none' | 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';
  vertAlign?: 'superscript' | 'subscript'; strike?: boolean; outline?: boolean;
}

export type FillPattern =
  | 'none' | 'solid'
  | 'darkVertical' | 'darkHorizontal' | 'darkGrid' | 'darkTrellis' | 'darkDown' | 'darkUp'
  | 'lightVertical' | 'lightHorizontal' | 'lightGrid' | 'lightTrellis' | 'lightDown' | 'lightUp'
  | 'darkGray' | 'mediumGray' | 'lightGray' | 'gray125' | 'gray0625';

export interface GradientStop { position: number; color: Color; }
export type Fill =
  | { type: 'pattern'; pattern: FillPattern; fgColor?: Color; bgColor?: Color }
  | { type: 'gradient'; gradient: 'angle'; degree: number; stops: GradientStop[] }
  | { type: 'gradient'; gradient: 'path'; center: { left: number; top: number }; stops: GradientStop[] };

export type BorderStyle =
  | 'thin' | 'dotted' | 'hair' | 'medium' | 'double' | 'thick' | 'dashed'
  | 'dashDot' | 'dashDotDot' | 'slantDashDot' | 'mediumDashed' | 'mediumDashDotDot' | 'mediumDashDot';
export interface Border { style?: BorderStyle; color?: Color; }
export interface BorderDiagonal extends Border { up?: boolean; down?: boolean; }
export interface Borders {
  top?: Border; left?: Border; bottom?: Border; right?: Border; diagonal?: BorderDiagonal;
}

export interface Alignment {
  horizontal?: 'left'|'center'|'right'|'fill'|'justify'|'centerContinuous'|'distributed';
  vertical?: 'top'|'middle'|'bottom'|'distributed'|'justify';
  wrapText?: boolean; shrinkToFit?: boolean; indent?: number;
  readingOrder?: 'rtl'|'ltr'; textRotation?: number | 'vertical';
}
export interface Protection { locked?: boolean; hidden?: boolean; }

export interface Style {
  numFmt?: string;
  font?: Font;
  alignment?: Alignment;
  protection?: Protection;
  border?: Borders;
  fill?: Fill;
}
```

**Style precedence helper (must match ExcelJS exactly — see API-AUDIT §4):**

```ts
/** Resolve a cell's effective style: row default and column default merged field-by-field,
 *  ROW WINS over COLUMN, then the cell's own style overrides both. Fields: numFmt, font,
 *  alignment, border, fill, protection. Returns a new object; never mutates inputs. */
export function resolveCellStyle(rowStyle?: Style, colStyle?: Style, cellStyle?: Style): Style;
```

## 5. `model/cell-value.ts` — the value union (the key redesign)

Replace ExcelJS's **12 strategy classes** with a discriminated union + pure functions. This is the single
biggest simplification and memory win called out in API-AUDIT §5.

```ts
export interface RichTextRun { text: string; font?: Font; }
export interface RichText { richText: RichTextRun[]; }
export interface Hyperlink { text: string; hyperlink: string; tooltip?: string; }
export interface CellError { error: ErrorCode; }
export interface CellFormula {
  formula?: string;          // master formula
  sharedFormula?: string;    // address of the master, for shared members
  ref?: string;              // shared/array range
  shareType?: 'shared' | 'array';
  result?: number | string | boolean | Date | CellError;
}

/** What callers pass to `cell.value = …` and read back. Ergonomic form. */
export type CellValue =
  | null | number | string | boolean | Date
  | CellError | RichText | Hyperlink | CellFormula;

/** Internal normalized storage on the cell model (discriminated). */
export type StoredValue =
  | { type: CellType.Null }
  | { type: CellType.Number;  value: number }
  | { type: CellType.String;  value: string }
  | { type: CellType.Boolean; value: boolean }
  | { type: CellType.Date;    value: Date }
  | { type: CellType.Error;   value: CellError }
  | { type: CellType.RichText; value: RichText }
  | { type: CellType.Hyperlink; value: Hyperlink }
  | { type: CellType.Formula; value: CellFormula }
  | { type: CellType.Merge;   master: string }
  | { type: CellType.SharedString; value: number };  // read-path only
```

Pure helpers (no class-per-type):

```ts
/** Duck-type a CellValue into a CellType (mirrors ExcelJS Value.getType order:
 *  null → string → number → boolean → date → hyperlink(text&hyperlink) → formula(formula|sharedFormula)
 *  → richText → error). */
export function detectType(v: CellValue): CellType;
export function toStored(v: CellValue): StoredValue;
export function fromStored(s: StoredValue): CellValue;
/** For a formula cell, the type of its `result`; else the cell's own type. */
export function effectiveType(s: StoredValue): CellType;
export function toText(s: StoredValue): string;    // display string
export function toCsv(s: StoredValue): string;     // CSV-escaped rendering
```

Notes: `Merge` stores only the master address; resolving a merge slave's value goes through the sheet.
Shared-formula → per-cell translation is a **Stage 1.3** concern (serializer) — here just store
`sharedFormula`/`ref`; do not eagerly translate on read.

## 6. `model/cell.ts`, `row.ts`, `column.ts`

```ts
export interface CellModel {
  address: string;
  type: CellType;
  value?: unknown; formula?: string; result?: unknown; // per stored variant
  style?: Style;
  // comment?, dataValidation? — deferred
}

export class Cell {
  constructor(row: Row, column: Column, address: string);
  readonly address: string;
  readonly row: number;            // 1-based
  readonly col: number;            // 1-based
  readonly $col$row: string;       // '$B$2'

  get value(): CellValue;
  set value(v: CellValue);         // setting on a merge slave writes the master
  readonly type: CellType;
  readonly effectiveType: CellType;
  readonly text: string;
  toString(): string;
  toCsvString(): string;

  // merge
  readonly isMerged: boolean;
  readonly master: Cell;           // self if unmerged
  merge(master: Cell, ignoreStyle?: boolean): void;
  unmerge(): void;

  // formula
  readonly formula?: string;
  readonly result?: unknown;
  readonly formulaType: FormulaType;

  // style — resolved from row+column at construction, overridable per-field
  style: Style;
  numFmt?: string; font?: Font; alignment?: Alignment;
  border?: Borders; fill?: Fill; protection?: Protection;

  get model(): CellModel;
  set model(m: CellModel);
}
```

```ts
export class Row {
  constructor(worksheet: Worksheet, number: number);
  readonly number: number;         // 1-based
  readonly worksheet: Worksheet;

  values: RowValues;               // get: sparse array; set: contiguous | sparse | {colKey: value}
  height?: number; hidden: boolean; outlineLevel: number;
  readonly collapsed: boolean;

  getCell(col: number | string /* letter | key */): Cell;
  findCell(col: number): Cell | undefined;
  eachCell(cb: (cell: Cell, colNumber: number) => void): void;
  eachCell(opts: { includeEmpty: boolean }, cb: (cell: Cell, colNumber: number) => void): void;

  readonly hasValues: boolean;
  readonly cellCount: number;
  readonly actualCellCount: number;
  readonly dimensions: { min: number; max: number } | null;

  style: Style;
  numFmt?: string; font?: Font; alignment?: Alignment; border?: Borders; fill?: Fill; protection?: Protection;
  // style setters fan out to existing cells (as ExcelJS _applyStyle)

  splice(start: number, count: number, ...inserts: CellValue[]): void;
  get model(): RowModel; set model(m: RowModel);
}

export type RowValues = CellValue[] | { [key: string]: CellValue } | undefined | null;
```

```ts
export class Column {
  constructor(worksheet: Worksheet, number: number, defn?: Partial<ColumnDefn> | false);
  readonly number: number; readonly letter: string;
  key?: string; header?: string | string[]; width?: number;
  hidden: boolean; outlineLevel: number; readonly collapsed: boolean;
  readonly isCustomWidth: boolean; readonly isDefault: boolean;
  style: Style;
  numFmt?: string; font?: Font; alignment?: Alignment; border?: Borders; fill?: Fill; protection?: Protection;
  values: CellValue[];             // get/set down the column
  eachCell(opts, cb): void;
  defn: ColumnDefn;

  equivalentTo(other: Column): boolean;
  /** Run-length compress an array of Columns into <col min max> ranges (and back). */
  static toModel(columns: Column[]): ColumnRange[] | undefined;
  static fromModel(ws: Worksheet, cols: ColumnRange[]): Column[] | null;
}
```

Default column width constant = `9`. Header assignment writes the top row cells' values (as ExcelJS).

## 7. `model/worksheet.ts` & `model/workbook.ts`

Trim to the audited essentials; leave feature hooks (tables/pivots/CF/images/protection/dataValidation)
as **`// Stage 2`** stubs or omit from types for now.

```ts
export class Worksheet {
  readonly id: number;
  name: string;                    // validated setter (see below)
  orderNo: number;
  state: 'visible' | 'hidden' | 'veryHidden';
  readonly workbook: Workbook;
  properties: WorksheetProperties; // tabColor, defaultRowHeight=15, outlineLevelCol/Row, dyDescent=55, ...
  views: WorksheetView[];

  // rows
  getRow(r: number): Row;
  getRows(start: number, length: number): Row[] | undefined;
  findRow(r: number): Row | undefined;
  addRow(data: CellValue[] | Record<string, CellValue>, style?: RowInheritOption): Row;
  addRows(rows: Array<CellValue[] | Record<string, CellValue>>, style?: RowInheritOption): Row[];
  insertRow(pos: number, value, style?: RowInheritOption): Row;
  spliceRows(start: number, count: number, ...inserts): void;
  duplicateRow(rowNum: number, count: number, insert?: boolean): void;
  eachRow(cb: (row: Row, n: number) => void): void;
  eachRow(opts: { includeEmpty: boolean }, cb: (row: Row, n: number) => void): void;
  readonly rowCount: number; readonly actualRowCount: number; readonly lastRow?: Row;

  // columns
  columns: Partial<ColumnDefn>[];
  getColumn(idOrKey: number | string): Column;
  readonly columnCount: number; readonly actualColumnCount: number; readonly lastColumn?: Column;
  spliceColumns(start: number, count: number, ...inserts: CellValue[][]): void;

  // cells & merges
  getCell(r: number | string, c?: number | string): Cell;
  findCell(r: number | string, c?: number | string): Cell | undefined;
  mergeCells(...cells: unknown[]): void;
  unMergeCells(...cells: unknown[]): void;
  readonly hasMerges: boolean;
  readonly dimensions: Range;

  get model(): WorksheetModel; set model(m: WorksheetModel);
}

/** Replaces ExcelJS's cryptic 'i'|'o'|'i+'|'o+'|'n' style DSL (API-AUDIT §2). */
export type RowInheritOption = { inheritFrom?: 'above' | 'below'; includeEmpty?: boolean } | undefined;
```

**`name` setter validation (throw typed errors — no `console.warn`):**
reject `* ? : / \ [ ]`, reject leading/trailing `'`, reject empty, reject reserved `"History"`, reject
duplicate (case-insensitive within the workbook); on >31 chars throw `WorksheetNameTooLongError` (decide:
throw vs. truncate — recommend **throw** for a clean API, unlike ExcelJS which warns+truncates).

```ts
export class Workbook {
  creator?: string; lastModifiedBy?: string;
  created: Date; modified: Date; lastPrinted?: Date;
  title?: string; subject?: string; keywords?: string; category?: string; description?: string;
  company?: string; manager?: string;
  properties: WorkbookProperties;   // { date1904?: boolean }
  calcProperties: CalculationProperties;
  views: WorkbookView[];

  readonly worksheets: Worksheet[]; // ordered, O(n) — NOT slice/sort/filter every access
  addWorksheet(name?: string, options?: Partial<AddWorksheetOptions>): Worksheet;
  getWorksheet(idOrName?: number | string): Worksheet | undefined;
  removeWorksheet(idOrName: number | string): void;
  eachSheet(cb: (ws: Worksheet, id: number) => void): void;

  get model(): WorkbookModel; set model(m: WorkbookModel);
}
```

Back sheet storage with a `Map<number, Worksheet>` + a maintained order list so `worksheets` is O(n) and
there's no 1-based sparse hole (API-AUDIT §5). Drop the deprecated `tabColor`-as-string branch entirely.

## 8. Test plan (vitest, colocated pattern like `tests/zip-layer.test.ts`)

- **col-cache:** `A↔1`, `Z↔26`, `AA↔27`, `XFD↔16384`; `encode/decode` round-trip fuzz (random row/col);
  `$A$1` and `'Sheet 1'!B2` decodeEx; invalid inputs throw `InvalidAddressError`; cache returns identical
  results across calls.
- **range:** construct from every `RangeSource` form; unordered inputs normalize; `contains`/`intersects`
  truth table; `forEachAddress` visits `count` cells in row-major order; `expand` grows the box.
- **cell-value:** `detectType` for each variant incl. edge cases (`''` is String not Null, `0`/`false`
  are Number/Boolean not Null, `{text,hyperlink}` → Hyperlink, `{formula}`/`{sharedFormula}` → Formula);
  `toStored`/`fromStored` round-trip; `effectiveType` of a formula follows its `result`; `toText`/`toCsv`
  per type (CSV quotes+escapes strings, renders formula result, empty for null/merge).
- **style:** `resolveCellStyle` — row-over-column precedence per field; cell overrides both; inputs never
  mutated; partial fonts merge (setting `font.bold` alone).
- **cell:** value set/get for every type; setting a value on a merge slave writes the master; `text`,
  `$col$row`, `address`/`row`/`col`; style shortcuts write into `style`; `model` round-trip.
- **row/column:** `values` as contiguous vs sparse vs keyed object; style fan-out to existing cells;
  `dimensions`/`hasValues`; `Column.toModel/fromModel` compresses & re-expands identical runs; header
  writes top cells.
- **worksheet:** `getRow`/`getCell` create on demand; `addRow` appends at next row; `mergeCells` +
  `unMergeCells` (incl. throw on overlapping merge); `dimensions` bounding box; name-validation throws for
  each illegal case; `spliceRows` shifts cells and re-anchors merges.
- **workbook:** `addWorksheet` naming/ids/order; `getWorksheet` by id and name; `removeWorksheet`;
  `worksheets` ordering after out-of-order adds/removes; full `Workbook` `model` round-trip deep-equals.

Add a coverage gate later; for now require every public method to have at least one assertion.

## 9. Explicitly deferred to Stage 1.3+

Serializer (model → `[Content_Types].xml`, `_rels/.rels`, `xl/workbook.xml`(+rels),
`xl/worksheets/sheetN.xml`, `xl/styles.xml`, optional `xl/sharedStrings.xml`, `docProps/*`, fixed
`theme1.xml`) → pack with the shipped zip layer → valid `.xlsx`; shared-strings pool; style dedup pools
with baseline seeds + numFmt id 164 base; date serial (1900/1904) conversion; the read path. See
OOXML-NOTES §5–§8.

---

*Grounded in `API-AUDIT.md` (ExcelJS 4.4.0 audit) and `OOXML-NOTES.md`. Clean-room; no ExcelJS code copied.*
