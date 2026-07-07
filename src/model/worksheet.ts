// Worksheet: the concrete WorksheetHost. Owns rows (a 1-based Map), columns (a Map plus a
// key->column index), and merge bookkeeping (a Map<masterAddress, Range>). Rows/Columns create
// their own cells; the Worksheet coordinates cross-cell concerns — merges, dimensions, and the
// row/column shifting in spliceRows/spliceColumns (docs/API-AUDIT.md §2/§5, the trickiest part).
//
// `Workbook` is a type-only import (back-reference only): Worksheet never constructs a Workbook, so
// this carries no runtime dependency and cannot cycle with workbook.ts's `import { Worksheet }`.

import { colLetterToNumber, decodeAddress, Range, type RangeSource } from "../address/index.js";
import {
  DuplicateWorksheetNameError,
  InvalidRangeError,
  InvalidWorksheetNameError,
  MergeConflictError,
  ReservedWorksheetNameError,
  WorksheetNameTooLongError,
} from "../errors.js";
import type { Cell } from "./cell.js";
import type { CellValue } from "./cell-value.js";
import { Column, type ColumnDefn, type ColumnRange } from "./column.js";
import { type CommentModel, Comments } from "./comment.js";
import type { ConditionalFormatting } from "./conditional-formatting.js";
import { type DataValidationModel, DataValidations } from "./data-validation.js";
import { CellType } from "./enums.js";
import { Row, type RowModel } from "./row.js";
import type { Color, Style } from "./style.js";
import type { RowValues } from "./types.js";
import type { Workbook } from "./workbook.js";
import type { WorksheetHost } from "./worksheet-host.js";

/** Sheet properties. `outlineLevelCol`/`outlineLevelRow` back the `collapsed` derivation on
 *  Column/Row; the rest is plain round-tripped data. */
export interface WorksheetProperties {
  tabColor?: Color;
  outlineLevelCol: number;
  outlineLevelRow: number;
  defaultRowHeight: number;
  defaultColWidth?: number;
  dyDescent: number;
  showGridLines?: boolean;
  [key: string]: unknown;
}

/** Freeze/split panes, zoom, gridlines — stored as plain data (round-trips through the model).
 *  `xSplit`/`ySplit`/`topLeftCell`/`activePane` back the `<pane>` child emitted when `state` is
 *  `"frozen"` or `"split"` (CT_Pane in the OOXML schema). */
export interface WorksheetView {
  state?: "normal" | "frozen" | "split";
  showGridLines?: boolean;
  zoomScale?: number;
  xSplit?: number;
  ySplit?: number;
  topLeftCell?: string;
  activePane?: "bottomRight" | "topRight" | "bottomLeft" | "topLeft";
  [key: string]: unknown;
}

/** Options accepted by `Workbook.addWorksheet`. */
export interface AddWorksheetOptions {
  state: "visible" | "hidden" | "veryHidden";
  properties: Partial<WorksheetProperties>;
  views: WorksheetView[];
  pageSetup: unknown;
  headerFooter: unknown;
}

/** Replaces ExcelJS's cryptic `'i' | 'o' | 'i+' | 'o+' | 'n'` style DSL (API-AUDIT §2). */
export type RowInheritOption =
  | { inheritFrom?: "above" | "below"; includeEmpty?: boolean }
  | undefined;

/** The serializable snapshot of a worksheet. */
export interface WorksheetModel {
  id: number;
  name: string;
  orderNo: number;
  state: "visible" | "hidden" | "veryHidden";
  properties: WorksheetProperties;
  views: WorksheetView[];
  columns?: ColumnRange[];
  rows: RowModel[];
  merges: string[];
  dataValidations?: DataValidationModel[];
  /** Cell comments, one entry per noted cell `ref` (`cell.note`). */
  comments?: CommentModel[];
  /** `<conditionalFormatting>` blocks: one entry per `ref`, each carrying its ordered rules. */
  conditionalFormattings?: ConditionalFormatting[];
  /** `<autoFilter ref="…">`, an A1 range like `"A1:D1"`. */
  autoFilter?: string;
  pageSetup?: unknown;
  headerFooter?: unknown;
}

type EachRowCb = (row: Row, n: number) => void;
type RowData = CellValue[] | Record<string, CellValue>;

/** Captured value+style of a row's cells, used to relocate a row during a splice/duplicate. */
interface RowSnapshot {
  height?: number;
  hidden: boolean;
  outlineLevel: number;
  style: Style;
  cells: Array<{ col: number; value: CellValue; style: Style }>;
}

const ILLEGAL_NAME_CHARS = /[*?:\\/[\]]/;

/**
 * Format validation shared by the `name` setter and `Workbook.addWorksheet` (which validates before
 * a Worksheet exists to own the setter). Duplicate detection needs the workbook and is done by the
 * caller. Not part of the public barrel.
 */
export function validateWorksheetNameFormat(name: string): void {
  if (name.length === 0) throw new InvalidWorksheetNameError(name, "empty");
  if (name.length > WorksheetNameTooLongError.MAX_LENGTH) throw new WorksheetNameTooLongError(name);
  if (ILLEGAL_NAME_CHARS.test(name))
    throw new InvalidWorksheetNameError(name, "illegal-characters");
  if (name.startsWith("'") || name.endsWith("'")) {
    throw new InvalidWorksheetNameError(name, "edge-quote");
  }
  if (name.toLowerCase() === "history") throw new ReservedWorksheetNameError(name);
}

/** Normalizes the several `mergeCells`/`unMergeCells` argument forms into a single `Range`. */
function normalizeMergeArgs(cells: unknown[]): Range {
  if (cells.length === 1) return new Range(cells[0] as RangeSource);
  if (cells.length === 2 && typeof cells[0] === "string" && typeof cells[1] === "string") {
    return new Range([cells[0], cells[1]]);
  }
  if (cells.length === 4 && cells.every((c) => typeof c === "number")) {
    return new Range(
      cells[0] as number,
      cells[1] as number,
      cells[2] as number,
      cells[3] as number,
    );
  }
  throw new InvalidRangeError(cells);
}

export class Worksheet implements WorksheetHost {
  orderNo: number;
  state: "visible" | "hidden" | "veryHidden" = "visible";
  readonly workbook: Workbook;
  properties: WorksheetProperties;
  views: WorksheetView[] = [];
  /** `<autoFilter ref="…">`, an A1 range like `"A1:D1"`. */
  autoFilter?: string;
  pageSetup?: unknown;
  headerFooter?: unknown;

  private _id: number;
  private _name: string;

  private readonly _rows = new Map<number, Row>();
  private readonly _columns = new Map<number, Column>();
  private readonly _columnsByKey = new Map<string, Column>();
  /** Active merge regions keyed by their master (top-left) cell address. */
  private readonly _merges = new Map<string, Range>();
  /** Data validations keyed by `sqref`; reached by cells through `WorksheetHost`. */
  private readonly _dataValidations = new DataValidations();
  /** Cell comments keyed by ref; reached by cells through `WorksheetHost`. */
  private readonly _comments = new Comments();
  /** Conditional-formatting blocks in insertion order (sheet-level, not cell-routed). */
  private _conditionalFormattings: ConditionalFormatting[] = [];

  constructor(workbook: Workbook, id: number, name: string, orderNo: number) {
    this.workbook = workbook;
    this._id = id;
    this._name = name;
    this.orderNo = orderNo;
    this.properties = {
      outlineLevelCol: 0,
      outlineLevelRow: 0,
      defaultRowHeight: 15,
      dyDescent: 55,
    };
  }

  // --- identity ---------------------------------------------------------------------------------

  get id(): number {
    return this._id;
  }

  get name(): string {
    return this._name;
  }

  set name(value: string) {
    validateWorksheetNameFormat(value);
    const lower = value.toLowerCase();
    for (const ws of this.workbook.worksheets) {
      if (ws !== this && ws.name.toLowerCase() === lower) {
        throw new DuplicateWorksheetNameError(value);
      }
    }
    this._name = value;
  }

  // --- rows -------------------------------------------------------------------------------------

  getRow(row: number): Row {
    let r = this._rows.get(row);
    if (!r) {
      r = new Row(this, row);
      this._rows.set(row, r);
    }
    return r;
  }

  getRows(start: number, length: number): Row[] | undefined {
    if (length < 1) return undefined;
    const rows: Row[] = [];
    for (let i = 0; i < length; i++) rows.push(this.getRow(start + i));
    return rows;
  }

  findRow(row: number): Row | undefined {
    return this._rows.get(row);
  }

  addRow(data: RowData, options?: RowInheritOption): Row {
    const row = this.getRow(this.rowCount + 1);
    row.values = data as RowValues;
    this.applyInheritance(row, options);
    return row;
  }

  addRows(rows: RowData[], options?: RowInheritOption): Row[] {
    return rows.map((data) => this.addRow(data, options));
  }

  insertRow(pos: number, value: RowData, options?: RowInheritOption): Row {
    this.spliceRows(pos, 0, value);
    const row = this.getRow(pos);
    this.applyInheritance(row, options);
    return row;
  }

  private applyInheritance(row: Row, options: RowInheritOption): void {
    if (!options?.inheritFrom) return;
    const sourceNumber = options.inheritFrom === "above" ? row.number - 1 : row.number + 1;
    const source = this._rows.get(sourceNumber);
    if (!source) return;
    row.style = source.style;
    source.eachCell({ includeEmpty: options.includeEmpty ?? false }, (cell, col) => {
      row.getCell(col).style = cell.style;
    });
  }

  duplicateRow(rowNum: number, count: number, insert = false): void {
    const snapshot = this.snapshotRow(this.getRow(rowNum));
    if (insert) {
      const blanks: RowData[] = Array.from({ length: count }, () => [] as CellValue[]);
      this.spliceRows(rowNum + 1, 0, ...blanks);
    }
    for (let i = 1; i <= count; i++) this.restoreRow(rowNum + i, snapshot);
  }

  spliceRows(start: number, count: number, ...inserts: RowData[]): void {
    const delta = inserts.length - count;

    const oldMerges = [...this._merges.values()].map((r) => ({
      top: r.top,
      left: r.left,
      bottom: r.bottom,
      right: r.right,
    }));

    const snapshots = new Map<number, RowSnapshot>();
    for (const [num, row] of this._rows) snapshots.set(num, this.snapshotRow(row));

    this._rows.clear();
    this._merges.clear();

    for (const [num, snap] of snapshots) {
      if (num < start) {
        this.restoreRow(num, snap);
      } else if (num >= start + count) {
        this.restoreRow(num + delta, snap);
      }
      // rows in [start, start+count) are deleted.
    }

    inserts.forEach((values, i) => {
      this.getRow(start + i).values = values as RowValues;
    });

    // Re-anchor merges: those entirely below the removed region shift by `delta`, those entirely
    // above survive unchanged, and any that straddle the removed region are dropped.
    for (const m of oldMerges) {
      if (m.top >= start + count) {
        this.applyMerge(new Range(m.top + delta, m.left, m.bottom + delta, m.right));
      } else if (m.bottom < start) {
        this.applyMerge(new Range(m.top, m.left, m.bottom, m.right));
      }
    }
  }

  private snapshotRow(row: Row): RowSnapshot {
    const cells: RowSnapshot["cells"] = [];
    row.eachCell((cell, col) => {
      cells.push({ col, value: cell.value, style: cell.style });
    });
    return {
      height: row.height,
      hidden: row.hidden,
      outlineLevel: row.outlineLevel,
      style: row.style,
      cells,
    };
  }

  private restoreRow(num: number, snap: RowSnapshot): void {
    const row = this.getRow(num);
    row.height = snap.height;
    row.hidden = snap.hidden;
    row.outlineLevel = snap.outlineLevel;
    row.style = snap.style;
    for (const c of snap.cells) {
      const cell = row.getCell(c.col);
      cell.value = c.value;
      cell.style = c.style;
    }
  }

  eachRow(cb: EachRowCb): void;
  eachRow(opts: { includeEmpty: boolean }, cb: EachRowCb): void;
  eachRow(optsOrCb: { includeEmpty: boolean } | EachRowCb, maybeCb?: EachRowCb): void {
    const includeEmpty = typeof optsOrCb === "object" ? optsOrCb.includeEmpty : false;
    const cb = typeof optsOrCb === "function" ? optsOrCb : (maybeCb as EachRowCb);

    if (includeEmpty) {
      const max = this.rowCount;
      for (let n = 1; n <= max; n++) cb(this.getRow(n), n);
    } else {
      const sorted = [...this._rows.entries()].sort((a, b) => a[0] - b[0]);
      for (const [n, row] of sorted) cb(row, n);
    }
  }

  getSheetValues(): RowValues[] {
    const result: RowValues[] = [];
    this.eachRow((row, n) => {
      result[n] = row.values;
    });
    return result;
  }

  get rowCount(): number {
    let max = 0;
    for (const n of this._rows.keys()) if (n > max) max = n;
    return max;
  }

  get actualRowCount(): number {
    let count = 0;
    for (const row of this._rows.values()) if (row.hasValues) count++;
    return count;
  }

  get lastRow(): Row | undefined {
    const rc = this.rowCount;
    return rc >= 1 ? this._rows.get(rc) : undefined;
  }

  // --- columns ----------------------------------------------------------------------------------

  get columns(): Column[] {
    const count = this.columnCount;
    const arr: Column[] = [];
    for (let n = 1; n <= count; n++) arr.push(this.getColumn(n));
    return arr;
  }

  set columns(defns: Partial<ColumnDefn>[]) {
    defns.forEach((defn, i) => {
      const col = this.getColumn(i + 1);
      col.defn = { ...col.defn, ...defn };
      if (defn.key !== undefined) this._columnsByKey.set(defn.key, col);
    });
  }

  getColumn(idOrKey: number | string): Column {
    if (typeof idOrKey === "number") return this.getOrCreateColumn(idOrKey);
    const byKey = this._columnsByKey.get(idOrKey);
    if (byKey) return byKey;
    return this.getOrCreateColumn(colLetterToNumber(idOrKey));
  }

  private getOrCreateColumn(n: number): Column {
    let c = this._columns.get(n);
    if (!c) {
      c = new Column(this, n);
      this._columns.set(n, c);
    }
    return c;
  }

  get columnCount(): number {
    let max = 0;
    for (const n of this._columns.keys()) if (n > max) max = n;
    this.eachRow((row) => {
      const d = row.dimensions;
      if (d && d.max > max) max = d.max;
    });
    return max;
  }

  get actualColumnCount(): number {
    const cols = new Set<number>();
    this.eachRow((row) => {
      row.eachCell((cell, col) => {
        if (cell.type !== CellType.Null) cols.add(col);
      });
    });
    return cols.size;
  }

  get lastColumn(): Column | undefined {
    const cc = this.columnCount;
    return cc >= 1 ? this.getColumn(cc) : undefined;
  }

  spliceColumns(start: number, count: number, ...inserts: CellValue[][]): void {
    const delta = inserts.length - count;

    this.eachRow((row) => {
      const rowInserts = inserts.map((colVals) => colVals[row.number] ?? null);
      row.splice(start, count, ...rowInserts);
    });

    // Shift stored column definitions to match the cell shift.
    const oldColumns = [...this._columns.entries()];
    this._columns.clear();
    this._columnsByKey.clear();
    for (const [n, col] of oldColumns) {
      let newNumber: number;
      if (n < start) newNumber = n;
      else if (n >= start + count) newNumber = n + delta;
      else continue; // deleted
      const moved = new Column(this, newNumber, col.defn);
      this._columns.set(newNumber, moved);
      if (moved.key !== undefined) this._columnsByKey.set(moved.key, moved);
    }

    // Re-anchor merges horizontally, mirroring spliceRows.
    const oldMerges = [...this._merges.values()].map((r) => ({
      top: r.top,
      left: r.left,
      bottom: r.bottom,
      right: r.right,
    }));
    this._merges.clear();
    for (const m of oldMerges) {
      if (m.left >= start + count) {
        this.applyMerge(new Range(m.top, m.left + delta, m.bottom, m.right + delta));
      } else if (m.right < start) {
        this.applyMerge(new Range(m.top, m.left, m.bottom, m.right));
      }
    }
  }

  // --- column keys ------------------------------------------------------------------------------

  getColumnKey(key: string): Column | undefined {
    return this._columnsByKey.get(key);
  }

  setColumnKey(key: string, column: Column): void {
    column.key = key;
    this._columnsByKey.set(key, column);
  }

  deleteColumnKey(key: string): void {
    const col = this._columnsByKey.get(key);
    if (col) col.key = undefined;
    this._columnsByKey.delete(key);
  }

  eachColumnKey(cb: (column: Column, key: string) => void): void {
    for (const [key, col] of this._columnsByKey) cb(col, key);
  }

  // --- cells ------------------------------------------------------------------------------------

  getCell(r: number | string, c?: number | string): Cell {
    if (c === undefined) {
      const d = decodeAddress(r as string);
      return this.getRow(d.row).getCell(d.col);
    }
    return this.getRow(r as number).getCell(c);
  }

  findCell(r: number | string, c?: number | string): Cell | undefined {
    let rowNum: number;
    let colNum: number;
    if (c === undefined) {
      const d = decodeAddress(r as string);
      rowNum = d.row;
      colNum = d.col;
    } else {
      rowNum = r as number;
      colNum = this.toColNumber(c);
    }
    const row = this._rows.get(rowNum);
    return row?.findCell(colNum);
  }

  private toColNumber(col: number | string): number {
    if (typeof col === "number") return col;
    const byKey = this._columnsByKey.get(col);
    if (byKey) return byKey.number;
    return colLetterToNumber(col);
  }

  // --- merges -----------------------------------------------------------------------------------

  mergeCells(...cells: unknown[]): void {
    const range = normalizeMergeArgs(cells);
    for (const existing of this._merges.values()) {
      if (range.intersects(existing)) throw new MergeConflictError(range.range);
    }
    this.applyMerge(range);
  }

  unMergeCells(...cells: unknown[]): void {
    const range = normalizeMergeArgs(cells);
    for (const [key, existing] of [...this._merges.entries()]) {
      if (!range.intersects(existing)) continue;
      existing.forEachAddress((_addr, row, col) => {
        if (row === existing.top && col === existing.left) return; // master keeps its value
        this.getRow(row).getCell(col).unmerge();
      });
      this._merges.delete(key);
    }
  }

  /** Links the slave cells of `range` to its top-left master and records the region. */
  private applyMerge(range: Range): void {
    const master = this.getCell(range.top, range.left);
    range.forEachAddress((_addr, row, col) => {
      if (row === range.top && col === range.left) return;
      this.getRow(row).getCell(col).merge(master);
    });
    this._merges.set(master.address, range);
  }

  get hasMerges(): boolean {
    return this._merges.size > 0;
  }

  // --- data validations -------------------------------------------------------------------------

  /** Range-keyed data-validation store (`add`/`find`/`remove`/`model`), mirroring ExcelJS. */
  get dataValidations(): DataValidations {
    return this._dataValidations;
  }

  // --- comments ---------------------------------------------------------------------------------

  /** Ref-keyed cell-comment store (`add`/`find`/`remove`/`model`); `cell.note` reads/writes it. */
  get comments(): Comments {
    return this._comments;
  }

  // --- conditional formatting -------------------------------------------------------------------

  /** The worksheet's conditional-formatting blocks, in insertion order. */
  get conditionalFormattings(): ConditionalFormatting[] {
    return this._conditionalFormattings;
  }

  /**
   * Register a `{ ref, rules }` block (ExcelJS shape). `ref` is validated through the `Range` decoder
   * — a malformed range throws a typed {@link InvalidRangeError}, never a bare `TypeError`.
   */
  addConditionalFormatting(cf: ConditionalFormatting): ConditionalFormatting {
    for (const part of cf.ref.trim().split(/\s+/)) {
      if (part === "") continue;
      // Constructing a Range validates the fragment and throws InvalidRangeError on garbage.
      new Range(part);
    }
    this._conditionalFormattings.push(cf);
    return cf;
  }

  /** Drop every conditional-formatting block whose `ref` matches (exact-string comparison). */
  removeConditionalFormatting(ref: string): void {
    this._conditionalFormattings = this._conditionalFormattings.filter((cf) => cf.ref !== ref);
  }

  get dimensions(): Range {
    let top = Number.POSITIVE_INFINITY;
    let left = Number.POSITIVE_INFINITY;
    let bottom = Number.NEGATIVE_INFINITY;
    let right = Number.NEGATIVE_INFINITY;
    this.eachRow((row, n) => {
      const d = row.dimensions;
      if (!d) return;
      if (n < top) top = n;
      if (n > bottom) bottom = n;
      if (d.min < left) left = d.min;
      if (d.max > right) right = d.max;
    });
    if (top === Number.POSITIVE_INFINITY) return new Range(1, 1, 1, 1);
    return new Range(top, left, bottom, right);
  }

  // --- model ------------------------------------------------------------------------------------

  get model(): WorksheetModel {
    const rows = [...this._rows.entries()].sort((a, b) => a[0] - b[0]).map(([, row]) => row.model);
    const columns = Column.toModel([...this._columns.values()]);

    const m: WorksheetModel = {
      id: this._id,
      name: this._name,
      orderNo: this.orderNo,
      state: this.state,
      properties: { ...this.properties },
      views: this.views.map((v) => ({ ...v })),
      rows,
      merges: [...this._merges.values()].map((r) => r.range),
    };
    if (columns) m.columns = columns;
    if (this._dataValidations.size > 0) m.dataValidations = this._dataValidations.model;
    if (this._comments.size > 0) m.comments = this._comments.model;
    if (this._conditionalFormattings.length > 0) {
      m.conditionalFormattings = this._conditionalFormattings.map((cf) => ({
        ref: cf.ref,
        rules: cf.rules.map((r) => ({ ...r })),
      }));
    }
    if (this.autoFilter !== undefined) m.autoFilter = this.autoFilter;
    if (this.pageSetup !== undefined) m.pageSetup = this.pageSetup;
    if (this.headerFooter !== undefined) m.headerFooter = this.headerFooter;
    return m;
  }

  set model(m: WorksheetModel) {
    this._id = m.id;
    this._name = m.name;
    this.orderNo = m.orderNo;
    this.state = m.state;
    this.properties = { ...m.properties };
    this.views = m.views.map((v) => ({ ...v }));
    this.autoFilter = m.autoFilter;
    this.pageSetup = m.pageSetup;
    this.headerFooter = m.headerFooter;

    this._rows.clear();
    this._columns.clear();
    this._columnsByKey.clear();
    this._merges.clear();
    this._dataValidations.model = m.dataValidations ?? [];
    this._comments.model = m.comments ?? [];
    this._conditionalFormattings = (m.conditionalFormattings ?? []).map((cf) => ({
      ref: cf.ref,
      rules: cf.rules.map((r) => ({ ...r })),
    }));

    if (m.columns) {
      const cols = Column.fromModel(this, m.columns) ?? [];
      for (const col of cols) {
        this._columns.set(col.number, col);
        if (col.key !== undefined) this._columnsByKey.set(col.key, col);
      }
    }

    for (const rowModel of m.rows) {
      this.getRow(rowModel.number).model = rowModel;
    }

    for (const mergeStr of m.merges) this.applyMerge(new Range(mergeStr));
  }
}
