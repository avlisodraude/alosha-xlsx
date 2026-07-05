// Column: a projection over cells, not a container (docs/API-AUDIT.md §3) — it has no cell
// storage of its own; `eachCell`/`values` reach cells via the worksheet host's rows. Style
// setters fan out through `eachCell`. `toModel`/`fromModel` run-length-compress adjacent columns
// sharing the same OOXML-relevant attributes (width/hidden/outlineLevel/style) into `<col min
// max>` ranges — `key`/`header` are pure in-memory conveniences with no `<col>` representation,
// so they're deliberately excluded from that comparison and from `ColumnRange`.

import { colNumberToLetter } from "../address/index.js";
import type { Cell } from "./cell.js";
import type { CellValue } from "./cell-value.js";
import type { Alignment, Borders, Fill, Font, Protection, Style } from "./style.js";
import type { WorksheetHost } from "./worksheet-host.js";

export const DEFAULT_COLUMN_WIDTH = 9;

export interface ColumnDefn {
  key?: string;
  header?: string | string[];
  width?: number;
  hidden?: boolean;
  outlineLevel?: number;
  style?: Style;
}

/** Run-length range as compressed by `Column.toModel` — OOXML `<col min max .../>` attributes only. */
export interface ColumnRange {
  min: number;
  max: number;
  width?: number;
  hidden?: boolean;
  outlineLevel?: number;
  style?: Style;
}

type EachCellCb = (cell: Cell, rowNumber: number) => void;

function deepEqual(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  if (typeof a !== "object" || typeof b !== "object" || a === null || b === null) return false;
  const aKeys = Object.keys(a as object);
  const bKeys = Object.keys(b as object);
  if (aKeys.length !== bKeys.length) return false;
  return aKeys.every((key) =>
    deepEqual((a as Record<string, unknown>)[key], (b as Record<string, unknown>)[key]),
  );
}

export class Column {
  key?: string;
  width?: number;
  hidden = false;
  outlineLevel = 0;

  private _header: string | string[] | undefined;
  private _style: Style = {};

  constructor(
    private readonly host: WorksheetHost,
    public readonly number: number,
    defn?: Partial<ColumnDefn> | false,
  ) {
    if (defn) {
      if (defn.key !== undefined) this.key = defn.key;
      if (defn.width !== undefined) this.width = defn.width;
      if (defn.hidden !== undefined) this.hidden = defn.hidden;
      if (defn.outlineLevel !== undefined) this.outlineLevel = defn.outlineLevel;
      if (defn.style !== undefined) this._style = defn.style;
      if (defn.header !== undefined) this.header = defn.header;
    }
  }

  get letter(): string {
    return colNumberToLetter(this.number);
  }

  get collapsed(): boolean {
    return this.outlineLevel > this.host.properties.outlineLevelCol;
  }

  get isCustomWidth(): boolean {
    return this.width !== undefined && this.width !== DEFAULT_COLUMN_WIDTH;
  }

  get isDefault(): boolean {
    return (
      !this.isCustomWidth &&
      !this.hidden &&
      this.outlineLevel === 0 &&
      this.key === undefined &&
      this._header === undefined &&
      Object.keys(this._style).length === 0
    );
  }

  get header(): string | string[] | undefined {
    return this._header;
  }

  /** Writes the header text(s) into this column's top cell(s): row 1 for a string, rows 1..N
   *  for a string array (SPEC §6 — "header assignment writes the top row cells' values"). */
  set header(h: string | string[] | undefined) {
    this._header = h;
    if (h === undefined) return;
    const lines = Array.isArray(h) ? h : [h];
    lines.forEach((text, i) => {
      this.host.getRow(i + 1).getCell(this.number).value = text;
    });
  }

  get values(): CellValue[] {
    const result: CellValue[] = [];
    this.host.eachRow((row, n) => {
      const cell = row.findCell(this.number);
      if (cell) result[n] = cell.value;
    });
    return result;
  }

  set values(vals: CellValue[]) {
    vals.forEach((val, i) => {
      if (i === 0 || val === undefined) return;
      this.host.getRow(i).getCell(this.number).value = val;
    });
  }

  eachCell(cb: EachCellCb): void;
  eachCell(opts: { includeEmpty: boolean }, cb: EachCellCb): void;
  eachCell(optsOrCb: { includeEmpty: boolean } | EachCellCb, maybeCb?: EachCellCb): void {
    const includeEmpty = typeof optsOrCb === "object" ? optsOrCb.includeEmpty : false;
    const cb = typeof optsOrCb === "function" ? optsOrCb : (maybeCb as EachCellCb);

    this.host.eachRow({ includeEmpty }, (row, n) => {
      const cell = includeEmpty ? row.getCell(this.number) : row.findCell(this.number);
      if (cell) cb(cell, n);
    });
  }

  private applyToCells(fn: (cell: Cell) => void): void {
    this.eachCell(fn);
  }

  get style(): Style {
    return this._style;
  }

  set style(s: Style) {
    this._style = s;
    this.applyToCells((cell) => {
      cell.style = s;
    });
  }

  get numFmt(): string | undefined {
    return this._style.numFmt;
  }

  set numFmt(v: string | undefined) {
    this._style = { ...this._style, numFmt: v };
    this.applyToCells((cell) => {
      cell.numFmt = v;
    });
  }

  get font(): Font | undefined {
    return this._style.font;
  }

  set font(v: Font | undefined) {
    this._style = { ...this._style, font: v };
    this.applyToCells((cell) => {
      cell.font = v;
    });
  }

  get alignment(): Alignment | undefined {
    return this._style.alignment;
  }

  set alignment(v: Alignment | undefined) {
    this._style = { ...this._style, alignment: v };
    this.applyToCells((cell) => {
      cell.alignment = v;
    });
  }

  get border(): Borders | undefined {
    return this._style.border;
  }

  set border(v: Borders | undefined) {
    this._style = { ...this._style, border: v };
    this.applyToCells((cell) => {
      cell.border = v;
    });
  }

  get fill(): Fill | undefined {
    return this._style.fill;
  }

  set fill(v: Fill | undefined) {
    this._style = { ...this._style, fill: v };
    this.applyToCells((cell) => {
      cell.fill = v;
    });
  }

  get protection(): Protection | undefined {
    return this._style.protection;
  }

  set protection(v: Protection | undefined) {
    this._style = { ...this._style, protection: v };
    this.applyToCells((cell) => {
      cell.protection = v;
    });
  }

  get defn(): ColumnDefn {
    const d: ColumnDefn = {};
    if (this.key !== undefined) d.key = this.key;
    if (this._header !== undefined) d.header = this._header;
    if (this.width !== undefined) d.width = this.width;
    if (this.hidden) d.hidden = true;
    if (this.outlineLevel !== 0) d.outlineLevel = this.outlineLevel;
    if (Object.keys(this._style).length > 0) d.style = this._style;
    return d;
  }

  set defn(d: ColumnDefn) {
    this.key = d.key;
    this.width = d.width;
    this.hidden = d.hidden ?? false;
    this.outlineLevel = d.outlineLevel ?? 0;
    this._style = d.style ?? {};
    this.header = d.header;
  }

  /** Whether two columns share the same OOXML-relevant attributes (used to compress runs). */
  equivalentTo(other: Column): boolean {
    return (
      this.width === other.width &&
      this.hidden === other.hidden &&
      this.outlineLevel === other.outlineLevel &&
      deepEqual(this._style, other._style)
    );
  }

  private static rangeOf(start: Column, end: Column): ColumnRange {
    const range: ColumnRange = { min: start.number, max: end.number };
    if (start.width !== undefined) range.width = start.width;
    if (start.hidden) range.hidden = true;
    if (start.outlineLevel !== 0) range.outlineLevel = start.outlineLevel;
    if (Object.keys(start._style).length > 0) range.style = start._style;
    return range;
  }

  /** Run-length compress an array of Columns into `<col min max>` ranges. */
  static toModel(columns: Column[]): ColumnRange[] | undefined {
    if (columns.length === 0) return undefined;

    const sorted = [...columns].sort((a, b) => a.number - b.number);
    const ranges: ColumnRange[] = [];
    let start = sorted[0] as Column;
    let prev = start;

    for (let i = 1; i < sorted.length; i++) {
      const col = sorted[i] as Column;
      if (col.number === prev.number + 1 && col.equivalentTo(prev)) {
        prev = col;
        continue;
      }
      ranges.push(Column.rangeOf(start, prev));
      start = col;
      prev = col;
    }
    ranges.push(Column.rangeOf(start, prev));
    return ranges;
  }

  /** Expands `<col min max>` ranges back into individual Columns. */
  static fromModel(host: WorksheetHost, cols: ColumnRange[]): Column[] | null {
    if (cols.length === 0) return null;

    const columns: Column[] = [];
    for (const range of cols) {
      for (let n = range.min; n <= range.max; n++) {
        columns.push(
          new Column(host, n, {
            width: range.width,
            hidden: range.hidden,
            outlineLevel: range.outlineLevel,
            style: range.style,
          }),
        );
      }
    }
    return columns;
  }
}
