// Row: a sparse, 1-indexed collection of Cells. Owns cell creation (the one runtime dependency
// among these modules: Row -> Cell), values get/set in all three RowValues forms, style fan-out
// to already-existing cells (docs/API-AUDIT.md §3's `_applyStyle`), and splice.

import { colLetterToNumber, decodeAddress, encodeAddress } from "../address/index.js";
import { InvalidRowNumberError } from "../errors.js";
import { Cell, type CellModel } from "./cell.js";
import type { CellValue } from "./cell-value.js";
import { CellType } from "./enums.js";
import type { Alignment, Borders, Fill, Font, Protection, Style } from "./style.js";
import type { RowValues } from "./types.js";
import type { WorksheetHost } from "./worksheet-host.js";

export interface RowModel {
  number: number;
  cells: CellModel[];
  height?: number;
  hidden?: boolean;
  outlineLevel?: number;
  style?: Style;
}

type EachCellCb = (cell: Cell, colNumber: number) => void;

export class Row {
  height?: number;
  hidden = false;
  outlineLevel = 0;

  /** Sparse, 1-indexed by column number; index 0 is never assigned. */
  private _cells: Cell[] = [];
  private _style: Style = {};

  constructor(
    public readonly worksheet: WorksheetHost,
    public readonly number: number,
  ) {}

  get collapsed(): boolean {
    return this.outlineLevel > this.worksheet.properties.outlineLevelRow;
  }

  private resolveColNumber(col: number | string): number {
    if (typeof col === "number") return col;
    const byKey = this.worksheet.getColumnKey(col);
    if (byKey) return byKey.number;
    return colLetterToNumber(col);
  }

  getCell(col: number | string): Cell {
    const colNumber = this.resolveColNumber(col);
    const existing = this._cells[colNumber];
    if (existing) return existing;

    const column = this.worksheet.getColumn(colNumber);
    const cell = new Cell(this, column, encodeAddress(this.number, colNumber));
    this._cells[colNumber] = cell;
    return cell;
  }

  findCell(col: number): Cell | undefined {
    return this._cells[col];
  }

  eachCell(cb: EachCellCb): void;
  eachCell(opts: { includeEmpty: boolean }, cb: EachCellCb): void;
  eachCell(optsOrCb: { includeEmpty: boolean } | EachCellCb, maybeCb?: EachCellCb): void {
    const includeEmpty = typeof optsOrCb === "object" ? optsOrCb.includeEmpty : false;
    const cb = typeof optsOrCb === "function" ? optsOrCb : (maybeCb as EachCellCb);

    if (includeEmpty) {
      const max = this._cells.length - 1;
      for (let col = 1; col <= max; col++) {
        cb(this.getCell(col), col);
      }
    } else {
      this._cells.forEach((cell, col) => {
        cb(cell, col);
      });
    }
  }

  get hasValues(): boolean {
    return this._cells.some((cell) => cell !== undefined && cell.type !== CellType.Null);
  }

  get cellCount(): number {
    return Math.max(this._cells.length - 1, 0);
  }

  get actualCellCount(): number {
    let count = 0;
    for (const cell of this._cells) if (cell !== undefined) count++;
    return count;
  }

  get dimensions(): { min: number; max: number } | null {
    let min = Number.POSITIVE_INFINITY;
    let max = Number.NEGATIVE_INFINITY;
    this._cells.forEach((_cell, col) => {
      if (col < min) min = col;
      if (col > max) max = col;
    });
    return min === Number.POSITIVE_INFINITY ? null : { min, max };
  }

  get values(): RowValues {
    const arr: CellValue[] = [];
    this._cells.forEach((cell, col) => {
      arr[col] = cell.value;
    });
    return arr;
  }

  set values(v: RowValues) {
    this._cells = [];
    if (v === undefined || v === null) return;

    if (Array.isArray(v)) {
      const offset = v[0] !== undefined ? 1 : 0;
      v.forEach((val, index) => {
        if (val === undefined) return;
        const col = index + offset;
        if (col < 1) return;
        this.getCell(col).value = val;
      });
      return;
    }

    for (const [key, val] of Object.entries(v)) {
      this.getCell(key).value = val;
    }
  }

  splice(start: number, count: number, ...inserts: CellValue[]): void {
    const currentMax = this._cells.length - 1;
    const insertCount = inserts.length;

    const tail: Array<{ value: CellValue; style: Style } | undefined> = [];
    for (let col = start + count; col <= currentMax; col++) {
      const cell = this._cells[col];
      tail[col - (start + count)] = cell ? { value: cell.value, style: cell.style } : undefined;
    }

    this._cells.length = Math.min(this._cells.length, start);

    inserts.forEach((val, i) => {
      this.getCell(start + i).value = val;
    });

    tail.forEach((entry, i) => {
      if (!entry) return;
      const cell = this.getCell(start + insertCount + i);
      cell.value = entry.value;
      cell.style = entry.style;
    });
  }

  private applyToCells(fn: (cell: Cell) => void): void {
    this._cells.forEach((cell) => {
      fn(cell);
    });
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

  get model(): RowModel {
    const cells: CellModel[] = [];
    this._cells.forEach((cell) => {
      cells.push(cell.model);
    });

    const m: RowModel = { number: this.number, cells };
    if (this.height !== undefined) m.height = this.height;
    if (this.hidden) m.hidden = true;
    if (this.outlineLevel !== 0) m.outlineLevel = this.outlineLevel;
    if (Object.keys(this._style).length > 0) m.style = this._style;
    return m;
  }

  set model(m: RowModel) {
    if (m.number !== this.number) {
      throw new InvalidRowNumberError(this.number, m.number);
    }
    this._cells = [];
    this.height = m.height;
    this.hidden = m.hidden ?? false;
    this.outlineLevel = m.outlineLevel ?? 0;
    this._style = m.style ?? {};

    for (const cellModel of m.cells) {
      const col = decodeAddress(cellModel.address).col;
      this.getCell(col).model = cellModel;
    }
  }
}
