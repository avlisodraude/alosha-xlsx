// Cell: value get/set through the CellValue <-> StoredValue helpers (no per-type strategy
// classes — see docs/API-AUDIT.md §5), merge bookkeeping, and the style shortcuts. Style is
// resolved from row+column defaults once at construction (docs/API-AUDIT.md §4's `_mergeStyle`
// note); further per-field writes go straight into the cell's own style object.
//
// `Row`/`Column` are type-only imports: Cell never constructs either (both are handed to it by
// their owning Row), so this carries no runtime dependency and cannot cycle with row.ts's
// `import { Cell }`.

import {
  type CellFormula,
  type CellValue,
  effectiveType as effectiveTypeOf,
  fromStored,
  type StoredValue,
  toCsv,
  toStored,
  toText,
} from "./cell-value.js";
import type { Column } from "./column.js";
import type { DataValidation } from "./data-validation.js";
import { CellType, FormulaType } from "./enums.js";
import type { Row } from "./row.js";
import {
  type Alignment,
  type Borders,
  type Fill,
  type Font,
  type Protection,
  resolveCellStyle,
  type Style,
} from "./style.js";

export interface CellModel {
  address: string;
  type: CellType;
  value?: unknown;
  formula?: string;
  result?: unknown;
  style?: Style;
}

export class Cell {
  private _stored: StoredValue = { type: CellType.Null };
  private _style: Style;
  private _master: Cell;

  constructor(
    private readonly rowRef: Row,
    private readonly columnRef: Column,
    public readonly address: string,
  ) {
    this._master = this;
    this._style = resolveCellStyle(rowRef.style, columnRef.style);
  }

  get row(): number {
    return this.rowRef.number;
  }

  get col(): number {
    return this.columnRef.number;
  }

  get $col$row(): string {
    return `$${this.columnRef.letter}$${this.rowRef.number}`;
  }

  /** The stored value this cell resolves through: itself, or (if merged) its master. */
  private get resolvedStored(): StoredValue {
    return this.isMerged ? this._master.resolvedStored : this._stored;
  }

  get value(): CellValue {
    return fromStored(this.resolvedStored);
  }

  set value(v: CellValue) {
    if (this.isMerged) {
      this._master.value = v;
      return;
    }
    this._stored = toStored(v);
  }

  /** This cell's own marker type — `Merge` for a slave, even though `value`/`text` delegate. */
  get type(): CellType {
    return this._stored.type;
  }

  get effectiveType(): CellType {
    return effectiveTypeOf(this.resolvedStored);
  }

  get text(): string {
    return toText(this.resolvedStored);
  }

  toString(): string {
    return this.text;
  }

  toCsvString(): string {
    return toCsv(this.resolvedStored);
  }

  get isMerged(): boolean {
    return this._master !== this;
  }

  /** The master cell of this cell's merge region; itself if unmerged. */
  get master(): Cell {
    return this._master;
  }

  merge(master: Cell, ignoreStyle = false): void {
    this._master = master;
    this._stored = { type: CellType.Merge, master: master.address };
    if (!ignoreStyle) this._style = master.style;
  }

  unmerge(): void {
    this._master = this;
    this._stored = { type: CellType.Null };
  }

  get formula(): string | undefined {
    const s = this.resolvedStored;
    return s.type === CellType.Formula ? s.value.formula : undefined;
  }

  get result(): unknown {
    const s = this.resolvedStored;
    return s.type === CellType.Formula ? s.value.result : undefined;
  }

  get formulaType(): FormulaType {
    const s = this.resolvedStored;
    if (s.type !== CellType.Formula) return FormulaType.None;
    return s.value.sharedFormula !== undefined ? FormulaType.Shared : FormulaType.Master;
  }

  get style(): Style {
    return this._style;
  }

  set style(s: Style) {
    this._style = s;
  }

  get numFmt(): string | undefined {
    return this._style.numFmt;
  }

  set numFmt(v: string | undefined) {
    this._style = { ...this._style, numFmt: v };
  }

  get font(): Font | undefined {
    return this._style.font;
  }

  set font(v: Font | undefined) {
    this._style = { ...this._style, font: v };
  }

  get alignment(): Alignment | undefined {
    return this._style.alignment;
  }

  set alignment(v: Alignment | undefined) {
    this._style = { ...this._style, alignment: v };
  }

  get border(): Borders | undefined {
    return this._style.border;
  }

  set border(v: Borders | undefined) {
    this._style = { ...this._style, border: v };
  }

  get fill(): Fill | undefined {
    return this._style.fill;
  }

  set fill(v: Fill | undefined) {
    this._style = { ...this._style, fill: v };
  }

  get protection(): Protection | undefined {
    return this._style.protection;
  }

  set protection(v: Protection | undefined) {
    this._style = { ...this._style, protection: v };
  }

  /** The data validation covering this cell, if any (resolved through the worksheet's store). */
  get dataValidation(): DataValidation | undefined {
    return this.rowRef.worksheet.dataValidations.find(this.address);
  }

  set dataValidation(dv: DataValidation) {
    this.rowRef.worksheet.dataValidations.add(this.address, dv);
  }

  get model(): CellModel {
    const m: CellModel = { address: this.address, type: this.type };
    switch (this._stored.type) {
      case CellType.Null:
      case CellType.Merge:
        break;
      case CellType.Formula:
        if (this._stored.value.formula !== undefined) m.formula = this._stored.value.formula;
        if (this._stored.value.result !== undefined) m.result = this._stored.value.result;
        break;
      default:
        m.value = fromStored(this._stored);
    }
    if (Object.keys(this._style).length > 0) m.style = this._style;
    return m;
  }

  set model(m: CellModel) {
    // Model assignment always targets this cell directly — an implicit unmerge, mirroring how
    // ExcelJS reloads a cell wholesale from serialized state.
    this._master = this;
    switch (m.type) {
      case CellType.Merge:
        this._stored = { type: CellType.Merge, master: this.address };
        break;
      case CellType.Formula:
        this._stored = toStored({
          formula: m.formula,
          result: m.result as CellFormula["result"],
        });
        break;
      case CellType.Null:
        this._stored = { type: CellType.Null };
        break;
      default:
        this.value = m.value as CellValue;
    }
    this._style = m.style ?? {};
  }
}
