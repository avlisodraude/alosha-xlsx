// The ExcelJS `Enums.ValueType`/`FormulaType` bridge (docs/COMPAT-SPEC.md §3). Our `CellType`
// shares names with ExcelJS's `ValueType` but not numbers; these maps translate between the two.
// `FormulaType` numbering already matches ExcelJS 1:1, so it's a straight re-export.

import { CellType } from "../model/index.js";

/** ExcelJS's `Enums.ValueType` numbering. */
export const ValueType = {
  Null: 0,
  Merge: 1,
  Number: 2,
  String: 3,
  Date: 4,
  Hyperlink: 5,
  Formula: 6,
  SharedString: 7,
  RichText: 8,
  Boolean: 9,
  Error: 10,
} as const;

export type ValueType = (typeof ValueType)[keyof typeof ValueType];

export { FormulaType } from "../model/index.js";

const TO_EXCELJS_VALUE_TYPE: Record<CellType, ValueType> = {
  [CellType.Null]: ValueType.Null,
  [CellType.Number]: ValueType.Number,
  [CellType.String]: ValueType.String,
  [CellType.Boolean]: ValueType.Boolean,
  [CellType.Date]: ValueType.Date,
  [CellType.Error]: ValueType.Error,
  [CellType.RichText]: ValueType.RichText,
  [CellType.Hyperlink]: ValueType.Hyperlink,
  [CellType.Formula]: ValueType.Formula,
  [CellType.Merge]: ValueType.Merge,
  [CellType.SharedString]: ValueType.SharedString,
};

const FROM_EXCELJS_VALUE_TYPE: Record<ValueType, CellType> = {
  [ValueType.Null]: CellType.Null,
  [ValueType.Merge]: CellType.Merge,
  [ValueType.Number]: CellType.Number,
  [ValueType.String]: CellType.String,
  [ValueType.Date]: CellType.Date,
  [ValueType.Hyperlink]: CellType.Hyperlink,
  [ValueType.Formula]: CellType.Formula,
  [ValueType.SharedString]: CellType.SharedString,
  [ValueType.RichText]: CellType.RichText,
  [ValueType.Boolean]: CellType.Boolean,
  [ValueType.Error]: CellType.Error,
};

/** `cell.type` in compat mode: our `CellType` → ExcelJS's `ValueType` number. */
export function toExcelJSValueType(cellType: CellType): ValueType {
  return TO_EXCELJS_VALUE_TYPE[cellType];
}

/** The inverse of {@link toExcelJSValueType}. */
export function fromExcelJSValueType(value: ValueType): CellType {
  return FROM_EXCELJS_VALUE_TYPE[value];
}
