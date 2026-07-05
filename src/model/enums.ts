// Cell-type and formula-type enums, plus the OOXML error-code union. Keeps the *set* of value
// types from ExcelJS (see docs/API-AUDIT.md §ValueType) but drops its numeric-wire coupling — the
// OOXML `t=` mapping lives in the Stage 1.3 serializer, not here.

export enum CellType {
  Null = 0,
  Number = 1,
  String = 2,
  Boolean = 3,
  Date = 4,
  Error = 5,
  RichText = 6,
  Hyperlink = 7,
  Formula = 8,
  Merge = 9,
  SharedString = 10,
}

export enum FormulaType {
  None = 0,
  Master = 1,
  Shared = 2,
}

export type ErrorCode = "#N/A" | "#REF!" | "#NAME?" | "#DIV/0!" | "#NULL!" | "#VALUE!" | "#NUM!";
