// Data validation: cell dropdowns and value constraints (docs/DATA-VALIDATION-SPEC.md §1).
//
// A validation is inherently range-based, so it is stored per worksheet keyed by its `sqref`
// (a single address, a range, or space-separated ranges) rather than on individual cells. The
// `DataValidations` manager mirrors ExcelJS's shape (`add`/`find`/`remove`/`model`) so the compat
// facade can pass straight through; `cell.dataValidation` is a thin accessor over it.

import { Range } from "../address/index.js";

export type DataValidationType =
  | "list"
  | "whole"
  | "decimal"
  | "date"
  | "time"
  | "textLength"
  | "custom";

export type DataValidationOperator =
  | "between"
  | "notBetween"
  | "equal"
  | "notEqual"
  | "greaterThan"
  | "lessThan"
  | "greaterThanOrEqual"
  | "lessThanOrEqual";

export type DataValidationErrorStyle = "stop" | "warning" | "information";

/** The types that carry a comparison `operator` (`list`/`custom` never do). Shared by the writer
 *  (which omits `operator` outside this set) and the reader (which only reads it back for these). */
export const DATA_VALIDATION_OPERATOR_TYPES = new Set<DataValidationType>([
  "whole",
  "decimal",
  "date",
  "time",
  "textLength",
]);

/** Mirrors ExcelJS's `cell.dataValidation` shape (so the compat facade is a pass-through). */
export interface DataValidation {
  type: DataValidationType;
  operator?: DataValidationOperator;
  allowBlank?: boolean;
  showInputMessage?: boolean;
  showErrorMessage?: boolean;
  /** `[formula1]` or `[formula1, formula2]`. For `list`: an inline `'"a,b,c"'` or a range `"$E$1:$E$5"`. */
  formulae?: (string | number)[];
  showDropDown?: boolean; // list only; NOTE the OOXML attribute is inverted (see spec §4)
  errorStyle?: DataValidationErrorStyle;
  errorTitle?: string;
  error?: string;
  promptTitle?: string;
  prompt?: string;
}

/** A validation plus the `sqref` it applies to — the serializable form on `WorksheetModel`. */
export type DataValidationModel = { sqref: string } & DataValidation;

/** Whether `address` (e.g. "A1") falls within a `sqref`, which may be space-separated ranges. */
function sqrefContains(sqref: string, address: string): boolean {
  for (const part of sqref.trim().split(/\s+/)) {
    if (part === "") continue;
    try {
      if (new Range(part).contains(address)) return true;
    } catch {
      // Malformed range fragment — skip it rather than throw.
    }
  }
  return false;
}

/**
 * Per-worksheet store of data validations keyed by `sqref`. Mirrors ExcelJS's `DataValidations`
 * so `worksheet.dataValidations` is a straight pass-through in the compat facade.
 */
export class DataValidations {
  private readonly _map = new Map<string, DataValidation>();

  /** Register `dv` for `sqref` (a single address, a range, or space-separated ranges). */
  add(sqref: string, dv: DataValidation): DataValidation {
    this._map.set(sqref, dv);
    return dv;
  }

  /** The validation whose `sqref` covers `address`, if any (exact key first, then range match). */
  find(address: string): DataValidation | undefined {
    const direct = this._map.get(address);
    if (direct) return direct;
    for (const [sqref, dv] of this._map) {
      if (sqrefContains(sqref, address)) return dv;
    }
    return undefined;
  }

  remove(sqref: string): void {
    this._map.delete(sqref);
  }

  get size(): number {
    return this._map.size;
  }

  clear(): void {
    this._map.clear();
  }

  get model(): DataValidationModel[] {
    return [...this._map.entries()].map(([sqref, dv]) => ({ sqref, ...dv }));
  }

  set model(models: DataValidationModel[]) {
    this._map.clear();
    for (const { sqref, ...dv } of models) this._map.set(sqref, dv);
  }
}
