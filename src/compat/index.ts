/**
 * @alosha/xlsx/compat — the ExcelJS drop-in shim entry point (docs/COMPAT-SPEC.md).
 *
 * Session 1 scaffold: enum bridge + subpath wiring only. The default export re-exports the plain
 * model classes as-is; the ExcelJS-faithful facade (cell.type emulation, style precedence,
 * unsupported-method stubs) lands in a later session — see COMPAT-SPEC.md §5.
 *
 * Re-exports from `../index.js` (not `../model/index.js`) so the `workbook.xlsx` accessor — attached
 * to `Workbook.prototype` as a side effect of loading the main entry — is present on compat workbooks.
 */

import { Workbook } from "../index.js";
import { FormulaType, ValueType } from "./value-type.js";

export * from "../index.js";
export { fromExcelJSValueType, toExcelJSValueType, ValueType } from "./value-type.js";

/** ExcelJS's `Enums` grouping (`import { Enums } from "exceljs"`). */
export const Enums = {
  ValueType,
  FormulaType,
} as const;

/** The ExcelJS-shaped default export: `import ExcelJS from "@alosha/xlsx/compat"`. */
const ExcelJS = {
  Workbook,
  ValueType,
  FormulaType,
  Enums,
};

export default ExcelJS;
