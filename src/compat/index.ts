/**
 * @alosha/xlsx/compat — the ExcelJS drop-in shim entry point (docs/COMPAT-SPEC.md).
 *
 * The default export is shaped like the `exceljs` namespace (`{ Workbook, ValueType, Enums, stream }`)
 * and the named `Workbook` is the ExcelJS-faithful facade (`CompatWorkbook`, §5): compat-constructed
 * workbooks report ExcelJS `ValueType` numbers, apply whole-object style precedence, and throw
 * `UnsupportedFeatureError` on the deferred surface (§6). An ExcelJS user changes one import.
 *
 * Re-exports from `../index.js` (not `../model/index.js`) so the `workbook.xlsx` accessor — attached
 * to `Workbook.prototype` as a side effect of loading the main entry — is present on compat workbooks.
 * The star export's plain `Workbook` is deliberately shadowed by the explicit `CompatWorkbook as
 * Workbook` re-export below (explicit named exports win over `export *`).
 */

import { FormulaType, ValueType } from "./value-type.js";
import { CompatWorkbook, compatStream } from "./workbook.js";

export * from "../index.js";
export { fromExcelJSValueType, toExcelJSValueType, ValueType } from "./value-type.js";
export { CompatWorkbook, CompatWorkbook as Workbook } from "./workbook.js";

/** ExcelJS's `Enums` grouping (`import { Enums } from "exceljs"`). */
export const Enums = {
  ValueType,
  FormulaType,
} as const;

/** The ExcelJS-shaped default export: `import ExcelJS from "@alosha/xlsx/compat"`. */
const ExcelJS = {
  Workbook: CompatWorkbook,
  ValueType,
  FormulaType,
  Enums,
  stream: compatStream,
};

export default ExcelJS;
