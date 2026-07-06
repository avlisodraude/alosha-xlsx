// Public barrel for the OOXML writer (Stage 1.3). The model -> `.xlsx` core lives behind
// `writeWorkbookBuffer`; `src/index.ts` layers the ExcelJS-style `workbook.xlsx.*` accessor on top.

export { dateToSerial, serialToDate } from "./date.js";
export { createSharedStrings, type SharedStrings } from "./shared-strings.js";
export { createStyles, type Styles } from "./styles.js";
export {
  type WriteOptions,
  writeWorkbookBuffer,
  writeWorkbookFile,
} from "./write-workbook.js";
