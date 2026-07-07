// Public barrel for the OOXML reader + writer (Stage 1.3 write, Stage 2 read). The model <-> `.xlsx`
// core lives behind `writeWorkbookBuffer`/`readWorkbookBuffer`; `src/index.ts` layers the ExcelJS-style
// `workbook.xlsx.*` accessor on top.

export { dateToSerial, serialToDate } from "./date.js";
export {
  type ReadOptions,
  readWorkbookBuffer,
  readWorkbookFile,
} from "./read-workbook.js";
export { createSharedStrings, type SharedStrings } from "./shared-strings.js";
export {
  type ByteSink,
  type NodeWritableLike,
  type StreamedRow,
  WorkbookWriter,
  type WorkbookWriterOptions,
  WorksheetWriter,
} from "./stream/index.js";
export { createStyles, type Styles } from "./styles.js";
export {
  type WriteOptions,
  writeWorkbookBuffer,
  writeWorkbookFile,
} from "./write-workbook.js";
