// Public barrel for the streaming writer (Track E). A self-contained module layered beside the buffered
// writer — see `workbook-writer.ts` for the ExcelJS-shaped `WorkbookWriter` API and the memory rationale.

export type { ByteSink, NodeWritableLike } from "./output.js";
export {
  type StreamedRow,
  WorkbookWriter,
  type WorkbookWriterOptions,
  WorksheetWriter,
} from "./workbook-writer.js";
