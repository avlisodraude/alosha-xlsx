/**
 * @alosha/xlsx — a modern TypeScript library for reading and writing XLSX
 * spreadsheets, inspired by exceljs.
 *
 * This is the public entry point: the Stage 1.2 document model plus the Stage 1.3 OOXML writer. The
 * writer is exposed two ways — the free functions from `./xlsx` (the environment-agnostic core) and
 * an ExcelJS-style `workbook.xlsx.writeBuffer/writeFile` accessor attached here at the composition
 * root, so the model layer keeps no dependency on the serializer.
 */

import { Workbook } from "./model/index.js";
import { type WriteOptions, writeWorkbookBuffer, writeWorkbookFile } from "./xlsx/index.js";

export * from "./address/index.js";
export * from "./errors.js";
export * from "./model/index.js";
export * from "./xlsx/index.js";

/** The ExcelJS-compatible serializer surface reached via `workbook.xlsx`. */
export interface XlsxAccessor {
  /** Serialize the workbook to `.xlsx` bytes. Async to mirror ExcelJS; wraps the sync core. */
  writeBuffer(options?: WriteOptions): Promise<Uint8Array>;
  /** Node-only: serialize and write to `path`. */
  writeFile(path: string, options?: WriteOptions): Promise<void>;
}

declare module "./model/workbook.js" {
  interface Workbook {
    /** ExcelJS-style serializer accessor (see {@link XlsxAccessor}). */
    readonly xlsx: XlsxAccessor;
  }
}

Object.defineProperty(Workbook.prototype, "xlsx", {
  configurable: true,
  get(this: Workbook): XlsxAccessor {
    return {
      writeBuffer: (options?: WriteOptions) => Promise.resolve(writeWorkbookBuffer(this, options)),
      writeFile: (path: string, options?: WriteOptions) => writeWorkbookFile(this, path, options),
    };
  },
});

/** The semantic version of the library. */
export const version = "0.0.0";
