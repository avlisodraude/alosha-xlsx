// Type-only contract. The package entry installs the runtime accessor on Workbook.prototype.
import type { Workbook } from "./model/workbook.js";
import type { ReadOptions } from "./xlsx/read-workbook.js";
import type { WriteOptions } from "./xlsx/write-workbook.js";

/** A Node Readable or another asynchronous source of byte chunks. */
export type ByteStream = AsyncIterable<Uint8Array>;

/** ExcelJS-compatible buffered I/O. Reads replace and return the receiving workbook. */
export interface XlsxAccessor {
  writeBuffer(options?: WriteOptions): Promise<Uint8Array>;
  writeFile(path: string, options?: WriteOptions): Promise<void>;
  load(data: Uint8Array, options?: ReadOptions): Promise<Workbook>;
  readFile(path: string, options?: ReadOptions): Promise<Workbook>;
  read(stream: ByteStream, options?: ReadOptions): Promise<Workbook>;
}
