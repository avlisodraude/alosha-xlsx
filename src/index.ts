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
import {
  type ReadOptions,
  readWorkbookBuffer,
  readWorkbookFile,
  type WriteOptions,
  writeWorkbookBuffer,
  writeWorkbookFile,
} from "./xlsx/index.js";

export * from "./address/index.js";
export * from "./errors.js";
export * from "./model/index.js";
export * from "./xlsx/index.js";

/** A source of chunks for {@link XlsxAccessor.read}; a Node `Readable` satisfies this structurally. */
type ByteStream = AsyncIterable<Uint8Array>;

/** The ExcelJS-compatible read/write surface reached via `workbook.xlsx`. `load`/`readFile`/`read`
 *  mutate the receiving workbook in place and return it, matching ExcelJS. */
export interface XlsxAccessor {
  /** Serialize the workbook to `.xlsx` bytes. Async to mirror ExcelJS; wraps the sync core. */
  writeBuffer(options?: WriteOptions): Promise<Uint8Array>;
  /** Node-only: serialize and write to `path`. */
  writeFile(path: string, options?: WriteOptions): Promise<void>;
  /** Parse `.xlsx` bytes into this workbook (replaces its contents), returning it. */
  load(data: Uint8Array, options?: ReadOptions): Promise<Workbook>;
  /** Node-only: read and parse a `.xlsx` from `path` into this workbook, returning it. */
  readFile(path: string, options?: ReadOptions): Promise<Workbook>;
  /** Drain a byte stream (e.g. a Node `Readable`) and parse it into this workbook, returning it. */
  read(stream: ByteStream, options?: ReadOptions): Promise<Workbook>;
}

declare module "./model/workbook.js" {
  interface Workbook {
    /** ExcelJS-style read/write accessor (see {@link XlsxAccessor}). */
    readonly xlsx: XlsxAccessor;
  }
}

/** Concatenate a byte stream's chunks into a single buffer (no `fs`; browser-safe). */
async function drainStream(stream: ByteStream): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  let total = 0;
  for await (const chunk of stream) {
    chunks.push(chunk);
    total += chunk.length;
  }
  const out = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.length;
  }
  return out;
}

Object.defineProperty(Workbook.prototype, "xlsx", {
  configurable: true,
  get(this: Workbook): XlsxAccessor {
    return {
      writeBuffer: (options?: WriteOptions) => Promise.resolve(writeWorkbookBuffer(this, options)),
      writeFile: (path: string, options?: WriteOptions) => writeWorkbookFile(this, path, options),
      load: (data: Uint8Array, options?: ReadOptions) => {
        this.model = readWorkbookBuffer(data, options).model;
        return Promise.resolve(this);
      },
      readFile: async (path: string, options?: ReadOptions) => {
        this.model = (await readWorkbookFile(path, options)).model;
        return this;
      },
      read: async (stream: ByteStream, options?: ReadOptions) => {
        this.model = readWorkbookBuffer(await drainStream(stream), options).model;
        return this;
      },
    };
  },
});

// Replaced at build time by tsup's `define` (and by vitest's `define` under test) with the real
// package.json version. `typeof` guards against a ReferenceError in any context where neither define
// ran, falling back to the sentinel.
declare const __PKG_VERSION__: string;

/** The semantic version of the library, injected from `package.json` at build time. */
export const version: string = typeof __PKG_VERSION__ === "undefined" ? "0.0.0" : __PKG_VERSION__;
