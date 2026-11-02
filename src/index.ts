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
import { checkReadLimit, resolveReadLimits } from "./read-limits.js";
import {
  type ReadOptions,
  readWorkbookBuffer,
  readWorkbookFile,
  type WriteOptions,
  writeWorkbookBuffer,
  writeWorkbookFile,
} from "./xlsx/index.js";
import type { ByteStream, XlsxAccessor } from "./xlsx-accessor.js";

export * from "./address/index.js";
export * from "./errors.js";
export * from "./model/index.js";
export { DEFAULT_READ_LIMITS, type ReadLimits } from "./read-limits.js";
export * from "./xlsx/index.js";
export type { XlsxAccessor } from "./xlsx-accessor.js";

/** Concatenate a byte stream's chunks into a single buffer (no `fs`; browser-safe). */
async function drainStream(stream: ByteStream, options?: ReadOptions): Promise<Uint8Array> {
  const limits = resolveReadLimits(options);
  const chunks: Uint8Array[] = [];
  let total = 0;
  for await (const chunk of stream) {
    chunks.push(chunk);
    total += chunk.length;
    checkReadLimit("maxInputBytes", limits.maxInputBytes, total);
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
      load: async (data: Uint8Array, options?: ReadOptions) => {
        this.model = readWorkbookBuffer(data, options).model;
        return this;
      },
      readFile: async (path: string, options?: ReadOptions) => {
        this.model = (await readWorkbookFile(path, options)).model;
        return this;
      },
      read: async (stream: ByteStream, options?: ReadOptions) => {
        this.model = readWorkbookBuffer(await drainStream(stream, options), options).model;
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
