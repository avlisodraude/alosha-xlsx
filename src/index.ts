/**
 * @alosha/xlsx — a modern TypeScript library for reading and writing XLSX
 * spreadsheets, inspired by exceljs.
 *
 * This is the public entry point: the core document model (Stage 1.2). Serialization (model →
 * OOXML → zip) lands in Stage 1.3.
 */

export * from "./address/index.js";
export * from "./errors.js";
export * from "./model/index.js";

/** The semantic version of the library. */
export const version = "0.0.0";
