// The narrow surface Row/Column need from a worksheet, so neither has to import the concrete
// Worksheet class (session 4). The real Worksheet will implement this interface.
//
// `getRow` is one method beyond the minimal column/key/eachRow surface: Column#header writes into
// the top row cell(s) (SPEC §6 — "header assignment writes the top row cells' values"), which is
// unreachable through eachRow alone (it only visits rows that already exist). Worksheet already
// exposes `getRow` as part of its own public API (SPEC §7), so implementing it here costs nothing.

import type { Column } from "./column.js";
import type { Row } from "./row.js";

export interface WorksheetHost {
  getColumn(col: number | string): Column;
  getColumnKey(key: string): Column | undefined;
  getRow(row: number): Row;
  eachRow(cb: (row: Row, n: number) => void): void;
  eachRow(opts: { includeEmpty: boolean }, cb: (row: Row, n: number) => void): void;
  readonly properties: { outlineLevelCol: number; outlineLevelRow: number } & Record<
    string,
    unknown
  >;
}
