import type { CellValue } from "./cell-value.js";

/** A row's cell values: a contiguous/sparse array, a column-key-keyed object, or absent. */
export type RowValues = CellValue[] | { [key: string]: CellValue } | undefined | null;
