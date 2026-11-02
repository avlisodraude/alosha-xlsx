import { InvalidAddressError, InvalidRangeError } from "../errors.js";

export const MAX_ROW = 1_048_576;
export const MAX_COL = 16_384;

export function assertRowNumber(row: number): void {
  if (!Number.isInteger(row) || row < 1 || row > MAX_ROW) {
    throw new InvalidAddressError(String(row));
  }
}

export function assertColumnNumber(col: number): void {
  if (!Number.isInteger(col) || col < 1 || col > MAX_COL) {
    throw new InvalidAddressError(String(col));
  }
}

/** Validate a splice's numeric extent before any model state is moved or removed. */
export function assertSpliceBounds(
  start: number,
  count: number,
  inserts: number,
  last: number,
  maximum: number,
): void {
  if (
    !Number.isInteger(count) ||
    count < 0 ||
    count > maximum - start + 1 ||
    (inserts > 0 && start + inserts - 1 > maximum) ||
    (last >= start + count && last + inserts - count > maximum)
  ) {
    throw new InvalidRangeError([start, count, inserts]);
  }
}
