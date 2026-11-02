// Excel serial-date <-> JS Date conversion (see docs/STAGE-1.3-SPEC.md §6, docs/OOXML-NOTES.md §8).
// Self-contained: no dayjs. Excel stores dates as day serials, and we honor both epoch systems plus
// the deliberate 1900 leap-year bug that Excel preserves for Lotus 1-2-3 compatibility.

import { InvalidCellValueError } from "../errors.js";
import { assertFiniteScalar } from "../model/cell-value.js";

const MS_PER_DAY = 86_400_000;

/** 1900 date system: this UTC instant is serial 0, so 1899-12-31 + serial days => the date.
 *  serial 1 = 1900-01-01, serial 59 = 1900-02-28. */
const EPOCH_1900 = Date.UTC(1899, 11, 31);

/** 1904 date system: serial 0 = 1904-01-01 (Mac Excel's default). */
const EPOCH_1904 = Date.UTC(1904, 0, 1);

/** Excel treats 1900 as a leap year, so serial 60 is the fictional 1900-02-29. Real dates on or after
 *  1900-03-01 are shifted one day up to make room for it; conversions bridge that gap here. */
const LEAP_BUG_SERIAL = 60;

/** Convert a `Date` to its Excel day serial (fractional part carries the time of day). */
export function dateToSerial(d: Date, date1904 = false): number {
  if (!(d instanceof Date)) throw new InvalidCellValueError(d);
  assertFiniteScalar(d);
  if (date1904) {
    return (d.getTime() - EPOCH_1904) / MS_PER_DAY;
  }
  const raw = (d.getTime() - EPOCH_1900) / MS_PER_DAY;
  // Real dates from 1900-03-01 (raw 60) onward sit one serial higher because of the phantom 02-29.
  return raw >= LEAP_BUG_SERIAL ? raw + 1 : raw;
}

/** Convert an Excel day serial back to a `Date` (UTC-based instant). */
export function serialToDate(serial: number, date1904 = false): Date {
  if (typeof serial !== "number") throw new InvalidCellValueError(serial);
  assertFiniteScalar(serial);
  if (date1904) {
    const date = new Date(EPOCH_1904 + serial * MS_PER_DAY);
    if (!Number.isFinite(date.getTime())) throw new InvalidCellValueError(serial);
    return date;
  }
  // Undo the phantom-day shift for serials past the bug; serial 60 itself resolves to 1900-02-29,
  // which JS normalizes to 1900-03-01 (the same instant serial 61 also maps to).
  const adjusted = serial > LEAP_BUG_SERIAL ? serial - 1 : serial;
  const date = new Date(EPOCH_1900 + adjusted * MS_PER_DAY);
  if (!Number.isFinite(date.getTime())) throw new InvalidCellValueError(serial);
  return date;
}
