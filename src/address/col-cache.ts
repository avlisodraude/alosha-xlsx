// A1 <-> {row,col} address math. Pure, hot-path, memoized. Port of ExcelJS's col-cache behavior
// (see docs/API-AUDIT.md §5) as typed free functions instead of a class.

import { InvalidAddressError } from "../errors.js";
import { assertColumnNumber, assertRowNumber, MAX_COL, MAX_ROW } from "./limits.js";

/** 1-based rows/cols everywhere in this module's public surface. */
export interface Decoded {
  row: number;
  col: number;
  address: string;
}

/** `decodeEx` result: adds the absolute `$col$row` form and an optional sheet qualifier. */
export interface DecodedEx extends Decoded {
  $col$row: string;
  sheetName?: string;
}

const LETTER_RE = /^[A-Za-z]+$/;
const ADDRESS_RE = /^([A-Za-z]+)(\d+)$/;
const SHEET_PREFIX_RE = /^(?:'([^']*)'|([^'!]+))!/;
const CELL_REF_RE = /^\$?([A-Za-z]+)\$?(\d+)$/;

/** Upper bound on {@link decodeAddress}'s memo. The column caches below are naturally bounded by
 *  Excel's 16,384-column cap, but this one is keyed by full cell address — distinct cells across
 *  large or many workbooks (up to 16,384 × 1,048,576) would otherwise grow it without limit, which
 *  was a real process-global leak. When the cap is reached the oldest entry is evicted (FIFO): a
 *  bulk read decodes each address once regardless, and realistic repeated-access working sets stay
 *  well under this bound. */
export const DECODE_ADDRESS_CACHE_MAX = 100_000;

const colLetterToNumberCache = new Map<string, number>();
const colNumberToLetterCache = new Map<number, string>();
const decodeAddressCache = new Map<string, Decoded>();

/** 'A'->1, 'Z'->26, 'AA'->27. Throws {@link InvalidAddressError} on non-letters. */
export function colLetterToNumber(letter: string): number {
  const cached = colLetterToNumberCache.get(letter);
  if (cached !== undefined) return cached;

  if (letter.length > 3 || !LETTER_RE.test(letter)) {
    throw new InvalidAddressError(letter);
  }

  const upper = letter.toUpperCase();
  let col = 0;
  for (let i = 0; i < upper.length; i++) {
    col = col * 26 + (upper.charCodeAt(i) - 64);
  }

  assertColumnNumber(col);
  colLetterToNumberCache.set(letter, col);
  return col;
}

/** 1->'A', 27->'AA'. */
export function colNumberToLetter(col: number): string {
  const cached = colNumberToLetterCache.get(col);
  if (cached !== undefined) return cached;

  assertColumnNumber(col);

  let n = col;
  let letters = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }

  colNumberToLetterCache.set(col, letters);
  return letters;
}

/** (row=1,col=2) -> 'B1'. */
export function encodeAddress(row: number, col: number): string {
  assertRowNumber(row);
  return `${colNumberToLetter(col)}${row}`;
}

/** 'B1' -> { row:1, col:2, address:'B1' }. Cached. */
export function decodeAddress(address: string): Decoded {
  const cached = decodeAddressCache.get(address);
  if (cached !== undefined) return cached;

  const match = ADDRESS_RE.exec(address);
  if (!match) {
    throw new InvalidAddressError(address);
  }

  const [, letters, digits] = match as unknown as [string, string, string];
  const col = colLetterToNumber(letters);
  const row = Number.parseInt(digits, 10);

  if (row < 1 || row > MAX_ROW || col < 1 || col > MAX_COL) {
    throw new InvalidAddressError(address);
  }

  const decoded: Decoded = { row, col, address: `${letters.toUpperCase()}${row}` };
  if (decodeAddressCache.size >= DECODE_ADDRESS_CACHE_MAX) {
    // FIFO eviction: a Map iterates keys in insertion order, so the first is the oldest.
    const oldest = decodeAddressCache.keys().next().value;
    if (oldest !== undefined) decodeAddressCache.delete(oldest);
  }
  decodeAddressCache.set(address, decoded);
  return decoded;
}

/** Handles '$B$1', "'Sheet 1'!B1", Sheet1!B1, absolute/relative `$` markers. */
export function decodeEx(address: string): DecodedEx {
  let sheetName: string | undefined;
  let rest = address;

  const sheetMatch = SHEET_PREFIX_RE.exec(address);
  if (sheetMatch) {
    sheetName = sheetMatch[1] ?? sheetMatch[2];
    rest = address.slice(sheetMatch[0].length);
  }

  const cellMatch = CELL_REF_RE.exec(rest);
  if (!cellMatch) {
    throw new InvalidAddressError(address);
  }

  const [, letters, digits] = cellMatch as unknown as [string, string, string];
  const decoded = decodeAddress(`${letters}${digits}`);

  return {
    ...decoded,
    $col$row: `$${colNumberToLetter(decoded.col)}$${decoded.row}`,
    ...(sheetName !== undefined ? { sheetName } : {}),
  };
}

/** Throws {@link InvalidAddressError} if not a well-formed A1 cell address. */
export function validateAddress(address: string): void {
  decodeAddress(address);
}
