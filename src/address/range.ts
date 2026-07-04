// Range: a normalized bounding box over cell addresses, plus set operations. Used by
// Worksheet.dimensions and merge bookkeeping (see docs/STAGE-1.2-SPEC.md §2).

import { InvalidRangeError } from "../errors.js";
import { type Decoded, decodeEx, encodeAddress } from "./col-cache.js";

export type RangeSource =
  | string // 'A1:B2' or 'A1'
  | [string, string] // ['A1','B2']
  | [number, number, number, number] // [top,left,bottom,right]
  | Range
  | Decoded;

interface Box {
  top: number;
  left: number;
  bottom: number;
  right: number;
  sheetName?: string;
}

function isDecodedLike(value: unknown): value is Decoded {
  return (
    typeof value === "object" &&
    value !== null &&
    !(value instanceof Range) &&
    typeof (value as Decoded).row === "number" &&
    typeof (value as Decoded).col === "number" &&
    typeof (value as Decoded).address === "string"
  );
}

function assertValidBox(box: Box, source: unknown): void {
  for (const n of [box.top, box.left, box.bottom, box.right]) {
    if (!Number.isInteger(n) || n < 1) {
      throw new InvalidRangeError(source);
    }
  }
}

function decodeTwoRefs(a: string, b: string, source: unknown): Box {
  try {
    const d1 = decodeEx(a);
    const d2 = decodeEx(b);
    return {
      top: d1.row,
      left: d1.col,
      bottom: d2.row,
      right: d2.col,
      sheetName: d1.sheetName ?? d2.sheetName,
    };
  } catch {
    throw new InvalidRangeError(source);
  }
}

function parseRangeString(source: string): Box {
  const parts = source.split(":");
  if (parts.length === 1) {
    return decodeTwoRefs(parts[0] as string, parts[0] as string, source);
  }
  if (parts.length === 2) {
    const [p0, p1] = parts as [string, string];
    if (p0 === "" || p1 === "") throw new InvalidRangeError(source);
    return decodeTwoRefs(p0, p1, source);
  }
  throw new InvalidRangeError(source);
}

function resolveSource(source: RangeSource): Box {
  if (source instanceof Range) {
    return {
      top: source.top,
      left: source.left,
      bottom: source.bottom,
      right: source.right,
      sheetName: source.sheetName,
    };
  }

  if (typeof source === "string") {
    return parseRangeString(source);
  }

  if (Array.isArray(source)) {
    if (source.length === 4 && source.every((n) => typeof n === "number")) {
      const [top, left, bottom, right] = source as [number, number, number, number];
      return { top, left, bottom, right };
    }
    if (source.length === 2 && source.every((s) => typeof s === "string")) {
      const [a, b] = source as [string, string];
      return decodeTwoRefs(a, b, source);
    }
    throw new InvalidRangeError(source);
  }

  if (isDecodedLike(source)) {
    return { top: source.row, left: source.col, bottom: source.row, right: source.col };
  }

  throw new InvalidRangeError(source);
}

/** A normalized bounding box (`top<=bottom`, `left<=right`) over 1-based cell addresses. */
export class Range {
  top: number;
  left: number;
  bottom: number;
  right: number;
  sheetName?: string;

  constructor(source?: RangeSource);
  constructor(top: number, left: number, bottom: number, right: number);
  constructor(sourceOrTop?: RangeSource | number, left?: number, bottom?: number, right?: number) {
    let box: Box;

    if (sourceOrTop === undefined) {
      box = { top: 1, left: 1, bottom: 1, right: 1 };
    } else if (typeof sourceOrTop === "number") {
      if (left === undefined || bottom === undefined || right === undefined) {
        throw new InvalidRangeError([sourceOrTop, left, bottom, right]);
      }
      box = { top: sourceOrTop, left, bottom, right };
      assertValidBox(box, [sourceOrTop, left, bottom, right]);
    } else {
      box = resolveSource(sourceOrTop);
      assertValidBox(box, sourceOrTop);
    }

    this.top = Math.min(box.top, box.bottom);
    this.bottom = Math.max(box.top, box.bottom);
    this.left = Math.min(box.left, box.right);
    this.right = Math.max(box.left, box.right);
    if (box.sheetName !== undefined) this.sheetName = box.sheetName;
  }

  /** Top-left address, e.g. 'A1'. */
  get tl(): string {
    return encodeAddress(this.top, this.left);
  }

  /** Bottom-right address, e.g. 'B2'. */
  get br(): string {
    return encodeAddress(this.bottom, this.right);
  }

  /** 'A1:B2'. */
  get range(): string {
    return `${this.tl}:${this.br}`;
  }

  /** Total cell count covered by the box. */
  get count(): number {
    return (this.bottom - this.top + 1) * (this.right - this.left + 1);
  }

  /** Grows the box to include the given box (union of bounds). */
  expand(top: number, left: number, bottom: number, right: number): void {
    this.top = Math.min(this.top, top);
    this.left = Math.min(this.left, left);
    this.bottom = Math.max(this.bottom, bottom);
    this.right = Math.max(this.right, right);
  }

  /** Whether the given address falls within this box. */
  contains(address: string): boolean {
    const d = decodeEx(address);
    return d.row >= this.top && d.row <= this.bottom && d.col >= this.left && d.col <= this.right;
  }

  /** Whether this box overlaps another box. */
  intersects(other: Range): boolean {
    return (
      this.left <= other.right &&
      other.left <= this.right &&
      this.top <= other.bottom &&
      other.top <= this.bottom
    );
  }

  /** Visits every address in the box, row-major (row-by-row, left-to-right within a row). */
  forEachAddress(cb: (address: string, row: number, col: number) => void): void {
    for (let row = this.top; row <= this.bottom; row++) {
      for (let col = this.left; col <= this.right; col++) {
        cb(encodeAddress(row, col), row, col);
      }
    }
  }
}
