// The cell value union + pure helpers. Replaces ExcelJS's 12 Value strategy classes with a
// discriminated union and free functions (see docs/API-AUDIT.md §5) — same behavior, no
// per-cell class allocation.
//
// `Font` is a type-only import: style.ts has no dependency back on this module, so this does not
// introduce an import cycle, and `import type` is erased entirely at build time.

import { InvalidCellValueError } from "../errors.js";
import { CellType, type ErrorCode } from "./enums.js";
import type { Font } from "./style.js";

export interface RichTextRun {
  text: string;
  font?: Font;
}

export interface RichText {
  richText: RichTextRun[];
}

export interface Hyperlink {
  text: string;
  hyperlink: string;
  tooltip?: string;
}

export interface CellError {
  error: ErrorCode;
}

export interface CellFormula {
  formula?: string;
  sharedFormula?: string;
  ref?: string;
  shareType?: "shared" | "array";
  result?: number | string | boolean | Date | CellError;
}

/** What callers pass to `cell.value = …` and read back. */
export type CellValue =
  | null
  | number
  | string
  | boolean
  | Date
  | CellError
  | RichText
  | Hyperlink
  | CellFormula;

/** Internal normalized storage on the cell model (discriminated on `type`). */
export type StoredValue =
  | { type: CellType.Null }
  | { type: CellType.Number; value: number }
  | { type: CellType.String; value: string }
  | { type: CellType.Boolean; value: boolean }
  | { type: CellType.Date; value: Date }
  | { type: CellType.Error; value: CellError }
  | { type: CellType.RichText; value: RichText }
  | { type: CellType.Hyperlink; value: Hyperlink }
  | { type: CellType.Formula; value: CellFormula }
  | { type: CellType.Merge; master: string }
  | { type: CellType.SharedString; value: number };

/**
 * Duck-type a `CellValue` into a `CellType`, mirroring ExcelJS's `Value.getType` check order:
 * null → string → number → boolean → date → hyperlink (`text`&`hyperlink`) → formula
 * (`formula`|`sharedFormula`) → richText → error. Throws {@link InvalidCellValueError} for
 * anything else.
 */
export function detectType(v: CellValue): CellType {
  if (v === null) return CellType.Null;
  if (typeof v === "string") return CellType.String;
  if (typeof v === "number") return CellType.Number;
  if (typeof v === "boolean") return CellType.Boolean;
  if (v instanceof Date) return CellType.Date;
  if ("text" in v && "hyperlink" in v) return CellType.Hyperlink;
  if ("formula" in v || "sharedFormula" in v) return CellType.Formula;
  if ("richText" in v) return CellType.RichText;
  if ("error" in v) return CellType.Error;
  throw new InvalidCellValueError(v);
}

/** Normalizes an ergonomic `CellValue` into its discriminated `StoredValue` form. */
export function toStored(v: CellValue): StoredValue {
  switch (detectType(v)) {
    case CellType.Null:
      return { type: CellType.Null };
    case CellType.String:
      return { type: CellType.String, value: v as string };
    case CellType.Number:
      return { type: CellType.Number, value: v as number };
    case CellType.Boolean:
      return { type: CellType.Boolean, value: v as boolean };
    case CellType.Date:
      return { type: CellType.Date, value: v as Date };
    case CellType.Hyperlink:
      return { type: CellType.Hyperlink, value: v as Hyperlink };
    case CellType.Formula:
      return { type: CellType.Formula, value: v as CellFormula };
    case CellType.RichText:
      return { type: CellType.RichText, value: v as RichText };
    case CellType.Error:
      return { type: CellType.Error, value: v as CellError };
    default:
      throw new InvalidCellValueError(v);
  }
}

/**
 * Recovers the ergonomic `CellValue` from a `StoredValue`. `Merge` (resolved via the master cell)
 * and `SharedString` (resolved via the shared-strings pool, Stage 1.3) have no `CellValue`
 * equivalent of their own and read back as `null`.
 */
export function fromStored(s: StoredValue): CellValue {
  switch (s.type) {
    case CellType.Null:
    case CellType.Merge:
    case CellType.SharedString:
      return null;
    case CellType.Number:
    case CellType.String:
    case CellType.Boolean:
    case CellType.Date:
    case CellType.Error:
    case CellType.RichText:
    case CellType.Hyperlink:
    case CellType.Formula:
      return s.value;
  }
}

/** For a formula cell, the type of its cached `result`; else the cell's own type. */
export function effectiveType(s: StoredValue): CellType {
  if (s.type !== CellType.Formula) return s.type;

  const { result } = s.value;
  if (result === undefined) return CellType.Null;
  if (typeof result === "number") return CellType.Number;
  if (typeof result === "string") return CellType.String;
  if (typeof result === "boolean") return CellType.Boolean;
  if (result instanceof Date) return CellType.Date;
  return CellType.Error;
}

function primitiveToText(value: number | string | boolean | Date | CellError): string {
  if (typeof value === "number") return String(value);
  if (typeof value === "string") return value;
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (value instanceof Date) return value.toISOString();
  return value.error;
}

/** Display string for a stored value (no number-format application — that's Stage 1.3). */
export function toText(s: StoredValue): string {
  switch (s.type) {
    case CellType.Null:
    case CellType.Merge:
    case CellType.SharedString:
      return "";
    case CellType.Number:
    case CellType.String:
    case CellType.Boolean:
    case CellType.Date:
    case CellType.Error:
      return primitiveToText(s.value);
    case CellType.RichText:
      return s.value.richText.map((run) => run.text).join("");
    case CellType.Hyperlink:
      return s.value.text;
    case CellType.Formula: {
      const { result } = s.value;
      return result === undefined ? "" : primitiveToText(result);
    }
  }
}

const CSV_SPECIAL_RE = /["\n\r,]/;

/** CSV-escaped rendering of a stored value: quotes the field and doubles internal quotes when it
 *  contains a comma, quote, or newline. */
export function toCsv(s: StoredValue): string {
  const text = toText(s);
  if (!CSV_SPECIAL_RE.test(text)) return text;
  return `"${text.replace(/"/g, '""')}"`;
}
