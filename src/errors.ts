/**
 * Typed error hierarchy for @alosha/xlsx.
 *
 * Every error thrown by the library extends {@link AloshaXlsxError}, so callers can do:
 *
 *   try { ws.name = 'a/b'; }
 *   catch (e) {
 *     if (e instanceof AloshaXlsxError) console.log(e.code); // stable, minify-safe
 *   }
 *
 * The `code` string is the stable programmatic handle (safe across bundling/minification, unlike
 * `name`/`instanceof` across realms). Prefer switching on `code`. Design goal (see docs/API-AUDIT.md §5):
 * replace ExcelJS's `console.warn`/`console.trace` side effects with thrown, inspectable errors.
 */

export type AloshaXlsxErrorCode =
  | "INVALID_ADDRESS"
  | "INVALID_RANGE"
  | "INVALID_WORKSHEET_NAME"
  | "RESERVED_WORKSHEET_NAME"
  | "DUPLICATE_WORKSHEET_NAME"
  | "WORKSHEET_NAME_TOO_LONG"
  | "MERGE_CONFLICT"
  | "INVALID_ROW_NUMBER"
  | "INVALID_CELL_VALUE"
  | "UNSUPPORTED_FEATURE"
  | "INVALID_PACKAGE"
  | "INVALID_PANE"
  | "INVALID_CONDITIONAL_FORMATTING";

/**
 * Base class for all library errors. Abstract — never thrown directly; always a concrete subclass.
 * Sets `name` from the concrete class and repairs the prototype chain so `instanceof` holds even when
 * compiled to older targets.
 */
export abstract class AloshaXlsxError extends Error {
  /** Stable, machine-readable discriminator. Prefer this over `name`/`instanceof`. */
  abstract readonly code: AloshaXlsxErrorCode;

  constructor(message: string) {
    super(message);
    // `new.target` is the concrete subclass being constructed.
    this.name = new.target.name;
    Object.setPrototypeOf(this, new.target.prototype);
  }
}

/** Narrowing type guard for library errors. */
export function isAloshaXlsxError(e: unknown): e is AloshaXlsxError {
  return e instanceof AloshaXlsxError;
}

// ---------------------------------------------------------------------------
// Addresses & ranges  (src/address/*)
// ---------------------------------------------------------------------------

/** A cell address or column letter could not be parsed (e.g. `''`, `'A0'`, `'1A'`, `'@'`). */
export class InvalidAddressError extends AloshaXlsxError {
  readonly code = "INVALID_ADDRESS" as const;
  constructor(public readonly address: string) {
    super(`Invalid cell address: ${JSON.stringify(address)}`);
  }
}

/** A range could not be parsed from the supplied source (e.g. `'A1:'`, malformed tuple). */
export class InvalidRangeError extends AloshaXlsxError {
  readonly code = "INVALID_RANGE" as const;
  constructor(public readonly range: unknown) {
    super(`Invalid range: ${JSON.stringify(range)}`);
  }
}

// ---------------------------------------------------------------------------
// Worksheet naming  (Worksheet#name setter — see docs/API-AUDIT.md §2)
// ---------------------------------------------------------------------------

/** Why a worksheet name was rejected (illegal-character/empty/edge-quote cases). */
export type InvalidWorksheetNameReason = "empty" | "illegal-characters" | "edge-quote";

/** Worksheet name is empty, contains `* ? : / \ [ ]`, or starts/ends with a single quote. */
export class InvalidWorksheetNameError extends AloshaXlsxError {
  readonly code = "INVALID_WORKSHEET_NAME" as const;
  constructor(
    public readonly worksheetName: string,
    public readonly reason: InvalidWorksheetNameReason,
  ) {
    super(InvalidWorksheetNameError.messageFor(worksheetName, reason));
  }

  private static messageFor(name: string, reason: InvalidWorksheetNameReason): string {
    switch (reason) {
      case "empty":
        return "Worksheet name cannot be empty.";
      case "edge-quote":
        return `Worksheet name cannot start or end with a single quote: ${JSON.stringify(name)}`;
      case "illegal-characters":
        return `Worksheet name ${JSON.stringify(name)} cannot include any of: * ? : \\ / [ ]`;
    }
  }
}

/** Worksheet name uses a reserved word (Excel reserves `"History"`). */
export class ReservedWorksheetNameError extends AloshaXlsxError {
  readonly code = "RESERVED_WORKSHEET_NAME" as const;
  constructor(public readonly worksheetName: string) {
    super(`Worksheet name is reserved and cannot be used: ${JSON.stringify(worksheetName)}`);
  }
}

/** A worksheet with this name already exists in the workbook (comparison is case-insensitive). */
export class DuplicateWorksheetNameError extends AloshaXlsxError {
  readonly code = "DUPLICATE_WORKSHEET_NAME" as const;
  constructor(public readonly worksheetName: string) {
    super(`Worksheet name already exists: ${JSON.stringify(worksheetName)}`);
  }
}

/**
 * Worksheet name exceeds Excel's limit. Unlike ExcelJS (which warns + silently truncates), the model
 * throws so the caller decides — see docs/STAGE-1.2-SPEC.md §7.
 */
export class WorksheetNameTooLongError extends AloshaXlsxError {
  readonly code = "WORKSHEET_NAME_TOO_LONG" as const;
  /** Excel's hard cap on sheet-name length. */
  static readonly MAX_LENGTH = 31;
  constructor(
    public readonly worksheetName: string,
    public readonly maxLength: number = WorksheetNameTooLongError.MAX_LENGTH,
  ) {
    super(
      `Worksheet name exceeds ${maxLength} characters (got ${worksheetName.length}): ` +
        `${JSON.stringify(worksheetName)}`,
    );
  }
}

// ---------------------------------------------------------------------------
// Cells, rows & merges  (src/model/*)
// ---------------------------------------------------------------------------

/** Attempted to merge a region that overlaps an existing merge (merges are atomic — API-AUDIT §2). */
export class MergeConflictError extends AloshaXlsxError {
  readonly code = "MERGE_CONFLICT" as const;
  constructor(public readonly range: string) {
    super(`Cannot merge cells: range ${JSON.stringify(range)} overlaps an existing merge.`);
  }
}

/** A row model's `number` does not match the row it is being assigned to (model round-trip guard). */
export class InvalidRowNumberError extends AloshaXlsxError {
  readonly code = "INVALID_ROW_NUMBER" as const;
  constructor(
    public readonly expected: number,
    public readonly actual: number,
  ) {
    super(`Invalid row number in model: expected ${expected}, got ${actual}.`);
  }
}

/** A value passed to `cell.value` is not a supported {@link import('./model/cell-value').CellValue}. */
export class InvalidCellValueError extends AloshaXlsxError {
  readonly code = "INVALID_CELL_VALUE" as const;
  constructor(public readonly value: unknown) {
    super(`Unsupported cell value: ${InvalidCellValueError.describe(value)}`);
  }

  private static describe(value: unknown): string {
    const t = value === null ? "null" : typeof value;
    return t === "object"
      ? `object ${Object.prototype.toString.call(value)}`
      : `${t} (${String(value)})`;
  }
}

/**
 * A feature not implemented in the current stage was invoked (pivot tables, conditional formatting,
 * drawings, tables, data validation, comments, protection — deferred to Stage 2; see STAGE-1.2-SPEC §9).
 * Use to fail loudly on stubs instead of silently no-op'ing.
 */
export class UnsupportedFeatureError extends AloshaXlsxError {
  readonly code = "UNSUPPORTED_FEATURE" as const;
  constructor(
    public readonly feature: string,
    /** Optional pointer appended to the message (e.g. the compat shim's tracking-issue note). */
    public readonly detail?: string,
  ) {
    super(
      detail === undefined
        ? `Feature not yet supported in @alosha/xlsx: ${feature}`
        : `${feature}: ${detail}`,
    );
  }
}

// ---------------------------------------------------------------------------
// Package structure  (src/xlsx/read/package.ts, STAGE-2-SPEC.md §3)
// ---------------------------------------------------------------------------

/**
 * The `.xlsx` package's OPC structure (zip / rels / parts) is missing or malformed: a truncated/corrupt
 * zip, a missing `_rels/.rels` or workbook part, or a relationship that cannot be resolved. Thrown by
 * `openPackage` before any SpreadsheetML parsing begins.
 */
export class InvalidPackageError extends AloshaXlsxError {
  readonly code = "INVALID_PACKAGE" as const;
  constructor(public readonly reason: string) {
    super(`Invalid .xlsx package: ${reason}`);
  }
}

/** A `<pane>` element's `xSplit`/`ySplit` attribute is present but not a finite number (CT_Pane). */
export class InvalidPaneError extends AloshaXlsxError {
  readonly code = "INVALID_PANE" as const;
  constructor(
    public readonly attribute: "xSplit" | "ySplit",
    public readonly value: string,
  ) {
    super(`Invalid <pane> ${attribute}: ${JSON.stringify(value)} is not a number.`);
  }
}

/** A `<cfRule>` element's `priority`/`rank` attribute is present but not a finite number (CT_CfRule). */
export class InvalidConditionalFormattingError extends AloshaXlsxError {
  readonly code = "INVALID_CONDITIONAL_FORMATTING" as const;
  constructor(
    public readonly attribute: "priority" | "rank",
    public readonly value: string,
  ) {
    super(`Invalid <cfRule> ${attribute}: ${JSON.stringify(value)} is not a number.`);
  }
}
