// Inverse of the Stage 1.3 worksheet writer (see docs/STAGE-2-SPEC.md §6, and src/xlsx/parts/worksheet.ts
// for the forward direction). Walks sheetN.xml with the pull parser and populates an already-created
// Worksheet: `<cols>` -> Column defaults, `<sheetData>` rows/cells -> values + styles, `<mergeCells>` ->
// merge regions (applied last so masters keep their values).
//
// The `t` attribute drives value decoding, exactly reversing the writer's mapping:
//   s -> shared-string index   inlineStr -> <is> text/runs   str -> formula string result
//   b -> boolean               e -> error code               n/absent -> number, or Date when the
//                                                             cell's number format is a date/time one.
// A `<f>` present makes it a formula cell whose cached `<v>` becomes the typed `result`.

import { decodeAddress } from "../../address/index.js";
import { InvalidCellValueError } from "../../errors.js";
import {
  type CellError,
  type CellValue,
  DATA_VALIDATION_OPERATOR_TYPES,
  type DataValidation,
  type DataValidationErrorStyle,
  type DataValidationOperator,
  type DataValidationType,
  type ErrorCode,
  type RichText,
  type Row,
  type Worksheet,
} from "../../model/index.js";
import { cursor, localName, type XmlCursor, type XmlEvent } from "../../xml/index.js";
import { serialToDate } from "../date.js";
import type { ResolvedRel } from "./package.js";
import { parseRichString } from "./shared-strings.js";
import type { StyleTable } from "./styles.js";

export interface WorksheetReadContext {
  sharedStrings: (string | RichText)[];
  styles: StyleTable;
  date1904: boolean;
  /** This worksheet's own relationships (rId -> {type, target, mode}), used to resolve
   *  `<hyperlinks>` entries. `undefined`/absent means hyperlinks are left as plain string cells. */
  rels?: Map<string, ResolvedRel>;
}

type OpenEvent = Extract<XmlEvent, { kind: "open" }>;

/** A parsed `<hyperlink>` sheet-level record, before its `r:id` is resolved to a target. */
interface HyperlinkRef {
  ref: string;
  rId: string;
  tooltip?: string;
}

/** A parsed `<dataValidation>`, paired with the `sqref` it registers under. */
interface DataValidationRef {
  sqref: string;
  dv: DataValidation;
}

/** Populate an already-created {@link Worksheet} from its sheetN.xml. */
export function parseWorksheet(ws: Worksheet, xml: string, ctx: WorksheetReadContext): void {
  const c = cursor(xml);
  let ev = c.next();
  while (ev && !(ev.kind === "open" && localName(ev.name) === "worksheet")) ev = c.next();
  if (ev?.kind !== "open" || ev.selfClosing) return;

  // Merges are collected and applied after every cell is placed, regardless of document order, so a
  // master cell's value is never clobbered by linking its slaves. Hyperlinks are collected the same
  // way and resolved last, once the display-text cells they point at already hold their string value.
  const merges: string[] = [];
  const hyperlinks: HyperlinkRef[] = [];
  const dataValidations: DataValidationRef[] = [];

  c.eachChild(ev.name, (child) => {
    switch (localName(child.name)) {
      case "cols":
        if (!child.selfClosing) parseCols(c, child.name, ws, ctx);
        break;
      case "sheetData":
        if (!child.selfClosing) parseSheetData(c, child.name, ws, ctx);
        break;
      case "mergeCells":
        if (!child.selfClosing) collectMerges(c, child.name, merges);
        break;
      case "dataValidations":
        if (!child.selfClosing) collectDataValidations(c, child.name, dataValidations);
        break;
      case "hyperlinks":
        if (!child.selfClosing) collectHyperlinks(c, child.name, hyperlinks);
        break;
      default:
        if (!child.selfClosing) c.textUntilClose(child.name);
    }
  });

  for (const ref of merges) ws.mergeCells(ref);
  for (const { sqref, dv } of dataValidations) ws.dataValidations.add(sqref, dv);
  if (ctx.rels) {
    for (const link of hyperlinks) applyHyperlink(ws, link, ctx.rels);
  }
}

function parseCols(
  c: XmlCursor,
  container: string,
  ws: Worksheet,
  ctx: WorksheetReadContext,
): void {
  c.eachChild(container, (col) => {
    if (localName(col.name) !== "col") {
      if (!col.selfClosing) c.textUntilClose(col.name);
      return;
    }
    const min = num(col.attrs.min);
    const max = num(col.attrs.max);
    if (min === undefined || max === undefined) return;

    const width = num(col.attrs.width);
    const hidden = truthy(col.attrs.hidden);
    const outline = num(col.attrs.outlineLevel);
    const styleIdx = num(col.attrs.style);
    const style = styleIdx ? ctx.styles.styleFor(styleIdx) : undefined;

    for (let n = min; n <= max; n++) {
      const column = ws.getColumn(n);
      if (width !== undefined) column.width = width;
      if (hidden) column.hidden = true;
      if (outline) column.outlineLevel = outline;
      if (style) column.style = style;
    }
  });
}

function parseSheetData(
  c: XmlCursor,
  container: string,
  ws: Worksheet,
  ctx: WorksheetReadContext,
): void {
  let expectedRow = 0;
  c.eachChild(container, (rowEl) => {
    if (localName(rowEl.name) !== "row") {
      if (!rowEl.selfClosing) c.textUntilClose(rowEl.name);
      return;
    }
    expectedRow = num(rowEl.attrs.r) ?? expectedRow + 1;
    parseRow(c, rowEl, expectedRow, ws, ctx);
  });
}

function parseRow(
  c: XmlCursor,
  rowEl: OpenEvent,
  rowNumber: number,
  ws: Worksheet,
  ctx: WorksheetReadContext,
): void {
  const row = ws.getRow(rowNumber);

  const height = num(rowEl.attrs.ht);
  if (height !== undefined) row.height = height;
  if (truthy(rowEl.attrs.hidden)) row.hidden = true;
  const outline = num(rowEl.attrs.outlineLevel);
  if (outline) row.outlineLevel = outline;
  // Set the row style before creating any cell so the (empty) fan-out to existing cells is a no-op;
  // cells created below then resolve this row style plus their column style at construction.
  const rowStyleIdx = num(rowEl.attrs.s);
  if (rowStyleIdx) row.style = ctx.styles.styleFor(rowStyleIdx);

  if (rowEl.selfClosing) return;

  let expectedCol = 0;
  c.eachChild(rowEl.name, (cellEl) => {
    if (localName(cellEl.name) !== "c") {
      if (!cellEl.selfClosing) c.textUntilClose(cellEl.name);
      return;
    }
    expectedCol = parseCell(c, cellEl, expectedCol + 1, row, ctx);
  });
}

/** Parse one `<c>`; returns the column index it occupied (so a following `<c>` without an `r` can be
 *  placed in the next column). */
function parseCell(
  c: XmlCursor,
  cellEl: OpenEvent,
  fallbackCol: number,
  row: Row,
  ctx: WorksheetReadContext,
): number {
  const col = cellEl.attrs.r !== undefined ? decodeAddress(cellEl.attrs.r).col : fallbackCol;
  const styleIdx = num(cellEl.attrs.s) ?? 0;
  const t = cellEl.attrs.t;

  let vText: string | undefined;
  let inlineValue: string | RichText | undefined;
  let formulaExpr: string | undefined;
  let hasFormula = false;

  if (!cellEl.selfClosing) {
    c.eachChild(cellEl.name, (child) => {
      switch (localName(child.name)) {
        case "v":
          vText = child.selfClosing ? "" : c.textUntilClose(child.name);
          break;
        case "f":
          hasFormula = true;
          formulaExpr = child.selfClosing ? "" : c.textUntilClose(child.name);
          break;
        case "is":
          inlineValue = child.selfClosing ? "" : parseRichString(c, child.name);
          break;
        default:
          if (!child.selfClosing) c.textUntilClose(child.name);
      }
    });
  }

  const value = hasFormula
    ? buildFormula(formulaExpr ?? "", t, vText, styleIdx, ctx)
    : decodeValue(t, vText, inlineValue, styleIdx, ctx);

  // An empty, unstyled, non-formula cell contributes nothing to the model — skip it.
  if (value === null && styleIdx === 0 && !hasFormula) return col;

  const cell = row.getCell(col);
  if (value !== null) cell.value = value;
  if (styleIdx !== 0) cell.style = ctx.styles.styleFor(styleIdx);
  return col;
}

/** Decode a non-formula cell's value from its `t` and body. Returns `null` for a valueless cell. */
function decodeValue(
  t: string | undefined,
  vText: string | undefined,
  inlineValue: string | RichText | undefined,
  styleIdx: number,
  ctx: WorksheetReadContext,
): CellValue {
  switch (t) {
    case "s": {
      if (vText === undefined) return null;
      return ctx.sharedStrings[Number(vText)] ?? "";
    }
    case "inlineStr":
      return inlineValue ?? "";
    case "str":
      return vText ?? "";
    case "b":
      return vText === undefined ? null : decodeBool(vText);
    case "e":
      return vText === undefined ? null : { error: vText as ErrorCode };
    default: {
      // "n" or the attribute absent: a number, promoted to a Date when the style is a date format.
      if (vText === undefined) return null;
      const n = Number(vText);
      return ctx.styles.isDateStyle(styleIdx) ? serialToDate(n, ctx.date1904) : n;
    }
  }
}

/** Build a formula cell: the bare `<f>` expression (matching the writer, which stripped the leading
 *  `=`) plus its cached `<v>` decoded into the typed `result`. */
function buildFormula(
  expr: string,
  t: string | undefined,
  vText: string | undefined,
  styleIdx: number,
  ctx: WorksheetReadContext,
): CellValue {
  const result = vText === undefined ? undefined : decodeScalar(t, vText, styleIdx, ctx);
  return result === undefined ? { formula: expr } : { formula: expr, result };
}

/** Decode a cached formula `<v>` into the scalar `result` union (no RichText — a rich shared-string
 *  result, which the writer never emits, flattens to its text). */
function decodeScalar(
  t: string | undefined,
  vText: string,
  styleIdx: number,
  ctx: WorksheetReadContext,
): number | string | boolean | Date | CellError {
  switch (t) {
    case "b":
      return decodeBool(vText);
    case "e":
      return { error: vText as ErrorCode };
    case "str":
      return vText;
    case "s": {
      const entry = ctx.sharedStrings[Number(vText)];
      if (entry === undefined) return "";
      return typeof entry === "string" ? entry : entry.richText.map((run) => run.text).join("");
    }
    default: {
      const n = Number(vText);
      return ctx.styles.isDateStyle(styleIdx) ? serialToDate(n, ctx.date1904) : n;
    }
  }
}

function collectMerges(c: XmlCursor, container: string, merges: string[]): void {
  c.eachChild(container, (mergeCell) => {
    if (localName(mergeCell.name) === "mergeCell" && mergeCell.attrs.ref !== undefined) {
      merges.push(mergeCell.attrs.ref);
    } else if (!mergeCell.selfClosing) {
      c.textUntilClose(mergeCell.name);
    }
  });
}

const KNOWN_DATA_VALIDATION_TYPES = new Set<DataValidationType>([
  "list",
  "whole",
  "decimal",
  "date",
  "time",
  "textLength",
  "custom",
]);

/** A `<formula1>`/`<formula2>` body: numeric for the operator types when it's a plain number
 *  literal (matching the writer's `String(number)`), otherwise kept as-is (a formula/cell ref, or
 *  the quoted inline list / range text for `list`). */
function parseFormulaValue(type: DataValidationType, text: string): string | number {
  if (DATA_VALIDATION_OPERATOR_TYPES.has(type) && /^-?\d+(\.\d+)?$/.test(text)) {
    return Number(text);
  }
  return text;
}

/** Collect `<dataValidation>` entries (paired with their `sqref`), applied to the worksheet's
 *  manager after the full child walk — order-independent, like merges/hyperlinks. An entry with
 *  no `sqref` or an unrecognized `type` is malformed and skipped, never thrown. */
function collectDataValidations(c: XmlCursor, container: string, out: DataValidationRef[]): void {
  c.eachChild(container, (child) => {
    if (localName(child.name) !== "dataValidation") {
      if (!child.selfClosing) c.textUntilClose(child.name);
      return;
    }

    const sqref = child.attrs.sqref;
    const type = child.attrs.type;
    if (
      sqref === undefined ||
      type === undefined ||
      !KNOWN_DATA_VALIDATION_TYPES.has(type as DataValidationType)
    ) {
      if (!child.selfClosing) c.textUntilClose(child.name);
      return;
    }
    const dvType = type as DataValidationType;

    let formula1: string | undefined;
    let formula2: string | undefined;
    if (!child.selfClosing) {
      c.eachChild(child.name, (f) => {
        switch (localName(f.name)) {
          case "formula1":
            formula1 = f.selfClosing ? "" : c.textUntilClose(f.name);
            break;
          case "formula2":
            formula2 = f.selfClosing ? "" : c.textUntilClose(f.name);
            break;
          default:
            if (!f.selfClosing) c.textUntilClose(f.name);
        }
      });
    }

    const dv: DataValidation = { type: dvType };
    if (child.attrs.operator !== undefined && DATA_VALIDATION_OPERATOR_TYPES.has(dvType)) {
      dv.operator = child.attrs.operator as DataValidationOperator;
    }
    if (truthy(child.attrs.allowBlank)) dv.allowBlank = true;
    if (truthy(child.attrs.showInputMessage)) dv.showInputMessage = true;
    if (truthy(child.attrs.showErrorMessage)) dv.showErrorMessage = true;
    // OOXML footgun: `showDropDown="1"` HIDES the dropdown, so its presence inverts to `false`;
    // its absence leaves the model's `showDropDown` unset (the default, shown).
    if (truthy(child.attrs.showDropDown)) dv.showDropDown = false;
    if (child.attrs.errorStyle !== undefined) {
      dv.errorStyle = child.attrs.errorStyle as DataValidationErrorStyle;
    }
    if (child.attrs.errorTitle !== undefined) dv.errorTitle = child.attrs.errorTitle;
    if (child.attrs.error !== undefined) dv.error = child.attrs.error;
    if (child.attrs.promptTitle !== undefined) dv.promptTitle = child.attrs.promptTitle;
    if (child.attrs.prompt !== undefined) dv.prompt = child.attrs.prompt;

    const formulae: (string | number)[] = [];
    if (formula1 !== undefined) formulae.push(parseFormulaValue(dvType, formula1));
    if (formula2 !== undefined) formulae.push(parseFormulaValue(dvType, formula2));
    if (formulae.length > 0) dv.formulae = formulae;

    out.push({ sqref, dv });
  });
}

function collectHyperlinks(c: XmlCursor, container: string, hyperlinks: HyperlinkRef[]): void {
  c.eachChild(container, (child) => {
    if (localName(child.name) !== "hyperlink") {
      if (!child.selfClosing) c.textUntilClose(child.name);
      return;
    }
    const ref = child.attrs.ref;
    const rId = relId(child);
    if (ref === undefined || rId === "") return;
    const tooltip = child.attrs.tooltip;
    hyperlinks.push(tooltip !== undefined ? { ref, rId, tooltip } : { ref, rId });
  });
}

/** Resolve a hyperlink record to its worksheet-rels target and set the model value on its (top-left,
 *  for a range `ref`) cell, keeping that cell's already-parsed string as the display `text`. A rId
 *  that doesn't resolve (malformed package) is silently skipped. */
function applyHyperlink(ws: Worksheet, link: HyperlinkRef, rels: Map<string, ResolvedRel>): void {
  const rel = rels.get(link.rId);
  if (!rel) return;

  const topLeft = link.ref.split(":")[0] ?? link.ref;
  const cell = ws.getCell(topLeft);
  const text = typeof cell.value === "string" ? cell.value : "";
  cell.value =
    link.tooltip !== undefined
      ? { text, hyperlink: rel.target, tooltip: link.tooltip }
      : { text, hyperlink: rel.target };
}

/** The `r:id` relationship reference on a `<hyperlink>` element. OOXML fixes the `r` prefix, but
 *  fall back to any `*:id` attribute so an unconventional namespace prefix still resolves. */
function relId(el: OpenEvent): string {
  const direct = el.attrs["r:id"];
  if (direct !== undefined) return direct;
  for (const key of Object.keys(el.attrs)) {
    if (localName(key) === "id") return el.attrs[key] ?? "";
  }
  return "";
}

/** A boolean cell's `<v>` is `1`/`0`; anything else is a malformed package, surfaced as a typed error. */
function decodeBool(raw: string): boolean {
  if (raw === "1" || raw === "true" || raw === "TRUE") return true;
  if (raw === "0" || raw === "false" || raw === "FALSE") return false;
  throw new InvalidCellValueError(raw);
}

/** Parse a numeric attribute; `undefined` when absent or not a finite number. */
function num(val: string | undefined): number | undefined {
  if (val === undefined) return undefined;
  const n = Number(val);
  return Number.isFinite(n) ? n : undefined;
}

/** A boolean-ish OOXML attribute (`hidden`): absent means false; `"1"`/`"true"` mean true. */
function truthy(val: string | undefined): boolean {
  return val === "1" || val === "true";
}
