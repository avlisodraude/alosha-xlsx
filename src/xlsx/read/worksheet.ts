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

import { decodeAddress, encodeAddress } from "../../address/index.js";
import {
  InvalidCellValueError,
  InvalidConditionalFormattingError,
  InvalidPaneError,
} from "../../errors.js";
import {
  type CellError,
  type CellFormula,
  type CellValue,
  type CfAverage,
  type CfCellIs,
  type CfColorScale,
  type CfDataBar,
  type CfDuplicate,
  type CfExpression,
  type CfIconSet,
  type CfOperator,
  type CfRuleBase,
  type CfText,
  type CfTop10,
  type CfValueObject,
  type ConditionalFormatting,
  type ConditionalRule,
  DATA_VALIDATION_OPERATOR_TYPES,
  type DataValidation,
  type DataValidationErrorStyle,
  type DataValidationOperator,
  type DataValidationType,
  type ErrorCode,
  type RichText,
  type Row,
  type Style,
  type Worksheet,
  type WorksheetView,
} from "../../model/index.js";
import { cursor, localName, type XmlCursor, type XmlEvent } from "../../xml/index.js";
import { serialToDate } from "../date.js";
import { REL_TYPE } from "../parts/constants.js";
import { parseComments } from "./comments.js";
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
  /** Fetches another package part's decoded text by its resolved path (as found in `rels`' targets),
   *  e.g. the `comments{n}.xml` the sheet's `comments` rel points at. `undefined`/absent, or a path
   *  with no such part, means comments are left unread (docs/COMMENTS-SPEC.md §2). */
  readPart?: (path: string) => string | undefined;
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
  const conditionalFormattings: ConditionalFormatting[] = [];
  let legacyDrawingRid: string | undefined;

  c.eachChild(ev.name, (child) => {
    switch (localName(child.name)) {
      case "sheetViews":
        if (!child.selfClosing) parseSheetViews(c, child.name, ws);
        break;
      case "cols":
        if (!child.selfClosing) parseCols(c, child.name, ws, ctx);
        break;
      case "sheetData":
        if (!child.selfClosing) parseSheetData(c, child.name, ws, ctx, new Map());
        break;
      case "autoFilter":
        if (child.attrs.ref !== undefined) ws.autoFilter = child.attrs.ref;
        if (!child.selfClosing) c.textUntilClose(child.name);
        break;
      case "mergeCells":
        if (!child.selfClosing) collectMerges(c, child.name, merges);
        break;
      case "conditionalFormatting":
        collectConditionalFormatting(c, child, ctx, conditionalFormattings);
        break;
      case "dataValidations":
        if (!child.selfClosing) collectDataValidations(c, child.name, dataValidations);
        break;
      case "hyperlinks":
        if (!child.selfClosing) collectHyperlinks(c, child.name, hyperlinks);
        break;
      case "legacyDrawing":
        legacyDrawingRid = relId(child);
        if (!child.selfClosing) c.textUntilClose(child.name);
        break;
      default:
        if (!child.selfClosing) c.textUntilClose(child.name);
    }
  });

  // Preserve each slave cell's just-parsed style: authoring-time mergeCells would overwrite it with
  // the master's, silently dropping e.g. a border set on a non-master cell of the range.
  for (const ref of merges) ws.mergeCellsPreservingStyles(ref);
  for (const cf of conditionalFormattings) {
    try {
      ws.addConditionalFormatting(cf);
    } catch {
      // A malformed sqref fails the model's Range validation (InvalidRangeError) — skip this block
      // rather than let one corrupt conditionalFormatting derail the rest of the read.
    }
  }
  for (const { sqref, dv } of dataValidations) ws.dataValidations.add(sqref, dv);
  if (ctx.rels) {
    for (const link of hyperlinks) applyHyperlink(ws, link, ctx.rels);
  }
  applyComments(ws, legacyDrawingRid, ctx);
}

/** Inverse of `renderSheetViews` (src/xlsx/parts/worksheet.ts): reads the sheet's single `<sheetView>`
 *  — `showGridLines="0"`, `zoomScale` (dropped when absent or `100`), and an optional `<pane>` child —
 *  into `views[0]`. A bare `<sheetView workbookViewId="0"/>` with no `<pane>` and no non-default attrs
 *  leaves `views` empty; a foreign file with more than one `<sheetView>` only keeps the first (our model
 *  carries a single view). */
function parseSheetViews(c: XmlCursor, container: string, ws: Worksheet): void {
  let seenSheetView = false;
  c.eachChild(container, (sheetViewEl) => {
    if (localName(sheetViewEl.name) !== "sheetView") {
      if (!sheetViewEl.selfClosing) c.textUntilClose(sheetViewEl.name);
      return;
    }
    if (seenSheetView) {
      if (!sheetViewEl.selfClosing) c.textUntilClose(sheetViewEl.name);
      return;
    }
    seenSheetView = true;

    const view: WorksheetView = {};
    if (sheetViewEl.attrs.showGridLines === "0") view.showGridLines = false;
    const zoomScale = num(sheetViewEl.attrs.zoomScale);
    if (zoomScale !== undefined && zoomScale !== 100) view.zoomScale = zoomScale;

    if (!sheetViewEl.selfClosing) {
      c.eachChild(sheetViewEl.name, (child) => {
        if (localName(child.name) === "pane") {
          Object.assign(view, parsePane(child));
          if (!child.selfClosing) c.textUntilClose(child.name);
        } else if (!child.selfClosing) {
          c.textUntilClose(child.name);
        }
      });
    }

    if (Object.keys(view).length > 0) ws.views = [view];
  });
}

/** Parse a `<pane>` element into the fields `renderSheetViews` emits: `state` verbatim, `xSplit`/
 *  `ySplit` coerced to numbers (absent attrs dropped, never materialised as `0`), `topLeftCell`/
 *  `activePane` copied through. A present-but-non-numeric split throws {@link InvalidPaneError} rather
 *  than silently producing `NaN`. */
function parsePane(paneEl: OpenEvent): WorksheetView {
  const view: WorksheetView = {};
  if (paneEl.attrs.state !== undefined) {
    view.state = paneEl.attrs.state as WorksheetView["state"];
  }
  const xSplit = paneNum(paneEl.attrs.xSplit, "xSplit");
  if (xSplit !== undefined) view.xSplit = xSplit;
  const ySplit = paneNum(paneEl.attrs.ySplit, "ySplit");
  if (ySplit !== undefined) view.ySplit = ySplit;
  if (paneEl.attrs.topLeftCell !== undefined) view.topLeftCell = paneEl.attrs.topLeftCell;
  if (paneEl.attrs.activePane !== undefined) {
    view.activePane = paneEl.attrs.activePane as WorksheetView["activePane"];
  }
  return view;
}

/** Parse a `<pane>` numeric attribute; `undefined` when absent, thrown when present but not finite. */
function paneNum(val: string | undefined, attr: "xSplit" | "ySplit"): number | undefined {
  if (val === undefined) return undefined;
  const n = Number(val);
  if (!Number.isFinite(n)) throw new InvalidPaneError(attr, val);
  return n;
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
  sharedMasters: Map<string, string>,
): void {
  let expectedRow = 0;
  c.eachChild(container, (rowEl) => {
    if (localName(rowEl.name) !== "row") {
      if (!rowEl.selfClosing) c.textUntilClose(rowEl.name);
      return;
    }
    expectedRow = num(rowEl.attrs.r) ?? expectedRow + 1;
    parseRow(c, rowEl, expectedRow, ws, ctx, sharedMasters);
  });
}

function parseRow(
  c: XmlCursor,
  rowEl: OpenEvent,
  rowNumber: number,
  ws: Worksheet,
  ctx: WorksheetReadContext,
  sharedMasters: Map<string, string>,
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
    expectedCol = parseCell(c, cellEl, expectedCol + 1, row, ctx, sharedMasters);
  });
}

/** The `t`/`si`/`ref` attributes on a `<f>` element — what distinguishes a shared/array master from a
 *  slave. `undefined` when the cell carries no `<f>`. */
interface FormulaAttrs {
  shareType?: string;
  si?: string;
  ref?: string;
}

/** Parse one `<c>`; returns the column index it occupied (so a following `<c>` without an `r` can be
 *  placed in the next column). */
function parseCell(
  c: XmlCursor,
  cellEl: OpenEvent,
  fallbackCol: number,
  row: Row,
  ctx: WorksheetReadContext,
  sharedMasters: Map<string, string>,
): number {
  const col = cellEl.attrs.r !== undefined ? decodeAddress(cellEl.attrs.r).col : fallbackCol;
  const styleIdx = num(cellEl.attrs.s) ?? 0;
  const t = cellEl.attrs.t;

  let vText: string | undefined;
  let inlineValue: string | RichText | undefined;
  let formulaExpr: string | undefined;
  let formulaAttrs: FormulaAttrs | undefined;
  let hasFormula = false;

  if (!cellEl.selfClosing) {
    c.eachChild(cellEl.name, (child) => {
      switch (localName(child.name)) {
        case "v":
          vText = child.selfClosing ? "" : c.textUntilClose(child.name);
          break;
        case "f":
          hasFormula = true;
          formulaAttrs = { shareType: child.attrs.t, si: child.attrs.si, ref: child.attrs.ref };
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

  const address = cellEl.attrs.r ?? encodeAddress(row.number, col);
  const value = hasFormula
    ? buildFormula(formulaExpr ?? "", formulaAttrs, t, vText, styleIdx, ctx, sharedMasters, address)
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
      const entry = ctx.sharedStrings[Number(vText)];
      if (entry === undefined) return "";
      // Clone rich text so cells that reference the same shared-string index don't alias one mutable
      // RichText object — editing one cell's runs must not silently mutate every other cell's.
      return typeof entry === "string" ? entry : structuredClone(entry);
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

/** Build a formula cell from its `<f>` (expression + `t`/`si`/`ref` attributes) and cached `<v>`,
 *  mirroring the writer's forward direction and ExcelJS's `reconcile`:
 *   - `t="array"`  → an array master: `{ formula, shareType:"array", ref, result }`.
 *   - `t="shared"` with `ref` → a shared master: recorded by `si` so its slaves can find it, kept as
 *     `{ formula, shareType:"shared", ref, result }`.
 *   - `t="shared"` without `ref` → a slave: resolved to its master's address (via `si`) as
 *     `{ sharedFormula: masterAddress, result }`; the master always precedes it in row-major order.
 *   - no `t` → a plain formula: `{ formula, result }`.
 *  The leading `=` was stripped on write, matching `expr`. */
function buildFormula(
  expr: string,
  attrs: FormulaAttrs | undefined,
  t: string | undefined,
  vText: string | undefined,
  styleIdx: number,
  ctx: WorksheetReadContext,
  sharedMasters: Map<string, string>,
  address: string,
): CellValue {
  const result = vText === undefined ? undefined : decodeScalar(t, vText, styleIdx, ctx);
  const withResult = (f: CellFormula): CellFormula => (result === undefined ? f : { ...f, result });

  if (attrs?.shareType === "array") {
    const f: CellFormula = { formula: expr, shareType: "array" };
    if (attrs.ref !== undefined) f.ref = attrs.ref;
    return withResult(f);
  }

  if (attrs?.shareType === "shared") {
    if (attrs.ref !== undefined) {
      // Master: register its address under `si` so following slaves resolve to it.
      if (attrs.si !== undefined) sharedMasters.set(attrs.si, address);
      return withResult({ formula: expr, shareType: "shared", ref: attrs.ref });
    }
    // Slave: point at the master's address rather than duplicating the expression.
    const master = attrs.si !== undefined ? sharedMasters.get(attrs.si) : undefined;
    if (master !== undefined) return withResult({ sharedFormula: master });
    // Master unseen (malformed / out-of-order file): degrade to whatever expression we have.
    return withResult({ formula: expr });
  }

  return withResult({ formula: expr });
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

// --- conditional formatting ---------------------------------------------------------------------------
// Inverse of renderConditionalFormatting/renderCfRule (src/xlsx/parts/worksheet.ts §2b). CT_Worksheet
// allows several sibling `<conditionalFormatting sqref="…">` blocks — unlike `<dataValidations>` there
// is no outer wrapper, so each block is handled as it is walked rather than collected via `eachChild`
// over a container.

const KNOWN_CF_TYPES = new Set<string>([
  "cellIs",
  "expression",
  "containsText",
  "notContainsText",
  "beginsWith",
  "endsWith",
  "top10",
  "aboveAverage",
  "duplicateValues",
  "uniqueValues",
  "colorScale",
  "dataBar",
  "iconSet",
]);

/** Collect a single `<conditionalFormatting sqref="…">` block, applied via `ws.addConditionalFormatting`
 *  after the full child walk (like merges/data validations). A block with no `sqref` is malformed and
 *  skipped; each `<cfRule>` child is inverted by {@link parseCfRule}, which itself skips (never throws
 *  for) an unrecognized `type`. */
function collectConditionalFormatting(
  c: XmlCursor,
  block: OpenEvent,
  ctx: WorksheetReadContext,
  out: ConditionalFormatting[],
): void {
  const ref = block.attrs.sqref;
  if (ref === undefined) {
    if (!block.selfClosing) c.textUntilClose(block.name);
    return;
  }

  const rules: ConditionalRule[] = [];
  if (!block.selfClosing) {
    c.eachChild(block.name, (ruleEl) => {
      if (localName(ruleEl.name) !== "cfRule") {
        if (!ruleEl.selfClosing) c.textUntilClose(ruleEl.name);
        return;
      }
      const rule = parseCfRule(c, ruleEl, ctx);
      if (rule !== undefined) rules.push(rule);
    });
  }

  out.push({ ref, rules });
}

/**
 * Invert one `<cfRule>` (CT_CfRule) into the model union, mirroring every shape `renderCfRule` emits.
 * An unrecognized `type`, or a rule missing a field its type requires (`cellIs`/`expression` with no
 * `<formula>`, the `containsText` family with no `text`, `top10` with no `rank`, `colorScale`/`dataBar`/
 * `iconSet` with no stops), is skipped rather than thrown — only a *malformed* `priority`/`rank`
 * (present but non-numeric) throws {@link InvalidConditionalFormattingError}, mirroring the pane
 * reader's numeric-attribute contract (`paneNum`).
 */
function parseCfRule(
  c: XmlCursor,
  el: OpenEvent,
  ctx: WorksheetReadContext,
): ConditionalRule | undefined {
  const type = el.attrs.type;
  if (type === undefined || !KNOWN_CF_TYPES.has(type)) {
    if (!el.selfClosing) c.textUntilClose(el.name);
    return undefined;
  }

  const priority = cfNum(el.attrs.priority, "priority");
  const stopIfTrue = truthy(el.attrs.stopIfTrue);

  switch (type) {
    case "cellIs": {
      const operator = el.attrs.operator as CfOperator | undefined;
      const formulae = collectCfFormulae(c, el).map(numericLiteralOr);
      if (operator === undefined || formulae.length === 0) return undefined;
      const rule: CfCellIs = { type: "cellIs", operator, formulae };
      applyBase(rule, priority, stopIfTrue);
      applyDxfStyle(rule, ctx, el.attrs.dxfId);
      return rule;
    }
    case "expression": {
      const [first] = collectCfFormulae(c, el);
      if (first === undefined) return undefined;
      const rule: CfExpression = { type: "expression", formulae: [first] };
      applyBase(rule, priority, stopIfTrue);
      applyDxfStyle(rule, ctx, el.attrs.dxfId);
      return rule;
    }
    case "containsText":
    case "notContainsText":
    case "beginsWith":
    case "endsWith": {
      const text = el.attrs.text;
      collectCfFormulae(c, el); // consume + discard the derived SEARCH/LEFT/RIGHT formula
      if (text === undefined) return undefined;
      const textType = type as "containsText" | "notContainsText" | "beginsWith" | "endsWith";
      const rule: CfText = { type: textType, text };
      applyBase(rule, priority, stopIfTrue);
      applyDxfStyle(rule, ctx, el.attrs.dxfId);
      return rule;
    }
    case "top10": {
      const rank = cfNum(el.attrs.rank, "rank");
      if (!el.selfClosing) c.textUntilClose(el.name);
      if (rank === undefined) return undefined;
      const rule: CfTop10 = { type: "top10", rank };
      if (truthy(el.attrs.percent)) rule.percent = true;
      if (truthy(el.attrs.bottom)) rule.bottom = true;
      applyBase(rule, priority, stopIfTrue);
      applyDxfStyle(rule, ctx, el.attrs.dxfId);
      return rule;
    }
    case "aboveAverage": {
      if (!el.selfClosing) c.textUntilClose(el.name);
      // `aboveAverage="0"` is the writer's encoding for `belowAverage` (same OOXML type).
      const isBelow = el.attrs.aboveAverage !== undefined && !truthy(el.attrs.aboveAverage);
      const rule: CfAverage = { type: isBelow ? "belowAverage" : "aboveAverage" };
      applyBase(rule, priority, stopIfTrue);
      applyDxfStyle(rule, ctx, el.attrs.dxfId);
      return rule;
    }
    case "duplicateValues":
    case "uniqueValues": {
      if (!el.selfClosing) c.textUntilClose(el.name);
      const dupType = type as "duplicateValues" | "uniqueValues";
      const rule: CfDuplicate = { type: dupType };
      applyBase(rule, priority, stopIfTrue);
      applyDxfStyle(rule, ctx, el.attrs.dxfId);
      return rule;
    }
    case "colorScale": {
      const wrapper = parseCfVisualWrapper(c, el, "colorScale");
      if (wrapper === undefined || wrapper.cfvo.length === 0 || wrapper.colors.length === 0) {
        return undefined;
      }
      const rule: CfColorScale = {
        type: "colorScale",
        cfvo: wrapper.cfvo,
        color: wrapper.colors,
      };
      applyBase(rule, priority, stopIfTrue);
      return rule;
    }
    case "dataBar": {
      const wrapper = parseCfVisualWrapper(c, el, "dataBar");
      const first = wrapper?.cfvo[0];
      const second = wrapper?.cfvo[1];
      const color = wrapper?.colors[0];
      if (first === undefined || second === undefined || color === undefined) return undefined;
      const rule: CfDataBar = { type: "dataBar", cfvo: [first, second], color };
      applyBase(rule, priority, stopIfTrue);
      return rule;
    }
    case "iconSet": {
      const wrapper = parseCfVisualWrapper(c, el, "iconSet");
      const iconSetName = wrapper?.attrs.iconSet;
      if (wrapper === undefined || iconSetName === undefined || wrapper.cfvo.length === 0) {
        return undefined;
      }
      const rule: CfIconSet = { type: "iconSet", iconSet: iconSetName, cfvo: wrapper.cfvo };
      applyBase(rule, priority, stopIfTrue);
      return rule;
    }
    default:
      if (!el.selfClosing) c.textUntilClose(el.name);
      return undefined;
  }
}

/** Collect every direct `<formula>` child's raw text (`cellIs`: 1–2, `expression`: 1, the
 *  `containsText` family: 1 — discarded there, since `text` carries the field and the formula is
 *  derived on write). */
function collectCfFormulae(c: XmlCursor, el: OpenEvent): string[] {
  const formulae: string[] = [];
  if (!el.selfClosing) {
    c.eachChild(el.name, (child) => {
      if (localName(child.name) === "formula") {
        formulae.push(child.selfClosing ? "" : c.textUntilClose(child.name));
      } else if (!child.selfClosing) {
        c.textUntilClose(child.name);
      }
    });
  }
  return formulae;
}

/** A `<formula>`/`<cfvo val>` body: a plain numeric literal (matching the writer's `String(number)`)
 *  decodes to a number; anything else (a formula, cell ref, or quoted text) is kept as-is. */
function numericLiteralOr(text: string): string | number {
  return /^-?\d+(\.\d+)?$/.test(text) ? Number(text) : text;
}

/** Parse a `<cfRule>` numeric attribute (`priority`/`rank`); `undefined` when absent, thrown when
 *  present but not finite — mirrors `paneNum`'s no-silent-NaN contract for `<pane>`. */
function cfNum(val: string | undefined, attr: "priority" | "rank"): number | undefined {
  if (val === undefined) return undefined;
  const n = Number(val);
  if (!Number.isFinite(n)) throw new InvalidConditionalFormattingError(attr, val);
  return n;
}

function applyBase(rule: CfRuleBase, priority: number | undefined, stopIfTrue: boolean): void {
  if (priority !== undefined) rule.priority = priority;
  if (stopIfTrue) rule.stopIfTrue = true;
}

/** Resolve `dxfId` through the styles pool into `rule.style`, mirroring `ctx.styles.addDxf` on write.
 *  An absent/non-numeric `dxfId`, or one resolving to an empty differential style, leaves `style`
 *  unset — matching a rule that was originally written with no `style`. */
function applyDxfStyle(
  rule: { style?: Partial<Style> },
  ctx: WorksheetReadContext,
  dxfIdAttr: string | undefined,
): void {
  const dxfId = num(dxfIdAttr);
  if (dxfId === undefined) return;
  const style = ctx.styles.dxfFor(dxfId);
  if (Object.keys(style).length > 0) rule.style = style;
}

interface CfVisualWrapper {
  attrs: Record<string, string>;
  cfvo: CfValueObject[];
  colors: string[];
}

/** Find the `<colorScale>`/`<dataBar>`/`<iconSet>` wrapper under a self-visualising `<cfRule>` and
 *  collect its `<cfvo>` stops (in order) and `<color rgb>` values (in order); `iconSet`'s own preset
 *  name lives on `attrs.iconSet`. `undefined` when the wrapper is missing entirely (e.g. a
 *  self-closing `<cfRule>`). */
function parseCfVisualWrapper(
  c: XmlCursor,
  cfRuleEl: OpenEvent,
  wrapperLocalName: string,
): CfVisualWrapper | undefined {
  let result: CfVisualWrapper | undefined;
  if (!cfRuleEl.selfClosing) {
    c.eachChild(cfRuleEl.name, (child) => {
      if (localName(child.name) !== wrapperLocalName) {
        if (!child.selfClosing) c.textUntilClose(child.name);
        return;
      }
      const cfvo: CfValueObject[] = [];
      const colors: string[] = [];
      if (!child.selfClosing) {
        c.eachChild(child.name, (grand) => {
          switch (localName(grand.name)) {
            case "cfvo": {
              const cfvoType = grand.attrs.type as CfValueObject["type"] | undefined;
              if (cfvoType !== undefined) {
                const val = grand.attrs.val;
                cfvo.push(
                  val !== undefined
                    ? { type: cfvoType, value: numericLiteralOr(val) }
                    : { type: cfvoType },
                );
              }
              if (!grand.selfClosing) c.textUntilClose(grand.name);
              break;
            }
            case "color":
              if (grand.attrs.rgb !== undefined) colors.push(grand.attrs.rgb);
              if (!grand.selfClosing) c.textUntilClose(grand.name);
              break;
            default:
              if (!grand.selfClosing) c.textUntilClose(grand.name);
          }
        });
      }
      result = { attrs: child.attrs, cfvo, colors };
    });
  }
  return result;
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
    if (DATA_VALIDATION_OPERATOR_TYPES.has(dvType)) {
      // ST_DataValidationOperator defaults to "between" when the attribute is absent (ECMA-376;
      // ExcelJS's own reader applies the same default), so a foreign file that omits it for a
      // between-style validation still resolves to the correct operator.
      dv.operator = (child.attrs.operator as DataValidationOperator | undefined) ?? "between";
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

/** Follow the sheet's `<legacyDrawing r:id>` (docs/COMMENTS-SPEC.md §2) to confirm a VML note anchor
 *  exists — the VML is anchor-only, so its content is never fetched — then resolve the sheet's
 *  separate `comments` rel (not referenced by any worksheet element, only discoverable by rel type,
 *  mirroring the writer which allocates it alongside but never embeds its rId) to `comments{n}.xml`
 *  and apply its records to `cell.note`. A missing `<legacyDrawing>`, an unresolved rel, or an absent
 *  part leaves every cell without a note; a comment with a malformed `ref` is skipped — none of this
 *  ever throws. */
function applyComments(
  ws: Worksheet,
  legacyDrawingRid: string | undefined,
  ctx: WorksheetReadContext,
): void {
  if (legacyDrawingRid === undefined || !ctx.rels || !ctx.readPart) return;
  if (!ctx.rels.get(legacyDrawingRid)) return;

  const commentsRel = [...ctx.rels.values()].find((r) => r.type === REL_TYPE.comments);
  if (!commentsRel) return;
  const commentsXml = ctx.readPart(commentsRel.target);
  if (commentsXml === undefined) return;

  for (const record of parseComments(commentsXml)) {
    try {
      ws.getCell(record.ref).note = { texts: record.texts, author: record.author };
    } catch {
      // Malformed ref (InvalidAddressError from decodeAddress) — skip this comment, keep the rest.
    }
  }
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
