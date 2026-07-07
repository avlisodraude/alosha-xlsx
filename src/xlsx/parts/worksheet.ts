// `xl/worksheets/sheetN.xml` — the grid (SPEC §5, OOXML-NOTES §3). Serialized from a worksheet's
// `model` snapshot plus `dimensions`. Rendering a cell interns its value into the shared-strings pool
// (when enabled) and its style into the styles pool, so this MUST run before those parts are rendered.
//
// Cell `t` mapping (from CellModel.type / a formula's cached-result type):
//   string   -> "s" (shared index) or "inlineStr"      number -> "n" (attr omitted, the default)
//   boolean  -> "b"                                     date   -> "n" with a date numFmt (serial value)
//   error    -> "e"                                     richtext -> shared string (or inline runs)
//   hyperlink-> its display text as a string            formula -> <f>expr</f> + typed cached <v>
//   merge-slave / styled-empty -> empty <c> carrying only its style   null (no style) -> skipped

import { Range } from "../../address/index.js";
import type {
  CellModel,
  CommentModel,
  ConditionalFormatting,
  ConditionalRule,
  DataValidationModel,
  WorksheetView,
} from "../../model/index.js";
import {
  CellType,
  DATA_VALIDATION_OPERATOR_TYPES,
  DEFAULT_COMMENT_AUTHOR,
  type RichText,
  type Worksheet,
} from "../../model/index.js";
import { XmlWriter } from "../../xml/index.js";
import { dateToSerial } from "../date.js";
import type { SharedStrings } from "../shared-strings.js";
import type { Styles } from "../styles.js";
import { type CommentRecord, renderComments } from "./comments.js";
import type { WorksheetRel } from "./constants.js";
import { CONTENT_TYPE, NS, REL_TYPE } from "./constants.js";
import { renderVmlDrawing } from "./vml-drawing.js";
import { renderWorksheetRels } from "./worksheet-rels.js";

export interface WorksheetContext {
  styles: Styles;
  sharedStrings: SharedStrings;
  useSharedStrings: boolean;
  date1904: boolean;
  /** Workbook-monotonic allocator for cfRule priorities with no explicit value. Supplied by the
   *  workbook writer so priorities stay unique across sheets; a per-render fallback covers isolated
   *  part tests that construct a context without it. */
  nextPriority?: () => number;
  /** 1-based position of this sheet in the workbook, set by the orchestrator per sheet. Lets child
   *  parts (Tracks C/D) name themselves deterministically (`drawing${n}.xml`, `comments${n}.xml`, …).
   *  Optional so isolated part tests can omit it. */
  sheetNumber?: number;
}

/** An extra child part a worksheet contributes to the package — an XML sibling (comments, drawing) or
 *  raw binary media (an image). The orchestrator writes it at `path` and folds its content-type into
 *  `[Content_Types].xml`. Empty in F0; Tracks C/D are the first producers. */
export interface ExtraPart {
  /** Archive path, e.g. `xl/drawings/drawing1.xml` or `xl/media/image1.png`. */
  path: string;
  /** XML text (encoded by the orchestrator) or raw media bytes (passed through untouched). */
  data: string | Uint8Array;
  /** Emits an `<Override>` for this exact part (for XML parts with a non-`.xml` content-type). */
  contentType?: string;
  /** Emits a `<Default Extension>` covering media by extension (deduped across the workbook). */
  defaultFor?: { extension: string; contentType: string };
}

/** The rendered worksheet, its unified `.rels` doc, and any extra child parts it owns. */
export interface RenderedWorksheet {
  xml: string;
  /** Rendered `xl/worksheets/_rels/sheetN.xml.rels`; `undefined` when the sheet has no rels. */
  rels?: string;
  /** Child parts (XML + media) the sheet contributes; `undefined` when it owns none. */
  extraParts?: ExtraPart[];
}

/** A worksheet's single relationship-id source: one rId namespace (`rId1, rId2, …`) and one `rels`
 *  accumulator shared by every feature (hyperlinks first, then Tracks C/D), so pushed rels never
 *  collide on rIds. */
export interface SheetRelAllocator {
  /** Hand out the next local `rId${n}` for this sheet. */
  next(): string;
  /** The accumulated rels, in allocation order; fed once to `renderWorksheetRels`. */
  readonly rels: WorksheetRel[];
}

function createSheetRelAllocator(): SheetRelAllocator {
  const rels: WorksheetRel[] = [];
  let n = 0;
  return {
    next: () => `rId${++n}`,
    rels,
  };
}

/** A cell hyperlink collected during the sheetData pass, emitted as a `<hyperlink>` after merges. */
interface HyperlinkRecord {
  ref: string;
  target: string;
  tooltip?: string;
}

/** Applied to a Date cell whose style carries no explicit number format, so Excel shows a date rather
 *  than the raw serial. Built-in id 14, so it adds no custom `<numFmt>` entry. */
const DEFAULT_DATE_NUMFMT = "mm-dd-yy";

export function renderWorksheet(ws: Worksheet, ctx: WorksheetContext): RenderedWorksheet {
  const model = ws.model;
  const w = new XmlWriter();
  w.open("worksheet", { xmlns: NS.spreadsheetml, "xmlns:r": NS.relationships });

  w.leaf("dimension", { ref: ws.dimensions.range });

  renderSheetViews(w, model.views, model.properties.showGridLines);

  w.leaf("sheetFormatPr", {
    defaultRowHeight: model.properties.defaultRowHeight ?? 15,
    defaultColWidth: model.properties.defaultColWidth,
  });

  // One rId source and one extra-parts accumulator for the whole sheet — hyperlinks are the first
  // rel contributor; Tracks C/D push more rels + parts through the same seam without colliding.
  const alloc = createSheetRelAllocator();
  const extraParts: ExtraPart[] = [];

  const hyperlinks: HyperlinkRecord[] = [];
  renderCols(w, model.columns, ctx);
  renderSheetData(w, model.rows, ctx, hyperlinks);
  // CT_Worksheet order: `<autoFilter>` sits after `<sheetData>`, before `<mergeCells>`;
  // `<dataValidations>` after `<mergeCells>`, `<hyperlinks>` after that.
  renderAutoFilter(w, model.autoFilter);
  renderMergeCells(w, model.merges);
  // CT_Worksheet order: `<conditionalFormatting>` sits between `<mergeCells>` and `<dataValidations>`.
  renderConditionalFormatting(w, model.conditionalFormattings, ctx);
  renderDataValidations(w, model.dataValidations);
  renderHyperlinks(w, hyperlinks, alloc);
  // `<legacyDrawing>` is the last child (CT_Worksheet: `…drawing, legacyDrawing`), so comments run
  // after hyperlinks — their rels take rIds after any hyperlink rels too.
  renderSheetComments(w, model.comments, ctx, alloc, extraParts);

  w.close("worksheet");
  const rels = alloc.rels.length > 0 ? renderWorksheetRels(alloc.rels) : undefined;
  return {
    xml: w.toString(),
    rels,
    extraParts: extraParts.length > 0 ? extraParts : undefined,
  };
}

/** Emit `<sheetViews><sheetView>…</sheetView></sheetViews>`, adding a `<pane>` child when the first
 *  view is frozen/split (CT_Pane). Only the first view backs the sheet's single `<sheetView>` — this
 *  mirrors Excel, which stores one view per sheet in practice. */
function renderSheetViews(
  w: XmlWriter,
  views: WorksheetView[],
  showGridLinesProp: boolean | undefined,
): void {
  const view = views[0];
  const showGridLines = showGridLinesProp === false ? false : undefined;

  w.open("sheetViews");
  if (view && (view.state === "frozen" || view.state === "split")) {
    w.open("sheetView", { workbookViewId: 0, showGridLines });
    w.leaf("pane", {
      xSplit: view.xSplit,
      ySplit: view.ySplit,
      topLeftCell: view.topLeftCell,
      activePane: view.activePane,
      state: view.state,
    });
    w.close("sheetView");
  } else {
    w.leaf("sheetView", { workbookViewId: 0, showGridLines });
  }
  w.close("sheetViews");
}

/** Emit `<autoFilter ref="…">` (CT_Worksheet: after `<sheetData>`, before `<mergeCells>`). */
function renderAutoFilter(w: XmlWriter, autoFilter: string | undefined): void {
  if (!autoFilter) return;
  w.leaf("autoFilter", { ref: autoFilter });
}

function renderCols(
  w: XmlWriter,
  columns: Worksheet["model"]["columns"],
  ctx: WorksheetContext,
): void {
  if (!columns || columns.length === 0) return;

  const fragments: string[] = [];
  for (const range of columns) {
    const styleId = range.style ? ctx.styles.addStyle(range.style) : 0;
    const hasWidth = range.width !== undefined;
    if (!hasWidth && !range.hidden && !range.outlineLevel && styleId === 0) continue;

    const col = new XmlWriter({ declaration: false });
    col.leaf("col", {
      min: range.min,
      max: range.max,
      width: range.width,
      customWidth: hasWidth ? true : undefined,
      style: styleId !== 0 ? styleId : undefined,
      hidden: range.hidden || undefined,
      outlineLevel: range.outlineLevel || undefined,
    });
    fragments.push(col.toString());
  }

  if (fragments.length === 0) return;
  w.open("cols");
  for (const f of fragments) w.raw(f);
  w.close("cols");
}

function renderSheetData(
  w: XmlWriter,
  rows: Worksheet["model"]["rows"],
  ctx: WorksheetContext,
  hyperlinks: HyperlinkRecord[],
): void {
  w.open("sheetData");
  for (const row of rows) {
    const styleId = row.style ? ctx.styles.addStyle(row.style) : 0;

    let cellsXml = "";
    for (const cell of row.cells) {
      cellsXml += renderCell(cell, ctx, hyperlinks);
    }

    const hasAttrs =
      row.height !== undefined || row.hidden || (row.outlineLevel ?? 0) !== 0 || styleId !== 0;
    if (cellsXml === "" && !hasAttrs) continue;

    w.open("row", {
      r: row.number,
      ht: row.height,
      customHeight: row.height !== undefined ? true : undefined,
      hidden: row.hidden || undefined,
      outlineLevel: row.outlineLevel || undefined,
      s: styleId !== 0 ? styleId : undefined,
      customFormat: styleId !== 0 ? true : undefined,
    });
    w.raw(cellsXml);
    w.close("row");
  }
  w.close("sheetData");
}

/** Render one `<c>` element, or `""` when the cell contributes nothing (a null/merge cell with the
 *  default style). Value interning into the shared-strings/styles pools happens here. */
function renderCell(cell: CellModel, ctx: WorksheetContext, hyperlinks: HyperlinkRecord[]): string {
  const styleId = cell.style ? ctx.styles.addStyle(cell.style) : 0;
  const r = cell.address;

  switch (cell.type) {
    case CellType.Null:
    case CellType.Merge:
      return styleId !== 0 ? leafCell(r, styleId) : "";

    case CellType.Number:
      return valueCell(r, styleId, undefined, String(cell.value as number));

    case CellType.Boolean:
      return valueCell(r, styleId, "b", (cell.value as boolean) ? "1" : "0");

    case CellType.Date: {
      const s = cell.style?.numFmt
        ? styleId
        : ctx.styles.addStyle({ ...cell.style, numFmt: DEFAULT_DATE_NUMFMT });
      return valueCell(r, s, undefined, String(dateToSerial(cell.value as Date, ctx.date1904)));
    }

    case CellType.Error:
      return valueCell(r, styleId, "e", (cell.value as { error: string }).error);

    case CellType.String:
      return stringCell(r, styleId, cell.value as string, ctx);

    case CellType.Hyperlink: {
      const link = cell.value as { text: string; hyperlink: string; tooltip?: string };
      hyperlinks.push({ ref: r, target: link.hyperlink, tooltip: link.tooltip });
      // Still emit the display text as a normal string cell — the link is a sheet-level record.
      return stringCell(r, styleId, link.text, ctx);
    }

    case CellType.RichText:
      return richTextCell(r, styleId, cell.value as RichText, ctx);

    case CellType.Formula:
      return formulaCell(r, styleId, cell, ctx);

    default:
      return styleId !== 0 ? leafCell(r, styleId) : "";
  }
}

/** `<c r s t><v>raw</v></c>` — `raw` is already a serialized number/token, escaped on write. */
function valueCell(r: string, s: number, t: string | undefined, raw: string): string {
  const w = new XmlWriter({ declaration: false });
  w.open("c", { r, s: s !== 0 ? s : undefined, t });
  w.el("v", raw);
  w.close("c");
  return w.toString();
}

/** `<c r s/>` — an empty cell carrying only its style (merge slave / styled blank). */
function leafCell(r: string, s: number): string {
  const w = new XmlWriter({ declaration: false });
  w.leaf("c", { r, s });
  return w.toString();
}

function stringCell(r: string, s: number, value: string, ctx: WorksheetContext): string {
  if (ctx.useSharedStrings) {
    return valueCell(r, s, "s", String(ctx.sharedStrings.add(value)));
  }
  const w = new XmlWriter({ declaration: false });
  w.open("c", { r, s: s !== 0 ? s : undefined, t: "inlineStr" });
  w.open("is");
  w.el("t", value);
  w.close("is");
  w.close("c");
  return w.toString();
}

function richTextCell(r: string, s: number, value: RichText, ctx: WorksheetContext): string {
  if (ctx.useSharedStrings) {
    return valueCell(r, s, "s", String(ctx.sharedStrings.add(value)));
  }
  // Inline fallback flattens the runs to plain text (formatting survives only via shared strings).
  const text = value.richText.map((run) => run.text).join("");
  return stringCell(r, s, text, { ...ctx, useSharedStrings: false });
}

function formulaCell(r: string, s: number, cell: CellModel, ctx: WorksheetContext): string {
  const expr = (cell.formula ?? "").replace(/^=/, "");
  const result = cell.result;

  let t: string | undefined;
  let cachedV: string | undefined;
  if (result === undefined) {
    cachedV = undefined;
  } else if (typeof result === "number") {
    cachedV = String(result);
  } else if (typeof result === "boolean") {
    t = "b";
    cachedV = result ? "1" : "0";
  } else if (result instanceof Date) {
    cachedV = String(dateToSerial(result, ctx.date1904));
  } else if (typeof result === "object" && result !== null && "error" in result) {
    t = "e";
    cachedV = (result as { error: string }).error;
  } else {
    t = "str";
    cachedV = String(result);
  }

  const w = new XmlWriter({ declaration: false });
  w.open("c", { r, s: s !== 0 ? s : undefined, t });
  w.el("f", expr);
  if (cachedV !== undefined) w.el("v", cachedV);
  w.close("c");
  return w.toString();
}

function renderMergeCells(w: XmlWriter, merges: string[]): void {
  if (merges.length === 0) return;
  w.open("mergeCells", { count: merges.length });
  for (const ref of merges) w.leaf("mergeCell", { ref });
  w.close("mergeCells");
}

/** Maps a text-family rule `type` to its `<cfRule operator>` (the operator vocabulary differs from the
 *  type: `notContainsText` uses operator `notContains`). */
const CF_TEXT_OPERATOR: Record<
  "containsText" | "notContainsText" | "beginsWith" | "endsWith",
  string
> = {
  containsText: "containsText",
  notContainsText: "notContains",
  beginsWith: "beginsWith",
  endsWith: "endsWith",
};

/** Emit the `<conditionalFormatting>` blocks (CT_Worksheet: after `<mergeCells>`, before
 *  `<dataValidations>`). One `<conditionalFormatting sqref>` per block, then its `<cfRule>`s. Rules
 *  with no explicit `priority` draw a workbook-monotonic one; dxf-bearing rules intern their `style`
 *  into the shared `<dxfs>` pool. See docs/CONDITIONAL-FORMATTING-SPEC.md §2b. */
function renderConditionalFormatting(
  w: XmlWriter,
  cfs: ConditionalFormatting[] | undefined,
  ctx: WorksheetContext,
): void {
  if (!cfs || cfs.length === 0) return;

  let local = 0;
  const nextPriority = ctx.nextPriority ?? (() => ++local);

  for (const cf of cfs) {
    // The text-family formulae anchor on the top-left cell of the ref (Excel's convention).
    const anchor = new Range(cf.ref.trim().split(/\s+/)[0] ?? cf.ref).tl;
    w.open("conditionalFormatting", { sqref: cf.ref });
    for (const rule of cf.rules) renderCfRule(w, rule, anchor, ctx, nextPriority);
    w.close("conditionalFormatting");
  }
}

/** Render a single `<cfRule>`. Self-visualising rules (colorScale/dataBar/iconSet) carry no `dxfId`;
 *  every other rule interns its optional differential `style` and references the resulting `dxfId`. */
function renderCfRule(
  w: XmlWriter,
  rule: ConditionalRule,
  anchor: string,
  ctx: WorksheetContext,
  nextPriority: () => number,
): void {
  const priority = rule.priority ?? nextPriority();
  const stopIfTrue = rule.stopIfTrue ? true : undefined;

  switch (rule.type) {
    case "cellIs": {
      const dxfId = ctx.styles.addDxf(rule.style);
      w.open("cfRule", { type: "cellIs", dxfId, priority, operator: rule.operator, stopIfTrue });
      for (const f of rule.formulae) w.el("formula", String(f));
      w.close("cfRule");
      break;
    }
    case "expression": {
      const dxfId = ctx.styles.addDxf(rule.style);
      w.open("cfRule", { type: "expression", dxfId, priority, stopIfTrue });
      w.el("formula", String(rule.formulae[0]));
      w.close("cfRule");
      break;
    }
    case "containsText":
    case "notContainsText":
    case "beginsWith":
    case "endsWith": {
      const dxfId = ctx.styles.addDxf(rule.style);
      w.open("cfRule", {
        type: rule.type,
        dxfId,
        priority,
        operator: CF_TEXT_OPERATOR[rule.type],
        text: rule.text,
        stopIfTrue,
      });
      w.el("formula", textRuleFormula(rule.type, rule.text, anchor));
      w.close("cfRule");
      break;
    }
    case "top10": {
      const dxfId = ctx.styles.addDxf(rule.style);
      w.leaf("cfRule", {
        type: "top10",
        dxfId,
        priority,
        percent: rule.percent ? true : undefined,
        bottom: rule.bottom ? true : undefined,
        rank: rule.rank,
        stopIfTrue,
      });
      break;
    }
    case "aboveAverage":
    case "belowAverage": {
      const dxfId = ctx.styles.addDxf(rule.style);
      w.leaf("cfRule", {
        type: "aboveAverage",
        dxfId,
        priority,
        // `belowAverage` is the same OOXML type with `aboveAverage="0"`.
        aboveAverage: rule.type === "belowAverage" ? false : undefined,
        stopIfTrue,
      });
      break;
    }
    case "duplicateValues":
    case "uniqueValues": {
      const dxfId = ctx.styles.addDxf(rule.style);
      w.leaf("cfRule", { type: rule.type, dxfId, priority, stopIfTrue });
      break;
    }
    case "colorScale": {
      w.open("cfRule", { type: "colorScale", priority, stopIfTrue });
      w.open("colorScale");
      for (const cfvo of rule.cfvo) renderCfvo(w, cfvo);
      for (const color of rule.color) w.leaf("color", { rgb: color });
      w.close("colorScale");
      w.close("cfRule");
      break;
    }
    case "dataBar": {
      w.open("cfRule", { type: "dataBar", priority, stopIfTrue });
      w.open("dataBar");
      for (const cfvo of rule.cfvo) renderCfvo(w, cfvo);
      w.leaf("color", { rgb: rule.color });
      w.close("dataBar");
      w.close("cfRule");
      break;
    }
    case "iconSet": {
      w.open("cfRule", { type: "iconSet", priority, stopIfTrue });
      w.open("iconSet", { iconSet: rule.iconSet });
      for (const cfvo of rule.cfvo) renderCfvo(w, cfvo);
      w.close("iconSet");
      w.close("cfRule");
      break;
    }
  }
}

/** A `<cfvo>`: the `val` attribute is omitted for `min`/`max` (which have no threshold value). */
function renderCfvo(w: XmlWriter, cfvo: { type: string; value?: number | string }): void {
  const val =
    cfvo.type === "min" || cfvo.type === "max" || cfvo.value === undefined
      ? undefined
      : String(cfvo.value);
  w.leaf("cfvo", { type: cfvo.type, val });
}

/** The Excel-generated `<formula>` backing a text-family rule, anchored on the ref's top-left cell. */
function textRuleFormula(
  type: "containsText" | "notContainsText" | "beginsWith" | "endsWith",
  text: string,
  anchor: string,
): string {
  switch (type) {
    case "containsText":
      return `NOT(ISERROR(SEARCH("${text}",${anchor})))`;
    case "notContainsText":
      return `ISERROR(SEARCH("${text}",${anchor}))`;
    case "beginsWith":
      return `LEFT(${anchor},LEN("${text}"))="${text}"`;
    case "endsWith":
      return `RIGHT(${anchor},LEN("${text}"))="${text}"`;
  }
}

/** Emit the `<dataValidations>` block (dropdowns + value constraints). Lives entirely in the sheet —
 *  no relationships or extra parts. See docs/DATA-VALIDATION-SPEC.md §3. */
function renderDataValidations(w: XmlWriter, validations: DataValidationModel[] | undefined): void {
  if (!validations || validations.length === 0) return;

  w.open("dataValidations", { count: validations.length });
  for (const dv of validations) {
    // `operator` is meaningful only for the numeric/length types; `formula2` only for a two-sided
    // comparison (`between`/`notBetween`).
    const operator = DATA_VALIDATION_OPERATOR_TYPES.has(dv.type) ? dv.operator : undefined;
    w.open("dataValidation", {
      type: (dv.type as string) === "none" ? undefined : dv.type,
      operator,
      allowBlank: dv.allowBlank || undefined,
      // OOXML footgun: `showDropDown="1"` HIDES the dropdown. So emit it only when the model
      // explicitly hides it (`showDropDown === false`); the default (unset/true) shows it.
      showDropDown: dv.showDropDown === false ? true : undefined,
      showInputMessage: dv.showInputMessage || undefined,
      showErrorMessage: dv.showErrorMessage || undefined,
      errorStyle: dv.errorStyle,
      errorTitle: dv.errorTitle,
      error: dv.error,
      promptTitle: dv.promptTitle,
      prompt: dv.prompt,
      sqref: dv.sqref,
    });

    const formulae = dv.formulae;
    if (formulae && formulae[0] !== undefined) w.el("formula1", String(formulae[0]));
    if (
      (operator === "between" || operator === "notBetween") &&
      formulae &&
      formulae[1] !== undefined
    ) {
      w.el("formula2", String(formulae[1]));
    }
    w.close("dataValidation");
  }
  w.close("dataValidations");
}

/** Emit the `<hyperlinks>` block, drawing each link's `r:id` from the shared sheet allocator and
 *  pushing a matching external rel — so the same rId backs both `<hyperlink r:id>` and its
 *  `<Relationship Id>`. Hyperlinks are the first consumer of the allocator's rId namespace. */
function renderHyperlinks(
  w: XmlWriter,
  hyperlinks: HyperlinkRecord[],
  alloc: SheetRelAllocator,
): void {
  if (hyperlinks.length === 0) return;

  w.open("hyperlinks");
  for (const h of hyperlinks) {
    const rId = alloc.next();
    w.leaf("hyperlink", { ref: h.ref, "r:id": rId, tooltip: h.tooltip });
    alloc.rels.push({ rId, type: REL_TYPE.hyperlink, target: h.target, mode: "External" });
  }
  w.close("hyperlinks");
}

/** Emit the legacy-comment wiring for a sheet with `cell.note`s (docs/COMMENTS-SPEC.md §1): push the
 *  `comments{n}.xml` note table and the `vmlDrawing{n}.vml` anchor part as F0 `ExtraPart`s, allocate a
 *  rel for each off the shared sheet allocator, and emit `<legacyDrawing r:id>` (the vmlDrawing rId) as
 *  the sheet's last child. Authors default to `"Author"` when unset (ExcelJS parity). */
function renderSheetComments(
  w: XmlWriter,
  comments: CommentModel[] | undefined,
  ctx: WorksheetContext,
  alloc: SheetRelAllocator,
  extraParts: ExtraPart[],
): void {
  if (!comments || comments.length === 0) return;
  const n = ctx.sheetNumber ?? 1;

  const records: CommentRecord[] = comments.map((c) => ({
    ref: c.ref,
    texts: c.texts,
    author: c.author ?? DEFAULT_COMMENT_AUTHOR,
  }));

  extraParts.push({
    path: `xl/comments${n}.xml`,
    data: renderComments(records),
    contentType: CONTENT_TYPE.comments,
  });
  extraParts.push({
    path: `xl/drawings/vmlDrawing${n}.vml`,
    data: renderVmlDrawing(records.map((r) => r.ref)),
    defaultFor: { extension: "vml", contentType: CONTENT_TYPE.vmlDrawing },
  });

  const commentsRId = alloc.next();
  alloc.rels.push({ rId: commentsRId, type: REL_TYPE.comments, target: `../comments${n}.xml` });
  const vmlRId = alloc.next();
  alloc.rels.push({
    rId: vmlRId,
    type: REL_TYPE.vmlDrawing,
    target: `../drawings/vmlDrawing${n}.vml`,
  });

  // The `<legacyDrawing>` points at the VML, not the comments part.
  w.leaf("legacyDrawing", { "r:id": vmlRId });
}
