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

import type { CellModel } from "../../model/index.js";
import { CellType, type RichText, type Worksheet } from "../../model/index.js";
import { XmlWriter } from "../../xml/index.js";
import { dateToSerial } from "../date.js";
import type { SharedStrings } from "../shared-strings.js";
import type { Styles } from "../styles.js";
import { NS } from "./constants.js";

export interface WorksheetContext {
  styles: Styles;
  sharedStrings: SharedStrings;
  useSharedStrings: boolean;
  date1904: boolean;
}

/** Applied to a Date cell whose style carries no explicit number format, so Excel shows a date rather
 *  than the raw serial. Built-in id 14, so it adds no custom `<numFmt>` entry. */
const DEFAULT_DATE_NUMFMT = "mm-dd-yy";

export function renderWorksheet(ws: Worksheet, ctx: WorksheetContext): string {
  const model = ws.model;
  const w = new XmlWriter();
  w.open("worksheet", { xmlns: NS.spreadsheetml, "xmlns:r": NS.relationships });

  w.leaf("dimension", { ref: ws.dimensions.range });

  w.open("sheetViews");
  w.leaf("sheetView", {
    workbookViewId: 0,
    showGridLines: model.properties.showGridLines === false ? false : undefined,
  });
  w.close("sheetViews");

  w.leaf("sheetFormatPr", {
    defaultRowHeight: model.properties.defaultRowHeight ?? 15,
    defaultColWidth: model.properties.defaultColWidth,
  });

  renderCols(w, model.columns, ctx);
  renderSheetData(w, model.rows, ctx);
  renderMergeCells(w, model.merges);

  w.close("worksheet");
  return w.toString();
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
): void {
  w.open("sheetData");
  for (const row of rows) {
    const styleId = row.style ? ctx.styles.addStyle(row.style) : 0;

    let cellsXml = "";
    for (const cell of row.cells) {
      cellsXml += renderCell(cell, ctx);
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
function renderCell(cell: CellModel, ctx: WorksheetContext): string {
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

    case CellType.Hyperlink:
      return stringCell(r, styleId, (cell.value as { text: string }).text, ctx);

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
