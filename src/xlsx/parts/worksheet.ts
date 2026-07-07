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

import type { CellModel, DataValidationModel, DataValidationType } from "../../model/index.js";
import { CellType, type RichText, type Worksheet } from "../../model/index.js";
import { XmlWriter } from "../../xml/index.js";
import { dateToSerial } from "../date.js";
import type { SharedStrings } from "../shared-strings.js";
import type { Styles } from "../styles.js";
import type { WorksheetRel } from "./constants.js";
import { NS } from "./constants.js";
import { renderWorksheetRels } from "./worksheet-rels.js";

export interface WorksheetContext {
  styles: Styles;
  sharedStrings: SharedStrings;
  useSharedStrings: boolean;
  date1904: boolean;
}

/** The rendered worksheet plus its `.rels` doc, if the sheet has any relationships (hyperlinks). */
export interface RenderedWorksheet {
  xml: string;
  /** Rendered `xl/worksheets/_rels/sheetN.xml.rels`; `undefined` when the sheet has no hyperlinks. */
  rels?: string;
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

  const hyperlinks: HyperlinkRecord[] = [];
  renderCols(w, model.columns, ctx);
  renderSheetData(w, model.rows, ctx, hyperlinks);
  renderMergeCells(w, model.merges);
  // CT_Worksheet order: `<dataValidations>` sits after `<mergeCells>`, `<hyperlinks>` after that.
  renderDataValidations(w, model.dataValidations);
  const rels = renderHyperlinks(w, hyperlinks);

  w.close("worksheet");
  return { xml: w.toString(), rels };
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

/** OOXML `CT_DataValidation` types that carry a comparison `operator` (`list`/`custom` do not). */
const OPERATOR_TYPES = new Set<DataValidationType>([
  "whole",
  "decimal",
  "date",
  "time",
  "textLength",
]);

/** Emit the `<dataValidations>` block (dropdowns + value constraints). Lives entirely in the sheet —
 *  no relationships or extra parts. See docs/DATA-VALIDATION-SPEC.md §3. */
function renderDataValidations(w: XmlWriter, validations: DataValidationModel[] | undefined): void {
  if (!validations || validations.length === 0) return;

  w.open("dataValidations", { count: validations.length });
  for (const dv of validations) {
    // `operator` is meaningful only for the numeric/length types; `formula2` only for a two-sided
    // comparison (`between`/`notBetween`).
    const operator = OPERATOR_TYPES.has(dv.type) ? dv.operator : undefined;
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

/** Emit the `<hyperlinks>` block, assigning each link a worksheet-local `rId${n}`. Returns the
 *  matching `.rels` doc (or `undefined` when the sheet has no hyperlinks), so the same rId space
 *  backs both the `<hyperlink r:id>` and the relationship `Id`. */
function renderHyperlinks(w: XmlWriter, hyperlinks: HyperlinkRecord[]): string | undefined {
  if (hyperlinks.length === 0) return undefined;

  const rels: WorksheetRel[] = [];
  w.open("hyperlinks");
  hyperlinks.forEach((h, i) => {
    const rId = `rId${i + 1}`;
    w.leaf("hyperlink", { ref: h.ref, "r:id": rId, tooltip: h.tooltip });
    rels.push({ rId, target: h.target, mode: "External" });
  });
  w.close("hyperlinks");

  return renderWorksheetRels(rels);
}
