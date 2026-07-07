// Inverse of the Track D writer (src/xlsx/parts/drawing.ts, docs/IMAGES-DRAWINGS-SPEC.md §3). Walks a
// `drawing{n}.xml` part and recovers one `DrawingImage` (a drawing-local `r:embed` rId plus the resolved
// anchor) per `<xdr:twoCellAnchor>`/`<xdr:oneCellAnchor>`. `<xdr:absoluteAnchor>` and non-`<xdr:pic>`
// shapes (charts, groups) are out of scope (spec §0) and are drained without producing an image.
//
// Anchor inversion (mirrors renderDrawing exactly):
//   twoCellAnchor: tl <- <xdr:from> verbatim; br <- <xdr:to> shifted back one (exclusive -> inclusive).
//   oneCellAnchor: tl <- <xdr:from> verbatim; ext <- <xdr:ext cx/cy> divided by the EMU-per-pixel factor.
// A present-but-non-numeric marker/ext value throws InvalidDrawingError (no silent NaN, mirroring the
// pane/cfRule readers); a missing/malformed anchor (no <a:blip r:embed>, or no from/col/row) is simply
// skipped — this reader never throws for content it cannot place.

import { InvalidDrawingError } from "../../errors.js";
import type { CellPosition } from "../../model/index.js";
import { cursor, localName, type XmlCursor, type XmlEvent } from "../../xml/index.js";
import type { DrawingImage } from "../parts/drawing.js";

/** English Metric Units per pixel — the fixed DrawingML conversion factor (matches the writer). */
const EMU_PER_PIXEL = 9525;

type OpenEvent = Extract<XmlEvent, { kind: "open" }>;

/** The four children of an `<xdr:from>`/`<xdr:to>` marker, as parsed (each optional — absent when the
 *  element is missing from the document). */
interface MarkerFields {
  col?: number;
  row?: number;
  colOff?: number;
  rowOff?: number;
}

/** Parse `xl/drawings/drawing{n}.xml` into its placed pictures, in document order. A malformed/rootless
 *  document (no `<xdr:wsDr>`) yields an empty array. */
export function parseDrawing(xml: string): DrawingImage[] {
  const images: DrawingImage[] = [];
  const c = cursor(xml);
  let ev = c.next();
  while (ev && !(ev.kind === "open" && localName(ev.name) === "wsDr")) ev = c.next();
  if (ev?.kind !== "open" || ev.selfClosing) return images;

  c.eachChild(ev.name, (anchorEl) => {
    const kind = localName(anchorEl.name);
    if (kind !== "twoCellAnchor" && kind !== "oneCellAnchor") {
      if (!anchorEl.selfClosing) c.textUntilClose(anchorEl.name);
      return;
    }
    if (anchorEl.selfClosing) return;

    const built = parseAnchorBody(c, anchorEl, kind);
    if (built) images.push(built);
  });

  return images;
}

/** Walk one anchor's children (order-independent, per the schema) collecting `from`/`to`/`ext`/the
 *  picture's `r:embed`, then invert into a `DrawingImage`. */
function parseAnchorBody(
  c: XmlCursor,
  anchorEl: OpenEvent,
  kind: "twoCellAnchor" | "oneCellAnchor",
): DrawingImage | undefined {
  let from: MarkerFields = {};
  let to: MarkerFields = {};
  let cx: number | undefined;
  let cy: number | undefined;
  let embedId: string | undefined;

  c.eachChild(anchorEl.name, (child) => {
    switch (localName(child.name)) {
      case "from":
        from = child.selfClosing ? {} : parseMarker(c, child.name);
        break;
      case "to":
        to = child.selfClosing ? {} : parseMarker(c, child.name);
        break;
      case "ext":
        cx = drawingNum(child.attrs.cx, "cx");
        cy = drawingNum(child.attrs.cy, "cy");
        if (!child.selfClosing) c.textUntilClose(child.name);
        break;
      case "pic":
        embedId = scanForBlip(c, child);
        break;
      default:
        if (!child.selfClosing) c.textUntilClose(child.name);
    }
  });

  if (embedId === undefined || from.col === undefined || from.row === undefined) return undefined;

  const tl: CellPosition = { col: from.col, row: from.row };
  if (from.colOff) tl.nativeColOff = from.colOff;
  if (from.rowOff) tl.nativeRowOff = from.rowOff;

  if (kind === "oneCellAnchor") {
    if (cx === undefined || cy === undefined) return undefined;
    return {
      embedId,
      anchor: { tl, ext: { width: cx / EMU_PER_PIXEL, height: cy / EMU_PER_PIXEL } },
    };
  }

  if (to.col === undefined || to.row === undefined) return undefined;
  const br: CellPosition = { col: to.col - 1, row: to.row - 1 };
  return { embedId, anchor: { tl, br } };
}

/** Parse a `<xdr:from>`/`<xdr:to>` marker's four text-content children. */
function parseMarker(c: XmlCursor, containerName: string): MarkerFields {
  const out: MarkerFields = {};
  c.eachChild(containerName, (child) => {
    const text = child.selfClosing ? "" : c.textUntilClose(child.name);
    switch (localName(child.name)) {
      case "col":
        out.col = drawingNum(text, "col");
        break;
      case "row":
        out.row = drawingNum(text, "row");
        break;
      case "colOff":
        out.colOff = drawingNum(text, "colOff");
        break;
      case "rowOff":
        out.rowOff = drawingNum(text, "rowOff");
        break;
      default:
        break;
    }
  });
  return out;
}

/** Recursively hunt an `<xdr:pic>` subtree for its `<a:blip r:embed>` value, fully draining every
 *  descendant along the way (required by the cursor's `eachChild` contract). `undefined` when no blip
 *  is found (a shape with no image data — out of scope, skipped rather than thrown). */
function scanForBlip(c: XmlCursor, el: OpenEvent): string | undefined {
  const self = localName(el.name) === "blip" ? el.attrs["r:embed"] : undefined;
  if (el.selfClosing) return self;

  let found = self;
  c.eachChild(el.name, (child) => {
    const nested = scanForBlip(c, child);
    if (found === undefined) found = nested;
  });
  return found;
}

/** Parse a marker/ext numeric value; `undefined` when absent (element/attribute missing, or empty
 *  text), thrown when present but not finite — mirrors `paneNum`/`cfNum`'s no-silent-NaN contract. */
function drawingNum(
  val: string | undefined,
  attr: "col" | "row" | "colOff" | "rowOff" | "cx" | "cy",
): number | undefined {
  if (val === undefined || val === "") return undefined;
  const n = Number(val);
  if (!Number.isFinite(n)) throw new InvalidDrawingError(attr, val);
  return n;
}
