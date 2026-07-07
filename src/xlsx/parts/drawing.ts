// `xl/drawings/drawing{n}.xml` — the DrawingML anchor part for a sheet's embedded images (Track D,
// docs/IMAGES-DRAWINGS-SPEC.md §2). One `<xdr:wsDr>` holding a `<xdr:twoCellAnchor>` (image spans a
// cell box) or `<xdr:oneCellAnchor>` (pinned at a cell, sized by an explicit pixel extent) per placed
// image. Each `<xdr:pic>` references its bytes through `<a:blip r:embed>`, whose rId resolves in the
// drawing's *own* `.rels` (a separate rId namespace from the worksheet's). Clean-room DrawingML.
//
// Anchor convention (from D1's model): stored positions are 0-based col/row and `br` is the *inclusive*
// bottom-right cell. DrawingML's `<xdr:to>` is exclusive, so it is emitted as `br.col+1`/`br.row+1`
// with zero corner offsets (e.g. the range `B2:D6` → from col=1,row=1, to col=4,row=6).
//
// EMU footgun: `colOff`/`rowOff` and the one-cell `<xdr:ext cx cy>` are in English Metric Units;
// 1 pixel = 9525 EMU. Pixel extents are converted ×9525; two-cell corner offsets stay 0 unless the
// model supplies `nativeColOff`/`nativeRowOff`.

import type { OneCellAnchor, ResolvedImageAnchor, TwoCellAnchor } from "../../model/index.js";
import { XmlWriter } from "../../xml/index.js";
import { NS } from "./constants.js";

/** English Metric Units per pixel — the fixed DrawingML conversion factor. */
const EMU_PER_PIXEL = 9525;

/** One placed image to render: its resolved anchor plus the `r:embed` rId in the drawing's own rels. */
export interface DrawingImage {
  /** The `<a:blip r:embed>` value — an rId local to this drawing's `.rels`. */
  embedId: string;
  anchor: ResolvedImageAnchor;
}

/** Render `xl/drawings/drawing{n}.xml` for a sheet's placed images (one anchor each, in order). */
export function renderDrawing(images: DrawingImage[]): string {
  const w = new XmlWriter();
  w.open("xdr:wsDr", {
    "xmlns:xdr": NS.drawingSpreadsheet,
    "xmlns:a": NS.drawingMain,
    "xmlns:r": NS.relationships,
  });
  images.forEach((img, i) => {
    // picId is 1-based and unique within the drawing (`id`/`name` on `<xdr:cNvPr>`).
    if ("ext" in img.anchor) {
      renderOneCellAnchor(w, img.anchor, img.embedId, i + 1);
    } else {
      renderTwoCellAnchor(w, img.anchor, img.embedId, i + 1);
    }
  });
  w.close("xdr:wsDr");
  return w.toString();
}

/** A `<xdr:from>`/`<xdr:to>` marker: a cell index plus an intra-cell EMU offset on each axis. */
function marker(
  w: XmlWriter,
  tag: string,
  col: number,
  colOff: number,
  row: number,
  rowOff: number,
): void {
  w.open(tag);
  w.el("xdr:col", String(col));
  w.el("xdr:colOff", String(colOff));
  w.el("xdr:row", String(row));
  w.el("xdr:rowOff", String(rowOff));
  w.close(tag);
}

/** `<xdr:twoCellAnchor>`: from `tl` to the exclusive `br+1` corner (zero corner offsets). */
function renderTwoCellAnchor(w: XmlWriter, a: TwoCellAnchor, embedId: string, picId: number): void {
  w.open("xdr:twoCellAnchor", { editAs: "oneCell" });
  marker(w, "xdr:from", a.tl.col, a.tl.nativeColOff ?? 0, a.tl.row, a.tl.nativeRowOff ?? 0);
  marker(w, "xdr:to", a.br.col + 1, 0, a.br.row + 1, 0);
  renderPic(w, embedId, picId);
  w.leaf("xdr:clientData");
  w.close("xdr:twoCellAnchor");
}

/** `<xdr:oneCellAnchor>`: pinned at `tl`, sized by an explicit `<xdr:ext>` in EMU. */
function renderOneCellAnchor(w: XmlWriter, a: OneCellAnchor, embedId: string, picId: number): void {
  w.open("xdr:oneCellAnchor");
  marker(w, "xdr:from", a.tl.col, a.tl.nativeColOff ?? 0, a.tl.row, a.tl.nativeRowOff ?? 0);
  w.leaf("xdr:ext", { cx: a.ext.width * EMU_PER_PIXEL, cy: a.ext.height * EMU_PER_PIXEL });
  renderPic(w, embedId, picId);
  w.leaf("xdr:clientData");
  w.close("xdr:oneCellAnchor");
}

/** The `<xdr:pic>` picture element shared by both anchor kinds — bytes via `<a:blip r:embed>`. */
function renderPic(w: XmlWriter, embedId: string, picId: number): void {
  w.open("xdr:pic");
  w.open("xdr:nvPicPr");
  w.leaf("xdr:cNvPr", { id: picId, name: `Picture ${picId}` });
  w.leaf("xdr:cNvPicPr");
  w.close("xdr:nvPicPr");
  w.open("xdr:blipFill");
  w.leaf("a:blip", { "r:embed": embedId });
  w.open("a:stretch");
  w.leaf("a:fillRect");
  w.close("a:stretch");
  w.close("xdr:blipFill");
  w.open("xdr:spPr");
  w.open("a:prstGeom", { prst: "rect" });
  w.leaf("a:avLst");
  w.close("a:prstGeom");
  w.close("xdr:spPr");
  w.close("xdr:pic");
}
