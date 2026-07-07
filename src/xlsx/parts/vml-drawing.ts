// `xl/drawings/vmlDrawing{n}.vml` — the legacy VML anchor part for cell comments
// (docs/COMMENTS-SPEC.md §1). Excel renders a note's box from this drawing, not from the modern
// drawing part; the worksheet wires it via `<legacyDrawing r:id>`.
//
// One `<v:shape>` per comment, each carrying an `<x:ClientData ObjectType="Note">` with the anchor
// cell's 0-based `<x:Row>`/`<x:Column>`. Box geometry is a fixed default — we author a sensible note
// box and ignore geometry on read (§0 "Out"). Pushed as an F0 `ExtraPart` with a `vml` `<Default>`.

import { decodeAddress } from "../../address/index.js";
import { XmlWriter } from "../../xml/index.js";

/** The three legacy Microsoft Office VML namespaces (clean-room, from the OOXML transitional spec). */
const VML_NS = {
  v: "urn:schemas-microsoft-com:vml",
  o: "urn:schemas-microsoft-com:office:office",
  x: "urn:schemas-microsoft-com:office:excel",
} as const;

/** Render `xl/drawings/vmlDrawing{n}.vml` — one note shape per `ref` (never called when empty). */
export function renderVmlDrawing(refs: string[]): string {
  // No XML declaration: VML drawing parts are conventionally raw `<xml>` roots.
  const w = new XmlWriter({ declaration: false });
  w.open("xml", { "xmlns:v": VML_NS.v, "xmlns:o": VML_NS.o, "xmlns:x": VML_NS.x });

  // Shared boilerplate: the shape layout and the note-box shapetype every note instance references.
  w.open("o:shapelayout", { "v:ext": "edit" });
  w.leaf("o:idmap", { "v:ext": "edit", data: "1" });
  w.close("o:shapelayout");
  w.open("v:shapetype", {
    id: "_x0000_t202",
    coordsize: "21600,21600",
    "o:spt": 202,
    path: "m,l,21600r21600,l21600,xe",
  });
  w.leaf("v:stroke", { joinstyle: "miter" });
  w.leaf("v:path", { gradientshapeok: "t", "o:connecttype": "rect" });
  w.close("v:shapetype");

  refs.forEach((ref, i) => {
    const { row, col } = decodeAddress(ref);
    const row0 = row - 1; // 0-based anchor row/column carried in ClientData.
    const col0 = col - 1;

    w.open("v:shape", {
      id: `_x0000_s${1025 + i}`,
      type: "#_x0000_t202",
      style: "position:absolute;visibility:hidden",
      fillcolor: "#ffffe1",
      "o:insetmode": "auto",
    });
    w.leaf("v:fill", { color2: "#ffffe1" });
    w.leaf("v:shadow", { on: "t", color: "black", obscured: "t" });
    w.leaf("v:path", { "o:connecttype": "none" });
    w.open("v:textbox", { style: "mso-direction-alt:auto" });
    w.close("v:textbox");
    w.open("x:ClientData", { ObjectType: "Note" });
    w.leaf("x:MoveWithCells");
    w.leaf("x:SizeWithCells");
    // Default box: two columns wide, four rows tall, one column to the right of the anchor.
    w.el("x:Anchor", `${col0 + 1}, 15, ${row0}, 10, ${col0 + 3}, 15, ${row0 + 4}, 4`);
    w.el("x:AutoFill", "False");
    w.el("x:Row", String(row0));
    w.el("x:Column", String(col0));
    w.close("x:ClientData");
    w.close("v:shape");
  });

  w.close("xml");
  return w.toString();
}
