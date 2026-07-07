// `xl/worksheets/_rels/sheetN.xml.rels` — a worksheet's own relationships: external hyperlinks and,
// from Tracks C/D, internal child parts (drawing, comments, vmlDrawing, image). The `Id`s here MUST
// equal the `r:id`s emitted in the sheet body; `worksheet.ts` allocates them in one pass through a
// single per-sheet allocator and passes the same list here. rIds live in the worksheet's local
// namespace (rId1, rId2, … per sheet), independent of the workbook rels. Mirrors `workbook-rels.ts`.

import { XmlWriter } from "../../xml/index.js";
import type { WorksheetRel } from "./constants.js";
import { NS } from "./constants.js";

export function renderWorksheetRels(rels: WorksheetRel[]): string {
  const w = new XmlWriter();
  w.open("Relationships", { xmlns: NS.packageRels });
  for (const rel of rels) {
    // `TargetMode` is dropped when `rel.mode` is undefined (internal rels) — XmlWriter omits
    // undefined attrs — so passing it directly renders external and internal rels correctly.
    w.leaf("Relationship", {
      Id: rel.rId,
      Type: rel.type,
      Target: rel.target,
      TargetMode: rel.mode,
    });
  }
  w.close("Relationships");
  return w.toString();
}
