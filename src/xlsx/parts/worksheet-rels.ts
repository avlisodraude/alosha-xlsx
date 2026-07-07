// `xl/worksheets/_rels/sheetN.xml.rels` — a worksheet's own relationships, currently only its
// external hyperlinks. The `Id`s here MUST equal the `r:id`s emitted in the sheet's `<hyperlinks>`
// block; `worksheet.ts` assigns them in one pass and passes the same list to both. rIds live in the
// worksheet's local namespace (rId1, rId2, … per sheet), independent of the workbook rels. Mirrors
// `workbook-rels.ts`.

import { XmlWriter } from "../../xml/index.js";
import type { WorksheetRel } from "./constants.js";
import { NS, REL_TYPE } from "./constants.js";

export function renderWorksheetRels(rels: WorksheetRel[]): string {
  const w = new XmlWriter();
  w.open("Relationships", { xmlns: NS.packageRels });
  for (const rel of rels) {
    w.leaf("Relationship", {
      Id: rel.rId,
      Type: REL_TYPE.hyperlink,
      Target: rel.target,
      TargetMode: rel.mode,
    });
  }
  w.close("Relationships");
  return w.toString();
}
