// `xl/_rels/workbook.xml.rels` — relationships from the workbook to its worksheets, styles,
// sharedStrings (when present) and theme. Worksheet rIds here MUST equal the `r:id`s emitted in
// workbook.xml; the orchestrator assigns them once and passes the same list to both (SPEC §5).

import { XmlWriter } from "../../xml/index.js";
import type { WorkbookRel } from "./constants.js";
import { NS } from "./constants.js";

export function renderWorkbookRels(rels: WorkbookRel[]): string {
  const w = new XmlWriter();
  w.open("Relationships", { xmlns: NS.packageRels });
  for (const rel of rels) {
    w.leaf("Relationship", { Id: rel.rId, Type: rel.type, Target: rel.target });
  }
  w.close("Relationships");
  return w.toString();
}
