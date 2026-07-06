// `_rels/.rels` — the package root relationships: the office document (workbook), and the two
// document-property parts (core + app). rIds are local to this file (SPEC §5).

import { XmlWriter } from "../../xml/index.js";
import { NS, PART_PATH, REL_TYPE } from "./constants.js";

export function renderRootRels(): string {
  const w = new XmlWriter();
  w.open("Relationships", { xmlns: NS.packageRels });
  w.leaf("Relationship", {
    Id: "rId1",
    Type: REL_TYPE.officeDocument,
    Target: PART_PATH.workbook,
  });
  w.leaf("Relationship", {
    Id: "rId2",
    Type: REL_TYPE.coreProps,
    Target: PART_PATH.coreProps,
  });
  w.leaf("Relationship", {
    Id: "rId3",
    Type: REL_TYPE.extProps,
    Target: PART_PATH.appProps,
  });
  w.close("Relationships");
  return w.toString();
}
