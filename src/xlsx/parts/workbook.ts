// `xl/workbook.xml` — the workbook part: the `<sheets>` list (each `<sheet>` links to its worksheet
// part via `r:id`), `<workbookPr date1904=…>` when the 1904 system is set, and a minimal `<bookViews>`.
// The `r:id`s here resolve against xl/_rels/workbook.xml.rels, so they must match the SheetRel rIds
// assigned once by the orchestrator (SPEC §5).

import type { Workbook } from "../../model/index.js";
import { XmlWriter } from "../../xml/index.js";
import type { SheetRel } from "./constants.js";
import { NS } from "./constants.js";

export function renderWorkbook(wb: Workbook, sheets: SheetRel[]): string {
  const w = new XmlWriter();
  w.open("workbook", { xmlns: NS.spreadsheetml, "xmlns:r": NS.relationships });

  if (wb.properties.date1904) {
    w.leaf("workbookPr", { date1904: true });
  }

  const activeTab = typeof wb.views[0]?.activeTab === "number" ? wb.views[0].activeTab : undefined;
  w.open("bookViews");
  w.leaf("workbookView", { activeTab });
  w.close("bookViews");

  w.open("sheets");
  for (const sheet of sheets) {
    w.leaf("sheet", {
      name: sheet.name,
      sheetId: sheet.sheetId,
      state: sheet.state === "visible" ? undefined : sheet.state,
      "r:id": sheet.rId,
    });
  }
  w.close("sheets");

  if (wb.calcProperties.fullCalcOnLoad) {
    w.leaf("calcPr", { calcId: 0, fullCalcOnLoad: true });
  }

  w.close("workbook");
  return w.toString();
}
