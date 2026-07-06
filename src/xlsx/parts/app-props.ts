// `docProps/app.xml` — extended (application) properties. The `HeadingPairs`/`TitlesOfParts`
// vectors describe the worksheet tab names; Company/Manager ride along when set (SPEC §5).

import type { Workbook } from "../../model/index.js";
import { XmlWriter } from "../../xml/index.js";
import { NS } from "./constants.js";

export function renderAppProps(wb: Workbook, sheetNames: string[]): string {
  const w = new XmlWriter();
  w.open("Properties", { xmlns: NS.extProps, "xmlns:vt": NS.docPropsVt });

  w.el("Application", "Microsoft Excel");
  w.el("DocSecurity", "0");
  w.el("ScaleCrop", "false");

  w.open("HeadingPairs");
  w.open("vt:vector", { size: 2, baseType: "variant" });
  w.open("vt:variant");
  w.el("vt:lpstr", "Worksheets");
  w.close("vt:variant");
  w.open("vt:variant");
  w.el("vt:i4", String(sheetNames.length));
  w.close("vt:variant");
  w.close("vt:vector");
  w.close("HeadingPairs");

  w.open("TitlesOfParts");
  w.open("vt:vector", { size: sheetNames.length, baseType: "lpstr" });
  for (const name of sheetNames) w.el("vt:lpstr", name);
  w.close("vt:vector");
  w.close("TitlesOfParts");

  if (wb.company !== undefined) w.el("Company", wb.company);
  if (wb.manager !== undefined) w.el("Manager", wb.manager);
  w.el("LinksUpToDate", "false");
  w.el("SharedDoc", "false");
  w.el("HyperlinksChanged", "false");
  w.el("AppVersion", "1.0000");

  w.close("Properties");
  return w.toString();
}
