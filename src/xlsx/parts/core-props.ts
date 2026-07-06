// `docProps/core.xml` — Dublin Core document metadata. Dates are W3CDTF (ISO 8601 without
// milliseconds), tagged `xsi:type="dcterms:W3CDTF"`. Optional fields are omitted when unset (SPEC §5).

import type { Workbook } from "../../model/index.js";
import { XmlWriter } from "../../xml/index.js";
import { NS } from "./constants.js";

/** W3CDTF: `YYYY-MM-DDTHH:mm:ssZ` (drop the `.sssZ` millisecond fraction `toISOString` adds). */
export function toW3CDTF(date: Date): string {
  return `${date.toISOString().replace(/\.\d{3}Z$/, "Z")}`;
}

export function renderCoreProps(wb: Workbook): string {
  const w = new XmlWriter();
  w.open("cp:coreProperties", {
    "xmlns:cp": NS.coreProps,
    "xmlns:dc": NS.dc,
    "xmlns:dcterms": NS.dcterms,
    "xmlns:dcmitype": NS.dcmitype,
    "xmlns:xsi": NS.xsi,
  });

  if (wb.creator !== undefined) w.el("dc:creator", wb.creator);
  if (wb.title !== undefined) w.el("dc:title", wb.title);
  if (wb.subject !== undefined) w.el("dc:subject", wb.subject);
  if (wb.description !== undefined) w.el("dc:description", wb.description);
  if (wb.keywords !== undefined) w.el("cp:keywords", wb.keywords);
  if (wb.category !== undefined) w.el("cp:category", wb.category);
  if (wb.lastModifiedBy !== undefined) w.el("cp:lastModifiedBy", wb.lastModifiedBy);
  w.el("dcterms:created", toW3CDTF(wb.created), { "xsi:type": "dcterms:W3CDTF" });
  w.el("dcterms:modified", toW3CDTF(wb.modified), { "xsi:type": "dcterms:W3CDTF" });

  w.close("cp:coreProperties");
  return w.toString();
}
