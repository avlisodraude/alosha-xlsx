// `[Content_Types].xml` — declares the MIME type of every part in the package. Two `<Default>`s cover
// the `.rels` and `.xml` extensions; one `<Override>` names each concrete part. The sharedStrings
// override is emitted only when that part exists (SPEC §5).

import { XmlWriter } from "../../xml/index.js";
import { CONTENT_TYPE, NS, PART_PATH } from "./constants.js";

export interface ContentTypesInput {
  /** Worksheet file names under `xl/worksheets/` (e.g. `["sheet1.xml", "sheet2.xml"]`). */
  worksheetFiles: string[];
  hasSharedStrings: boolean;
}

export function renderContentTypes(input: ContentTypesInput): string {
  const w = new XmlWriter();
  w.open("Types", { xmlns: NS.contentTypes });

  w.leaf("Default", { Extension: "rels", ContentType: CONTENT_TYPE.rels });
  w.leaf("Default", { Extension: "xml", ContentType: CONTENT_TYPE.xml });

  w.leaf("Override", { PartName: `/${PART_PATH.workbook}`, ContentType: CONTENT_TYPE.workbook });
  for (const file of input.worksheetFiles) {
    w.leaf("Override", {
      PartName: `/xl/worksheets/${file}`,
      ContentType: CONTENT_TYPE.worksheet,
    });
  }
  w.leaf("Override", { PartName: `/${PART_PATH.styles}`, ContentType: CONTENT_TYPE.styles });
  if (input.hasSharedStrings) {
    w.leaf("Override", {
      PartName: `/${PART_PATH.sharedStrings}`,
      ContentType: CONTENT_TYPE.sharedStrings,
    });
  }
  w.leaf("Override", { PartName: `/${PART_PATH.theme}`, ContentType: CONTENT_TYPE.theme });
  w.leaf("Override", { PartName: `/${PART_PATH.coreProps}`, ContentType: CONTENT_TYPE.coreProps });
  w.leaf("Override", { PartName: `/${PART_PATH.appProps}`, ContentType: CONTENT_TYPE.extProps });

  w.close("Types");
  return w.toString();
}
