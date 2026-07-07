// `[Content_Types].xml` — declares the MIME type of every part in the package. Two base `<Default>`s
// cover the `.rels` and `.xml` extensions; one `<Override>` names each concrete part. The sharedStrings
// override is emitted only when that part exists (SPEC §5). Worksheet child parts (Tracks C/D) push
// their media `<Default>`s (e.g. `png`, `vml`) and per-part `<Override>`s through `extraDefaults`/
// `extraOverrides`; the schema requires every `<Default>` before any `<Override>`, kept invariant here.

import { XmlWriter } from "../../xml/index.js";
import { CONTENT_TYPE, NS, PART_PATH } from "./constants.js";

export interface ContentTypesInput {
  /** Worksheet file names under `xl/worksheets/` (e.g. `["sheet1.xml", "sheet2.xml"]`). */
  worksheetFiles: string[];
  hasSharedStrings: boolean;
  /** Media `<Default>`s from worksheet child parts, deduped by lowercase extension (e.g. png, vml). */
  extraDefaults?: { extension: string; contentType: string }[];
  /** Per-part `<Override>`s from worksheet child parts; `partName` is absolute (`/xl/drawings/…`). */
  extraOverrides?: { partName: string; contentType: string }[];
}

export function renderContentTypes(input: ContentTypesInput): string {
  const w = new XmlWriter();
  w.open("Types", { xmlns: NS.contentTypes });

  // All `<Default>`s first. Base two, then the deduped extras (a repeated extension yields one entry;
  // the base `rels`/`xml` are pre-seeded so an extra never re-declares them).
  const seenExt = new Set<string>(["rels", "xml"]);
  w.leaf("Default", { Extension: "rels", ContentType: CONTENT_TYPE.rels });
  w.leaf("Default", { Extension: "xml", ContentType: CONTENT_TYPE.xml });
  for (const d of input.extraDefaults ?? []) {
    const extension = d.extension.toLowerCase();
    if (seenExt.has(extension)) continue;
    seenExt.add(extension);
    w.leaf("Default", { Extension: extension, ContentType: d.contentType });
  }

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
  for (const o of input.extraOverrides ?? []) {
    w.leaf("Override", { PartName: o.partName, ContentType: o.contentType });
  }

  w.close("Types");
  return w.toString();
}
