// The read orchestrator (SPEC §8): `.xlsx` bytes -> parts -> a Stage 1.2 `Workbook`. The inverse of
// write-workbook.ts, mirroring its write order in reverse: open the package, parse the pooled parts
// (styles, sharedStrings), read the workbook header for the sheet list + `date1904`, then populate each
// worksheet with the shared read context, and finally lay the doc-props metadata back on top.

import { Workbook } from "../model/index.js";
import { PART_PATH, REL_TYPE } from "./parts/constants.js";
import { openPackage } from "./read/package.js";
import { applyAppProps, applyCoreProps } from "./read/props.js";
import { parseSharedStrings } from "./read/shared-strings.js";
import { parseStyles } from "./read/styles.js";
import { parseWorkbook } from "./read/workbook.js";
import { parseWorksheet } from "./read/worksheet.js";

/** Reserved for future read options (e.g. `ignoreStyles`). None are required for v1. */
// biome-ignore lint/suspicious/noEmptyInterface: intentionally empty; a stable extension point.
export interface ReadOptions {}

/** Parse a buffered `.xlsx` into a `Workbook`. Environment-agnostic (no `fs`); see `readWorkbookFile`. */
export function readWorkbookBuffer(bytes: Uint8Array, _options?: ReadOptions): Workbook {
  const pkg = openPackage(bytes);
  const rels = pkg.workbookRels();

  // The pooled parts, located by relationship type so a foreign producer's file names still resolve.
  let stylesXml: string | undefined;
  let sharedStringsXml: string | undefined;
  for (const { type, target } of rels.values()) {
    if (type === REL_TYPE.styles) stylesXml = pkg.text(target);
    else if (type === REL_TYPE.sharedStrings) sharedStringsXml = pkg.text(target);
  }

  const styles = parseStyles(stylesXml);
  const sharedStrings = sharedStringsXml !== undefined ? parseSharedStrings(sharedStringsXml) : [];

  const workbookXml = pkg.text(pkg.workbookPath()) ?? "";
  const header = parseWorkbook(workbookXml);

  const wb = new Workbook();
  if (header.date1904) wb.properties.date1904 = true;
  if (header.activeTab !== undefined) wb.views = [{ activeTab: header.activeTab }];

  const ctx = { sharedStrings, styles, date1904: header.date1904, readPart: pkg.text };

  for (const sheet of header.sheets) {
    const ws = wb.addWorksheet(sheet.name, { state: sheet.state });
    const target = rels.get(sheet.rId)?.target;
    if (target === undefined) continue;
    const sheetXml = pkg.text(target);
    if (sheetXml !== undefined) {
      parseWorksheet(ws, sheetXml, { ...ctx, rels: pkg.partRels(target) });
    }
  }

  // Doc-props sit at their conventional paths (both our writer and Excel/LibreOffice use these).
  const coreXml = pkg.text(PART_PATH.coreProps);
  if (coreXml !== undefined) applyCoreProps(wb, coreXml);
  const appXml = pkg.text(PART_PATH.appProps);
  if (appXml !== undefined) applyAppProps(wb, appXml);

  return wb;
}

/** Node-only convenience: read and parse a `.xlsx` from disk. Isolated here so browser bundles that
 *  only use {@link readWorkbookBuffer} never pull in `node:fs`. */
export async function readWorkbookFile(path: string, options?: ReadOptions): Promise<Workbook> {
  const { readFile } = await import("node:fs/promises");
  const bytes = await readFile(path);
  return readWorkbookBuffer(bytes, options);
}
