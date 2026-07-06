// The write orchestrator (SPEC §7): in-memory Workbook -> the OOXML part set -> zip -> Uint8Array.
// Single pass, mirroring OOXML-NOTES §2's write order: worksheets first (they populate the shared-
// strings and styles pools as cells emit), then the pooled parts, then the wiring parts.
//
// Relationship-id discipline: sheet rIds (rId1..N) are assigned once and shared between workbook.xml
// and workbook.xml.rels; styles/sharedStrings/theme take the ids after them. sharedStrings only gets
// an id (and a part) when the pool is non-empty.

import type { Workbook } from "../model/index.js";
import { writeZipArchive, type ZipEntries } from "../zip-layer.js";
import { renderAppProps } from "./parts/app-props.js";
import type { SheetRel, WorkbookRel } from "./parts/constants.js";
import { PART_PATH, REL_TYPE } from "./parts/constants.js";
import { renderContentTypes } from "./parts/content-types.js";
import { renderCoreProps } from "./parts/core-props.js";
import { renderRootRels } from "./parts/root-rels.js";
import { renderTheme } from "./parts/theme.js";
import { renderWorkbook } from "./parts/workbook.js";
import { renderWorkbookRels } from "./parts/workbook-rels.js";
import { renderWorksheet } from "./parts/worksheet.js";
import { createSharedStrings } from "./shared-strings.js";
import { createStyles, type Styles } from "./styles.js";

export interface WriteOptions {
  /** Emit a shared-string table (string cells store an index). When `false`, strings render inline.
   *  Default `true`. */
  useSharedStrings?: boolean;
  /** Intern cell/row/column styles into styles.xml. When `false`, cells render unstyled (`s` omitted)
   *  and styles.xml carries only the mandatory baseline. Default `true`. */
  useStyles?: boolean;
}

/** Serialize a workbook to `.xlsx` bytes. Environment-agnostic (no `fs`); see `writeWorkbookFile`. */
export function writeWorkbookBuffer(wb: Workbook, options?: WriteOptions): Uint8Array {
  const useSharedStrings = options?.useSharedStrings !== false;
  const useStyles = options?.useStyles !== false;
  const date1904 = wb.properties.date1904 === true;

  const worksheets = wb.worksheets;
  const sheetRels: SheetRel[] = worksheets.map((ws, i) => ({
    rId: `rId${i + 1}`,
    fileName: `sheet${i + 1}.xml`,
    sheetId: ws.id,
    name: ws.name,
    state: ws.state,
  }));

  const sharedStrings = createSharedStrings();
  const styles = createStyles();
  // When styles are disabled, cells intern nothing (always `s=0`); styles.xml still ships the baseline.
  const cellStyles: Styles = useStyles
    ? styles
    : { addStyle: () => 0, toXml: styles.toXml.bind(styles) };

  const ctx = { styles: cellStyles, sharedStrings, useSharedStrings, date1904 };

  // 1. Worksheets first — this is what fills the shared-strings and styles pools.
  const worksheetXml = new Map<string, string>();
  worksheets.forEach((ws, i) => {
    worksheetXml.set(sheetRels[i]!.fileName, renderWorksheet(ws, ctx));
  });

  // 2. Now the pools are populated: decide whether a sharedStrings part exists.
  const sharedStringsXml = useSharedStrings ? sharedStrings.toXml() : undefined;
  const hasSharedStrings = sharedStringsXml !== undefined;

  // 3. Assign the remaining rIds after the sheets, then wire the workbook rels.
  let counter = worksheets.length;
  const stylesRId = `rId${++counter}`;
  const sharedRId = hasSharedStrings ? `rId${++counter}` : undefined;
  const themeRId = `rId${++counter}`;

  const workbookRels: WorkbookRel[] = sheetRels.map((s) => ({
    rId: s.rId,
    type: REL_TYPE.worksheet,
    target: `worksheets/${s.fileName}`,
  }));
  workbookRels.push({ rId: stylesRId, type: REL_TYPE.styles, target: "styles.xml" });
  if (sharedRId) {
    workbookRels.push({
      rId: sharedRId,
      type: REL_TYPE.sharedStrings,
      target: "sharedStrings.xml",
    });
  }
  workbookRels.push({ rId: themeRId, type: REL_TYPE.theme, target: "theme/theme1.xml" });

  // 4. Assemble every part.
  const enc = new TextEncoder();
  const entries: ZipEntries = new Map();
  const put = (path: string, xml: string) => entries.set(path, enc.encode(xml));

  put(
    PART_PATH.contentTypes,
    renderContentTypes({ worksheetFiles: sheetRels.map((s) => s.fileName), hasSharedStrings }),
  );
  put(PART_PATH.rootRels, renderRootRels());
  put(PART_PATH.coreProps, renderCoreProps(wb));
  put(
    PART_PATH.appProps,
    renderAppProps(
      wb,
      sheetRels.map((s) => s.name),
    ),
  );
  put(PART_PATH.workbook, renderWorkbook(wb, sheetRels));
  put(PART_PATH.workbookRels, renderWorkbookRels(workbookRels));
  put(PART_PATH.styles, styles.toXml());
  if (sharedStringsXml !== undefined) put(PART_PATH.sharedStrings, sharedStringsXml);
  put(PART_PATH.theme, renderTheme());
  for (const [file, xml] of worksheetXml) put(`xl/worksheets/${file}`, xml);

  return writeZipArchive(entries);
}

/** Node-only convenience: serialize and write to disk. Isolated here so browser bundles that only
 *  use {@link writeWorkbookBuffer} never pull in `node:fs`. */
export async function writeWorkbookFile(
  wb: Workbook,
  path: string,
  options?: WriteOptions,
): Promise<void> {
  const bytes = writeWorkbookBuffer(wb, options);
  const { writeFile } = await import("node:fs/promises");
  await writeFile(path, bytes);
}
