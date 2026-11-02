// The write orchestrator (SPEC §7): in-memory Workbook -> the OOXML part set -> zip -> Uint8Array.
// Single pass, mirroring OOXML-NOTES §2's write order: worksheets first (they populate the shared-
// strings and styles pools as cells emit), then the pooled parts, then the wiring parts.
//
// Relationship-id discipline: sheet rIds (rId1..N) are assigned once and shared between workbook.xml
// and workbook.xml.rels; styles/sharedStrings/theme take the ids after them. sharedStrings only gets
// an id (and a part) when the pool is non-empty.

import type { Workbook } from "../model/index.js";
import { writeFileBytes } from "../platform/io.js";
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
import {
  type RenderedWorksheet,
  renderWorksheet,
  type WorksheetContext,
} from "./parts/worksheet.js";
import { createSharedStrings } from "./shared-strings.js";
import { createStyles, type Styles } from "./styles.js";

/** The worksheet child parts gathered across every sheet, ready to fold into the archive and its
 *  `[Content_Types].xml`. See {@link collectWorksheetParts}. */
export interface CollectedParts {
  /** Archive path -> bytes (already-encoded XML parts and raw media). */
  media: Map<string, Uint8Array>;
  extraDefaults: { extension: string; contentType: string }[];
  extraOverrides: { partName: string; contentType: string }[];
}

/** Walk each rendered worksheet's `extraParts`, encoding string data via `enc` and passing raw bytes
 *  through untouched into `media`, and gather each part's `contentType` (as an `<Override>`, partName
 *  `/${path}`) and `defaultFor` (as a `<Default>`). Pure — the testable seam Tracks C/D build on. */
export function collectWorksheetParts(
  rendered: Iterable<RenderedWorksheet>,
  enc: InstanceType<typeof TextEncoder>,
): CollectedParts {
  const media = new Map<string, Uint8Array>();
  const extraDefaults: { extension: string; contentType: string }[] = [];
  const extraOverrides: { partName: string; contentType: string }[] = [];

  for (const { extraParts } of rendered) {
    if (!extraParts) continue;
    for (const part of extraParts) {
      media.set(part.path, typeof part.data === "string" ? enc.encode(part.data) : part.data);
      if (part.contentType) {
        extraOverrides.push({ partName: `/${part.path}`, contentType: part.contentType });
      }
      if (part.defaultFor) extraDefaults.push(part.defaultFor);
    }
  }
  return { media, extraDefaults, extraOverrides };
}

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
    : { addStyle: () => 0, addDxf: () => 0, toXml: styles.toXml.bind(styles) };

  // Conditional-format rules with no explicit priority draw from a single workbook-monotonic counter,
  // shared across every sheet's render pass (mirrors ExcelJS).
  let cfPriority = 0;
  const ctx: WorksheetContext = {
    styles: cellStyles,
    sharedStrings,
    useSharedStrings,
    date1904,
    nextPriority: () => ++cfPriority,
  };

  // 1. Worksheets first — this is what fills the shared-strings and styles pools. `sheetNumber` is the
  //    1-based position so any child parts can name themselves deterministically.
  const worksheetXml = new Map<string, RenderedWorksheet>();
  for (const [i, ws] of worksheets.entries()) {
    ctx.sheetNumber = i + 1;
    const fileName = `sheet${i + 1}.xml`;
    worksheetXml.set(fileName, renderWorksheet(ws, ctx));
  }

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

  // Gather worksheet child parts (empty unless a feature contributed any) so their content-type
  // Defaults/Overrides can fold into `[Content_Types].xml` and their bytes into the archive.
  const collected = collectWorksheetParts(worksheetXml.values(), enc);

  put(
    PART_PATH.contentTypes,
    renderContentTypes({
      worksheetFiles: sheetRels.map((s) => s.fileName),
      hasSharedStrings,
      extraDefaults: collected.extraDefaults,
      extraOverrides: collected.extraOverrides,
    }),
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
  for (const [file, { xml, rels }] of worksheetXml) {
    put(`xl/worksheets/${file}`, xml);
    // The worksheet `.rels` file is covered by the existing `<Default Extension="rels" …>` in
    // [Content_Types].xml, so no content-types change is needed.
    if (rels !== undefined) put(`xl/worksheets/_rels/${file}.rels`, rels);
  }
  // Worksheet child parts (XML already encoded, media raw) land at their declared paths — once, not
  // per sheet. Empty in the common case, leaving the archive byte-identical.
  for (const [path, bytes] of collected.media) entries.set(path, bytes);

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
  await writeFileBytes(path, bytes);
}
