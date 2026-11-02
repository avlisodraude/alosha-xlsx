// The read orchestrator (SPEC §8): `.xlsx` bytes -> parts -> a Stage 1.2 `Workbook`. The inverse of
// write-workbook.ts, mirroring its write order in reverse: open the package, parse the pooled parts
// (styles, sharedStrings), read the workbook header for the sheet list + `date1904`, then populate each
// worksheet with the shared read context, and finally lay the doc-props metadata back on top.

import type { ImageExtension, Worksheet } from "../model/index.js";
import { Workbook } from "../model/index.js";
import { withModelReadBudget } from "../model/read-budget.js";
import { readFileBytes } from "../platform/io.js";
import { ModelReadBudget, type ReadLimits, resolveReadLimits } from "../read-limits.js";
import { IMAGE_CONTENT_TYPE, PART_PATH, REL_TYPE } from "./parts/constants.js";
import { parseDrawing } from "./read/drawing.js";
import { openPackage, type PackageParts, type ResolvedRel } from "./read/package.js";
import { applyAppProps, applyCoreProps } from "./read/props.js";
import { parseSharedStrings } from "./read/shared-strings.js";
import { parseStyles } from "./read/styles.js";
import { parseWorkbook } from "./read/workbook.js";
import { parseWorksheet } from "./read/worksheet.js";

/** Input/archive and model allocation limits; omitted limits use DEFAULT_READ_LIMITS. */
export type ReadOptions = ReadLimits;

/** Parse a buffered `.xlsx` into a `Workbook`. Environment-agnostic (no `fs`); see `readWorkbookFile`. */
export function readWorkbookBuffer(bytes: Uint8Array, options?: ReadOptions): Workbook {
  const budget = new ModelReadBudget(resolveReadLimits(options));
  const pkg = openPackage(bytes, options);
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

  // Images are interned once per distinct media target and shared across sheets that reference the
  // same drawing media (mirrors the workbook-wide `ImageRegistry` the writer reads from).
  const mediaCache = new Map<string, number>();

  for (const sheet of header.sheets) {
    const ws = wb.addWorksheet(sheet.name, { state: sheet.state });
    const target = rels.get(sheet.rId)?.target;
    if (target === undefined) continue;
    const sheetXml = pkg.text(target);
    const sheetRels = pkg.partRels(target);
    withModelReadBudget(ws, budget, () => {
      if (sheetXml !== undefined) {
        parseWorksheet(ws, sheetXml, { ...ctx, rels: sheetRels });
      }
      applyDrawing(wb, ws, pkg, sheetRels, mediaCache);
    });
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
  const limits = resolveReadLimits(options);
  const bytes = await readFileBytes(path, limits.maxInputBytes);
  return readWorkbookBuffer(bytes, options);
}

/** Inverse of the writer's `renderImages` (Track D, docs/IMAGES-DRAWINGS-SPEC.md §3): follow the
 *  sheet's `<drawing r:id>` rel to `drawing{n}.xml`, parse its anchors, resolve each anchor's `r:embed`
 *  through the drawing's own `.rels` to a media part, and place the bytes on the workbook registry +
 *  this sheet. Any missing/unresolvable part (no drawing rel, no drawing part, no matching drawing rel,
 *  no media entry, an unrecognized extension) silently drops that image rather than throwing — a
 *  malformed package should degrade to "no image", not fail the whole read. */
function applyDrawing(
  wb: Workbook,
  ws: Worksheet,
  pkg: PackageParts,
  sheetRels: Map<string, ResolvedRel>,
  mediaCache: Map<string, number>,
): void {
  let drawingTarget: string | undefined;
  for (const rel of sheetRels.values()) {
    if (rel.type === REL_TYPE.drawing) {
      drawingTarget = rel.target;
      break;
    }
  }
  if (drawingTarget === undefined) return;

  const drawingXml = pkg.text(drawingTarget);
  if (drawingXml === undefined) return;

  const images = parseDrawing(drawingXml);
  if (images.length === 0) return;

  const drawingRels = pkg.partRels(drawingTarget);

  for (const img of images) {
    const rel = drawingRels.get(img.embedId);
    if (!rel || rel.mode === "External") continue;
    const bytes = pkg.entry(rel.target);
    if (bytes === undefined) continue;
    const extension = imageExtensionOf(rel.target);
    if (extension === undefined) continue;

    let imageId = mediaCache.get(rel.target);
    if (imageId === undefined) {
      imageId = wb.addImage({ buffer: bytes, extension });
      mediaCache.set(rel.target, imageId);
    }
    ws.addImage(imageId, img.anchor);
  }
}

/** The media part's extension, lowercased, iff it's one of the formats {@link ImageExtension} supports
 *  (`IMAGE_CONTENT_TYPE`'s keys) — `undefined` for an unsupported/extension-less foreign media part. */
function imageExtensionOf(path: string): ImageExtension | undefined {
  const ext = path.slice(path.lastIndexOf(".") + 1).toLowerCase();
  return ext in IMAGE_CONTENT_TYPE ? (ext as ImageExtension) : undefined;
}
