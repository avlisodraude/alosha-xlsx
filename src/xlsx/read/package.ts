// Unzip + part discovery (see docs/STAGE-2-SPEC.md §3). Mirrors the write side's `zip-layer` +
// `parts/root-rels.ts`/`parts/workbook-rels.ts` in reverse: everything is resolved through the OPC
// relationships rather than fixed paths, so a foreign producer using different part names still works.

import { InvalidPackageError, ReadLimitError, UnsupportedFeatureError } from "../../errors.js";
import { type ReadLimits, resolveReadLimits } from "../../read-limits.js";
import { cursor } from "../../xml/index.js";
import { readZipArchive, type ZipEntries } from "../../zip-layer.js";
import { CONTENT_TYPE, NS, PART_PATH, REL_TYPE } from "../parts/constants.js";

/** A relationship resolved from some part's `.rels` sibling. `target` is package-root-relative for
 *  `mode: "Internal"`; left verbatim (a URL/mailto, not a package path) for `mode: "External"`. */
export interface ResolvedRel {
  type: string;
  target: string;
  mode: "Internal" | "External";
}

/** Resolved OOXML package parts, discovered via the OPC relationships. */
export interface PackageParts {
  /** Raw bytes of a part, or `undefined` if the archive has no such entry. */
  entry(path: string): Uint8Array | undefined;
  /** UTF-8 decoded text of a part, or `undefined` if the archive has no such entry. */
  text(path: string): string | undefined;
  /** rId -> target resolved relative to the workbook part's directory (e.g. `rId1` -> `"xl/worksheets/sheet1.xml"`). */
  workbookRels(): Map<string, { type: string; target: string }>;
  /** Absolute path of the workbook part from `_rels/.rels` (normally `xl/workbook.xml`). */
  workbookPath(): string;
  /** rId -> relationship info from an arbitrary part's own `.rels` sibling (e.g. a worksheet's
   *  `xl/worksheets/_rels/sheetN.xml.rels`, holding its hyperlink relationships). Returns an empty
   *  map when the part has no `.rels` sibling. */
  partRels(partPath: string): Map<string, ResolvedRel>;
}

interface RelEntry {
  id: string;
  type: string;
  target: string;
  mode: "Internal" | "External";
}

const decoder = new TextDecoder();

/** Unzips `.xlsx` bytes and resolves the workbook part + its relationships. */
export function openPackage(bytes: Uint8Array, options?: ReadLimits): PackageParts {
  const zip = readZip(bytes, options);

  const entry = (path: string): Uint8Array | undefined => zip.get(path);
  const text = (path: string): string | undefined => {
    const raw = zip.get(path);
    return raw === undefined ? undefined : decoder.decode(raw);
  };

  const rootRelsXml = text(PART_PATH.rootRels);
  if (rootRelsXml === undefined) {
    throw new InvalidPackageError(`missing ${PART_PATH.rootRels}`);
  }
  const officeDocRel = parseRels(rootRelsXml).find((r) => r.type === REL_TYPE.officeDocument);
  if (!officeDocRel) {
    throw new InvalidPackageError(`${PART_PATH.rootRels} has no officeDocument relationship`);
  }
  const workbookPath = resolveRelTarget("", officeDocRel.target);

  const contentTypesXml = text(PART_PATH.contentTypes);
  if (contentTypesXml === undefined) {
    throw new UnsupportedFeatureError(
      `package is missing ${PART_PATH.contentTypes}; not a recognizable OOXML package`,
    );
  }
  assertSpreadsheetPackage(contentTypesXml, workbookPath);

  if (entry(workbookPath) === undefined) {
    throw new InvalidPackageError(`workbook part not found at ${workbookPath}`);
  }

  const workbookDir = dirname(workbookPath);
  const workbookRelsXml = text(relsPathFor(workbookPath));
  const relsMap = new Map<string, { type: string; target: string }>();
  if (workbookRelsXml !== undefined) {
    for (const rel of parseRels(workbookRelsXml)) {
      relsMap.set(rel.id, { type: rel.type, target: resolveRelTarget(workbookDir, rel.target) });
    }
  }

  /** rId -> relationship info for an arbitrary part, resolving `Internal` targets against that
   *  part's own directory and leaving `External` targets (URLs) untouched. */
  const partRels = (partPath: string): Map<string, ResolvedRel> => {
    const dir = dirname(partPath);
    const relsXml = text(relsPathFor(partPath));
    const map = new Map<string, ResolvedRel>();
    if (relsXml === undefined) return map;
    for (const rel of parseRels(relsXml)) {
      const target = rel.mode === "External" ? rel.target : resolveRelTarget(dir, rel.target);
      map.set(rel.id, { type: rel.type, target, mode: rel.mode });
    }
    return map;
  };

  return {
    entry,
    text,
    workbookRels: () => relsMap,
    workbookPath: () => workbookPath,
    partRels,
  };
}

function readZip(bytes: Uint8Array, options?: ReadLimits): ZipEntries {
  const limits = resolveReadLimits(options);
  try {
    return readZipArchive(bytes, limits);
  } catch (e) {
    if (e instanceof ReadLimitError) throw e;
    throw new InvalidPackageError(`could not read zip archive: ${(e as Error).message}`);
  }
}

/** Parses a `.rels` document (`_rels/.rels` or `xl/_rels/workbook.xml.rels`-shaped) into its entries. */
function parseRels(xml: string): RelEntry[] {
  const c = cursor(xml, [NS.packageRels]);
  let ev = c.next();
  while (ev && !(ev.kind === "open" && ev.name === "Relationships")) ev = c.next();
  if (ev?.kind !== "open") return [];

  const rels: RelEntry[] = [];
  if (!ev.selfClosing) {
    c.eachChild(ev.name, (child) => {
      if (child.name !== "Relationship") return;
      rels.push({
        id: child.attrs.Id ?? "",
        type: child.attrs.Type ?? "",
        target: child.attrs.Target ?? "",
        mode: child.attrs.TargetMode === "External" ? "External" : "Internal",
      });
    });
  }
  return rels;
}

/** Sanity-checks `[Content_Types].xml` declares the workbook part as a spreadsheet main part. Fails
 *  fast on foreign OPC packages (e.g. a `.docx`) rather than parsing garbage worksheets later. */
function assertSpreadsheetPackage(contentTypesXml: string, workbookPath: string): void {
  const c = cursor(contentTypesXml, [NS.contentTypes]);
  let ev = c.next();
  while (ev && !(ev.kind === "open" && ev.name === "Types")) ev = c.next();
  if (ev?.kind !== "open") {
    throw new UnsupportedFeatureError(
      `${PART_PATH.contentTypes} has no <Types> root; not a recognizable OOXML package`,
    );
  }

  let isSpreadsheet = false;
  if (!ev.selfClosing) {
    c.eachChild(ev.name, (child) => {
      if (child.name !== "Override") return;
      const partName = child.attrs.PartName?.replace(/^\//, "");
      if (partName === workbookPath && child.attrs.ContentType === CONTENT_TYPE.workbook) {
        isSpreadsheet = true;
      }
    });
  }
  if (!isSpreadsheet) {
    throw new UnsupportedFeatureError(
      `package does not declare ${workbookPath} as a spreadsheet workbook in ${PART_PATH.contentTypes}`,
    );
  }
}

function dirname(path: string): string {
  const idx = path.lastIndexOf("/");
  return idx === -1 ? "" : path.slice(0, idx);
}

/** The `.rels` sibling of a part: `xl/workbook.xml` -> `xl/_rels/workbook.xml.rels`. */
function relsPathFor(partPath: string): string {
  const dir = dirname(partPath);
  const base = dir ? partPath.slice(dir.length + 1) : partPath;
  return dir ? `${dir}/_rels/${base}.rels` : `_rels/${base}.rels`;
}

/** Resolves a relationship `Target` against the directory containing the part that references it.
 *  Handles package-root-absolute targets (`/xl/media/image1.png`) and `../` traversal. */
function resolveRelTarget(baseDir: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts = baseDir ? baseDir.split("/") : [];
  for (const seg of target.split("/")) {
    if (seg === "" || seg === ".") continue;
    if (seg === "..") parts.pop();
    else parts.push(seg);
  }
  return parts.join("/");
}
