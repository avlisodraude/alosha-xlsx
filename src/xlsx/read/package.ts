// Unzip + part discovery (see docs/STAGE-2-SPEC.md §3). Mirrors the write side's `zip-layer` +
// `parts/root-rels.ts`/`parts/workbook-rels.ts` in reverse: everything is resolved through the OPC
// relationships rather than fixed paths, so a foreign producer using different part names still works.

import { InvalidPackageError, UnsupportedFeatureError } from "../../errors.js";
import { cursor } from "../../xml/index.js";
import { readZipArchive, type ZipEntries } from "../../zip-layer.js";
import { CONTENT_TYPE, PART_PATH, REL_TYPE } from "../parts/constants.js";

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
}

interface RelEntry {
  id: string;
  type: string;
  target: string;
}

const decoder = new TextDecoder();

/** Unzips `.xlsx` bytes and resolves the workbook part + its relationships. */
export function openPackage(bytes: Uint8Array): PackageParts {
  const zip = readZip(bytes);

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

  return {
    entry,
    text,
    workbookRels: () => relsMap,
    workbookPath: () => workbookPath,
  };
}

function readZip(bytes: Uint8Array): ZipEntries {
  try {
    return readZipArchive(bytes);
  } catch (e) {
    throw new InvalidPackageError(`could not read zip archive: ${(e as Error).message}`);
  }
}

/** Parses a `.rels` document (`_rels/.rels` or `xl/_rels/workbook.xml.rels`-shaped) into its entries. */
function parseRels(xml: string): RelEntry[] {
  const c = cursor(xml);
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
      });
    });
  }
  return rels;
}

/** Sanity-checks `[Content_Types].xml` declares the workbook part as a spreadsheet main part. Fails
 *  fast on foreign OPC packages (e.g. a `.docx`) rather than parsing garbage worksheets later. */
function assertSpreadsheetPackage(contentTypesXml: string, workbookPath: string): void {
  const c = cursor(contentTypesXml);
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
