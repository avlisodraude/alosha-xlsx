import type { WarningReporter } from "../../read-options.js";
import { cursor, localName, parseXml } from "../../xml/index.js";
import { namespaceEvents } from "../../xml/namespaces.js";
import { NS } from "../parts/constants.js";
import type { PackageParts } from "./package.js";

const WORKBOOK_LOSS = new Set([
  "definedNames",
  "externalReferences",
  "workbookProtection",
  "customWorkbookViews",
  "calcPr",
]);
const SHEET_LOSS = new Set([
  "sheetProtection",
  "pageSetup",
  "pageMargins",
  "printOptions",
  "headerFooter",
  "rowBreaks",
  "colBreaks",
  "tableParts",
  "oleObjects",
  "controls",
  "extLst",
  "sortState",
  "filterColumn",
  "selection",
]);

/** Opt-in preflight for known discarded parts and markup. This intentionally does not promise
 * complete schema validation or lossless preservation of every producer-specific extension. */
export function inspectImport(pkg: PackageParts, report: WarningReporter): void {
  const declared = new Map<string, string>();
  const types = cursor(pkg.text("[Content_Types].xml") ?? "", [NS.contentTypes]);
  for (let event = types.next(); event; event = types.next()) {
    if (event.kind !== "open" || event.name !== "Override") continue;
    const contentType = event.attrs.ContentType ?? "";
    const feature = /drawingml\.chart\+xml/.test(contentType)
      ? "charts"
      : /spreadsheetml\.table\+xml/.test(contentType)
        ? "tables"
        : /pivot(?:Table|Cache)/.test(contentType)
          ? "pivot tables"
          : /threadedcomments/i.test(contentType)
            ? "threaded comments"
            : /custom-properties/.test(contentType)
              ? "custom document properties"
              : undefined;
    if (feature && event.attrs.PartName)
      declared.set(event.attrs.PartName.replace(/^\//, ""), feature);
  }
  for (const part of pkg.paths()) {
    const feature =
      declared.get(part) ??
      (/(^|\/)charts\//.test(part)
        ? "charts"
        : /(^|\/)tables\//.test(part)
          ? "tables"
          : /(^|\/)pivot(?:Tables|Cache)\//.test(part)
            ? "pivot tables"
            : /(^|\/)threadedComments\//.test(part)
              ? "threaded comments"
              : undefined);
    if (feature && !part.endsWith(".rels"))
      report({
        code: "UNSUPPORTED_FEATURE",
        part,
        feature,
        message: `${feature} are not retained when this workbook is saved`,
      });
    if (!part.endsWith(".xml")) continue;
    const xml = pkg.text(part) ?? "";
    const workbook = part === pkg.workbookPath();
    const worksheet = /<(?:(?:[\w.-]+):)?worksheet(?:\s|\/?>)/.test(xml);
    if (!workbook && !worksheet) continue;
    for (const _event of namespaceEvents(parseXml(xml), [NS.spreadsheetml], (name, uri) => {
      report({
        code: "UNSUPPORTED_FEATURE",
        part,
        feature: "XML namespace",
        message: `${name} in namespace ${uri ?? "(none)"} is not read`,
      });
    })) {
      /* consume to inspect scoped namespace declarations */
    }
    for (const event of parseXml(xml)) {
      if (event.kind !== "open") continue;
      const name = localName(event.name);
      if ((workbook ? WORKBOOK_LOSS : SHEET_LOSS).has(name))
        report({
          code: "UNSUPPORTED_FEATURE",
          part,
          feature: name,
          message: `${name} is not retained by the document model`,
        });
    }
  }
  // Include package-level metadata rels as well as workbook/sheet/drawing rels.
  for (const host of ["", ...[...pkg.paths()].filter((path) => path.endsWith(".xml"))]) {
    for (const [ref, rel] of pkg.partRels(host)) {
      if (rel.mode === "Internal" && pkg.entry(rel.target) === undefined)
        report({
          code: "UNRESOLVED_RELATIONSHIP",
          part: host || "_rels/.rels",
          feature: "relationship",
          ref,
          message: `Relationship ${ref} targets missing part ${rel.target}`,
        });
    }
  }
}

export function checkPartRoot(
  xml: string,
  part: string,
  expected: string,
  report?: WarningReporter,
): void {
  if (!report) return;
  const c = cursor(xml, [NS.spreadsheetml]);
  let event = c.next();
  while (event && event.kind !== "open") event = c.next();
  if (event?.kind !== "open" || localName(event.name) !== expected)
    report({
      code: "MALFORMED_PART",
      part,
      feature: "XML root",
      message: `Expected a ${expected} root in the spreadsheet namespace`,
    });
}
