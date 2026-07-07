// Shared OPC/OOXML constants (namespaces, content types, relationship-type URIs) used across the
// part writers. These are the fixed identifiers from ECMA-376 / the Open Packaging Conventions —
// clean-room, not copied from any implementation.

/** XML namespaces. */
export const NS = {
  spreadsheetml: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  relationships: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  contentTypes: "http://schemas.openxmlformats.org/package/2006/content-types",
  packageRels: "http://schemas.openxmlformats.org/package/2006/relationships",
  coreProps: "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
  extProps: "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
  docPropsVt: "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
  dc: "http://purl.org/dc/elements/1.1/",
  dcterms: "http://purl.org/dc/terms/",
  dcmitype: "http://purl.org/dc/dcmitype/",
  xsi: "http://www.w3.org/2001/XMLSchema-instance",
} as const;

/** Part content types (`[Content_Types].xml` overrides + the two defaults). */
export const CONTENT_TYPE = {
  rels: "application/vnd.openxmlformats-package.relationships+xml",
  xml: "application/xml",
  workbook: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
  worksheet: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
  styles: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
  sharedStrings: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
  theme: "application/vnd.openxmlformats-officedocument.theme+xml",
  coreProps: "application/vnd.openxmlformats-package.core-properties+xml",
  extProps: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
} as const;

/** Relationship-type URIs (the `Type=` attribute on each `<Relationship>`). */
export const REL_TYPE = {
  officeDocument:
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
  coreProps:
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
  extProps:
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
  worksheet: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
  styles: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
  sharedStrings:
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
  theme: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
  hyperlink: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
} as const;

/** Archive paths for the fixed (non-worksheet) parts. Worksheet paths are `xl/worksheets/sheetN.xml`. */
export const PART_PATH = {
  contentTypes: "[Content_Types].xml",
  rootRels: "_rels/.rels",
  coreProps: "docProps/core.xml",
  appProps: "docProps/app.xml",
  workbook: "xl/workbook.xml",
  workbookRels: "xl/_rels/workbook.xml.rels",
  styles: "xl/styles.xml",
  sharedStrings: "xl/sharedStrings.xml",
  theme: "xl/theme/theme1.xml",
} as const;

/** A resolved relationship-id assignment shared by workbook.xml and workbook.xml.rels so the
 *  `r:id`s in the former always match the `Id`s in the latter. */
export interface WorkbookRel {
  rId: string;
  type: string;
  /** Target relative to `xl/` (e.g. `worksheets/sheet1.xml`). */
  target: string;
}

/** A single worksheet-level external relationship (e.g. a hyperlink). Its `rId` is local to the
 *  worksheet's own `.rels` namespace and must match the `r:id` emitted in the sheet's `<hyperlinks>`. */
export interface WorksheetRel {
  rId: string;
  /** The link destination (e.g. `https://example.com`, `mailto:a@b.c`). */
  target: string;
  /** `TargetMode` attribute; external hyperlinks are always `"External"`. */
  mode: "External";
}

/** A worksheet's identity within the package: its position-derived file, its `r:id`, and model data. */
export interface SheetRel {
  rId: string;
  /** File name under `xl/worksheets/` (e.g. `sheet1.xml`). */
  fileName: string;
  /** `sheetId` attribute in workbook.xml (the model's stable worksheet id). */
  sheetId: number;
  name: string;
  state: "visible" | "hidden" | "veryHidden";
}
