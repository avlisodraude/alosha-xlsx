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
  drawing: "application/vnd.openxmlformats-officedocument.drawing+xml",
  comments: "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
  vmlDrawing: "application/vnd.openxmlformats-officedocument.vmlDrawing",
} as const;

/** Image content-type by lowercase file extension, for the media `<Default>`s (Track D). */
export const IMAGE_CONTENT_TYPE = {
  png: "image/png",
  jpeg: "image/jpeg",
  jpg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
} as const satisfies Record<string, string>;

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
  drawing: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
  comments: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
  vmlDrawing: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
  image: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
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

/** A single worksheet-level relationship — a hyperlink, or an internal child part (drawing, comments,
 *  vmlDrawing, image). Its `rId` is local to the worksheet's own `.rels` namespace and must match the
 *  `r:id` emitted in the sheet body (e.g. `<hyperlink r:id>`, `<drawing r:id>`). */
export interface WorksheetRel {
  rId: string;
  /** A `REL_TYPE.*` URI (hyperlink, drawing, comments, …). */
  type: string;
  /** External URL, or an internal path relative to `xl/worksheets/` (e.g. `../drawings/drawing1.xml`). */
  target: string;
  /** `TargetMode` attribute — `"External"` for hyperlinks, omitted for internal rels. */
  mode?: "External";
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
