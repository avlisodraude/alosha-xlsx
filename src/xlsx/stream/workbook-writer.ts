// Streaming workbook writer (SPEC §1): an incremental, ExcelJS-shaped API —
//
//   const wb = new WorkbookWriter({ stream /* or */ filename, useSharedStrings });
//   const ws = wb.addWorksheet("Sheet1");
//   ws.addRow([1, "a", new Date()]).commit();  // this row's XML is deflated to the sink now
//   ws.commit();                               // close this sheet's worksheet part
//   await wb.commit();                         // write styles/sharedStrings/workbook/… and end the zip
//
// It holds none of the buffered writer's model in memory: rows are rendered one at a time via the shared
// `renderRowXml` and pushed straight into an open zip entry. The one thing OOXML forces us to retain is
// the sharedStrings table (sheets reference it by index but it is written last), so this writer defaults
// `useSharedStrings: false` (inline `<is>` strings) — then a committed row buffers nothing. The styles
// pool is small and stays resident. This is a standalone constructor, deliberately off the `Workbook`
// prototype (which keeps its buffered `workbook.xlsx` accessor).

import { colLetterToNumber } from "../../address/index.js";
import { DuplicateWorksheetNameError } from "../../errors.js";
import type { RowValues, WorksheetHost } from "../../model/index.js";
import { Column, DataValidations, Row, Workbook } from "../../model/index.js";
import { XmlWriter } from "../../xml/index.js";
import { renderAppProps } from "../parts/app-props.js";
import type { SheetRel, WorkbookRel } from "../parts/constants.js";
import { NS, PART_PATH, REL_TYPE } from "../parts/constants.js";
import { renderContentTypes } from "../parts/content-types.js";
import { renderCoreProps } from "../parts/core-props.js";
import { renderRootRels } from "../parts/root-rels.js";
import { renderTheme } from "../parts/theme.js";
import { renderWorkbook } from "../parts/workbook.js";
import { renderWorkbookRels } from "../parts/workbook-rels.js";
import { renderWorksheetRels } from "../parts/worksheet-rels.js";
import {
  createSheetRelAllocator,
  type ExtraPart,
  type HyperlinkRecord,
  renderHyperlinks,
  renderRowXml,
  renderSheetViews,
  type WorksheetContext,
} from "../parts/worksheet.js";
import { createSharedStrings } from "../shared-strings.js";
import { createStyles, type Styles } from "../styles.js";
import { drainToFile, drainToWritable, type NodeWritableLike, QueueSink } from "./output.js";
import { StreamingZip } from "./zip.js";

type SheetState = "visible" | "hidden" | "veryHidden";

/** A row handed back by {@link WorksheetWriter.addRow}. It is a full model {@link Row} — set cell
 *  values/styles or row `height`/`hidden` on it — plus a `commit()` that flushes its XML to the sink. */
export type StreamedRow = Row & { commit(): void };

export interface WorkbookWriterOptions {
  /** A Node `Writable` (e.g. `fs.createWriteStream(path)`) to pipe the archive into. Mutually
   *  exclusive with {@link filename}; when neither is set the writer is itself an
   *  `AsyncIterable<Uint8Array>`. */
  stream?: NodeWritableLike;
  /** A file path to write the archive to (Node only; `node:fs` is imported lazily). */
  filename?: string;
  /** Emit a shared-string table. **Defaults to `false`** (inline strings) so committed rows retain
   *  nothing — the opposite of the buffered writer's `true` default. */
  useSharedStrings?: boolean;
  /** Intern styles into styles.xml (see the buffered `WriteOptions.useStyles`). Default `true`. */
  useStyles?: boolean;
}

/** The narrow {@link WorksheetHost} a transient {@link Row} needs to build its cells — no row storage,
 *  no merges, just column projections so `Row.getCell` can resolve a cell's style. */
class RowBuilderHost implements WorksheetHost {
  readonly properties = { outlineLevelCol: 0, outlineLevelRow: 0 } as WorksheetHost["properties"];
  readonly dataValidations = new DataValidations();
  private readonly columns = new Map<number, Column>();

  getColumn(col: number | string): Column {
    const n = typeof col === "number" ? col : colLetterToNumber(col);
    let c = this.columns.get(n);
    if (!c) {
      c = new Column(this, n);
      this.columns.set(n, c);
    }
    return c;
  }

  getColumnKey(): Column | undefined {
    return undefined;
  }

  getRow(): Row {
    throw new Error("RowBuilderHost: getRow is not supported in the streaming writer");
  }

  eachRow(): void {
    /* no rows are retained */
  }
}

/** Streams a single worksheet part. Emits the prologue (`<worksheet>…<sheetData>`) on the first
 *  committed row (or on {@link commit} if the sheet is empty), a `<row>` per committed row, and the
 *  epilogue (`</sheetData>`, any `<hyperlinks>`, `</worksheet>`) plus the sheet `.rels` on commit. */
export class WorksheetWriter {
  private readonly host = new RowBuilderHost();
  private readonly hyperlinks: HyperlinkRecord[] = [];
  private rowCount = 0;
  private started = false;
  private committed = false;

  constructor(
    private readonly workbook: WorkbookWriter,
    readonly rel: SheetRel,
  ) {}

  get name(): string {
    return this.rel.name;
  }

  get isCommitted(): boolean {
    return this.committed;
  }

  /** Build the next row from `data`; its XML is deflated to the sink when you call `.commit()`. */
  addRow(data: RowValues): StreamedRow {
    if (this.committed) throw new Error(`worksheet "${this.name}" is already committed`);
    const number = ++this.rowCount;
    const row = new Row(this.host, number) as StreamedRow;
    row.values = data;
    row.commit = () => this.commitRow(row);
    return row;
  }

  private commitRow(row: Row): void {
    if (this.committed) throw new Error(`worksheet "${this.name}" is already committed`);
    this.ensureStarted();
    const xml = renderRowXml(row.model, this.workbook.context, this.hyperlinks);
    if (xml.length > 0) this.workbook.pushSheetChunk(xml);
  }

  /** Emit the sheet prologue and open its zip entry, exactly once. */
  private ensureStarted(): void {
    if (this.started) return;
    this.started = true;
    this.workbook.beginSheet(this);
    this.workbook.startSheetEntry(`xl/worksheets/${this.rel.fileName}`);

    const w = new XmlWriter();
    w.open("worksheet", { xmlns: NS.spreadsheetml, "xmlns:r": NS.relationships });
    // Default sheet view + format; `<dimension>` is intentionally omitted (it must precede the views,
    // which we have already streamed by the time the last row is known — the reader recomputes it).
    renderSheetViews(w, [], undefined);
    w.leaf("sheetFormatPr", { defaultRowHeight: 15 });
    w.open("sheetData");
    this.workbook.pushSheetChunk(w.toString());
  }

  /** Finalise this sheet: close `<sheetData>`, flush `<hyperlinks>` + the sheet `.rels`, close the part. */
  commit(): void {
    if (this.committed) return;
    this.ensureStarted();

    // Mirror renderWorksheet's epilogue on the F0 seam: one per-sheet rId allocator backs both the
    // hyperlink `r:id`s and the sheet `.rels`, and an (empty) extraParts accumulator stands in for the
    // child parts Track E2 will contribute (drawings/comments) without reshaping this path.
    const alloc = createSheetRelAllocator();
    const extraParts: ExtraPart[] = [];

    const w = new XmlWriter({ declaration: false });
    w.close("sheetData");
    renderHyperlinks(w, this.hyperlinks, alloc);
    w.close("worksheet");
    this.workbook.pushSheetChunk(w.toString(), true);

    if (alloc.rels.length > 0) {
      this.workbook.addPart(
        `xl/worksheets/_rels/${this.rel.fileName}.rels`,
        renderWorksheetRels(alloc.rels),
      );
    }
    // Child parts (XML + media) land at their declared archive paths, once per sheet. Empty in E1.
    for (const part of extraParts) {
      if (typeof part.data === "string") this.workbook.addPart(part.path, part.data);
      else this.workbook.addBinaryPart(part.path, part.data);
    }

    this.committed = true;
    this.workbook.endSheet(this);
  }
}

export class WorkbookWriter implements AsyncIterable<Uint8Array> {
  readonly useSharedStrings: boolean;
  readonly useStyles: boolean;

  /** Backs the doc-metadata accessors below; never gains worksheets (the part renderers read only its
   *  properties, and take the sheet list explicitly). */
  private readonly meta = new Workbook();
  private readonly encoder = new TextEncoder();
  private readonly sink = new QueueSink();
  private readonly zip = new StreamingZip(this.sink);
  private readonly sharedStrings = createSharedStrings();
  private readonly styles = createStyles();
  private readonly cellStyles: Styles;
  private cfPriority = 0;
  private ctx?: WorksheetContext;

  private readonly rels: SheetRel[] = [];
  private readonly writers: WorksheetWriter[] = [];
  private active?: WorksheetWriter;
  private committed = false;
  /** Present when a `filename`/`stream` destination is being pumped in the background. */
  private readonly drain?: Promise<void>;

  constructor(options?: WorkbookWriterOptions) {
    this.useSharedStrings = options?.useSharedStrings === true;
    this.useStyles = options?.useStyles !== false;
    // When styles are disabled, cells intern nothing (always `s=0`); styles.xml still ships its baseline.
    this.cellStyles = this.useStyles
      ? this.styles
      : { addStyle: () => 0, addDxf: () => 0, toXml: this.styles.toXml.bind(this.styles) };

    if (options?.filename !== undefined) {
      this.drain = drainToFile(this.sink, options.filename);
    } else if (options?.stream !== undefined) {
      this.drain = drainToWritable(this.sink, options.stream);
    }
  }

  // --- worksheet lifecycle ----------------------------------------------------------------------

  addWorksheet(name?: string, options?: { state?: SheetState }): WorksheetWriter {
    if (this.committed) throw new Error("workbook is already committed");
    const index = this.rels.length;
    const resolvedName = name ?? `Sheet${index + 1}`;
    const lower = resolvedName.toLowerCase();
    for (const r of this.rels) {
      if (r.name.toLowerCase() === lower) throw new DuplicateWorksheetNameError(resolvedName);
    }
    const rel: SheetRel = {
      rId: `rId${index + 1}`,
      fileName: `sheet${index + 1}.xml`,
      sheetId: index + 1,
      name: resolvedName,
      state: options?.state ?? "visible",
    };
    this.rels.push(rel);
    const ws = new WorksheetWriter(this, rel);
    this.writers.push(ws);
    return ws;
  }

  /** Finalise every part after the sheets: `[Content_Types]`, rels, styles, sharedStrings (if used),
   *  workbook + props, theme — then end the zip and (for a piped destination) await the flush. */
  async commit(): Promise<void> {
    if (this.committed) throw new Error("workbook is already committed");
    // Be forgiving if a caller forgot to commit a sheet (ExcelJS requires it explicitly).
    for (const ws of this.writers) {
      if (!ws.isCommitted) ws.commit();
    }
    this.committed = true;

    const sharedStringsXml = this.useSharedStrings ? this.sharedStrings.toXml() : undefined;
    const hasSharedStrings = sharedStringsXml !== undefined;

    // Assign the remaining rIds after the sheets, matching the buffered writer's discipline.
    let counter = this.rels.length;
    const stylesRId = `rId${++counter}`;
    const sharedRId = hasSharedStrings ? `rId${++counter}` : undefined;
    const themeRId = `rId${++counter}`;

    const workbookRels: WorkbookRel[] = this.rels.map((s) => ({
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

    const put = (path: string, xml: string): void => this.addPart(path, xml);
    put(
      PART_PATH.contentTypes,
      renderContentTypes({ worksheetFiles: this.rels.map((s) => s.fileName), hasSharedStrings }),
    );
    put(PART_PATH.rootRels, renderRootRels());
    put(PART_PATH.coreProps, renderCoreProps(this.meta));
    put(
      PART_PATH.appProps,
      renderAppProps(
        this.meta,
        this.rels.map((s) => s.name),
      ),
    );
    put(PART_PATH.workbook, renderWorkbook(this.meta, this.rels));
    put(PART_PATH.workbookRels, renderWorkbookRels(workbookRels));
    put(PART_PATH.styles, this.styles.toXml());
    if (sharedStringsXml !== undefined) put(PART_PATH.sharedStrings, sharedStringsXml);
    put(PART_PATH.theme, renderTheme());

    this.zip.end();
    if (this.drain) await this.drain;
  }

  [Symbol.asyncIterator](): AsyncIterator<Uint8Array> {
    return this.sink[Symbol.asyncIterator]();
  }

  // --- internal wiring used by WorksheetWriter (same-module surface) -----------------------------

  /** The shared render context (styles/sharedStrings pools, `date1904`, cf-priority allocator). Built
   *  lazily so `properties.date1904` can be set before the first row is committed. */
  get context(): WorksheetContext {
    if (!this.ctx) {
      this.ctx = {
        styles: this.cellStyles,
        sharedStrings: this.sharedStrings,
        useSharedStrings: this.useSharedStrings,
        date1904: this.meta.properties.date1904 === true,
        nextPriority: () => ++this.cfPriority,
      };
    }
    return this.ctx;
  }

  beginSheet(ws: WorksheetWriter): void {
    if (this.active && this.active !== ws) {
      throw new Error(
        `worksheet "${this.active.name}" is still open — commit it before writing to "${ws.name}"`,
      );
    }
    this.active = ws;
  }

  endSheet(ws: WorksheetWriter): void {
    if (this.active === ws) this.active = undefined;
  }

  startSheetEntry(path: string): void {
    this.zip.startEntry(path);
  }

  pushSheetChunk(xml: string, final = false): void {
    this.zip.push(this.encoder.encode(xml), final);
  }

  addPart(path: string, xml: string): void {
    this.zip.addFile(path, this.encoder.encode(xml));
  }

  /** Add a raw-bytes part (worksheet media). The E2 seam for image `extraParts`; unused in E1. */
  addBinaryPart(path: string, data: Uint8Array): void {
    this.zip.addFile(path, data);
  }

  // --- doc metadata (delegates to the backing model; mirrors ExcelJS's WorkbookWriter) ----------

  get creator(): string | undefined {
    return this.meta.creator;
  }
  set creator(v: string | undefined) {
    this.meta.creator = v;
  }
  get lastModifiedBy(): string | undefined {
    return this.meta.lastModifiedBy;
  }
  set lastModifiedBy(v: string | undefined) {
    this.meta.lastModifiedBy = v;
  }
  get created(): Date {
    return this.meta.created;
  }
  set created(v: Date) {
    this.meta.created = v;
  }
  get modified(): Date {
    return this.meta.modified;
  }
  set modified(v: Date) {
    this.meta.modified = v;
  }
  get title(): string | undefined {
    return this.meta.title;
  }
  set title(v: string | undefined) {
    this.meta.title = v;
  }
  get subject(): string | undefined {
    return this.meta.subject;
  }
  set subject(v: string | undefined) {
    this.meta.subject = v;
  }
  get keywords(): string | undefined {
    return this.meta.keywords;
  }
  set keywords(v: string | undefined) {
    this.meta.keywords = v;
  }
  get category(): string | undefined {
    return this.meta.category;
  }
  set category(v: string | undefined) {
    this.meta.category = v;
  }
  get description(): string | undefined {
    return this.meta.description;
  }
  set description(v: string | undefined) {
    this.meta.description = v;
  }
  get company(): string | undefined {
    return this.meta.company;
  }
  set company(v: string | undefined) {
    this.meta.company = v;
  }
  get manager(): string | undefined {
    return this.meta.manager;
  }
  set manager(v: string | undefined) {
    this.meta.manager = v;
  }
  /** Workbook-wide flags (notably `date1904`); set before committing the first row. */
  get properties(): Workbook["properties"] {
    return this.meta.properties;
  }
  set properties(v: Workbook["properties"]) {
    this.meta.properties = v;
  }
  get calcProperties(): Workbook["calcProperties"] {
    return this.meta.calcProperties;
  }
  set calcProperties(v: Workbook["calcProperties"]) {
    this.meta.calcProperties = v;
  }
  get views(): Workbook["views"] {
    return this.meta.views;
  }
  set views(v: Workbook["views"]) {
    this.meta.views = v;
  }
}
