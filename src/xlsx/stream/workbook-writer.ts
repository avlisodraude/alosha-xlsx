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

import { DuplicateWorksheetNameError } from "../../errors.js";
import type {
  AddImageOptions,
  Column,
  ColumnDefn,
  ConditionalFormatting,
  DataValidations,
  ImageAnchor,
  RowValues,
  WorksheetImage,
  WorksheetView,
} from "../../model/index.js";
import { Row, Workbook, Worksheet } from "../../model/index.js";
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
import {
  createSheetRelAllocator,
  type ExtraPart,
  type HyperlinkRecord,
  renderAutoFilter,
  renderCols,
  renderConditionalFormatting,
  renderDataValidations,
  renderHyperlinks,
  renderImages,
  renderMergeCells,
  renderRowXml,
  renderSheetComments,
  renderSheetViews,
  type WorksheetContext,
} from "../parts/worksheet.js";
import { renderWorksheetRels } from "../parts/worksheet-rels.js";
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

/** Streams a single worksheet part. Emits the prologue (`<worksheet>…<sheetData>`) on the first
 *  committed row (or on {@link commit} if the sheet is empty), a `<row>` per committed row, and the
 *  epilogue (trailing elements + child parts, in CT_Worksheet order) plus the sheet `.rels` on commit.
 *
 *  A real model {@link Worksheet} backs it (`sheet`), serving three roles at once: the {@link WorksheetHost}
 *  a transient {@link Row} resolves cell styles / comments / validations through, the store for the
 *  sheet-level feature config the epilogue serialises (merges, panes, auto-filter, conditional formatting,
 *  images), and — since comments are ref-keyed on it — the collector for notes set across committed rows.
 *  Rows are never added to it (they are built with `new Row(sheet, n)` and streamed straight out), so it
 *  never retains the row body: memory stays bounded regardless of row count. */
export class WorksheetWriter {
  /** The model worksheet: row host + sheet-level feature store. Rows are streamed, never stored here. */
  private readonly sheet: Worksheet;
  private readonly hyperlinks: HyperlinkRecord[] = [];
  private rowCount = 0;
  private started = false;
  private committed = false;

  constructor(
    private readonly workbook: WorkbookWriter,
    readonly rel: SheetRel,
    sheet: Worksheet,
  ) {
    this.sheet = sheet;
  }

  get name(): string {
    return this.rel.name;
  }

  get isCommitted(): boolean {
    return this.committed;
  }

  // --- sheet-level feature config (delegates to the backing model worksheet) ---------------------
  // Set these before `commit()`; they surface in the epilogue. Panes/auto-filter are also read by the
  // prologue, so configure them before the first row commits (like ExcelJS's WorksheetWriter).

  /** Frozen/split panes + gridline visibility; read by the prologue, so set before the first row. */
  get views(): WorksheetView[] {
    return this.sheet.views;
  }
  set views(v: WorksheetView[]) {
    this.sheet.views = v;
  }

  /** `<autoFilter ref>` (e.g. `"A1:D1"`), emitted in the epilogue. */
  get autoFilter(): string | undefined {
    return this.sheet.autoFilter;
  }
  set autoFilter(ref: string | undefined) {
    this.sheet.autoFilter = ref;
  }

  /** Mutable sheet properties (default row height, gridlines, …); read by the prologue. */
  get properties(): Worksheet["properties"] {
    return this.sheet.properties;
  }

  /** Column widths/styles/keys (`<cols>`), same before-first-row rule as `views`/`properties` above —
   *  it's read by the prologue, so a column set after the first row commits has no effect on the
   *  emitted archive (silently ignored, not thrown — matching `views`). */
  get columns(): Column[] {
    return this.sheet.columns;
  }
  set columns(defns: Partial<ColumnDefn>[]) {
    this.sheet.columns = defns;
  }

  /** Get (or lazily create) a column by 1-based index or assigned key; same before-first-row caveat
   *  as {@link columns} for anything that affects `<cols>` (width/hidden/outlineLevel/style). */
  getColumn(idOrKey: number | string): Column {
    return this.sheet.getColumn(idOrKey);
  }

  /** Merge a range (same argument forms as the buffered `Worksheet.mergeCells`). */
  mergeCells(...cells: unknown[]): void {
    this.sheet.mergeCells(...(cells as Parameters<Worksheet["mergeCells"]>));
  }

  /** Add a `<conditionalFormatting>` block, flushed in the epilogue. */
  addConditionalFormatting(cf: ConditionalFormatting): ConditionalFormatting {
    return this.sheet.addConditionalFormatting(cf);
  }

  /** Ref-keyed data-validation store (also reached via `cell.dataValidation`), flushed in the epilogue. */
  get dataValidations(): DataValidations {
    return this.sheet.dataValidations;
  }

  /** Place a workbook-registered image (`imageId` from {@link WorkbookWriter.addImage}); its drawing
   *  part + media bytes are flushed at `commit()`. */
  addImage(imageId: number, anchor: ImageAnchor): WorksheetImage {
    return this.sheet.addImage(imageId, anchor);
  }

  /** Build the next row from `data`; its XML is deflated to the sink when you call `.commit()`. Set
   *  cell styles, and `cell.note` comments, on the returned row before committing it. */
  addRow(data: RowValues): StreamedRow {
    if (this.committed) throw new Error(`worksheet "${this.name}" is already committed`);
    const number = ++this.rowCount;
    // Host the transient row on the model worksheet so styles/comments/validations resolve, but never
    // register it (no `getRow`/`addRow` on the model) — the row body is not retained.
    const row = new Row(this.sheet, number) as StreamedRow;
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

  /** Add and commit every row from `rows`, awaiting {@link WorkbookWriter.flush} every `flushEvery`
   *  rows (default 10,000) so a huge sheet streams to a `filename`/`stream` destination with flat
   *  memory instead of buffering the whole compressed archive. Accepts sync or async iterables, so a
   *  DB cursor or generator can feed it directly. For per-row control (setting styles/notes before
   *  commit) use {@link addRow} and periodic {@link WorkbookWriter.flush} instead. */
  async writeRows(
    rows: Iterable<RowValues> | AsyncIterable<RowValues>,
    opts?: { flushEvery?: number },
  ): Promise<void> {
    const flushEvery = Math.max(1, opts?.flushEvery ?? 10_000);
    let n = 0;
    for await (const data of rows) {
      this.addRow(data).commit();
      if (++n % flushEvery === 0) await this.workbook.flush();
    }
  }

  /** Emit the sheet prologue and open its zip entry, exactly once. */
  private ensureStarted(): void {
    if (this.started) return;
    this.started = true;
    this.workbook.beginSheet(this);
    this.workbook.startSheetEntry(`xl/worksheets/${this.rel.fileName}`);

    const w = new XmlWriter();
    w.open("worksheet", { xmlns: NS.spreadsheetml, "xmlns:r": NS.relationships });
    // `<dimension>` is intentionally omitted (it must precede the views, but the sheet's extent is not
    // known until the last row — the reader recomputes it). Panes + gridlines come from the config.
    const props = this.sheet.properties;
    renderSheetViews(w, this.sheet.views, props.showGridLines);
    w.leaf("sheetFormatPr", {
      defaultRowHeight: props.defaultRowHeight ?? 15,
      defaultColWidth: props.defaultColWidth,
    });
    // CT_Worksheet order: `<cols>` sits after `<sheetFormatPr>`, before `<sheetData>`.
    renderCols(w, this.sheet.model.columns, this.workbook.context);
    w.open("sheetData");
    this.workbook.pushSheetChunk(w.toString());
  }

  /** Finalise this sheet: close `<sheetData>`, emit every trailing feature in CT_Worksheet order,
   *  flush the sheet `.rels` + child parts (comments/VML, drawing/media), close the part. */
  commit(): void {
    if (this.committed) return;
    this.ensureStarted();

    // Mirror renderWorksheet's epilogue on the F0 seam: one per-sheet rId allocator backs the
    // hyperlink/drawing/comment `r:id`s and the sheet `.rels`; one extraParts accumulator gathers the
    // child parts (drawings + media, comments + VML) this sheet contributes.
    const ctx = this.workbook.context;
    // 1-based sheet position (== rel.sheetId by construction) — lets child parts name themselves
    // deterministically (`drawing{n}.xml`, `comments{n}.xml`); `renderSheetComments` reads it off ctx.
    ctx.sheetNumber = this.rel.sheetId;
    const alloc = createSheetRelAllocator();
    const extraParts: ExtraPart[] = [];
    const model = this.sheet.model;

    const w = new XmlWriter({ declaration: false });
    w.close("sheetData");
    // CT_Worksheet order (matches the buffered renderWorksheet epilogue exactly).
    renderAutoFilter(w, model.autoFilter);
    renderMergeCells(w, model.merges);
    renderConditionalFormatting(w, model.conditionalFormattings, ctx);
    renderDataValidations(w, model.dataValidations);
    renderHyperlinks(w, this.hyperlinks, alloc);
    renderImages(w, this.sheet, alloc, extraParts, ctx.sheetNumber);
    renderSheetComments(w, model.comments, ctx, alloc, extraParts);
    w.close("worksheet");
    this.workbook.pushSheetChunk(w.toString(), true);

    if (alloc.rels.length > 0) {
      this.workbook.addPart(
        `xl/worksheets/_rels/${this.rel.fileName}.rels`,
        renderWorksheetRels(alloc.rels),
      );
    }
    // Child parts (XML + media) land at their declared archive paths, once per sheet; their
    // content-type Defaults/Overrides fold into `[Content_Types].xml`.
    for (const part of extraParts) this.workbook.addExtraPart(part);

    this.committed = true;
    this.workbook.endSheet(this);
  }
}

export class WorkbookWriter implements AsyncIterable<Uint8Array> {
  readonly useSharedStrings: boolean;
  readonly useStyles: boolean;

  /** Backs the doc-metadata accessors below and the image registry (`addImage`/`getImage`) each sheet
   *  resolves placements against; never gains worksheets (the part renderers read only its properties
   *  and image pool, and take the sheet list explicitly). */
  private readonly meta = new Workbook();
  private readonly encoder = new TextEncoder();
  private readonly sink = new QueueSink();
  private readonly zip = new StreamingZip(this.sink);
  private readonly sharedStrings = createSharedStrings();
  private readonly styles = createStyles();
  private readonly cellStyles: Styles;
  private cfPriority = 0;
  private ctx?: WorksheetContext;
  /** Content-type Defaults/Overrides gathered from committed sheets' child parts (drawings/media,
   *  comments/VML), folded into `[Content_Types].xml` at {@link commit}. */
  private readonly extraDefaults: { extension: string; contentType: string }[] = [];
  private readonly extraOverrides: { partName: string; contentType: string }[] = [];
  /** Archive paths already written via {@link addExtraPart} — a workbook-registry image placed on
   *  multiple sheets re-pushes the same `xl/media/imageN.ext` ExtraPart once per placement (each
   *  sheet's `renderImages` resolves it independently); this catches the repeat before it becomes a
   *  second ZIP local-file-header for one part name (an OPC violation). The buffered writer gets this
   *  for free from `collectWorksheetParts`' `Map.set` (last-write-wins on one key). */
  private readonly writtenPartPaths = new Set<string>();

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
    const state = options?.state ?? "visible";
    const rel: SheetRel = {
      rId: `rId${index + 1}`,
      fileName: `sheet${index + 1}.xml`,
      sheetId: index + 1,
      name: resolvedName,
      state,
    };
    this.rels.push(rel);
    // A real model worksheet backs the writer (row host + feature store). It is bound to `meta` so its
    // image placements resolve against the workbook image registry, but it is never added to `meta`'s
    // sheet list (the part renderers take the sheet list explicitly via `rels`).
    const sheet = new Worksheet(this.meta, rel.sheetId, resolvedName, index);
    sheet.state = state;
    const ws = new WorksheetWriter(this, rel, sheet);
    this.writers.push(ws);
    return ws;
  }

  /** Intern image bytes once and return the numeric `imageId` for {@link WorksheetWriter.addImage}
   *  (mirrors the buffered `Workbook.addImage`). */
  addImage(options: AddImageOptions): number {
    return this.meta.addImage(options);
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
      renderContentTypes({
        worksheetFiles: this.rels.map((s) => s.fileName),
        hasSharedStrings,
        extraDefaults: this.extraDefaults,
        extraOverrides: this.extraOverrides,
      }),
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

  /** Await the background pump catching up: resolves once every chunk produced so far has been handed
   *  to the `filename`/`stream` destination. A fully synchronous `addRow().commit()` loop never yields
   *  the event loop, so without this the whole compressed archive accumulates in the sink's buffer
   *  until {@link commit}. Call it periodically in a large write loop — or use
   *  {@link WorksheetWriter.writeRows} — to keep peak memory flat. A no-op when the writer is consumed
   *  as an `AsyncIterable` (the consumer drives draining) or once committed. */
  async flush(): Promise<void> {
    if (this.committed || !this.drain) return;
    await this.sink.whenDrained();
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

  /** Write one worksheet child part (drawing/comments XML or raw media) into the archive and record
   *  its content-type Default/Override so `[Content_Types].xml` can declare it. Mirrors the buffered
   *  writer's `collectWorksheetParts`. Defaults are deduped by extension inside `renderContentTypes`.
   *  A repeat `path` (same registry image placed on a second sheet) is skipped entirely — its content
   *  is identical by construction (same `imageId` -> same registry buffer), so re-zipping it would only
   *  duplicate the local-file-header, not change what gets read back. */
  addExtraPart(part: ExtraPart): void {
    if (this.writtenPartPaths.has(part.path)) return;
    this.writtenPartPaths.add(part.path);
    if (typeof part.data === "string") this.addPart(part.path, part.data);
    else this.zip.addFile(part.path, part.data);
    if (part.contentType) {
      this.extraOverrides.push({ partName: `/${part.path}`, contentType: part.contentType });
    }
    if (part.defaultFor) this.extraDefaults.push(part.defaultFor);
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
