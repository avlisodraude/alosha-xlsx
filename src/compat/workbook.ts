// The ExcelJS-faithful facade over the core model (docs/COMPAT-SPEC.md §5).
//
// Mechanism (§5 leaves it to the implementer, with one hard rule: it MUST NOT leak ExcelJS
// semantics into the core `@alosha/xlsx` entry): a **Proxy tree**. `CompatWorkbook`'s constructor
// returns a Proxy over a plain core `Workbook`; every model object reachable from it (worksheets,
// rows, columns, cells) is wrapped on the way out. The proxies delegate everything by default and
// override only the four documented divergences plus the deferred surface:
//
//   1. `cell.type` / `cell.effectiveType` report ExcelJS `ValueType` numbers (§3, Decision D1 = full
//      emulation) instead of our `CellType` numbers.
//   2. cell style setters (`font`/`fill`/`border`/`alignment`/`protection`/`numFmt`) route through
//      `applyExcelJsStyle` — ExcelJS whole-object-per-category precedence (§4).
//   3. `worksheet.autoFilter =` accepts ExcelJS's `{from,to}` object form (each an address string or
//      `{row,column}`) and normalizes it to our `"A1:D1"` range string (PANES-AUTOFILTER-SPEC §4).
//   4. the deferred surface (§6: stream/csv/tables/conditional-formatting/comments/
//      defined-names/protection) throws `UnsupportedFeatureError` at the exact call, naming the
//      feature, rather than silently no-op'ing.
//
// `workbook.addImage`/`getImage`/`images`, `worksheet.addImage`/`getImages` (IMAGES-DRAWINGS-SPEC.md
// §4) route straight to the core registry — ExcelJS's `{buffer}` image form and its three anchor
// shapes (`"B2:D6"` string / `{tl,br}` / `{tl,ext}`) already match the core `ImageAnchor` type, so no
// bridging is needed beyond documenting the unsupported `{base64}`/`{filename}` inputs and the
// permanently-stubbed `worksheet.addBackgroundImage` (header/footer images are out of scope, §0).
//
// Everything else — `addRow(array | object)`, `row.values` (1-based, empty `[0]`), `worksheet.columns
// = [{header, key, width}]`, the `workbook.xlsx.*` I/O accessor, `cell.dataValidation` and
// `worksheet.dataValidations` (DATA-VALIDATION-SPEC §5 — the core shape already matches ExcelJS's),
// and `worksheet.views` (PANES-AUTOFILTER-SPEC §4 — ExcelJS's frozen-view fields are already our field
// names) — is already ExcelJS-shaped in the core model (COMPAT-SPEC.md §2 "already compatible"), so
// the proxies pass it straight through. The core classes are never mutated; the divergence lives
// entirely in these compat-only wrappers.
//
// Deferred still (DATA-VALIDATION-SPEC §8): a cross-sheet list ref (the x14/`<extLst>` form) is not
// supported — only same-sheet ranges and inline lists round-trip.

import { encodeAddress } from "../address/index.js";
import { UnsupportedFeatureError } from "../errors.js";
import { Cell, Column, Row, Worksheet } from "../index.js";
import {
  CellType,
  type Comment,
  type ConditionalFormatting,
  type ImageAnchor,
  type ImageExtension,
  type RichTextRun,
  type RowValues,
  type Style,
  type WorksheetView,
} from "../model/index.js";
import { Workbook } from "../model/workbook.js";
import {
  WorkbookWriter as CoreStreamWorkbookWriter,
  type WorksheetWriter as CoreStreamWorksheetWriter,
  type NodeWritableLike,
} from "../xlsx/stream/index.js";
import type { XlsxAccessor } from "../xlsx-accessor.js";
import { applyExcelJsStyle } from "./style-compat.js";
import { toExcelJSValueType } from "./value-type.js";

/** The pointer appended to every compat `UnsupportedFeatureError` (§6). */
export const COMPAT_DEFER =
  "not yet supported in @alosha/xlsx/compat; see https://github.com/avlisodraude/alosha-xlsx/issues";

/** Sentinel prop that unwraps a facade back to its raw target (used by {@link unwrap}). */
const RAW = Symbol("alosha.compat.raw");

/** The cell style categories whose setters must apply ExcelJS whole-object precedence (§4). */
const STYLE_SETTERS = new Set(["numFmt", "font", "fill", "border", "alignment", "protection"]);

// --- deferred surface (§6) ----------------------------------------------------------------------

/** A method stub that always throws, naming `feature`. */
function thrower(feature: string): (...args: never[]) => never {
  return () => {
    throw new UnsupportedFeatureError(feature, COMPAT_DEFER);
  };
}

/** A namespace-shaped object (e.g. `workbook.csv`) whose every method throws. */
function stubNamespace(prefix: string, methods: readonly string[]): Record<string, () => never> {
  const ns: Record<string, () => never> = {};
  for (const m of methods) ns[m] = thrower(`${prefix}.${m}`);
  return ns;
}

const CSV_STUB = stubNamespace("workbook.csv", [
  "read",
  "readFile",
  "write",
  "writeFile",
  "writeBuffer",
]);
const DEFINED_NAMES_STUB = stubNamespace("workbook.definedNames", [
  "add",
  "remove",
  "getRanges",
  "forEach",
  "getNames",
]);

/** Worksheet methods that don't exist in the core model and must fail loud in compat mode. */
const WORKSHEET_UNSUPPORTED = new Set([
  "addTable",
  "getTable",
  "getTables",
  "protect",
  "unprotect",
]);

/** `stream.xlsx.WorkbookReader` has no streaming counterpart yet (see docs/STREAMING.md) — still a
 *  throwing stub. `WorkbookWriter` is bridged for real below (see {@link compatStream}). */
class StreamWorkbookReaderStub {
  constructor() {
    throw new UnsupportedFeatureError("stream.xlsx.WorkbookReader", COMPAT_DEFER);
  }
}

// --- facade plumbing ----------------------------------------------------------------------------

// One proxy per underlying object, so facade identity is stable (`ws.getCell('A1') === ws.getCell
// ('A1')`) and re-wrapping is idempotent.
const workbookProxies = new WeakMap<Workbook, Workbook>();
const worksheetProxies = new WeakMap<Worksheet, Worksheet>();
const rowProxies = new WeakMap<Row, Row>();
const columnProxies = new WeakMap<Column, Column>();
const cellProxies = new WeakMap<Cell, Cell>();

/** Wrap a value crossing the facade boundary on the way out: model objects become their proxies,
 *  arrays are wrapped element-wise (preserving sparse holes, e.g. `row.values`), primitives pass. */
function wrap(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(wrap);
  if (value instanceof Cell) return proxyCell(value);
  if (value instanceof Row) return proxyRow(value);
  if (value instanceof Column) return proxyColumn(value);
  if (value instanceof Worksheet) return proxyWorksheet(value);
  if (value instanceof Workbook) return proxyWorkbook(value);
  return value;
}

/** The inverse of {@link wrap} for a single value: a facade proxy back to its raw target. */
function unwrap(value: unknown): unknown {
  if (value !== null && typeof value === "object") {
    const raw = (value as { [RAW]?: unknown })[RAW];
    if (raw !== undefined) return raw;
  }
  return value;
}

/** Transform one argument passed *into* a model method: unwrap facade args, and wrap the arguments
 *  a user callback receives (so `eachRow`/`eachCell` hand back proxied rows/cells, not raw ones). */
function mapArgIn(arg: unknown): unknown {
  if (typeof arg === "function") {
    const cb = arg as (...cbArgs: unknown[]) => unknown;
    return (...cbArgs: unknown[]) => cb(...cbArgs.map(wrap));
  }
  return unwrap(arg);
}

/** Default `get` behaviour shared by every facade: read the raw property, bind-and-wrap methods so
 *  their return value (and callback args) cross the boundary, else wrap the plain value. */
function reflectGet(target: object, prop: string | symbol): unknown {
  const value = Reflect.get(target, prop, target);
  if (typeof value === "function") {
    const fn = value as (...args: unknown[]) => unknown;
    return (...args: unknown[]) => wrap(fn.apply(target, args.map(mapArgIn)));
  }
  return wrap(value);
}

/** ExcelJS parity for `eachRow`/`eachCell`: without `includeEmpty`, ExcelJS drops rows/cells that are
 *  empty (never held a value, or were explicitly set to `null`), while our core model still yields any
 *  row/cell ever touched via `getRow`/`getCell`. This bridges that gap on the compat iteration path —
 *  `isEmpty` tests the raw model item, and surviving items are wrapped like every other facade boundary
 *  (so callbacks receive proxied rows/cells, not raw ones). `includeEmpty: true` passes straight through
 *  to the core's full enumeration. */
function eachWithEmptyFilter<T>(
  each: (opts: { includeEmpty: boolean }, cb: (item: T, n: number) => void) => void,
  isEmpty: (item: T) => boolean,
  optsOrCb: { includeEmpty?: boolean } | ((item: unknown, n: number) => void),
  maybeCb?: (item: unknown, n: number) => void,
): void {
  const includeEmpty = typeof optsOrCb === "object" ? optsOrCb.includeEmpty === true : false;
  const cb = (typeof optsOrCb === "function" ? optsOrCb : maybeCb) as (
    item: unknown,
    n: number,
  ) => void;
  each({ includeEmpty }, (item, n) => {
    if (!includeEmpty && isEmpty(item)) return;
    cb(wrap(item), n);
  });
}

/** ExcelJS accepts an `{from,to}` address (each an `"A1"` string or `{row,column}`) for `autoFilter`
 *  alongside a plain `"A1:D1"` string (docs/PANES-AUTOFILTER-SPEC.md §4); the core model only ever
 *  stores the string range, and the writer emits `ref` verbatim, so the object form must be
 *  normalized here before it reaches the core. */
function normalizeAutoFilter(value: unknown): unknown {
  if (value !== null && typeof value === "object" && "from" in value && "to" in value) {
    const { from, to } = value as {
      from: string | { row: number; column: number };
      to: string | { row: number; column: number };
    };
    const address = (a: string | { row: number; column: number }): string =>
      typeof a === "string" ? a : encodeAddress(a.row, a.column);
    return `${address(from)}:${address(to)}`;
  }
  return value;
}

/** ExcelJS represents a `colorScale`/`dataBar` colour as `{ argb }` (mirroring its cell-style `Color`
 *  shape); our core model stores `CfColor` as a bare ARGB string (conditional-format colours are
 *  argb-only — no theme/indexed support). Bridge at the `addConditionalFormatting` boundary so an
 *  ExcelJS-idiomatic call passes straight through; `argbToCfColor` is the inverse, applied when a
 *  compat caller reads `worksheet.conditionalFormattings` back without a write/read round trip (§4
 *  colour-shape bridge). `iconSet` never carries a colour (icons use Excel's built-in presets), so
 *  only `colorScale`/`dataBar` need touching. */
function cfColorToArgb(color: unknown): unknown {
  if (color !== null && typeof color === "object" && "argb" in color) {
    return (color as { argb?: unknown }).argb;
  }
  return color;
}

function argbToCfColor(argb: unknown): unknown {
  return typeof argb === "string" ? { argb } : argb;
}

/** ExcelJS represents the text-contains family as a single `type: "containsText"` with the real
 *  variant carried on `operator`; the core model uses a literal `type` per variant instead (§4
 *  text-family bridge, closing MIGRATION.md D3's first divergence). Maps ExcelJS's `operator` value
 *  to the core's `type`. */
const CF_TEXT_OPERATOR_TO_CORE_TYPE: Record<string, string> = {
  containsText: "containsText",
  notContains: "notContainsText",
  beginsWith: "beginsWith",
  endsWith: "endsWith",
};

/** The inverse of {@link CF_TEXT_OPERATOR_TO_CORE_TYPE}, applied when denormalizing a core rule back
 *  to ExcelJS's `{type: "containsText", operator}` shape. */
const CF_TEXT_CORE_TYPE_TO_OPERATOR: Record<string, string> = {
  containsText: "containsText",
  notContainsText: "notContains",
  beginsWith: "beginsWith",
  endsWith: "endsWith",
};

function normalizeCfRule(rule: unknown): unknown {
  if (rule === null || typeof rule !== "object") return rule;
  const r = rule as Record<string, unknown>;
  if (r.type === "colorScale" && Array.isArray(r.color)) {
    return { ...r, color: r.color.map(cfColorToArgb) };
  }
  if (r.type === "dataBar" && r.color !== undefined) {
    return { ...r, color: cfColorToArgb(r.color) };
  }
  if (r.type === "containsText" && typeof r.operator === "string") {
    const coreType = CF_TEXT_OPERATOR_TO_CORE_TYPE[r.operator];
    if (coreType !== undefined) {
      const rest = { ...r };
      delete rest.operator;
      return { ...rest, type: coreType };
    }
  }
  return rule;
}

function denormalizeCfRule(rule: unknown): unknown {
  if (rule === null || typeof rule !== "object") return rule;
  const r = rule as Record<string, unknown>;
  if (r.type === "colorScale" && Array.isArray(r.color)) {
    return { ...r, color: r.color.map(argbToCfColor) };
  }
  if (r.type === "dataBar" && r.color !== undefined) {
    return { ...r, color: argbToCfColor(r.color) };
  }
  if (typeof r.type === "string" && r.type in CF_TEXT_CORE_TYPE_TO_OPERATOR) {
    return { ...r, type: "containsText", operator: CF_TEXT_CORE_TYPE_TO_OPERATOR[r.type] };
  }
  return rule;
}

/** Normalize one `{ ref, rules }` block passed into `addConditionalFormatting`; non-CF-shaped input
 *  passes through untouched so a malformed call still reaches the core (and its own validation). */
function normalizeConditionalFormatting(cf: unknown): unknown {
  if (cf === null || typeof cf !== "object") return cf;
  const c = cf as { rules?: unknown };
  if (!Array.isArray(c.rules)) return cf;
  return { ...c, rules: c.rules.map(normalizeCfRule) };
}

/** The inverse, applied to every block a compat caller reads off `worksheet.conditionalFormattings` —
 *  clones rather than mutates, so the core's stored (normalized) blocks stay writer-ready. */
function denormalizeConditionalFormattings(cfs: readonly ConditionalFormatting[]): unknown[] {
  return cfs.map((cf) => {
    const c = cf as unknown as { rules?: unknown };
    if (!Array.isArray(c.rules)) return cf;
    return { ...c, rules: c.rules.map(denormalizeCfRule) };
  });
}

/** ExcelJS's `cell.note` carries rich runs as a flat `texts: [{font, text}]` array (docs/COMMENTS-
 *  SPEC.md §3); our core `Comment` wraps them in a `RichText` (`{ richText: [...] }`) to match the
 *  richText cell-value shape everywhere else. Bridge at the `cell.note` setter boundary: a bare
 *  string passes straight through, `{ texts: string, ... }` passes through with its runs re-wrapped,
 *  and anything else (already our own `Comment` shape, e.g. `{ texts: { richText: [...] } }` from
 *  native-API interop) passes through untouched. */
function toCoreComment(value: unknown): Comment {
  if (typeof value === "string") return { texts: value };
  const v = value as { texts: unknown; author?: string };
  const author = v.author;
  if (typeof v.texts === "string") {
    return author === undefined ? { texts: v.texts } : { texts: v.texts, author };
  }
  if (Array.isArray(v.texts)) {
    const richText = { richText: v.texts as RichTextRun[] };
    return author === undefined ? { texts: richText } : { texts: richText, author };
  }
  return value as Comment;
}

/** The inverse of {@link toCoreComment}, applied when a compat caller reads `cell.note` back: a
 *  plain-string comment stays a string, and a rich one unwraps its `RichText` back into ExcelJS's
 *  flat `texts` run array. */
function toExcelJsNote(comment: Comment | undefined): unknown {
  if (comment === undefined) return undefined;
  if (typeof comment.texts === "string") {
    return comment.author === undefined
      ? comment.texts
      : { texts: comment.texts, author: comment.author };
  }
  const texts = comment.texts.richText;
  return comment.author === undefined ? { texts } : { texts, author: comment.author };
}

// --- images (IMAGES-DRAWINGS-SPEC.md §4) ---------------------------------------------------------

/** ExcelJS's `Image` input to `workbook.addImage`: `{extension, buffer}` / `{base64}` / `{filename}`.
 *  The core registry only interns raw bytes, so only the `buffer` form is bridged — `base64` and
 *  `filename` are documented-unsupported inputs (fail loud rather than silently drop the image). */
function toAddImageOptions(image: unknown): { buffer: Uint8Array; extension: ImageExtension } {
  const img = image as {
    extension: ImageExtension;
    buffer?: Uint8Array;
    base64?: string;
    filename?: string;
  };
  if (img.buffer !== undefined) return { buffer: img.buffer, extension: img.extension };
  if (img.base64 !== undefined) {
    throw new UnsupportedFeatureError("workbook.addImage({ base64 })", COMPAT_DEFER);
  }
  if (img.filename !== undefined) {
    throw new UnsupportedFeatureError("workbook.addImage({ filename })", COMPAT_DEFER);
  }
  throw new UnsupportedFeatureError("workbook.addImage (missing buffer)", COMPAT_DEFER);
}

/** ExcelJS's `workbook.addImage`/`getImage` return the interned image with a `type: "image"` tag
 *  appended (`lib/doc/workbook.js`); mirror that shape so compat callers see the same fields. */
function toExcelJsMedia(image: { buffer: Uint8Array; extension: ImageExtension }): unknown {
  return { type: "image", ...image };
}

// --- streaming writer (docs/STREAMING.md; ExcelJS's `stream.xlsx.WorkbookWriter`) ----------------
//
// Bridges onto our own `xlsx/stream/workbook-writer.ts` `WorkbookWriter`, which is already
// ExcelJS-shaped: `addWorksheet()` → `ws.addRow(...).commit()` → `ws.commit()` → `await
// workbook.commit()`. Rows/cells crossing the boundary are real model `Row`/`Cell` instances, so
// reusing {@link wrap} gives them `cell.type` ExcelJS numbers, whole-object style setters, and
// `cell.note` bridging for free — the same as the buffered facade, with no extra plumbing. Only the
// constructor/`addWorksheet` *options* need bridging (or a loud `UnsupportedFeatureError` where our
// streaming writer has no equivalent), since those are plain data ExcelJS accepts that our own
// `WorkbookWriterOptions`/`addWorksheet` signature doesn't structurally share.

/** Real-ExcelJS `stream.xlsx.WorkbookWriter` constructor options with no equivalent on our streaming
 *  writer: `zip` (archiver-specific tuning) and `lastPrinted` (unlike the buffered `Workbook`, the
 *  streaming writer's doc-metadata surface doesn't track it). */
const STREAM_WORKBOOK_UNSUPPORTED_OPTIONS = ["zip", "lastPrinted"] as const;

/** Real-ExcelJS `addWorksheet(name, options)` keys with no equivalent on our streaming worksheet
 *  writer: `pageSetup`/`headerFooter` (printing config; unimplemented for streaming) and a per-sheet
 *  `useSharedStrings` override (only a workbook-wide setting exists). */
const STREAM_WORKSHEET_UNSUPPORTED_OPTIONS = [
  "pageSetup",
  "headerFooter",
  "useSharedStrings",
] as const;

class CompatWorksheetWriter {
  #raw: CoreStreamWorksheetWriter;
  constructor(raw: CoreStreamWorksheetWriter) {
    this.#raw = raw;
  }

  get name(): string {
    return this.#raw.name;
  }
  /** ExcelJS's field is `committed`, not our `isCommitted`. */
  get committed(): boolean {
    return this.#raw.isCommitted;
  }
  get views(): WorksheetView[] {
    return this.#raw.views;
  }
  set views(v: WorksheetView[]) {
    this.#raw.views = v;
  }
  get autoFilter(): string | undefined {
    return this.#raw.autoFilter;
  }
  set autoFilter(v: unknown) {
    this.#raw.autoFilter = normalizeAutoFilter(v) as string | undefined;
  }
  get properties(): CoreStreamWorksheetWriter["properties"] {
    return this.#raw.properties;
  }
  get dataValidations(): CoreStreamWorksheetWriter["dataValidations"] {
    return this.#raw.dataValidations;
  }

  mergeCells(...cells: unknown[]): void {
    this.#raw.mergeCells(...cells);
  }

  /** §4 colour-shape + text-family bridge, same as the buffered worksheet's `addConditionalFormatting`. */
  addConditionalFormatting(cf: unknown): unknown {
    const added = this.#raw.addConditionalFormatting(
      normalizeConditionalFormatting(cf) as ConditionalFormatting,
    );
    return wrap({ ...added, rules: added.rules.map(denormalizeCfRule) });
  }

  addImage(imageId: number, anchor: ImageAnchor): void {
    this.#raw.addImage(imageId, anchor);
  }

  /** The returned row is a real model `Row` (plus `commit()`); `wrap` proxies it exactly like the
   *  buffered facade's rows (ExcelJS `cell.type` numbers, whole-object style setters, `cell.note`). */
  addRow(data: RowValues): Row & { commit(): void; commitAsync(): Promise<void> } {
    return wrap(this.#raw.addRow(data)) as Row & { commit(): void; commitAsync(): Promise<void> };
  }

  async addRowAsync(data: RowValues): Promise<void> {
    await this.#raw.addRowAsync(data);
  }
  async writeRows(
    rows: Iterable<RowValues> | AsyncIterable<RowValues>,
    options?: { flushEvery?: number },
  ): Promise<void> {
    await this.#raw.writeRows(rows, options);
  }

  commit(): void {
    this.#raw.commit();
  }
}

class CompatWorkbookWriter implements AsyncIterable<Uint8Array> {
  #raw: CoreStreamWorkbookWriter;

  get bufferedBytes(): number {
    return this.#raw.bufferedBytes;
  }
  abort(reason?: unknown): void {
    this.#raw.abort(reason);
  }
  async flush(): Promise<void> {
    await this.#raw.flush();
  }
  async ready(): Promise<void> {
    await this.#raw.ready();
  }

  constructor(options?: unknown) {
    const opts = (options ?? {}) as Record<string, unknown>;
    for (const key of STREAM_WORKBOOK_UNSUPPORTED_OPTIONS) {
      if (key in opts) {
        throw new UnsupportedFeatureError(
          `stream.xlsx.WorkbookWriter option "${key}"`,
          COMPAT_DEFER,
        );
      }
    }
    this.#raw = new CoreStreamWorkbookWriter({
      filename: opts.filename as string | undefined,
      stream: opts.stream as NodeWritableLike | undefined,
      useSharedStrings: opts.useSharedStrings as boolean | undefined,
      useStyles: opts.useStyles as boolean | undefined,
      signal: opts.signal as AbortSignal | undefined,
      highWaterMarkBytes: opts.highWaterMarkBytes as number | undefined,
      maxBufferedBytes: opts.maxBufferedBytes as number | undefined,
    });
    if (opts.creator !== undefined) this.#raw.creator = opts.creator as string;
    if (opts.lastModifiedBy !== undefined) this.#raw.lastModifiedBy = opts.lastModifiedBy as string;
    if (opts.created !== undefined) this.#raw.created = opts.created as Date;
    if (opts.modified !== undefined) this.#raw.modified = opts.modified as Date;
  }

  addWorksheet(name?: string, options?: unknown): CompatWorksheetWriter {
    const opts = (options ?? {}) as Record<string, unknown>;
    for (const key of STREAM_WORKSHEET_UNSUPPORTED_OPTIONS) {
      if (key in opts) {
        throw new UnsupportedFeatureError(`worksheet option "${key}" (streaming)`, COMPAT_DEFER);
      }
    }
    const state = opts.state as "visible" | "hidden" | "veryHidden" | undefined;
    const ws = this.#raw.addWorksheet(name, state === undefined ? undefined : { state });

    // `properties` (and the deprecated top-level `tabColor` shorthand ExcelJS itself folds into it)
    // and `views`/`autoFilter` all have a post-construction equivalent on our streaming worksheet, so
    // apply them here instead of at core-`addWorksheet` time.
    const properties = { ...(opts.properties as Record<string, unknown> | undefined) };
    if (opts.tabColor !== undefined) properties.tabColor = opts.tabColor;
    if (Object.keys(properties).length > 0) Object.assign(ws.properties, properties);
    if (opts.views !== undefined) ws.views = opts.views as WorksheetView[];
    if (opts.autoFilter !== undefined) {
      ws.autoFilter = normalizeAutoFilter(opts.autoFilter) as string | undefined;
    }
    return new CompatWorksheetWriter(ws);
  }

  addImage(image: unknown): number {
    return this.#raw.addImage(toAddImageOptions(image));
  }

  async commit(): Promise<void> {
    await this.#raw.commit();
  }

  [Symbol.asyncIterator](): AsyncIterator<Uint8Array> {
    return this.#raw[Symbol.asyncIterator]();
  }

  get creator(): string | undefined {
    return this.#raw.creator;
  }
  set creator(v: string | undefined) {
    this.#raw.creator = v;
  }
  get lastModifiedBy(): string | undefined {
    return this.#raw.lastModifiedBy;
  }
  set lastModifiedBy(v: string | undefined) {
    this.#raw.lastModifiedBy = v;
  }
  get created(): Date {
    return this.#raw.created;
  }
  set created(v: Date) {
    this.#raw.created = v;
  }
  get modified(): Date {
    return this.#raw.modified;
  }
  set modified(v: Date) {
    this.#raw.modified = v;
  }
  get title(): string | undefined {
    return this.#raw.title;
  }
  set title(v: string | undefined) {
    this.#raw.title = v;
  }
  get subject(): string | undefined {
    return this.#raw.subject;
  }
  set subject(v: string | undefined) {
    this.#raw.subject = v;
  }
  get keywords(): string | undefined {
    return this.#raw.keywords;
  }
  set keywords(v: string | undefined) {
    this.#raw.keywords = v;
  }
  get category(): string | undefined {
    return this.#raw.category;
  }
  set category(v: string | undefined) {
    this.#raw.category = v;
  }
  get description(): string | undefined {
    return this.#raw.description;
  }
  set description(v: string | undefined) {
    this.#raw.description = v;
  }
  get company(): string | undefined {
    return this.#raw.company;
  }
  set company(v: string | undefined) {
    this.#raw.company = v;
  }
  get manager(): string | undefined {
    return this.#raw.manager;
  }
  set manager(v: string | undefined) {
    this.#raw.manager = v;
  }
  get properties(): CoreStreamWorkbookWriter["properties"] {
    return this.#raw.properties;
  }
  set properties(v: CoreStreamWorkbookWriter["properties"]) {
    this.#raw.properties = v;
  }
  get calcProperties(): CoreStreamWorkbookWriter["calcProperties"] {
    return this.#raw.calcProperties;
  }
  set calcProperties(v: CoreStreamWorkbookWriter["calcProperties"]) {
    this.#raw.calcProperties = v;
  }
  get views(): CoreStreamWorkbookWriter["views"] {
    return this.#raw.views;
  }
  set views(v: CoreStreamWorkbookWriter["views"]) {
    this.#raw.views = v;
  }
}

/** The ExcelJS namespace's `stream.xlsx.*` entry points. `WorkbookWriter` is a real bridge (above);
 *  `WorkbookReader` stays a throwing stub (no streaming reader yet). */
export const compatStream = {
  xlsx: {
    WorkbookWriter: CompatWorkbookWriter,
    WorkbookReader: StreamWorkbookReaderStub,
  },
} as const;

// --- the per-level proxies ----------------------------------------------------------------------

function proxyCell(cell: Cell): Cell {
  const cached = cellProxies.get(cell);
  if (cached) return cached;
  const proxy = new Proxy(cell, {
    get(target, prop) {
      if (prop === RAW) return target;
      // Decision D1: compat cells always report ExcelJS ValueType numbers (§3).
      if (prop === "type") return toExcelJSValueType(target.type);
      if (prop === "effectiveType") return toExcelJSValueType(target.effectiveType);
      // COMMENTS-SPEC.md §3: bridge our RichText-wrapped Comment to ExcelJS's flat texts array.
      if (prop === "note") return toExcelJsNote(target.note);
      return reflectGet(target, prop);
    },
    set(target, prop, value) {
      // §4: assigning a style category replaces it whole-object, never field-merges.
      if (typeof prop === "string" && STYLE_SETTERS.has(prop)) {
        const patch: Style = {};
        (patch as Record<string, unknown>)[prop] = value;
        applyExcelJsStyle(target, patch);
        return true;
      }
      if (prop === "note") {
        // `undefined` clears the note (core `Comments.remove`); anything else is normalized.
        target.note = value === undefined ? undefined : toCoreComment(value);
        return true;
      }
      return Reflect.set(target, prop, value, target);
    },
  });
  cellProxies.set(cell, proxy);
  return proxy;
}

function proxyRow(row: Row): Row {
  const cached = rowProxies.get(row);
  if (cached) return cached;
  const proxy = new Proxy(row, {
    get(target, prop) {
      if (prop === RAW) return target;
      // ExcelJS parity: default `eachCell` skips empty cells (§ MIGRATION "eachRow/eachCell").
      if (prop === "eachCell") {
        return (
          optsOrCb: { includeEmpty?: boolean } | ((cell: unknown, n: number) => void),
          maybeCb?: (cell: unknown, n: number) => void,
        ) =>
          eachWithEmptyFilter<Cell>(
            (o, cb) => target.eachCell(o, cb),
            (cell) => cell.type === CellType.Null,
            optsOrCb,
            maybeCb,
          );
      }
      return reflectGet(target, prop);
    },
    set(target, prop, value) {
      return Reflect.set(target, prop, value, target);
    },
  });
  rowProxies.set(row, proxy);
  return proxy;
}

function proxyColumn(column: Column): Column {
  const cached = columnProxies.get(column);
  if (cached) return cached;
  const proxy = new Proxy(column, {
    get(target, prop) {
      if (prop === RAW) return target;
      // ExcelJS parity: default `eachCell` skips empty cells (§ MIGRATION "eachRow/eachCell").
      if (prop === "eachCell") {
        return (
          optsOrCb: { includeEmpty?: boolean } | ((cell: unknown, n: number) => void),
          maybeCb?: (cell: unknown, n: number) => void,
        ) =>
          eachWithEmptyFilter<Cell>(
            (o, cb) => target.eachCell(o, cb),
            (cell) => cell.type === CellType.Null,
            optsOrCb,
            maybeCb,
          );
      }
      return reflectGet(target, prop);
    },
    set(target, prop, value) {
      return Reflect.set(target, prop, value, target);
    },
  });
  columnProxies.set(column, proxy);
  return proxy;
}

function proxyWorksheet(worksheet: Worksheet): Worksheet {
  const cached = worksheetProxies.get(worksheet);
  if (cached) return cached;
  const proxy = new Proxy(worksheet, {
    get(target, prop) {
      if (prop === RAW) return target;
      if (typeof prop === "string" && WORKSHEET_UNSUPPORTED.has(prop)) {
        return thrower(`worksheet.${prop}`);
      }
      // `addImage(imageId, anchor)` needs no bridging — ExcelJS's `"B2:D6"` string / `{tl,br}` /
      // `{tl,ext}` anchor forms already match the core `ImageAnchor` type, so it falls through to the
      // default `reflectGet` below. `getImages()` has no core equivalent method (only the `images`
      // getter), so it's reshaped here into ExcelJS's `{type,imageId,range}[]` array.
      if (prop === "getImages") {
        return () =>
          target.images.map((img) => ({ type: "image", imageId: img.imageId, range: img.range }));
      }
      // Header/footer images are out of scope (IMAGES-DRAWINGS-SPEC.md §0) — a permanent, documented
      // stub rather than a "not yet" deferral.
      if (prop === "addBackgroundImage") return thrower("worksheet.addBackgroundImage");
      // §4 colour-shape bridge: ExcelJS passes colorScale/dataBar colours as `{argb}` objects; the
      // core model stores a bare ARGB string. Normalize on the way in, denormalize on the way out
      // (so a read with no write/read round trip still reports the ExcelJS shape).
      if (prop === "addConditionalFormatting") {
        return (cf: unknown) => {
          const added = target.addConditionalFormatting(
            normalizeConditionalFormatting(cf) as ConditionalFormatting,
          );
          return wrap({ ...added, rules: added.rules.map(denormalizeCfRule) });
        };
      }
      if (prop === "conditionalFormattings") {
        return denormalizeConditionalFormattings(target.conditionalFormattings);
      }
      // §4 removeConditionalFormatting bridge: ExcelJS accepts an index (splice), a predicate run
      // over the denormalized (ExcelJS-shaped) blocks, or no argument (clear all); the core only
      // exposes by-ref removal, so resolve the index/predicate against a snapshot of
      // `conditionalFormattings` first, then drive the core method with each matched block's `ref`
      // (closing MIGRATION.md D3's second divergence).
      if (prop === "removeConditionalFormatting") {
        return (filter?: unknown) => {
          const cfs = target.conditionalFormattings;
          if (typeof filter === "number") {
            const cf = cfs.at(filter);
            if (cf) target.removeConditionalFormatting(cf.ref);
            return;
          }
          if (typeof filter === "function") {
            const predicate = filter as (cf: unknown, index: number, array: unknown[]) => unknown;
            const denorm = denormalizeConditionalFormattings(cfs);
            cfs.forEach((cf, i) => {
              if (!predicate(denorm[i], i, denorm)) target.removeConditionalFormatting(cf.ref);
            });
            return;
          }
          for (const cf of cfs) target.removeConditionalFormatting(cf.ref);
        };
      }
      // D7 spliceColumns bridge: ExcelJS's insert arrays are 0-based (`insert[i]` fills row `i+1`),
      // while the core follows the 1-based `Column.values` convention (`colVals[n]` fills row `n`,
      // index 0 unused). Shift each insert array up one slot so ported ExcelJS code lands on the right
      // rows. `start`/`count` are plain numbers; cell values are unwrapped like any other model arg.
      if (prop === "spliceColumns") {
        return (start: number, count: number, ...inserts: unknown[]) => {
          const bridged = inserts.map((col) =>
            Array.isArray(col) ? [null, ...col.map(unwrap)] : col,
          );
          return wrap(
            (target.spliceColumns as (...a: unknown[]) => unknown)(start, count, ...bridged),
          );
        };
      }
      // ExcelJS parity: default `eachRow` skips empty rows (§ MIGRATION "eachRow/eachCell").
      if (prop === "eachRow") {
        return (
          optsOrCb: { includeEmpty?: boolean } | ((row: unknown, n: number) => void),
          maybeCb?: (row: unknown, n: number) => void,
        ) =>
          eachWithEmptyFilter<Row>(
            (o, cb) => target.eachRow(o, cb),
            (row) => !row.hasValues,
            optsOrCb,
            maybeCb,
          );
      }
      return reflectGet(target, prop);
    },
    set(target, prop, value) {
      // §4: ExcelJS accepts an `{from,to}` object for `autoFilter`; the core only stores the range
      // string, and `renderAutoFilter` writes `ref` verbatim, so normalize it here.
      if (prop === "autoFilter") {
        return Reflect.set(target, prop, normalizeAutoFilter(value), target);
      }
      return Reflect.set(target, prop, value, target);
    },
  });
  worksheetProxies.set(worksheet, proxy);
  return proxy;
}

function proxyWorkbook(workbook: Workbook): Workbook {
  const cached = workbookProxies.get(workbook);
  if (cached) return cached;
  const proxy = new Proxy(workbook, {
    get(target, prop) {
      if (prop === RAW) return target;
      // The I/O accessor is already ExcelJS-shaped; only rebind load/read/readFile so they resolve
      // to the *facade* (not the raw workbook), keeping the returned handle inside compat mode.
      if (prop === "xlsx") {
        const acc = Reflect.get(target, "xlsx", target) as XlsxAccessor;
        return {
          writeBuffer: (...args: Parameters<XlsxAccessor["writeBuffer"]>) =>
            acc.writeBuffer(...args),
          writeFile: (...args: Parameters<XlsxAccessor["writeFile"]>) => acc.writeFile(...args),
          load: async (...args: Parameters<XlsxAccessor["load"]>) => {
            await acc.load(...args);
            return proxy;
          },
          readFile: async (...args: Parameters<XlsxAccessor["readFile"]>) => {
            await acc.readFile(...args);
            return proxy;
          },
          read: async (...args: Parameters<XlsxAccessor["read"]>) => {
            await acc.read(...args);
            return proxy;
          },
        } satisfies XlsxAccessor;
      }
      if (prop === "csv") return CSV_STUB;
      if (prop === "definedNames") return DEFINED_NAMES_STUB;
      if (prop === "addImage") {
        return (image: unknown) => target.addImage(toAddImageOptions(image));
      }
      if (prop === "getImage") {
        return (id: number) => {
          const image = target.getImage(id);
          return image === undefined ? undefined : toExcelJsMedia(image);
        };
      }
      if (prop === "media") return target.images.map(toExcelJsMedia);
      return reflectGet(target, prop);
    },
    set(target, prop, value) {
      return Reflect.set(target, prop, value, target);
    },
  });
  workbookProxies.set(workbook, proxy);
  return proxy;
}

/**
 * The compat `Workbook`: an ExcelJS drop-in. `new CompatWorkbook()` builds a plain core workbook and
 * returns a facade Proxy over it — so `instanceof` still holds (the proxy's target *is* a Workbook),
 * while every reachable cell/row/column/worksheet reports ExcelJS semantics (§5).
 */
export class CompatWorkbook extends Workbook {
  constructor() {
    super();
    // A constructor may return an object in place of `this`; we deliberately return the facade
    // wrapping the fully-initialised core instance — the whole mechanism of the shim (§5). Cast:
    // the proxy's target is this CompatWorkbook, so `instanceof` still holds.
    // biome-ignore lint/correctness/noConstructorReturn: returning the facade Proxy is intentional.
    return proxyWorkbook(this) as CompatWorkbook;
  }
}
