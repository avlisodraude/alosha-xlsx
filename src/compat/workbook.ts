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
//   4. the deferred surface (§6: stream/csv/images/tables/conditional-formatting/comments/
//      defined-names/protection) throws `UnsupportedFeatureError` at the exact call, naming the
//      feature, rather than silently no-op'ing.
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
import type { XlsxAccessor } from "../index.js";
import { Cell, Column, Row, Workbook, Worksheet } from "../index.js";
import {
  CellType,
  type Comment,
  type ConditionalFormatting,
  type RichTextRun,
  type Style,
} from "../model/index.js";
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
  "addImage",
  "getImages",
  "addBackgroundImage",
  "addTable",
  "getTable",
  "getTables",
  "protect",
  "unprotect",
]);

/** Streaming entry points hung off the ExcelJS namespace (`ExcelJS.stream.xlsx.*`); each throws. */
export const streamStub = {
  xlsx: {
    WorkbookWriter: class {
      constructor() {
        throw new UnsupportedFeatureError("stream.xlsx.WorkbookWriter", COMPAT_DEFER);
      }
    },
    WorkbookReader: class {
      constructor() {
        throw new UnsupportedFeatureError("stream.xlsx.WorkbookReader", COMPAT_DEFER);
      }
    },
  },
} as const;

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

function normalizeCfRule(rule: unknown): unknown {
  if (rule === null || typeof rule !== "object") return rule;
  const r = rule as Record<string, unknown>;
  if (r.type === "colorScale" && Array.isArray(r.color)) {
    return { ...r, color: r.color.map(cfColorToArgb) };
  }
  if (r.type === "dataBar" && r.color !== undefined) {
    return { ...r, color: cfColorToArgb(r.color) };
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
        target.note = toCoreComment(value);
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
      if (prop === "addImage" || prop === "getImage") return thrower(`workbook.${prop}`);
      if (prop === "media") {
        throw new UnsupportedFeatureError("workbook.media (images)", COMPAT_DEFER);
      }
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
