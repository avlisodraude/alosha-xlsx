// The ExcelJS-faithful facade over the core model (docs/COMPAT-SPEC.md §5).
//
// Mechanism (§5 leaves it to the implementer, with one hard rule: it MUST NOT leak ExcelJS
// semantics into the core `@alosha/xlsx` entry): a **Proxy tree**. `CompatWorkbook`'s constructor
// returns a Proxy over a plain core `Workbook`; every model object reachable from it (worksheets,
// rows, columns, cells) is wrapped on the way out. The proxies delegate everything by default and
// override only the three documented divergences plus the deferred surface:
//
//   1. `cell.type` / `cell.effectiveType` report ExcelJS `ValueType` numbers (§3, Decision D1 = full
//      emulation) instead of our `CellType` numbers.
//   2. cell style setters (`font`/`fill`/`border`/`alignment`/`protection`/`numFmt`) route through
//      `applyExcelJsStyle` — ExcelJS whole-object-per-category precedence (§4).
//   3. the deferred surface (§6: stream/csv/images/tables/conditional-formatting/data-validation/
//      comments/defined-names/protection) throws `UnsupportedFeatureError` at the exact call, naming
//      the feature, rather than silently no-op'ing.
//
// Everything else — `addRow(array | object)`, `row.values` (1-based, empty `[0]`), `worksheet.columns
// = [{header, key, width}]`, the `workbook.xlsx.*` I/O accessor — is already ExcelJS-shaped in the
// core model (COMPAT-SPEC.md §2 "already compatible"), so the proxies pass it straight through. The
// core classes are never mutated; the divergence lives entirely in these compat-only wrappers.

import { UnsupportedFeatureError } from "../errors.js";
import type { XlsxAccessor } from "../index.js";
import { Cell, Column, Row, Workbook, Worksheet } from "../index.js";
import type { Style } from "../model/index.js";
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
  "addConditionalFormatting",
  "removeConditionalFormatting",
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
        throw new UnsupportedFeatureError("cell.note (comments)", COMPAT_DEFER);
      }
      if (prop === "dataValidation") {
        throw new UnsupportedFeatureError("cell.dataValidation (data validation)", COMPAT_DEFER);
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
      if (prop === "dataValidations") {
        throw new UnsupportedFeatureError(
          "worksheet.dataValidations (data validation)",
          COMPAT_DEFER,
        );
      }
      return reflectGet(target, prop);
    },
    set(target, prop, value) {
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
