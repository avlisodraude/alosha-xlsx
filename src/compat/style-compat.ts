// The style-precedence divergence the compat shim must emulate (docs/COMPAT-SPEC.md §4).
//
// The core resolver (`resolveCellStyle`) does a **deep field-level merge**: a lone `cell.font.bold`
// coexists with an inherited `font.name`/`font.size`. ExcelJS instead replaces the **whole**
// category — assigning `cell.font` discards every inherited font field. A migrated spreadsheet must
// render identically, so the compat facade's style setters route through `applyExcelJsStyle`, which
// applies ExcelJS whole-object-per-category semantics.
//
// This only affects the authoring path: on read, styles come back already fully-resolved per cell,
// so there is no divergence to bridge there. The core `resolveCellStyle` is deliberately left
// untouched — the divergence is emulated here, scoped to compat-constructed workbooks.

import type { Cell, Style } from "../model/index.js";

/** The six independently-inheritable style categories, in the order the core resolver lists them. */
const STYLE_CATEGORIES = ["numFmt", "font", "fill", "border", "alignment", "protection"] as const;

/**
 * Apply a style `patch` to `cell` with ExcelJS whole-object precedence: every category **present**
 * in the patch replaces the cell's inherited category **outright** (no field-level merge), while
 * categories absent from the patch keep their current (possibly inherited) value. A category set to
 * `undefined` in the patch clears it. The underlying model store is otherwise unchanged — the
 * serializer still dedups the final styles.
 */
export function applyExcelJsStyle(cell: Cell, patch: Style): void {
  const inherited = cell.style;
  const next: Style = {};
  for (const key of STYLE_CATEGORIES) {
    // A category the patch mentions comes from the patch (whole-object replacement); otherwise it is
    // inherited from the cell's current style. Either way we copy the value as-is, never merging its
    // fields into the inherited one — that field-merge is exactly the core behaviour we diverge from.
    copyCategory(next, Object.hasOwn(patch, key) ? patch : inherited, key);
  }
  cell.style = next;
}

/** Copy one style category from `source` to `target`, skipping it entirely when unset. */
function copyCategory<K extends keyof Style>(target: Style, source: Style, key: K): void {
  const value = source[key];
  if (value !== undefined) target[key] = value;
}
