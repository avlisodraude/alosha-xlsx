// Style & format types (see docs/STAGE-1.2-SPEC.md §4 / docs/API-AUDIT.md §4). All fields are
// optional — a cell may set only e.g. `font.bold`, leaving the rest to inherit from row/column.

export interface Color {
  argb?: string;
  theme?: number;
  indexed?: number;
}

export interface Font {
  name?: string;
  size?: number;
  family?: number;
  scheme?: "minor" | "major" | "none";
  charset?: number;
  color?: Color;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | "none" | "single" | "double" | "singleAccounting" | "doubleAccounting";
  vertAlign?: "superscript" | "subscript";
  strike?: boolean;
  outline?: boolean;
}

export type FillPattern =
  | "none"
  | "solid"
  | "darkVertical"
  | "darkHorizontal"
  | "darkGrid"
  | "darkTrellis"
  | "darkDown"
  | "darkUp"
  | "lightVertical"
  | "lightHorizontal"
  | "lightGrid"
  | "lightTrellis"
  | "lightDown"
  | "lightUp"
  | "darkGray"
  | "mediumGray"
  | "lightGray"
  | "gray125"
  | "gray0625";

export interface GradientStop {
  position: number;
  color: Color;
}

export type Fill =
  | { type: "pattern"; pattern: FillPattern; fgColor?: Color; bgColor?: Color }
  | { type: "gradient"; gradient: "angle"; degree: number; stops: GradientStop[] }
  | {
      type: "gradient";
      gradient: "path";
      center: { left: number; top: number };
      stops: GradientStop[];
    };

export type BorderStyle =
  | "thin"
  | "dotted"
  | "hair"
  | "medium"
  | "double"
  | "thick"
  | "dashed"
  | "dashDot"
  | "dashDotDot"
  | "slantDashDot"
  | "mediumDashed"
  | "mediumDashDotDot"
  | "mediumDashDot";

export interface Border {
  style?: BorderStyle;
  color?: Color;
}

export interface BorderDiagonal extends Border {
  up?: boolean;
  down?: boolean;
}

export interface Borders {
  top?: Border;
  left?: Border;
  bottom?: Border;
  right?: Border;
  diagonal?: BorderDiagonal;
}

export interface Alignment {
  horizontal?:
    | "left"
    | "center"
    | "right"
    | "fill"
    | "justify"
    | "centerContinuous"
    | "distributed";
  vertical?: "top" | "middle" | "bottom" | "distributed" | "justify";
  wrapText?: boolean;
  shrinkToFit?: boolean;
  indent?: number;
  readingOrder?: "rtl" | "ltr";
  textRotation?: number | "vertical";
}

export interface Protection {
  locked?: boolean;
  hidden?: boolean;
}

export interface Style {
  numFmt?: string;
  font?: Font;
  alignment?: Alignment;
  protection?: Protection;
  border?: Borders;
  fill?: Fill;
}

/** Shallow-merges an object-typed style field: column, then row (wins), then cell (wins over both). */
function mergeField<T extends object>(
  colField: T | undefined,
  rowField: T | undefined,
  cellField: T | undefined,
): T | undefined {
  if (colField === undefined && rowField === undefined && cellField === undefined) {
    return undefined;
  }
  return { ...colField, ...rowField, ...cellField } as T;
}

/**
 * Resolve a cell's effective style: row default and column default merged field-by-field — row
 * wins over column — then the cell's own style overrides both (see docs/API-AUDIT.md §4's
 * `_mergeStyle` precedence note). `fill` is a discriminated union so it is replaced wholesale
 * rather than merged (mixing a `pattern` fill's fields with a `gradient` fill's would be
 * incoherent). Returns a new object; never mutates any input.
 */
export function resolveCellStyle(rowStyle?: Style, colStyle?: Style, cellStyle?: Style): Style {
  const result: Style = {};

  const numFmt = cellStyle?.numFmt ?? rowStyle?.numFmt ?? colStyle?.numFmt;
  if (numFmt !== undefined) result.numFmt = numFmt;

  const font = mergeField(colStyle?.font, rowStyle?.font, cellStyle?.font);
  if (font !== undefined) result.font = font;

  const alignment = mergeField(colStyle?.alignment, rowStyle?.alignment, cellStyle?.alignment);
  if (alignment !== undefined) result.alignment = alignment;

  const protection = mergeField(colStyle?.protection, rowStyle?.protection, cellStyle?.protection);
  if (protection !== undefined) result.protection = protection;

  const border = mergeField(colStyle?.border, rowStyle?.border, cellStyle?.border);
  if (border !== undefined) result.border = border;

  const fill = cellStyle?.fill ?? rowStyle?.fill ?? colStyle?.fill;
  if (fill !== undefined) result.fill = fill;

  return result;
}
