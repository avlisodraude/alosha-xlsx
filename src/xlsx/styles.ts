// The style dedup pools (see docs/STAGE-1.3-SPEC.md §4, docs/OOXML-NOTES.md §5). xl/styles.xml is a set
// of indexed sub-pools (numFmts, fonts, fills, borders, cellXfs); a cell references a composed cellXf by
// integer id via its `s=""` attribute. `addStyle` decomposes an already-resolved Style (Stage 1.2 did the
// row/column/cell precedence), interns each sub-object structurally, then interns the composed <xf>.
//
// The baseline seeds — fonts[0] Calibri, fills[0] none + fills[1] gray125, borders[0] empty, one default
// cellStyleXf/cellXf — are mandatory: Excel prompts to "repair" the file if these indices are wrong.

import type {
  Alignment,
  Border,
  Borders,
  Color,
  Fill,
  Font,
  Protection,
  Style,
} from "../model/index.js";
import { XmlWriter } from "../xml/index.js";

const SPREADSHEETML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/** Custom number-format codes allocate ids from here upward; 0–163 are reserved for built-ins. */
const FIRST_CUSTOM_NUMFMT_ID = 164;

/** The default font (cellXf/font id 0). Matches Excel's out-of-the-box Calibri 11, theme text colour. */
const DEFAULT_FONT: Font = {
  size: 11,
  color: { theme: 1 },
  name: "Calibri",
  family: 2,
  scheme: "minor",
};

/** cellStyleXfs holds a single default <xf> (id 0); every cellXf points back at it via `xfId="0"`. */
const DEFAULT_CELL_STYLE_XF = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>';

/**
 * Built-in number formats: reserved code → id. A cell whose `numFmt` matches one of these reuses the
 * reserved id and emits no `<numFmt>` entry; anything else becomes a custom code at id ≥ 164. Strings must
 * match exactly to dedup, so these mirror the canonical ECMA-376 built-in codes.
 */
export const BUILTIN_NUMFMT_BY_CODE = new Map<string, number>([
  ["General", 0],
  ["0", 1],
  ["0.00", 2],
  ["#,##0", 3],
  ["#,##0.00", 4],
  ["0%", 9],
  ["0.00%", 10],
  ["0.00E+00", 11],
  ["# ?/?", 12],
  ["# ??/??", 13],
  ["mm-dd-yy", 14],
  ["d-mmm-yy", 15],
  ["d-mmm", 16],
  ["mmm-yy", 17],
  ["h:mm AM/PM", 18],
  ["h:mm:ss AM/PM", 19],
  ["h:mm", 20],
  ["h:mm:ss", 21],
  ["m/d/yy h:mm", 22],
  ["#,##0 ;(#,##0)", 37],
  ["#,##0 ;[Red](#,##0)", 38],
  ["#,##0.00;(#,##0.00)", 39],
  ["#,##0.00;[Red](#,##0.00)", 40],
  ["mm:ss", 45],
  ["[h]:mm:ss", 46],
  ["mmss.0", 47],
  ["##0.0E+0", 48],
  ["@", 49],
]);

export interface Styles {
  /** Intern a resolved cell Style; returns its cellXf index (the value for a cell's `s=""` attr). A cell
   *  with no style — or a structurally-default one — returns 0, the seeded default cellXf. */
  addStyle(style: Style | undefined): number;
  /** Intern a differential style into the `<dxfs>` pool; returns its dxf index (a conditional-format
   *  rule's `dxfId`). Unlike {@link addStyle} there is no baseline seed — dxf ids start at 0. */
  addDxf(style: Partial<Style> | undefined): number;
  /** Render the complete xl/styles.xml. Always non-empty (the baseline seeds are always present). */
  toXml(): string;
}

export function createStyles(): Styles {
  return new StylesBuilder();
}

/** A structural-dedup pool: identical rendered XML collapses to one entry, its array index the id. */
class InternPool {
  private readonly ids = new Map<string, number>();
  readonly xml: string[] = [];

  constructor(seeds: string[]) {
    for (const seed of seeds) this.intern(seed);
  }

  intern(rendered: string): number {
    const existing = this.ids.get(rendered);
    if (existing !== undefined) return existing;
    const id = this.xml.length;
    this.ids.set(rendered, id);
    this.xml.push(rendered);
    return id;
  }

  get count(): number {
    return this.xml.length;
  }
}

class StylesBuilder implements Styles {
  private readonly fonts = new InternPool([renderFont(DEFAULT_FONT)]);
  private readonly fills = new InternPool([
    renderFill({ type: "pattern", pattern: "none" }),
    renderFill({ type: "pattern", pattern: "gray125" }),
  ]);
  private readonly borders = new InternPool([renderBorder(undefined)]);
  private readonly cellXfs = new InternPool([
    renderXf({ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 }),
  ]);
  /** Differential styles for conditional formatting. No baseline seed: dxf ids start at 0. */
  private readonly dxfs = new InternPool([]);

  /** Custom number-format code → allocated id (≥ 164), and the `<numFmt>` fragments in id order. */
  private readonly customNumFmtIds = new Map<string, number>();
  private readonly numFmtXml: string[] = [];
  private nextNumFmtId = FIRST_CUSTOM_NUMFMT_ID;

  /** Fast path for the common case of one Style object shared across many cells. */
  private readonly cache = new WeakMap<Style, number>();

  addStyle(style: Style | undefined): number {
    if (style === undefined) return 0;
    const cached = this.cache.get(style);
    if (cached !== undefined) return cached;

    const numFmtId = this.internNumFmt(style.numFmt);
    const fontId = style.font ? this.fonts.intern(renderFont(style.font)) : 0;
    const fillId = style.fill ? this.fills.intern(renderFill(style.fill)) : 0;
    const borderId = style.border ? this.borders.intern(renderBorder(style.border)) : 0;

    const id = this.cellXfs.intern(
      renderXf({
        numFmtId,
        fontId,
        fillId,
        borderId,
        alignment: style.alignment,
        protection: style.protection,
      }),
    );
    this.cache.set(style, id);
    return id;
  }

  addDxf(style: Partial<Style> | undefined): number {
    return this.dxfs.intern(renderDxf(style));
  }

  private internNumFmt(code: string | undefined): number {
    if (code === undefined) return 0;
    const builtin = BUILTIN_NUMFMT_BY_CODE.get(code);
    if (builtin !== undefined) return builtin;
    const existing = this.customNumFmtIds.get(code);
    if (existing !== undefined) return existing;
    const id = this.nextNumFmtId++;
    this.customNumFmtIds.set(code, id);
    this.numFmtXml.push(renderNumFmt(id, code));
    return id;
  }

  toXml(): string {
    const w = new XmlWriter();
    w.open("styleSheet", { xmlns: SPREADSHEETML_NS });

    if (this.numFmtXml.length > 0) {
      w.open("numFmts", { count: this.numFmtXml.length });
      for (const x of this.numFmtXml) w.raw(x);
      w.close("numFmts");
    }

    w.open("fonts", { count: this.fonts.count });
    for (const x of this.fonts.xml) w.raw(x);
    w.close("fonts");

    w.open("fills", { count: this.fills.count });
    for (const x of this.fills.xml) w.raw(x);
    w.close("fills");

    w.open("borders", { count: this.borders.count });
    for (const x of this.borders.xml) w.raw(x);
    w.close("borders");

    w.open("cellStyleXfs", { count: 1 });
    w.raw(DEFAULT_CELL_STYLE_XF);
    w.close("cellStyleXfs");

    w.open("cellXfs", { count: this.cellXfs.count });
    for (const x of this.cellXfs.xml) w.raw(x);
    w.close("cellXfs");

    w.open("cellStyles", { count: 1 });
    w.leaf("cellStyle", { name: "Normal", xfId: 0, builtinId: 0 });
    w.close("cellStyles");

    if (this.dxfs.count > 0) {
      w.open("dxfs", { count: this.dxfs.count });
      for (const x of this.dxfs.xml) w.raw(x);
      w.close("dxfs");
    } else {
      w.leaf("dxfs", { count: 0 });
    }

    w.leaf("tableStyles", {
      count: 0,
      defaultTableStyle: "TableStyleMedium2",
      defaultPivotStyle: "PivotStyleLight16",
    });

    w.close("styleSheet");
    return w.toString();
  }
}

// --- Pure renderers ---------------------------------------------------------------------------------

/** `<font>` per ECMA-376 CT_Font element order (distinct from the run-property CT_RPrElt order used by
 *  shared strings: here the name element is `<name>`, and strike/underline sit before size). */
export function renderFont(font: Font): string {
  const w = new XmlWriter({ declaration: false });
  w.open("font");
  if (font.bold) w.leaf("b");
  if (font.italic) w.leaf("i");
  if (font.strike) w.leaf("strike");
  if (font.outline) w.leaf("outline");
  renderUnderline(w, font.underline);
  if (font.vertAlign) w.leaf("vertAlign", { val: font.vertAlign });
  if (font.size !== undefined) w.leaf("sz", { val: font.size });
  if (font.color) renderColor(w, "color", font.color);
  if (font.name !== undefined) w.leaf("name", { val: font.name });
  if (font.family !== undefined) w.leaf("family", { val: font.family });
  if (font.charset !== undefined) w.leaf("charset", { val: font.charset });
  if (font.scheme) w.leaf("scheme", { val: font.scheme });
  w.close("font");
  return w.toString();
}

/** `<fill>` — a `<patternFill>` (with optional fg/bg colours) or a linear/path `<gradientFill>`. */
export function renderFill(fill: Fill): string {
  const w = new XmlWriter({ declaration: false });
  w.open("fill");
  if (fill.type === "pattern") {
    if (fill.fgColor || fill.bgColor) {
      w.open("patternFill", { patternType: fill.pattern });
      if (fill.fgColor) renderColor(w, "fgColor", fill.fgColor);
      if (fill.bgColor) renderColor(w, "bgColor", fill.bgColor);
      w.close("patternFill");
    } else {
      w.leaf("patternFill", { patternType: fill.pattern });
    }
  } else if (fill.gradient === "angle") {
    w.open("gradientFill", { degree: fill.degree });
    renderStops(w, fill.stops);
    w.close("gradientFill");
  } else {
    // A path gradient's single centre maps to the OOXML left/right & top/bottom insets.
    w.open("gradientFill", {
      type: "path",
      left: fill.center.left,
      right: fill.center.left,
      top: fill.center.top,
      bottom: fill.center.top,
    });
    renderStops(w, fill.stops);
    w.close("gradientFill");
  }
  w.close("fill");
  return w.toString();
}

function renderStops(w: XmlWriter, stops: { position: number; color: Color }[]): void {
  for (const stop of stops) {
    w.open("stop", { position: stop.position });
    renderColor(w, "color", stop.color);
    w.close("stop");
  }
}

/** `<border>` — CT_Border always lists all four sides plus the diagonal, empty ones self-closed. */
export function renderBorder(borders: Borders | undefined): string {
  const w = new XmlWriter({ declaration: false });
  w.open("border", {
    diagonalUp: borders?.diagonal?.up || undefined,
    diagonalDown: borders?.diagonal?.down || undefined,
  });
  renderBorderSide(w, "left", borders?.left);
  renderBorderSide(w, "right", borders?.right);
  renderBorderSide(w, "top", borders?.top);
  renderBorderSide(w, "bottom", borders?.bottom);
  renderBorderSide(w, "diagonal", borders?.diagonal);
  w.close("border");
  return w.toString();
}

function renderBorderSide(w: XmlWriter, tag: string, side: Border | undefined): void {
  if (!side || side.style === undefined) {
    w.leaf(tag);
    return;
  }
  if (!side.color) {
    w.leaf(tag, { style: side.style });
    return;
  }
  w.open(tag, { style: side.style });
  renderColor(w, "color", side.color);
  w.close(tag);
}

/** `<numFmt numFmtId="164" formatCode="…"/>` — only ever called for custom (id ≥ 164) codes. */
export function renderNumFmt(id: number, code: string): string {
  const w = new XmlWriter({ declaration: false });
  w.leaf("numFmt", { numFmtId: id, formatCode: code });
  return w.toString();
}

interface XfParts {
  numFmtId: number;
  fontId: number;
  fillId: number;
  borderId: number;
  alignment?: Alignment;
  protection?: Protection;
}

/** `<xf …>` — the composed cell format. Each `applyX` flag is set only when its component is non-default,
 *  and `<alignment>`/`<protection>` children are emitted only when they carry real content. */
export function renderXf(xf: XfParts): string {
  const alignmentXml = xf.alignment ? renderAlignment(xf.alignment) : "";
  const protectionXml = xf.protection ? renderProtection(xf.protection) : "";
  const attrs = {
    numFmtId: xf.numFmtId,
    fontId: xf.fontId,
    fillId: xf.fillId,
    borderId: xf.borderId,
    xfId: 0,
    applyNumberFormat: xf.numFmtId !== 0 || undefined,
    applyFont: xf.fontId !== 0 || undefined,
    applyFill: xf.fillId !== 0 || undefined,
    applyBorder: xf.borderId !== 0 || undefined,
    applyAlignment: alignmentXml !== "" || undefined,
    applyProtection: protectionXml !== "" || undefined,
  };

  const w = new XmlWriter({ declaration: false });
  if (alignmentXml === "" && protectionXml === "") {
    w.leaf("xf", attrs);
  } else {
    w.open("xf", attrs);
    if (alignmentXml) w.raw(alignmentXml);
    if (protectionXml) w.raw(protectionXml);
    w.close("xf");
  }
  return w.toString();
}

/**
 * `<dxf>` — a differential style for conditional formatting (CT_Dxf). Child order is
 * `font, numFmt, fill, alignment, border, protection`; only the present sub-objects are emitted.
 * `renderFont`/`renderBorder` are reused as-is, but the fill and numFmt need dxf-specific renderers
 * (see {@link renderDxfFill}). An empty style yields a bare `<dxf/>`, which Excel tolerates.
 */
export function renderDxf(style: Partial<Style> | undefined): string {
  const w = new XmlWriter({ declaration: false });
  const font = style?.font ? renderFont(style.font) : "";
  const numFmt = style?.numFmt !== undefined ? renderDxfNumFmt(style.numFmt) : "";
  const fill = style?.fill ? renderDxfFill(style.fill) : "";
  const alignment = style?.alignment ? renderAlignment(style.alignment) : "";
  const border = style?.border ? renderBorder(style.border) : "";
  const protection = style?.protection ? renderProtection(style.protection) : "";

  if (
    font === "" &&
    numFmt === "" &&
    fill === "" &&
    alignment === "" &&
    border === "" &&
    protection === ""
  ) {
    w.leaf("dxf");
    return w.toString();
  }
  w.open("dxf");
  if (font) w.raw(font);
  if (numFmt) w.raw(numFmt);
  if (fill) w.raw(fill);
  if (alignment) w.raw(alignment);
  if (border) w.raw(border);
  if (protection) w.raw(protection);
  w.close("dxf");
  return w.toString();
}

/** A dxf `<numFmt>`: the `formatCode` is what matters (the reader inverts by it); the `numFmtId` reuses
 *  a built-in id when the code is one, else the first custom slot. */
function renderDxfNumFmt(code: string): string {
  const id = BUILTIN_NUMFMT_BY_CODE.get(code) ?? FIRST_CUSTOM_NUMFMT_ID;
  const w = new XmlWriter({ declaration: false });
  w.leaf("numFmt", { numFmtId: id, formatCode: code });
  return w.toString();
}

/** A dxf `<fill>`. **Footgun:** unlike a cellXf solid fill (colour in `<fgColor>`), a dxf highlight
 *  carries its colour in `<patternFill><bgColor>` — so the fill's colour (from either `bgColor` or
 *  `fgColor`) is emitted as `bgColor`. A non-solid pattern keeps its `patternType`. */
function renderDxfFill(fill: Fill): string {
  const w = new XmlWriter({ declaration: false });
  w.open("fill");
  if (fill.type === "pattern") {
    const color = fill.bgColor ?? fill.fgColor;
    const patternType =
      fill.pattern && fill.pattern !== "solid" && fill.pattern !== "none"
        ? fill.pattern
        : undefined;
    if (color || patternType) {
      w.open("patternFill", { patternType });
      if (color) renderColor(w, "bgColor", color);
      w.close("patternFill");
    } else {
      w.leaf("patternFill");
    }
  } else {
    // A gradient fill is not meaningful as a dxf highlight; emit an empty patternFill to stay valid.
    w.leaf("patternFill");
  }
  w.close("fill");
  return w.toString();
}

/** `<alignment>`, or `""` when the alignment carries no meaningful attributes. */
function renderAlignment(a: Alignment): string {
  const attrs = {
    horizontal: a.horizontal,
    vertical: a.vertical,
    textRotation: convertTextRotation(a.textRotation),
    wrapText: a.wrapText || undefined,
    indent: a.indent,
    shrinkToFit: a.shrinkToFit || undefined,
    readingOrder: convertReadingOrder(a.readingOrder),
  };
  if (!hasDefinedValue(attrs)) return "";
  const w = new XmlWriter({ declaration: false });
  w.leaf("alignment", attrs);
  return w.toString();
}

/** `<protection>`, or `""` when neither locked nor hidden is set (false is meaningful, so it is kept). */
function renderProtection(p: Protection): string {
  const attrs = { locked: p.locked, hidden: p.hidden };
  if (!hasDefinedValue(attrs)) return "";
  const w = new XmlWriter({ declaration: false });
  w.leaf("protection", attrs);
  return w.toString();
}

function renderColor(w: XmlWriter, tag: string, color: Color): void {
  if (color.argb !== undefined) w.leaf(tag, { rgb: color.argb });
  else if (color.theme !== undefined) w.leaf(tag, { theme: color.theme });
  else if (color.indexed !== undefined) w.leaf(tag, { indexed: color.indexed });
}

function renderUnderline(w: XmlWriter, underline: Font["underline"]): void {
  if (underline === undefined || underline === false || underline === "none") return;
  if (underline === true) {
    w.leaf("u");
    return;
  }
  w.leaf("u", { val: underline });
}

/** Excel stores rotation as 0–90 (counter-clockwise), 91–180 (= 90 − angle, clockwise), or 255 (stacked
 *  "vertical"). Out-of-range numbers are dropped rather than producing an invalid attribute. */
function convertTextRotation(textRotation: Alignment["textRotation"]): number | undefined {
  if (textRotation === undefined) return undefined;
  if (textRotation === "vertical") return 255;
  const r = Math.round(textRotation);
  if (r >= 0 && r <= 90) return r;
  if (r < 0 && r >= -90) return 90 - r;
  return undefined;
}

function convertReadingOrder(readingOrder: Alignment["readingOrder"]): number | undefined {
  if (readingOrder === "ltr") return 1;
  if (readingOrder === "rtl") return 2;
  return undefined;
}

function hasDefinedValue(attrs: Record<string, unknown>): boolean {
  for (const key of Object.keys(attrs)) {
    if (attrs[key] !== undefined) return true;
  }
  return false;
}
