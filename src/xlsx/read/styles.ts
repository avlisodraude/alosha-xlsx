// Inverse of the Stage 1.3 styles builder (see docs/STAGE-2-SPEC.md §5, and src/xlsx/styles.ts for the
// forward direction). xl/styles.xml is a set of indexed sub-pools; a cell's `s=""` is a cellXf index.
// We rebuild every pool into an array, then compose a Style per cellXf — emitting a sub-object only when
// its `applyX` flag is set or its id is non-default, exactly mirroring the writer. This keeps a
// write -> read -> write cycle structurally identical: the reconstructed Style re-interns to the same
// pool ids in the same order because each sub-object renders byte-for-byte the same XML it was parsed
// from.

import type {
  Alignment,
  Border,
  BorderDiagonal,
  BorderStyle,
  Borders,
  Color,
  Fill,
  FillPattern,
  Font,
  Protection,
  Style,
} from "../../model/index.js";
import { parseXml } from "../../xml/index.js";
import { BUILTIN_NUMFMT_BY_CODE } from "../styles.js";

/** The reverse of {@link BUILTIN_NUMFMT_BY_CODE}: reserved numFmtId -> canonical format code. */
const BUILTIN_NUMFMT_BY_ID = new Map<number, string>();
for (const [code, id] of BUILTIN_NUMFMT_BY_CODE) BUILTIN_NUMFMT_BY_ID.set(id, code);

export interface StyleTable {
  /** Reconstructed Style for a cell's `s` index; index 0 (or out of range) -> {} (the default). */
  styleFor(index: number): Style;
  /** Whether a cellXf's resolved numFmt is a date/time format (drives number-vs-date on read). */
  isDateStyle(index: number): boolean;
}

/** Parse xl/styles.xml into a {@link StyleTable}. A missing part yields an all-default table. */
export function parseStyles(xml: string | undefined): StyleTable {
  if (xml === undefined) return EMPTY_TABLE;
  const root = parseTree(xml);
  if (root === undefined) return EMPTY_TABLE;

  const customNumFmts = parseNumFmts(findChild(root, "numFmts"));
  const fonts = mapChildren(findChild(root, "fonts"), "font", parseFont);
  const fills = mapChildren(findChild(root, "fills"), "fill", parseFill);
  const borders = mapChildren(findChild(root, "borders"), "border", parseBorder);
  const xfs = mapChildren(findChild(root, "cellXfs"), "xf", parseXf);

  const numFmtCode = (id: number): string | undefined =>
    customNumFmts.get(id) ?? BUILTIN_NUMFMT_BY_ID.get(id);

  const styleFor = (index: number): Style => {
    if (!Number.isInteger(index) || index <= 0 || index >= xfs.length) return {};
    const xf = xfs[index];
    if (xf === undefined) return {};

    const style: Style = {};

    if (xf.numFmtId !== 0 || xf.applyNumberFormat) {
      const code = numFmtCode(xf.numFmtId);
      if (code !== undefined && code !== "General") style.numFmt = code;
    }
    if (xf.fontId !== 0 || xf.applyFont) {
      const font = fonts[xf.fontId];
      if (font !== undefined && hasKeys(font)) style.font = font;
    }
    if (xf.fillId !== 0 || xf.applyFill) {
      const fill = fills[xf.fillId];
      if (fill !== undefined) style.fill = fill;
    }
    if (xf.borderId !== 0 || xf.applyBorder) {
      const border = borders[xf.borderId];
      if (border !== undefined && hasKeys(border)) style.border = border;
    }
    if (xf.alignment !== undefined && hasKeys(xf.alignment)) style.alignment = xf.alignment;
    if (xf.protection !== undefined && hasKeys(xf.protection)) style.protection = xf.protection;

    return style;
  };

  const isDateStyle = (index: number): boolean => {
    if (!Number.isInteger(index) || index < 0 || index >= xfs.length) return false;
    const xf = xfs[index];
    if (xf === undefined) return false;
    const id = xf.numFmtId;
    if ((id >= 14 && id <= 22) || (id >= 45 && id <= 47)) return true;
    const code = numFmtCode(id);
    return code !== undefined && isDateFormatCode(code);
  };

  return { styleFor, isDateStyle };
}

const EMPTY_TABLE: StyleTable = {
  styleFor: () => ({}),
  isDateStyle: () => false,
};

// --- number-format date detection -------------------------------------------------------------------

/**
 * Whether a format code renders a date/time. True when it carries a date/time token (`y m d h s`, an
 * elapsed-time bracket like `[h]`/`[mm]`, and by extension AM/PM which always co-occurs with `m`)
 * outside quoted literals, escapes and colour/condition brackets (`[Red]`, `[>100]`, `[$-409]`). Callers
 * short-circuit the reserved built-in date ids (14–22, 45–47); this handles custom codes. Mirrors the
 * ambiguity resolution Excel and ExcelJS apply to decide number-vs-date on read.
 */
export function isDateFormatCode(code: string): boolean {
  const n = code.length;
  let i = 0;
  while (i < n) {
    const ch = code[i];
    if (ch === undefined) break;
    if (ch === "\\") {
      i += 2; // an escaped literal char is never a token
      continue;
    }
    if (ch === '"') {
      i++;
      while (i < n && code[i] !== '"') i++;
      i++; // skip the closing quote
      continue;
    }
    if (ch === "[") {
      const close = code.indexOf("]", i);
      const first = code[i + 1];
      // `[h]`/`[hh]`/`[mm]`/`[ss]` are elapsed-time tokens; `[Red]`/`[>100]`/`[$-x]` are not.
      if (first !== undefined && "hHmMsS".includes(first)) return true;
      i = close === -1 ? n : close + 1;
      continue;
    }
    if ("yYmMdDhHsS".includes(ch)) return true;
    i++;
  }
  return false;
}

// --- xf composition ---------------------------------------------------------------------------------

interface XfData {
  numFmtId: number;
  fontId: number;
  fillId: number;
  borderId: number;
  applyNumberFormat: boolean;
  applyFont: boolean;
  applyFill: boolean;
  applyBorder: boolean;
  alignment?: Alignment;
  protection?: Protection;
}

function parseXf(el: Element): XfData {
  const xf: XfData = {
    numFmtId: numAttr(el, "numFmtId", 0),
    fontId: numAttr(el, "fontId", 0),
    fillId: numAttr(el, "fillId", 0),
    borderId: numAttr(el, "borderId", 0),
    applyNumberFormat: attrTrue(el.attrs.applyNumberFormat),
    applyFont: attrTrue(el.attrs.applyFont),
    applyFill: attrTrue(el.attrs.applyFill),
    applyBorder: attrTrue(el.attrs.applyBorder),
  };
  const alignmentEl = findChild(el, "alignment");
  if (alignmentEl !== undefined) xf.alignment = parseAlignment(alignmentEl);
  const protectionEl = findChild(el, "protection");
  if (protectionEl !== undefined) xf.protection = parseProtection(protectionEl);
  return xf;
}

function parseAlignment(el: Element): Alignment {
  const a: Alignment = {};
  if (el.attrs.horizontal !== undefined)
    a.horizontal = el.attrs.horizontal as Alignment["horizontal"];
  if (el.attrs.vertical !== undefined) a.vertical = el.attrs.vertical as Alignment["vertical"];
  if (el.attrs.textRotation !== undefined) {
    const rotation = inverseTextRotation(Number(el.attrs.textRotation));
    if (rotation !== undefined) a.textRotation = rotation;
  }
  if (attrTrue(el.attrs.wrapText)) a.wrapText = true;
  if (el.attrs.indent !== undefined) a.indent = Number(el.attrs.indent);
  if (attrTrue(el.attrs.shrinkToFit)) a.shrinkToFit = true;
  if (el.attrs.readingOrder === "1") a.readingOrder = "ltr";
  else if (el.attrs.readingOrder === "2") a.readingOrder = "rtl";
  return a;
}

/** Inverse of the writer's rotation encoding: 255 -> "vertical", 0–90 unchanged, 91–180 -> 90 − r. */
function inverseTextRotation(val: number): number | "vertical" | undefined {
  if (val === 255) return "vertical";
  if (val >= 0 && val <= 90) return val;
  if (val > 90 && val <= 180) return 90 - val;
  return undefined;
}

function parseProtection(el: Element): Protection {
  const p: Protection = {};
  if (el.attrs.locked !== undefined) p.locked = attrTrue(el.attrs.locked);
  if (el.attrs.hidden !== undefined) p.hidden = attrTrue(el.attrs.hidden);
  return p;
}

// --- font / fill / border pools ---------------------------------------------------------------------

function parseFont(el: Element): Font {
  const font: Font = {};
  for (const c of el.children) {
    switch (c.name) {
      case "b":
        if (elementOn(c.attrs.val)) font.bold = true;
        break;
      case "i":
        if (elementOn(c.attrs.val)) font.italic = true;
        break;
      case "strike":
        if (elementOn(c.attrs.val)) font.strike = true;
        break;
      case "outline":
        if (elementOn(c.attrs.val)) font.outline = true;
        break;
      case "u":
        font.underline = c.attrs.val === undefined ? true : (c.attrs.val as Font["underline"]);
        break;
      case "vertAlign":
        if (c.attrs.val !== undefined) font.vertAlign = c.attrs.val as Font["vertAlign"];
        break;
      case "sz":
        if (c.attrs.val !== undefined) font.size = Number(c.attrs.val);
        break;
      case "color": {
        const color = parseColor(c);
        if (color !== undefined) font.color = color;
        break;
      }
      case "name":
        if (c.attrs.val !== undefined) font.name = c.attrs.val;
        break;
      case "family":
        if (c.attrs.val !== undefined) font.family = Number(c.attrs.val);
        break;
      case "charset":
        if (c.attrs.val !== undefined) font.charset = Number(c.attrs.val);
        break;
      case "scheme":
        if (c.attrs.val !== undefined) font.scheme = c.attrs.val as Font["scheme"];
        break;
    }
  }
  return font;
}

function parseFill(el: Element): Fill {
  const patternFill = findChild(el, "patternFill");
  if (patternFill !== undefined) {
    const fill: Fill = {
      type: "pattern",
      pattern: (patternFill.attrs.patternType ?? "none") as FillPattern,
    };
    const fg = findChild(patternFill, "fgColor");
    if (fg !== undefined) {
      const color = parseColor(fg);
      if (color !== undefined) fill.fgColor = color;
    }
    const bg = findChild(patternFill, "bgColor");
    if (bg !== undefined) {
      const color = parseColor(bg);
      if (color !== undefined) fill.bgColor = color;
    }
    return fill;
  }

  const gradientFill = findChild(el, "gradientFill");
  if (gradientFill !== undefined) {
    const stops = findChildren(gradientFill, "stop").map((stop) => {
      const colorEl = findChild(stop, "color");
      return {
        position: Number(stop.attrs.position ?? 0),
        color: (colorEl !== undefined ? parseColor(colorEl) : undefined) ?? {},
      };
    });
    if (gradientFill.attrs.type === "path") {
      return {
        type: "gradient",
        gradient: "path",
        center: {
          left: Number(gradientFill.attrs.left ?? 0),
          top: Number(gradientFill.attrs.top ?? 0),
        },
        stops,
      };
    }
    return {
      type: "gradient",
      gradient: "angle",
      degree: Number(gradientFill.attrs.degree ?? 0),
      stops,
    };
  }

  // No recognizable child — fall back to an empty pattern fill so the pool index stays aligned.
  return { type: "pattern", pattern: "none" };
}

function parseBorder(el: Element): Borders {
  const borders: Borders = {};
  const left = parseBorderSide(findChild(el, "left"));
  if (left !== undefined) borders.left = left;
  const right = parseBorderSide(findChild(el, "right"));
  if (right !== undefined) borders.right = right;
  const top = parseBorderSide(findChild(el, "top"));
  if (top !== undefined) borders.top = top;
  const bottom = parseBorderSide(findChild(el, "bottom"));
  if (bottom !== undefined) borders.bottom = bottom;

  const diagonalSide = parseBorderSide(findChild(el, "diagonal"));
  const up = attrTrue(el.attrs.diagonalUp);
  const down = attrTrue(el.attrs.diagonalDown);
  if (diagonalSide !== undefined || up || down) {
    const diagonal: BorderDiagonal = { ...diagonalSide };
    if (up) diagonal.up = true;
    if (down) diagonal.down = true;
    borders.diagonal = diagonal;
  }
  return borders;
}

function parseBorderSide(el: Element | undefined): Border | undefined {
  if (el === undefined || el.attrs.style === undefined) return undefined;
  const border: Border = { style: el.attrs.style as BorderStyle };
  const colorEl = findChild(el, "color");
  if (colorEl !== undefined) {
    const color = parseColor(colorEl);
    if (color !== undefined) border.color = color;
  }
  return border;
}

/** Inverse of `renderColor`: rgb -> argb, else theme, else indexed (writer emits exactly one). */
function parseColor(el: Element): Color | undefined {
  if (el.attrs.rgb !== undefined) return { argb: el.attrs.rgb };
  if (el.attrs.theme !== undefined) return { theme: Number(el.attrs.theme) };
  if (el.attrs.indexed !== undefined) return { indexed: Number(el.attrs.indexed) };
  return undefined;
}

function parseNumFmts(el: Element | undefined): Map<number, string> {
  const map = new Map<number, string>();
  if (el === undefined) return map;
  for (const c of findChildren(el, "numFmt")) {
    const id = c.attrs.numFmtId;
    const code = c.attrs.formatCode;
    if (id !== undefined && code !== undefined) map.set(Number(id), code);
  }
  return map;
}

// --- minimal element tree over the pull parser ------------------------------------------------------

interface Element {
  name: string;
  attrs: Record<string, string>;
  children: Element[];
}

/** Build a shallow element tree from styles.xml. Text nodes are ignored — every datum lives in an
 *  attribute — which keeps this immune to insignificant whitespace between elements. */
function parseTree(xml: string): Element | undefined {
  const root: Element = { name: "#root", attrs: {}, children: [] };
  const stack: Element[] = [root];
  for (const ev of parseXml(xml)) {
    const top = stack[stack.length - 1];
    if (top === undefined) break;
    if (ev.kind === "open") {
      const el: Element = { name: ev.name, attrs: ev.attrs, children: [] };
      top.children.push(el);
      if (!ev.selfClosing) stack.push(el);
    } else if (ev.kind === "close" && stack.length > 1) {
      stack.pop();
    }
  }
  return root.children[0];
}

function findChild(el: Element, name: string): Element | undefined {
  return el.children.find((c) => c.name === name);
}

function findChildren(el: Element, name: string): Element[] {
  return el.children.filter((c) => c.name === name);
}

/** Map every `name` child of a pool container through `fn`, preserving order so array index == id. */
function mapChildren<T>(el: Element | undefined, name: string, fn: (child: Element) => T): T[] {
  if (el === undefined) return [];
  return findChildren(el, name).map(fn);
}

function numAttr(el: Element, name: string, fallback: number): number {
  const raw = el.attrs[name];
  return raw === undefined ? fallback : Number(raw);
}

/** A boolean toggle *element* like `<b/>`: its mere presence means on, so an absent `val` is `true`;
 *  only an explicit `val="0"`/`"false"`/`"off"` turns it off. */
function elementOn(val: string | undefined): boolean {
  return val !== "0" && val !== "false" && val !== "off";
}

/** A boolean *attribute* (`applyFont`, `wrapText`, `diagonalUp`, `locked`…): absent means `false`; only
 *  `"1"`/`"true"`/`"on"` mean `true`. The inverse convention to {@link elementOn}. */
function attrTrue(val: string | undefined): boolean {
  return val === "1" || val === "true" || val === "on";
}

function hasKeys(obj: object): boolean {
  for (const _ in obj) return true;
  return false;
}
