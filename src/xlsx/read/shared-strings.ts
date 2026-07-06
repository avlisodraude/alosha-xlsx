// Parses `xl/sharedStrings.xml` back into the pool the writer built from (see
// docs/STAGE-2-SPEC.md §4). Inverts `../shared-strings.ts`: a `<si>` holding a lone `<t>` becomes a
// plain string; otherwise its `<r>` runs become a `RichText`, each run's `<rPr>` inverting
// `renderRunProps` into a `Font`.

import type { Font, RichText, RichTextRun } from "../../model/index.js";
import { cursor, localName, type XmlCursor } from "../../xml/index.js";

/** Parses `xl/sharedStrings.xml`. An absent/empty part (no `<sst>` root) yields `[]`. */
export function parseSharedStrings(xml: string): (string | RichText)[] {
  const c = cursor(xml);
  let ev = c.next();
  while (ev && !(ev.kind === "open" && localName(ev.name) === "sst")) ev = c.next();
  if (ev?.kind !== "open" || ev.selfClosing) return [];

  const entries: (string | RichText)[] = [];
  c.eachChild(ev.name, (child) => {
    if (localName(child.name) !== "si") return;
    entries.push(child.selfClosing ? "" : parseSi(c, child.name));
  });
  return entries;
}

/** Parses one `<si>`'s children: a lone `<t>` is a plain string, `<r>` runs make a `RichText`. */
function parseSi(c: XmlCursor, siName: string): string | RichText {
  let text: string | undefined;
  const runs: RichTextRun[] = [];

  c.eachChild(siName, (child) => {
    const name = localName(child.name);
    if (name === "t") {
      text = child.selfClosing ? "" : c.textUntilClose(child.name);
    } else if (name === "r") {
      runs.push(child.selfClosing ? { text: "" } : parseRun(c, child.name));
    } else if (!child.selfClosing) {
      c.textUntilClose(child.name); // skip rPh/phoneticPr etc.
    }
  });

  return runs.length > 0 ? { richText: runs } : (text ?? "");
}

/** Parses one `<r>` run: its `<rPr>` (via {@link parseRunProps}) and `<t>` text. */
function parseRun(c: XmlCursor, rName: string): RichTextRun {
  let text = "";
  let font: Font | undefined;

  c.eachChild(rName, (child) => {
    const name = localName(child.name);
    if (name === "t") {
      text = child.selfClosing ? "" : c.textUntilClose(child.name);
    } else if (name === "rPr") {
      font = child.selfClosing ? {} : parseRunProps(c, child.name);
    } else if (!child.selfClosing) {
      c.textUntilClose(child.name);
    }
  });

  return font ? { text, font } : { text };
}

/** Inverts `renderRunProps`: `<rPr>`'s children back into a `Font`. */
function parseRunProps(c: XmlCursor, rPrName: string): Font {
  const font: Font = {};

  c.eachChild(rPrName, (child) => {
    const name = localName(child.name);
    switch (name) {
      case "rFont":
        font.name = child.attrs.val;
        break;
      case "charset":
        font.charset = Number(child.attrs.val);
        break;
      case "family":
        font.family = Number(child.attrs.val);
        break;
      case "b":
        font.bold = readBool(child.attrs.val);
        break;
      case "i":
        font.italic = readBool(child.attrs.val);
        break;
      case "strike":
        font.strike = readBool(child.attrs.val);
        break;
      case "outline":
        font.outline = readBool(child.attrs.val);
        break;
      case "color":
        font.color = parseColor(child.attrs);
        break;
      case "sz":
        font.size = Number(child.attrs.val);
        break;
      case "u":
        font.underline = (child.attrs.val as Font["underline"]) ?? true;
        break;
      case "vertAlign":
        font.vertAlign = child.attrs.val as Font["vertAlign"];
        break;
      case "scheme":
        font.scheme = child.attrs.val as Font["scheme"];
        break;
    }
    if (!child.selfClosing) c.textUntilClose(child.name);
  });

  return font;
}

function parseColor(attrs: Record<string, string>): Font["color"] {
  if (attrs.rgb !== undefined) return { argb: attrs.rgb };
  if (attrs.theme !== undefined) return { theme: Number(attrs.theme) };
  if (attrs.indexed !== undefined) return { indexed: Number(attrs.indexed) };
  return undefined;
}

/** CT_BooleanProperty: an absent `val` means true (the writer only ever emits the bare element),
 *  `"0"`/`"false"` mean false, anything else true. */
function readBool(val: string | undefined): boolean {
  return val === undefined || (val !== "0" && val !== "false");
}
