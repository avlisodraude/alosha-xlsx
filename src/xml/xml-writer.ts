// A tiny, allocation-light XML string builder for the OOXML writer (see docs/STAGE-1.3-SPEC.md §2).
// Hand-rolled on purpose: fflate is the only runtime dependency and stays that way, so we escape and
// concatenate strings rather than pulling in a DOM/XML library. No streaming, no validation — the
// callers know the schema.

import { encodeXString, isXStringAttribute, isXStringElement } from "./xstring.js";

/** Emitted at the head of every part unless `declaration: false` is passed. */
const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** Attribute values accepted by {@link XmlWriter}; `undefined` (and `null`) attributes are dropped. */
type AttrValue = string | number | boolean | undefined | null;

/** Matches, in one pass: the entity chars, every XML 1.0 control char, and lone (unpaired) surrogate
 *  halves. `\t \n \r` are deliberately excluded — they are legal XML content. */
const TEXT_ESCAPE_RE =
  // biome-ignore lint/suspicious/noControlCharactersInRegex: matching XML 1.0's illegal control chars is the point.
  /[&<>]|[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g;

/** Same as {@link TEXT_ESCAPE_RE} plus `"` and `'` for attribute values. */
const ATTR_ESCAPE_RE =
  // biome-ignore lint/suspicious/noControlCharactersInRegex: matching XML 1.0's illegal control chars is the point.
  /[&<>"']|[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g;

/** Replaces one match from {@link TEXT_ESCAPE_RE}/{@link ATTR_ESCAPE_RE}: entity chars get their named
 *  entity; a char illegal in XML 1.0 is re-encoded as Excel's `_xHHHH_` token (the same convention
 *  Excel itself uses for e.g. pasted PDF text) so it round-trips through our own reader instead of
 *  being silently dropped. */
function escapeMatch(match: string): string {
  switch (match) {
    case "&":
      return "&amp;";
    case "<":
      return "&lt;";
    case ">":
      return "&gt;";
    case '"':
      return "&quot;";
    case "'":
      return "&apos;";
    default:
      return `_x${match.charCodeAt(0).toString(16).toUpperCase().padStart(4, "0")}_`;
  }
}

/** Escapes text content: `&`, `<`, `>`, plus any character illegal in XML 1.0 (control chars, lone
 *  surrogates, the U+FFFE/U+FFFF noncharacters), which is re-encoded via Excel's `_xHHHH_` convention
 *  rather than stripped. Single regex pass. */
export function escapeXmlText(s: string): string {
  return s.replace(TEXT_ESCAPE_RE, escapeMatch);
}

/** Escapes an attribute value: the text set plus `"` and `'` (both may appear inside a quoted attr). */
export function escapeXmlAttr(s: string): string {
  return s.replace(ATTR_ESCAPE_RE, escapeMatch);
}

/** A string needs `xml:space="preserve"` when it has leading or trailing whitespace that Excel would
 *  otherwise trim off a `<t>` shared/inline-string element. */
function needsWhitespacePreserve(s: string): boolean {
  return /^\s|\s$/.test(s);
}

export class XmlWriter {
  private buffer = "";

  constructor(opts?: { declaration?: boolean }) {
    if (opts?.declaration !== false) {
      this.buffer = XML_DECLARATION;
    }
  }

  /** Serialize an attribute record, omitting `undefined`/`null` and rendering booleans as `1`/`0`. */
  private renderAttrs(attrs?: Record<string, AttrValue>, tag = ""): string {
    if (!attrs) return "";
    let out = "";
    for (const key of Object.keys(attrs)) {
      const value = attrs[key];
      if (value === undefined || value === null) continue;
      const raw = typeof value === "boolean" ? (value ? "1" : "0") : String(value);
      const encoded = isXStringAttribute(tag, key) ? encodeXString(raw) : raw;
      out += ` ${key}="${escapeXmlAttr(encoded)}"`;
    }
    return out;
  }

  /** Open an element: `<tag …>`. */
  open(tag: string, attrs?: Record<string, AttrValue>): this {
    this.buffer += `<${tag}${this.renderAttrs(attrs, tag)}>`;
    return this;
  }

  /** Self-closing element with attributes only: `<tag …/>`. */
  leaf(tag: string, attrs?: Record<string, AttrValue>): this {
    this.buffer += `<${tag}${this.renderAttrs(attrs, tag)}/>`;
    return this;
  }

  /** Element with escaped text content: `<tag …>text</tag>`. A `<t>` whose text has leading/trailing
   *  whitespace automatically gains `xml:space="preserve"` (explicit attrs win if they set it). */
  el(tag: string, text: string, attrs?: Record<string, AttrValue>): this {
    const finalAttrs =
      tag === "t" && needsWhitespacePreserve(text) ? { "xml:space": "preserve", ...attrs } : attrs;
    const encoded = isXStringElement(tag) ? encodeXString(text) : text;
    this.buffer += `<${tag}${this.renderAttrs(finalAttrs, tag)}>${escapeXmlText(encoded)}</${tag}>`;
    return this;
  }

  /** Append escaped text content at the current position. */
  text(value: string): this {
    this.buffer += escapeXmlText(value);
    return this;
  }

  /** Append already-serialized markup verbatim (escape hatch for composing sub-builders). */
  raw(xml: string): this {
    this.buffer += xml;
    return this;
  }

  /** Close an element: `</tag>`. */
  close(tag: string): this {
    this.buffer += `</${tag}>`;
    return this;
  }

  toString(): string {
    return this.buffer;
  }

  toBytes(): Uint8Array {
    return new TextEncoder().encode(this.buffer);
  }
}
