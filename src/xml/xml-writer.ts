// A tiny, allocation-light XML string builder for the OOXML writer (see docs/STAGE-1.3-SPEC.md §2).
// Hand-rolled on purpose: fflate is the only runtime dependency and stays that way, so we escape and
// concatenate strings rather than pulling in a DOM/XML library. No streaming, no validation — the
// callers know the schema.

/** Emitted at the head of every part unless `declaration: false` is passed. */
const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** Attribute values accepted by {@link XmlWriter}; `undefined` (and `null`) attributes are dropped. */
type AttrValue = string | number | boolean | undefined | null;

/** Escapes text content: `&`, `<`, `>`. `&` must be replaced first so the others' entities survive. */
export function escapeXmlText(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/** Escapes an attribute value: the text set plus `"` and `'` (both may appear inside a quoted attr). */
export function escapeXmlAttr(s: string): string {
  return escapeXmlText(s).replace(/"/g, "&quot;").replace(/'/g, "&apos;");
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
  private renderAttrs(attrs?: Record<string, AttrValue>): string {
    if (!attrs) return "";
    let out = "";
    for (const key of Object.keys(attrs)) {
      const value = attrs[key];
      if (value === undefined || value === null) continue;
      const raw = typeof value === "boolean" ? (value ? "1" : "0") : String(value);
      out += ` ${key}="${escapeXmlAttr(raw)}"`;
    }
    return out;
  }

  /** Open an element: `<tag …>`. */
  open(tag: string, attrs?: Record<string, AttrValue>): this {
    this.buffer += `<${tag}${this.renderAttrs(attrs)}>`;
    return this;
  }

  /** Self-closing element with attributes only: `<tag …/>`. */
  leaf(tag: string, attrs?: Record<string, AttrValue>): this {
    this.buffer += `<${tag}${this.renderAttrs(attrs)}/>`;
    return this;
  }

  /** Element with escaped text content: `<tag …>text</tag>`. A `<t>` whose text has leading/trailing
   *  whitespace automatically gains `xml:space="preserve"` (explicit attrs win if they set it). */
  el(tag: string, text: string, attrs?: Record<string, AttrValue>): this {
    const finalAttrs =
      tag === "t" && needsWhitespacePreserve(text) ? { "xml:space": "preserve", ...attrs } : attrs;
    this.buffer += `<${tag}${this.renderAttrs(finalAttrs)}>${escapeXmlText(text)}</${tag}>`;
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
