// A tiny, non-validating pull parser for the OOXML reader (see docs/STAGE-2-SPEC.md §2). Hand-rolled
// for the same reason xml-writer.ts is: fflate stays the only runtime dependency, so we scan strings
// rather than pulling in a DOM/SAX library. Not a general XML engine — it targets the well-formed
// subset Excel/LibreOffice/our own writer emit (no DTDs, no external entities).

/** One token from the document in source order. Prefixed names are kept verbatim ("r:id",
 *  "cp:coreProperties") — OOXML fixes its namespace prefixes, so literal matching is sufficient. */
export type XmlEvent =
  | { kind: "open"; name: string; attrs: Record<string, string>; selfClosing: boolean }
  | { kind: "text"; text: string }
  | { kind: "close"; name: string };

const NAMED_ENTITIES: Record<string, string> = {
  amp: "&",
  lt: "<",
  gt: ">",
  quot: '"',
  apos: "'",
};

const ENTITY_RE = /&(#x[0-9a-fA-F]+|#[0-9]+|[a-zA-Z][a-zA-Z0-9]*);/g;

/** Decodes the five named XML entities plus numeric `&#nn;`/`&#xhh;` references. Unknown entities are
 *  left untouched (better a literal `&foo;` survive than we silently drop it). */
function decodeEntities(s: string): string {
  if (!s.includes("&")) return s;
  return s.replace(ENTITY_RE, (match, ent: string) => {
    if (ent.startsWith("#x") || ent.startsWith("#X")) {
      const code = Number.parseInt(ent.slice(2), 16);
      return Number.isFinite(code) ? String.fromCodePoint(code) : match;
    }
    if (ent.startsWith("#")) {
      const code = Number.parseInt(ent.slice(1), 10);
      return Number.isFinite(code) ? String.fromCodePoint(code) : match;
    }
    const named = NAMED_ENTITIES[ent];
    return named !== undefined ? named : match;
  });
}

function isSpace(c: string | undefined): boolean {
  return c === " " || c === "\t" || c === "\n" || c === "\r";
}

/** Finds the unquoted `>` that ends a tag starting at `start` (the char right after `<`). Attribute
 *  values may legally contain a literal `>`, so quote state must be tracked while scanning. */
function findTagEnd(xml: string, start: number): number {
  let i = start;
  let quote: string | undefined;
  while (i < xml.length) {
    const c = xml[i];
    if (quote) {
      if (c === quote) quote = undefined;
    } else if (c === '"' || c === "'") {
      quote = c;
    } else if (c === ">") {
      return i;
    }
    i++;
  }
  return xml.length;
}

/** Parses `name attr1="val1" attr2='val2'` (the tag's inner content, `<`/`>`/trailing `/` stripped). */
function parseOpenTag(inner: string): { name: string; attrs: Record<string, string> } {
  const len = inner.length;
  let i = 0;
  while (i < len && !isSpace(inner[i])) i++;
  const name = inner.slice(0, i);
  const attrs: Record<string, string> = {};

  while (i < len) {
    while (i < len && isSpace(inner[i])) i++;
    if (i >= len) break;
    let j = i;
    while (j < len && inner[j] !== "=" && !isSpace(inner[j])) j++;
    const attrName = inner.slice(i, j);
    if (!attrName) break;
    while (j < len && isSpace(inner[j])) j++;
    if (inner[j] === "=") {
      j++;
      while (j < len && isSpace(inner[j])) j++;
      const quote = inner[j];
      j++;
      const valStart = j;
      while (j < len && inner[j] !== quote) j++;
      attrs[attrName] = decodeEntities(inner.slice(valStart, j));
      j++; // skip closing quote
      i = j;
    } else {
      attrs[attrName] = "";
      i = j;
    }
  }

  return { name, attrs };
}

/** Streams events in document order. Skips the `<?xml …?>` declaration (and any other processing
 *  instruction), `<!-- comments -->` and `<!DOCTYPE …>`-shaped declarations; emits `<![CDATA[…]]>`
 *  content as literal text. Whitespace in text nodes is never trimmed or collapsed. */
export function* parseXml(xml: string): Generator<XmlEvent> {
  const len = xml.length;
  let i = 0;

  while (i < len) {
    if (xml[i] === "<") {
      if (xml.startsWith("<?", i)) {
        const end = xml.indexOf("?>", i + 2);
        i = end === -1 ? len : end + 2;
        continue;
      }
      if (xml.startsWith("<!--", i)) {
        const end = xml.indexOf("-->", i + 4);
        i = end === -1 ? len : end + 3;
        continue;
      }
      if (xml.startsWith("<![CDATA[", i)) {
        const end = xml.indexOf("]]>", i + 9);
        const text = end === -1 ? xml.slice(i + 9) : xml.slice(i + 9, end);
        i = end === -1 ? len : end + 3;
        if (text.length > 0) yield { kind: "text", text };
        continue;
      }
      if (xml.startsWith("<!", i)) {
        const end = xml.indexOf(">", i + 2);
        i = end === -1 ? len : end + 1;
        continue;
      }
      if (xml[i + 1] === "/") {
        const end = xml.indexOf(">", i + 2);
        const name = xml.slice(i + 2, end === -1 ? len : end).trim();
        i = end === -1 ? len : end + 1;
        yield { kind: "close", name };
        continue;
      }

      const tagEnd = findTagEnd(xml, i + 1);
      const selfClosing = xml[tagEnd - 1] === "/";
      const inner = xml.slice(i + 1, selfClosing ? tagEnd - 1 : tagEnd);
      const { name, attrs } = parseOpenTag(inner);
      i = tagEnd === len ? len : tagEnd + 1;
      yield { kind: "open", name, attrs, selfClosing };
      continue;
    }

    const next = xml.indexOf("<", i);
    const raw = next === -1 ? xml.slice(i) : xml.slice(i, next);
    i = next === -1 ? len : next;
    if (raw.length > 0) yield { kind: "text", text: decodeEntities(raw) };
  }
}

/** Strips a namespace prefix (`"cp:coreProperties"` → `"coreProperties"`) for the few comparisons that
 *  must be prefix-agnostic. Most matching is literal — OOXML's prefixes are fixed by convention. */
export function localName(name: string): string {
  const idx = name.indexOf(":");
  return idx === -1 ? name : name.slice(idx + 1);
}

/** Convenience cursor over the event stream for recursive-descent readers. */
export interface XmlCursor {
  next(): XmlEvent | undefined;
  peek(): XmlEvent | undefined;
  /** Collects raw text content until the matching close of `name` (assumes its `open` was already
   *  consumed). Only text events contribute; child markup is skipped over, not included. */
  textUntilClose(name: string): string;
  /** Invokes `onOpen` for each direct child element of the current open element (whose `open` was
   *  already consumed by the caller), then consumes the container's own close. `onOpen` is responsible
   *  for fully consuming its element (including any nested close) before returning, unless it is
   *  self-closing. */
  eachChild(name: string, onOpen: (ev: Extract<XmlEvent, { kind: "open" }>) => void): void;
}

/** Builds a {@link XmlCursor} over `xml`. */
export function cursor(xml: string): XmlCursor {
  const gen = parseXml(xml);
  let lookahead: XmlEvent | undefined;
  let hasLookahead = false;

  function pull(): XmlEvent | undefined {
    const result = gen.next();
    return result.done ? undefined : result.value;
  }

  function next(): XmlEvent | undefined {
    if (hasLookahead) {
      hasLookahead = false;
      const ev = lookahead;
      lookahead = undefined;
      return ev;
    }
    return pull();
  }

  function peek(): XmlEvent | undefined {
    if (!hasLookahead) {
      lookahead = pull();
      hasLookahead = true;
    }
    return lookahead;
  }

  function textUntilClose(name: string): string {
    let text = "";
    let depth = 1;
    for (let ev = next(); ev; ev = next()) {
      if (ev.kind === "text") {
        text += ev.text;
      } else if (ev.kind === "open") {
        if (ev.name === name && !ev.selfClosing) depth++;
      } else if (ev.kind === "close" && ev.name === name) {
        depth--;
        if (depth === 0) break;
      }
    }
    return text;
  }

  function eachChild(
    name: string,
    onOpen: (ev: Extract<XmlEvent, { kind: "open" }>) => void,
  ): void {
    for (let ev = next(); ev; ev = next()) {
      if (ev.kind === "close" && ev.name === name) break;
      if (ev.kind === "open") onOpen(ev);
    }
  }

  return { next, peek, textUntilClose, eachChild };
}
