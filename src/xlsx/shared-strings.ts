// The shared-string pool (see docs/STAGE-1.3-SPEC.md §3, docs/OOXML-NOTES.md §4). String cells store an
// index into xl/sharedStrings.xml so a value repeated across the workbook is stored once. Plain strings
// dedup by value; rich text dedups by a stable serialization of its runs.

import type { Color, Font, RichText } from "../model/index.js";
import { XmlWriter } from "../xml/index.js";

const SPREADSHEETML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

export interface SharedStrings {
  /** Intern a plain string or rich text; returns its 0-based index. Repeated values return the same
   *  index (but still count as a reference). */
  add(value: string | RichText): number;
  /** Total references added (the sum of `add` calls). */
  readonly count: number;
  /** Distinct entries in the pool. */
  readonly uniqueCount: number;
  /** Render xl/sharedStrings.xml, or `undefined` when the pool is empty (the part is then omitted). */
  toXml(): string | undefined;
}

export function createSharedStrings(): SharedStrings {
  return new SharedStringsPool();
}

class SharedStringsPool implements SharedStrings {
  private readonly index = new Map<string, number>();
  private readonly entries: (string | RichText)[] = [];
  private references = 0;

  add(value: string | RichText): number {
    this.references += 1;
    const key = keyFor(value);
    const existing = this.index.get(key);
    if (existing !== undefined) return existing;
    const idx = this.entries.length;
    this.index.set(key, idx);
    this.entries.push(value);
    return idx;
  }

  get count(): number {
    return this.references;
  }

  get uniqueCount(): number {
    return this.entries.length;
  }

  toXml(): string | undefined {
    if (this.entries.length === 0) return undefined;
    const w = new XmlWriter();
    w.open("sst", {
      xmlns: SPREADSHEETML_NS,
      count: this.references,
      uniqueCount: this.entries.length,
    });
    for (const entry of this.entries) {
      w.open("si");
      if (typeof entry === "string") {
        w.el("t", entry);
      } else {
        for (const run of entry.richText) {
          w.open("r");
          const props = renderRunProps(run.font);
          if (props) w.raw(props);
          w.el("t", run.text);
          w.close("r");
        }
      }
      w.close("si");
    }
    w.close("sst");
    return w.toString();
  }
}

/** Dedup key: plain strings key on their value, rich text on a stable serialization of its runs.
 *  The `s:`/`r:` prefixes keep a plain string from ever colliding with a rich-text serialization. */
function keyFor(value: string | RichText): string {
  return typeof value === "string" ? `s:${value}` : `r:${stableStringify(value.richText)}`;
}

/** Deterministic JSON-ish serialization with object keys sorted and `undefined`-valued keys dropped, so
 *  two structurally identical runs (regardless of property insertion order) produce the same key. */
function stableStringify(value: unknown): string {
  if (value === null || typeof value !== "object") {
    return JSON.stringify(value);
  }
  if (Array.isArray(value)) {
    return `[${value.map(stableStringify).join(",")}]`;
  }
  const obj = value as Record<string, unknown>;
  const keys = Object.keys(obj)
    .filter((k) => obj[k] !== undefined)
    .sort();
  return `{${keys.map((k) => `${JSON.stringify(k)}:${stableStringify(obj[k])}`).join(",")}}`;
}

/** Render a run's `<rPr>` properties from its font, or `""` when there is nothing to emit. Element
 *  order follows the ECMA-376 CT_RPrElt schema so Excel accepts the file without repair. Shared with
 *  the comments writer (docs/COMMENTS-SPEC.md §1), whose `<text>` runs reuse the same `<r>`/`<rPr>`. */
export function renderRunProps(font: Font | undefined): string {
  if (!font) return "";
  const w = new XmlWriter({ declaration: false });
  if (font.name !== undefined) w.leaf("rFont", { val: font.name });
  if (font.charset !== undefined) w.leaf("charset", { val: font.charset });
  if (font.family !== undefined) w.leaf("family", { val: font.family });
  if (font.bold) w.leaf("b");
  if (font.italic) w.leaf("i");
  if (font.strike) w.leaf("strike");
  if (font.outline) w.leaf("outline");
  if (font.color) renderColor(w, font.color);
  if (font.size !== undefined) w.leaf("sz", { val: font.size });
  renderUnderline(w, font.underline);
  if (font.vertAlign) w.leaf("vertAlign", { val: font.vertAlign });
  if (font.scheme) w.leaf("scheme", { val: font.scheme });
  const inner = w.toString();
  return inner === "" ? "" : `<rPr>${inner}</rPr>`;
}

function renderColor(w: XmlWriter, color: Color): void {
  if (color.argb !== undefined) w.leaf("color", { rgb: color.argb });
  else if (color.theme !== undefined) w.leaf("color", { theme: color.theme });
  else if (color.indexed !== undefined) w.leaf("color", { indexed: color.indexed });
}

function renderUnderline(w: XmlWriter, underline: Font["underline"]): void {
  if (underline === undefined || underline === false || underline === "none") return;
  if (underline === true) {
    w.leaf("u");
    return;
  }
  w.leaf("u", { val: underline });
}
