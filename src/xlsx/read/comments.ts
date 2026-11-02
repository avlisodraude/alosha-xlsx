// Inverse of the Track C1 comments.xml writer (src/xlsx/parts/comments.ts, docs/COMMENTS-SPEC.md §2).
// `<authors>` resolves each `<comment authorId>` to its author string; `<text>` is a CT_Rst container —
// the same grammar as a shared/inline string (lone `<t>` vs `<r>` runs) — so it reuses `parseRichString`.

import type { RichText } from "../../model/index.js";
import { cursor, localName, type XmlCursor } from "../../xml/index.js";
import { NS } from "../parts/constants.js";
import { parseRichString } from "./shared-strings.js";

/** One parsed `<comment>`: its cell `ref`, resolved author, and text (plain or rich). */
export interface CommentRecord {
  ref: string;
  texts: RichText | string;
  author: string;
}

/** Parses `xl/comments{n}.xml`. A malformed/rootless document yields no records; a `<comment>` with no
 *  `ref` or an `authorId` that doesn't resolve against `<authors>` is skipped. */
export function parseComments(xml: string): CommentRecord[] {
  const c = cursor(xml, [NS.spreadsheetml]);
  let ev = c.next();
  while (ev && !(ev.kind === "open" && localName(ev.name) === "comments")) ev = c.next();
  if (ev?.kind !== "open" || ev.selfClosing) return [];

  const authors: string[] = [];
  const records: CommentRecord[] = [];

  c.eachChild(ev.name, (child) => {
    switch (localName(child.name)) {
      case "authors":
        if (!child.selfClosing) collectAuthors(c, child.name, authors);
        break;
      case "commentList":
        if (!child.selfClosing) collectCommentList(c, child.name, authors, records);
        break;
      default:
        if (!child.selfClosing) c.textUntilClose(child.name);
    }
  });

  return records;
}

function collectAuthors(c: XmlCursor, container: string, out: string[]): void {
  c.eachChild(container, (child) => {
    if (localName(child.name) !== "author") {
      if (!child.selfClosing) c.textUntilClose(child.name);
      return;
    }
    out.push(child.selfClosing ? "" : c.textUntilClose(child.name));
  });
}

function collectCommentList(
  c: XmlCursor,
  container: string,
  authors: string[],
  out: CommentRecord[],
): void {
  c.eachChild(container, (child) => {
    if (localName(child.name) !== "comment") {
      if (!child.selfClosing) c.textUntilClose(child.name);
      return;
    }

    const ref = child.attrs.ref;
    const author = authors[Number(child.attrs.authorId)];
    if (ref === undefined || author === undefined) {
      if (!child.selfClosing) c.textUntilClose(child.name);
      return;
    }

    let texts: RichText | string = "";
    if (!child.selfClosing) {
      c.eachChild(child.name, (grand) => {
        if (localName(grand.name) === "text") {
          texts = grand.selfClosing ? "" : parseRichString(c, grand.name);
        } else if (!grand.selfClosing) {
          c.textUntilClose(grand.name);
        }
      });
    }

    out.push({ ref, texts, author });
  });
}
