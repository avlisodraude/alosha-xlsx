// `xl/comments{n}.xml` — the note/author table backing `cell.note` (docs/COMMENTS-SPEC.md §1).
//
// An `<authors>` table (distinct authors, stable first-seen order) plus a `<commentList>` of
// `<comment ref authorId>` entries, each wrapping a rich `<text>`. The `<text>` runs reuse the
// shared-string run renderer (`renderRunProps` + `<r><t>`); a plain-string comment still renders as
// a single run rather than a lone `<t>` — exceljs's comment parser only reads `<r>` children (see
// COMMENTS-SPEC.md §3), so this keeps plain strings readable by both Excel and exceljs. The reader
// (src/xlsx/read/comments.ts) tolerates either form via the shared `parseRichString`. Pushed as an
// F0 `ExtraPart` with `contentType: CONTENT_TYPE.comments`.

import type { RichText } from "../../model/index.js";
import { XmlWriter } from "../../xml/index.js";
import { renderRunProps } from "../shared-strings.js";
import { NS } from "./constants.js";

/** One collected comment: its cell `ref`, text (plain or rich), and resolved author. */
export interface CommentRecord {
  ref: string;
  texts: RichText | string;
  author: string;
}

/** Render `xl/comments{n}.xml` from the sheet's collected comments (never called when empty). */
export function renderComments(comments: CommentRecord[]): string {
  // Distinct authors in first-seen order; `authorId` on each `<comment>` indexes this table.
  const authorIds = new Map<string, number>();
  const authors: string[] = [];
  for (const c of comments) {
    if (!authorIds.has(c.author)) {
      authorIds.set(c.author, authors.length);
      authors.push(c.author);
    }
  }

  const w = new XmlWriter();
  w.open("comments", { xmlns: NS.spreadsheetml });

  w.open("authors");
  for (const author of authors) w.el("author", author);
  w.close("authors");

  w.open("commentList");
  for (const c of comments) {
    w.open("comment", { ref: c.ref, authorId: authorIds.get(c.author) });
    w.open("text");
    renderCommentText(w, c.texts);
    w.close("text");
    w.close("comment");
  }
  w.close("commentList");

  w.close("comments");
  return w.toString();
}

/** A plain string renders as a single `<r><t>` run (not a lone `<t>` — exceljs only parses `<r>`
 *  children, see the module header); rich text renders one `<r>`/`<rPr>`/`<t>` per run. */
function renderCommentText(w: XmlWriter, texts: RichText | string): void {
  if (typeof texts === "string") {
    w.open("r");
    w.el("t", texts);
    w.close("r");
    return;
  }
  for (const run of texts.richText) {
    w.open("r");
    const props = renderRunProps(run.font);
    if (props) w.raw(props);
    w.el("t", run.text);
    w.close("r");
  }
}
