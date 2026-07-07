// Cell comments (ExcelJS `cell.note`, docs/COMMENTS-SPEC.md §1).
//
// A comment belongs to a single cell, so it is stored per worksheet keyed by that cell's ref (unlike
// data validations, which are range-keyed). The `Comments` manager mirrors the `DataValidations`
// shape (`add`/`find`/`remove`/`model`) so `cell.note` is a thin accessor over it and the compat
// facade can pass straight through (C3). `RichText` is a type-only import — no runtime dependency.

import type { RichText } from "./cell-value.js";

/** Mirrors ExcelJS's `cell.note` object shape. `texts` is a plain string or a rich-text run list. */
export interface Comment {
  texts: RichText | string;
  author?: string;
}

/** A comment plus the cell `ref` it applies to — the serializable form on `WorksheetModel`. */
export type CommentModel = { ref: string } & Comment;

/** Author written when a comment carries none (ExcelJS parity). */
export const DEFAULT_COMMENT_AUTHOR = "Author";

/**
 * Per-worksheet store of cell comments keyed by cell ref. Mirrors {@link DataValidations} so
 * `worksheet.comments` is a straight pass-through in the compat facade and `cell.note` is a thin
 * accessor over it.
 */
export class Comments {
  private readonly _map = new Map<string, Comment>();

  /** Register `comment` for `ref` (a single cell address, e.g. "B2"). */
  add(ref: string, comment: Comment): Comment {
    this._map.set(ref, comment);
    return comment;
  }

  /** The comment on `ref`, if any (exact-ref lookup — comments are single-cell). */
  find(ref: string): Comment | undefined {
    return this._map.get(ref);
  }

  remove(ref: string): void {
    this._map.delete(ref);
  }

  get size(): number {
    return this._map.size;
  }

  clear(): void {
    this._map.clear();
  }

  get model(): CommentModel[] {
    return [...this._map.entries()].map(([ref, comment]) => ({ ref, ...comment }));
  }

  set model(models: CommentModel[]) {
    this._map.clear();
    for (const { ref, ...comment } of models) this._map.set(ref, comment);
  }
}
