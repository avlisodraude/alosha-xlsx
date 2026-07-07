// Embedded images (Track D, docs/IMAGES-DRAWINGS-SPEC.md §1). An image is registered once at the
// workbook — its raw bytes interned in an `ImageRegistry` keyed by a numeric `imageId` — then placed
// on a sheet by that id plus an anchor. The registry owns the bytes (byte-identical round-trip is the
// definition of done); a worksheet only stores `{ imageId, range }`.
//
// Anchors come in three input shapes (`ImageAnchor`) but are stored in a single resolved shape
// (`ResolvedImageAnchor`): either a two-cell box (`tl`/`br`) or a one-cell + pixel extent (`tl`/`ext`).
// All stored positions are **0-based** col/row, matching DrawingML's `<xdr:from>`/`<xdr:to>` indices;
// the `"B2:D6"` string form is decoded through the 1-based `Range` and shifted down by one.

import { Range } from "../address/index.js";
import { InvalidImageAnchorError } from "../errors.js";

/**
 * Media formats supported in v1 (SVG / EMF / WMF are deferred — see spec §0). This union MUST stay in
 * sync with the keys of `IMAGE_CONTENT_TYPE` (`src/xlsx/parts/constants.ts`) so the writer's
 * `IMAGE_CONTENT_TYPE[ext]` lookup type-checks without a cast; kept as a literal here rather than
 * imported so the model layer carries no dependency on the xlsx writer.
 */
export type ImageExtension = "png" | "jpeg" | "jpg" | "gif" | "bmp";

/** Arguments to `workbook.addImage` — the raw bytes plus their format. */
export interface AddImageOptions {
  buffer: Uint8Array;
  extension: ImageExtension;
}

/** An image interned at the workbook: its bytes and format. The serializable registry entry. */
export interface WorkbookImage {
  buffer: Uint8Array;
  extension: ImageExtension;
}

/** A 0-based cell position with optional intra-cell pixel offsets (`colOff`/`rowOff` in EMU come later). */
export interface CellPosition {
  col: number;
  row: number;
  nativeColOff?: number;
  nativeRowOff?: number;
}

/** A two-cell anchor: the image spans from `tl` to `br` (both 0-based). */
export interface TwoCellAnchor {
  tl: CellPosition;
  br: CellPosition;
}

/** A one-cell anchor: pinned at `tl`, sized by an explicit pixel extent. */
export interface OneCellAnchor {
  tl: CellPosition;
  ext: { width: number; height: number };
}

/** The stored anchor shape after resolution — always one of the two normalized forms. */
export type ResolvedImageAnchor = TwoCellAnchor | OneCellAnchor;

/**
 * Accepted `worksheet.addImage` anchor inputs:
 * - a `"B2:D6"` two-cell range string,
 * - an explicit two-cell `{ tl, br }`,
 * - a one-cell `{ tl, ext }` with a pixel extent.
 */
export type ImageAnchor = string | TwoCellAnchor | OneCellAnchor;

/** An image placed on a worksheet: which registry image, anchored where. */
export interface WorksheetImage {
  imageId: number;
  range: ResolvedImageAnchor;
}

/** Whether `p` is a well-formed 0-based cell position (non-negative integer col/row). */
function isCellPosition(p: unknown): p is CellPosition {
  if (typeof p !== "object" || p === null) return false;
  const { col, row } = p as Partial<CellPosition>;
  return (
    Number.isInteger(col) && (col as number) >= 0 && Number.isInteger(row) && (row as number) >= 0
  );
}

/** Deep-copies a `CellPosition`, preserving only the offsets that are present. */
function cloneCellPosition(p: CellPosition): CellPosition {
  const out: CellPosition = { col: p.col, row: p.row };
  if (p.nativeColOff !== undefined) out.nativeColOff = p.nativeColOff;
  if (p.nativeRowOff !== undefined) out.nativeRowOff = p.nativeRowOff;
  return out;
}

/** Deep-copies a resolved anchor so stored anchors never alias caller-supplied objects. */
export function cloneAnchor(anchor: ResolvedImageAnchor): ResolvedImageAnchor {
  if ("ext" in anchor) {
    return { tl: cloneCellPosition(anchor.tl), ext: { ...anchor.ext } };
  }
  return { tl: cloneCellPosition(anchor.tl), br: cloneCellPosition(anchor.br) };
}

/**
 * Normalizes any {@link ImageAnchor} into a {@link ResolvedImageAnchor}. The string form is decoded
 * through the 1-based {@link Range} then shifted to 0-based `tl`/`br`. A malformed anchor throws a
 * typed {@link InvalidImageAnchorError} (the string form may surface an `InvalidRangeError` from the
 * decoder) — never a bare `TypeError`.
 */
export function resolveImageAnchor(anchor: ImageAnchor): ResolvedImageAnchor {
  if (typeof anchor === "string") {
    // `Range` validates the fragment and throws InvalidRangeError on garbage; 1-based → 0-based.
    const range = new Range(anchor);
    return {
      tl: { col: range.left - 1, row: range.top - 1 },
      br: { col: range.right - 1, row: range.bottom - 1 },
    };
  }
  if (typeof anchor === "object" && anchor !== null && isCellPosition(anchor.tl)) {
    if ("ext" in anchor && anchor.ext) {
      const { width, height } = anchor.ext;
      if (Number.isFinite(width) && width > 0 && Number.isFinite(height) && height > 0) {
        return { tl: cloneCellPosition(anchor.tl), ext: { width, height } };
      }
    } else if ("br" in anchor && isCellPosition(anchor.br)) {
      return { tl: cloneCellPosition(anchor.tl), br: cloneCellPosition(anchor.br) };
    }
  }
  throw new InvalidImageAnchorError(anchor);
}

/**
 * Workbook-wide image store. `add` interns bytes and returns their numeric `imageId` (the 0-based
 * insertion index, matching ExcelJS); `get` fetches the entry back. Round-trips through the workbook
 * model via `model`/`model=`. Dedup is intentionally omitted in v1 (spec §1).
 */
export class ImageRegistry {
  private readonly _images: WorkbookImage[] = [];

  /** Interns the bytes and returns the assigned `imageId`. */
  add(options: AddImageOptions): number {
    const id = this._images.length;
    this._images.push({ buffer: options.buffer, extension: options.extension });
    return id;
  }

  /** The registered image for `imageId`, or `undefined` if none. */
  get(imageId: number): WorkbookImage | undefined {
    return this._images[imageId];
  }

  get size(): number {
    return this._images.length;
  }

  get model(): WorkbookImage[] {
    return this._images.map((img) => ({ buffer: img.buffer, extension: img.extension }));
  }

  set model(models: WorkbookImage[]) {
    this._images.length = 0;
    for (const m of models) this._images.push({ buffer: m.buffer, extension: m.extension });
  }
}
