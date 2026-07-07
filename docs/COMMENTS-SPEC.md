# Comments Build Spec (Track C → 0.8.0)

> Cell comments (ExcelJS `cell.note`): a `xl/comments{n}.xml` note/author table plus a legacy
> `xl/drawings/vmlDrawing{n}.vml` anchor part, referenced from the worksheet via `<legacyDrawing r:id>`.
> First real consumer of the **F0 multi-part seam** ([`MULTIPART-PLUMBING-SPEC.md`](./MULTIPART-PLUMBING-SPEC.md)).
> Companion to [`ADVANCED-ROADMAP.md`](./ADVANCED-ROADMAP.md) Track C. Baseline: **v0.7.0 + F0** (main).
> Conventions unchanged: TS strict, ESM, `Uint8Array`/`TextDecoder`, no `console.*`, fflate the only
> runtime dep, clean-room OOXML.
>
> **Runs in parallel with Track D (images) and Track E (streaming).** C and D each add one trailing element
> to `renderWorksheet` (`<legacyDrawing>` vs `<drawing>`) — see §6 for the shared-line note.

## 0. Scope

**In:** a comment on a cell — `worksheet.getCell("B2").note = "text"` or the object form
`{ texts: RichText | string, author?: string }`. Written as `comments{n}.xml` (author table + `<comment
ref authorId>` with a rich `<text>`) plus the legacy `vmlDrawing{n}.vml` anchor shapes, wired through
`<legacyDrawing r:id>`. Read back into `cell.note`.

**Out (defer):** threaded/modern comments (`xl/threadedComments/` + persons), comment box geometry/margins/
`editAs`/visibility (we emit a sensible default VML shape and ignore its geometry on read), and rich VML
styling beyond the default note box.

**Definition of done:** (a) a write→read round-trip reproduces each comment's `ref`, author, and text
(string and rich forms); (b) an ExcelJS-authored file with a comment parses with author + text intact
(`exceljs` devDep-guarded); (c) the archive opens in Excel without repair (comments visible on hover);
(d) a workbook with no comments is byte-identical to F0 output.

## 1. Session C1 — model + writer

### Model (`src/model/comment.ts` + worksheet manager)

Mirror the data-validation precedent ([[DATA-VALIDATION-SPEC]] §1): a worksheet-level `Comments` manager
keyed by cell ref, plus a `cell.note` accessor routed through `WorksheetHost`.

```ts
export interface Comment { texts: RichText | string; author?: string; }   // ExcelJS cell.note shape
```

`Worksheet` owns a private `Comments` (Map<ref, Comment>); `Cell.note` get/set routes through
`this.rowRef.worksheet.comments` (find/add by `this.address`), exactly as `Cell.dataValidation` does.
Round-trip via a new optional `WorksheetModel.comments`. Default author when unset: `"Author"` (ExcelJS
parity). Export from the model barrel + `src/index.ts`. Add `readonly comments: Comments` to `WorksheetHost`
(type-only import, no cycle — same as `dataValidations`).

### Writer (uses the F0 seam)

During the sheetData pass in `renderWorksheet`, collect a `CommentRecord[]` the way hyperlinks are collected
(`{ ref, texts, author }`). After the pass, when the sheet has comments:

- **`src/xlsx/parts/comments.ts`** → render `xl/comments${ctx.sheetNumber}.xml`:
  ```xml
  <comments xmlns="…/spreadsheetml/2006/main"><authors><author>Author</author>…</authors>
    <commentList><comment ref="B2" authorId="0"><text><r><t>hello</t></r></text></comment>…</commentList>
  </comments>
  ```
  Build the `<authors>` table from the distinct authors (stable order); `authorId` indexes it. `<text>`
  reuses the shared-string run renderer (`renderRunProps`/`<r><t>`), or a lone `<t>` for a plain string.
- **`src/xlsx/parts/vml-drawing.ts`** → render `xl/drawings/vmlDrawing${ctx.sheetNumber}.vml`: the legacy
  `<xml><o:shapelayout>…<v:shapetype…><v:shape>` boilerplate, one `<v:shape>` per comment with a
  `<x:ClientData ObjectType="Note">` carrying the anchor `<x:Row>`/`<x:Column>` (0-based from the ref) and
  `<x:AutoFill>False</x:AutoFill>`. A fixed default box size is fine.
- Push both as `ExtraPart`s (F0): comments part with `contentType: CONTENT_TYPE.comments`; the VML part with
  `defaultFor: { extension: "vml", contentType: CONTENT_TYPE.vmlDrawing }`.
- Allocate **two** rels off the sheet's `SheetRelAllocator`: `{ type: REL_TYPE.comments, target:
  "../comments${n}.xml" }` and `{ type: REL_TYPE.vmlDrawing, target: "../drawings/vmlDrawing${n}.vml" }`.
- Emit `<legacyDrawing r:id="…"/>` (the vmlDrawing rId) as the **last** child of `renderWorksheet`
  (CT_Worksheet: `…drawing, legacyDrawing`).

**Tests:** golden fragments for comments.xml (string + rich + multi-author), the VML shape + ClientData
anchor, the two rels + content-types (Override + `vml` Default), the `<legacyDrawing>` last-child position,
and a no-comments-sheet regression (byte-identical to F0).

## 2. Session C2 — reader

Follow the worksheet's `<legacyDrawing>`/rels to the `comments{n}.xml` part (the VML is anchor-only —
confirm presence, don't parse geometry). Parse `<authors>` + `<commentList>` back into per-cell
`cell.note` (`{ texts, author }`), resolving `authorId` → author string. A plain lone `<t>` → string
`texts`; runs → `RichText`. Missing/unresolvable parts → no note, never throw. Collect during the child
walk, apply after cells are placed (like hyperlinks). Malformed `ref` → skip that comment.

**Tests:** golden inversions (string/rich/multi-author); a write→read round-trip in
`tests/xlsx/read-roundtrip.test.ts`; malformed-part tolerance.

## 3. Session C3 — compat + release 0.8.0

- **Compat (`src/compat/workbook.ts`):** replace the `cell.note` throw (~L281-282) with a pass-through to the
  core `Cell.note` (get/set). ExcelJS `cell.note` accepts a string or `{ texts, ... }`; normalise to our
  `Comment` shape (map ExcelJS `texts:[{font,text}]` runs to our `RichText`; accept a bare string).
- **Tests:** compat parity (a noted cell built via `exceljs` and the facade, both reloaded via `exceljs`,
  author+text diffed) + an `exceljs`-guarded foreign-file test. Document any ExcelJS normalisation quirk
  in `MIGRATION.md` rather than force-matching (tooltip/CF precedent).
- **Docs + release:** move comments to SUPPORTED in README + `MIGRATION.md` (note deferred threaded comments
  + box geometry). `pnpm changeset` minor → **0.8.0**; `npm pack` ships both entries; `--dry-run` clean.
  Run the full gate, commit, push. **No publish/tag/release — owner-only.**

## 4. Sessions & models

| # | Session | Model | Depends | Parallel? |
|---|---------|-------|---------|-----------|
| C1 | Model + writer (comments.xml + VML + wiring) | **Sonnet 5** | F0 on main | ‖ Track D, Track E |
| C2 | Reader | **Sonnet 5** | C1 | ‖ D, E |
| C3 | Compat un-stub + tests + release 0.8.0 | **Fable 5** | C2 | ‖ D, E |

Sequential within the track. Branch `feat/comments` off main; stack C1→C3, merge at release.

## 5. Parallel-with-D note (the shared line)

Both C (`<legacyDrawing>`) and D (`<drawing>`) add a trailing element to `renderWorksheet`, and CT_Worksheet
requires `<drawing>` **before** `<legacyDrawing>`. When both tracks land, whichever merges second rebases so
the two elements sit in that order. Everything else (new part files, model files, rels types) is disjoint.

*Grounded in the v0.7.0 + F0 tree: the `ExtraPart`/`SheetRelAllocator`/`sheetNumber` seam in
`src/xlsx/parts/worksheet.ts`, `REL_TYPE.comments`/`.vmlDrawing` + `CONTENT_TYPE.*` in `constants.ts`, the
`Cell.dataValidation`/`WorksheetHost` manager pattern, and the `cell.note` throw-stub in
`src/compat/workbook.ts`. Clean-room CT_Comments / VML shapes only.*
