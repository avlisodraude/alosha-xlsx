# Images / Drawings Build Spec (Track D → 0.9.0)

> Embedded images (ExcelJS `workbook.addImage` + `worksheet.addImage`): a `xl/drawings/drawing{n}.xml`
> DrawingML anchor part, the raw bytes in `xl/media/image{k}.{ext}`, a drawing-level `.rels`, and a
> `<drawing r:id>` in the worksheet. Second real consumer of the **F0 multi-part seam**
> ([`MULTIPART-PLUMBING-SPEC.md`](./MULTIPART-PLUMBING-SPEC.md)). Companion to
> [`ADVANCED-ROADMAP.md`](./ADVANCED-ROADMAP.md) Track D. Baseline: **v0.7.0 + F0** (main). Conventions
> unchanged: TS strict, ESM, `Uint8Array`/`TextDecoder`, no `console.*`, fflate the only runtime dep,
> clean-room OOXML.
>
> **Runs in parallel with Track C (comments) and Track E (streaming).** The EMU anchor math is the fiddly
> bit → **Opus** on the model+writer sessions.

## 0. Scope

**In:** register an image once at the workbook — `workbook.addImage({ buffer: Uint8Array, extension:
"png"|"jpeg"|"gif" })` → an image id — then place it on a sheet: `worksheet.addImage(id, "B2:D6")`
(two-cell anchor) or `worksheet.addImage(id, { tl: { col, row }, ext: { width, height } })` (one-cell
anchor, pixel extents). Written as `drawing{n}.xml` + `media/image{k}.{ext}` + the drawing's `.rels`, wired
through `<drawing r:id>`. Read back into the registry + per-sheet anchors, exposing the raw bytes.

**Out (defer):** `oneCellAnchor` vs `absoluteAnchor` beyond the two forms above, image
cropping/rotation/effects, hyperlinked images, header/footer images (`addBackgroundImage` stays a
documented stub), SVG, and EMF/WMF.

**Definition of done:** (a) a write→read round-trip reproduces the image **bytes** (byte-identical) and its
anchor; (b) an ExcelJS-authored file with an embedded image parses with the bytes intact (`exceljs`
devDep-guarded); (c) the archive opens in Excel without repair; (d) a workbook with no images is
byte-identical to F0 output.

## 1. Session D1 — model + registry (Opus 4.8)

- **Workbook image registry** (`src/model/image.ts` + workbook wiring): `workbook.addImage({ buffer,
  extension })` interns the bytes and returns a numeric `imageId`; a private `Image[]` on `Workbook`
  round-trips via the workbook model. Dedup identical buffers optionally (not required for v1).
- **Per-sheet anchors:** `worksheet.addImage(imageId, anchor)` where
  ```ts
  type ImageAnchor =
    | string                                            // "B2:D6" two-cell range
    | { tl: CellPosition; br: CellPosition }            // explicit two-cell
    | { tl: CellPosition; ext: { width: number; height: number } };  // one-cell + pixel extents
  type CellPosition = { col: number; row: number; nativeColOff?: number; nativeRowOff?: number }; // 0-based
  ```
  Stored as a `WorksheetImage[]` on the worksheet, round-tripping via `WorksheetModel`. Parse the string
  form through the existing `Range` decoder into `tl`/`br` (0-based col/row). **No `WorksheetHost` change.**

**Tests:** registry add/dedup/id assignment, string-anchor → tl/br decode, model snapshot round-trip,
invalid-anchor typed error.

## 2. Session D2 — writer (Opus 4.8, uses the F0 seam)

When a sheet has images, in `renderWorksheet`:

- **`src/xlsx/parts/drawing.ts`** → render `xl/drawings/drawing${ctx.sheetNumber}.xml`:
  ```xml
  <xdr:wsDr xmlns:xdr="…/drawing/2006/spreadsheetDrawing" xmlns:a="…/drawingml/2006/main">
    <xdr:twoCellAnchor editAs="oneCell">
      <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
      <xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
      <xdr:pic>
        <xdr:nvPicPr><xdr:cNvPr id="1" name="Picture 1"/><xdr:cNvPicPr/></xdr:nvPicPr>
        <xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
        <xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
      </xdr:pic>
      <xdr:clientData/>
    </xdr:twoCellAnchor>
  </xdr:wsDr>
  ```
  For the one-cell/`ext` form emit `<xdr:oneCellAnchor>` with `<xdr:ext cx="…" cy="…"/>` in **EMU**.
  **EMU footgun:** 1 pixel = **9525 EMU**; `colOff`/`rowOff` and `ext` cx/cy are EMU — convert pixel
  extents ×9525, and keep two-cell corner offsets at 0 unless `nativeColOff`/`nativeRowOff` are given.
  The `r:embed` value is the drawing-rels rId for the image.
- **Media:** push the raw bytes as an `ExtraPart` — `{ path: "xl/media/image${k}.${ext}", data: bytes,
  defaultFor: { extension: ext, contentType: IMAGE_CONTENT_TYPE[ext] } }`. The workbook-unique `k` comes
  from the registry.
- **Drawing part + its own rels:** push the drawing XML as an `ExtraPart` with `contentType:
  CONTENT_TYPE.drawing`, and push a **second** `ExtraPart` for the drawing's rels
  (`xl/drawings/_rels/drawing${n}.xml.rels`, plain string, covered by the `rels` Default) whose
  `<Relationship Type=image Target=../media/image${k}.${ext}/>` matches the blip `r:embed`.
- **Worksheet rel + element:** allocate `{ type: REL_TYPE.drawing, target: "../drawings/drawing${n}.xml" }`
  off the sheet's `SheetRelAllocator` and emit `<drawing r:id="…"/>` **before** any `<legacyDrawing>`
  (CT_Worksheet order — see the Track C note).

**Tests:** golden drawing XML (two-cell + one-cell/ext EMU), a **byte-identity** test that the media
survives the zip round-trip, the drawing `.rels` ↔ `r:embed` match, content-types (Override + image
Default), the `<drawing>` position, and a no-images regression.

## 3. Session D3 — reader (Sonnet 5)

Follow the worksheet's `<drawing>` rel → `drawing{n}.xml` → each anchor's `<a:blip r:embed>` → the
drawing's `.rels` → the media path → the bytes. Reconstruct the workbook image registry (assign ids) and
each sheet's `WorksheetImage[]` anchors (EMU → back to the model's pixel/offset shape). Expose the raw
bytes. Missing/unresolvable drawing or media → no image, never throw.

**Tests:** golden inversions (two-cell + one-cell); a round-trip in `tests/xlsx/read-roundtrip.test.ts`
asserting the media bytes are byte-identical and the anchor survives.

## 4. Session D4 — compat + release 0.9.0 (Fable 5)

- **Compat (`src/compat/workbook.ts`):** un-stub `workbook.addImage`/`getImage` (~L430), `workbook.media`
  (~L431), and remove `addImage`/`getImages`/`addBackgroundImage` from `WORKSHEET_UNSUPPORTED` (~L81-83) —
  route to core, matching ExcelJS signatures (`addImage` accepts `{ buffer }` / `{ base64 }` / `{ filename }`;
  support at least `buffer`, document the rest). `addBackgroundImage` may stay a documented stub.
- **Tests:** compat parity (image via `exceljs` + facade, bytes diffed) + `exceljs`-guarded foreign-file.
- **Docs + release:** move images to SUPPORTED in README + `MIGRATION.md` (note deferred cropping/rotation,
  background/HF images, SVG/EMF). `pnpm changeset` minor → **0.9.0**; both entries in `npm pack`; `--dry-run`
  clean; full gate; commit; push. **No publish/tag/release — owner-only.**

## 5. Sessions & models

| # | Session | Model | Depends | Parallel? |
|---|---------|-------|---------|-----------|
| D1 | Model + image registry + anchors | **Opus 4.8** | F0 on main | ‖ Track C, Track E |
| D2 | Writer (drawing.xml + media + drawing rels + wiring) | **Opus 4.8** | D1 | ‖ C, E |
| D3 | Reader | **Sonnet 5** | D2 | ‖ C, E |
| D4 | Compat un-stub + tests + release 0.9.0 | **Fable 5** | D3 | ‖ C, E |

Sequential within the track. Branch `feat/images` off main; stack D1→D4, merge at release. See the Track C
spec §5 for the `<drawing>`-before-`<legacyDrawing>` rebase when both land.

*Grounded in the v0.7.0 + F0 tree: the `ExtraPart` (incl. `Uint8Array` media + `defaultFor`) /
`SheetRelAllocator` / `sheetNumber` seam, `REL_TYPE.drawing`/`.image` + `CONTENT_TYPE.drawing` +
`IMAGE_CONTENT_TYPE` in `constants.ts`, and the `addImage`/`media` throw-stubs in `src/compat/workbook.ts`.
Clean-room DrawingML (CT_Drawing / xdr anchors) shapes only.*
