# Multi-part Worksheet Plumbing Build Spec (F0 — shared foundation)

> Infrastructure only — **no user-facing feature**. Generalises the package machinery so a worksheet can
> own **extra child parts (XML *and* binary media)** with their own content-types and relationships, then
> proves it with synthetic tests. Unblocks Track C (comments) and Track D (images/drawings) to run in
> parallel afterward. Companion to [`ADVANCED-ROADMAP.md`](./ADVANCED-ROADMAP.md) F0. Baseline: **v0.7.0**
> (`421a22c`). Conventions unchanged: TS strict, ESM, `Uint8Array`/`TextDecoder`, no `console.*`, fflate
> the only runtime dep, clean-room OOXML.
>
> **One Opus 4.8 session.** It's small in surface but high blast-radius (touches the write orchestrator and
> the worksheet rels path that hyperlinks already depend on), so it gets Opus and a green full-suite gate.

## 0. Why this exists (the three gaps today)

1. **`worksheet-rels.ts`** hardcodes `Type: REL_TYPE.hyperlink` and always emits `TargetMode` — it can
   only describe external hyperlinks. Comments/drawings are **internal** rels (no `TargetMode`) of
   **different** types, and a single sheet may carry hyperlinks *and* a drawing *and* a legacy comment
   drawing in the **same** `.rels` file, sharing one local rId namespace.
2. **`content-types.ts`** emits a fixed set — no way to add an `<Override>` for `comments1.xml` /
   `drawing1.xml`, nor a `<Default Extension="png"…>` / `vml` for binary media.
3. **`write-workbook.ts`** only `put()`s XML strings; there's no path for a worksheet to contribute extra
   sibling parts, and none for raw `Uint8Array` media (`ZipEntries` is already `Map<string,Uint8Array>`,
   so the map itself is fine — the *plumbing to fill it* is what's missing).

## 1. Definition of done

- `renderWorksheetRels` renders a mixed list: an external hyperlink (Type=hyperlink, `TargetMode="External"`)
  and an internal drawing/comment rel (its own Type, **no** `TargetMode`) in one `.rels`, sharing rIds.
- `renderContentTypes` accepts extra `<Default>`s (deduped by extension) and extra `<Override>`s, emitted
  in schema order (all Defaults before all Overrides).
- The orchestrator writes a worksheet's extra parts — both encoded XML strings and raw `Uint8Array` media —
  into the archive at their declared paths, and folds their content-type Overrides/Defaults into
  `[Content_Types].xml`.
- A worksheet's `.rels` is built from **one** accumulated `WorksheetRel[]` per sheet (hyperlinks are just
  the first contributors), with a single local rId allocator — so C/D can push more rels without colliding.
- **No behavioural change to existing output:** a workbook with hyperlinks (and none of the new parts)
  serializes byte-identically to 0.7.0. All 651 existing tests stay green.

## 2. `constants.ts` — generalise the rel + add the OOXML constants C/D will consume

Widen `WorksheetRel` so it can describe any worksheet relationship:

```ts
export interface WorksheetRel {
  rId: string;                 // local to the sheet's own .rels namespace (rId1, rId2, …)
  type: string;                // a REL_TYPE.* URI (was implicitly hyperlink)
  target: string;              // external URL, or an internal path relative to xl/worksheets/ (e.g. "../drawings/drawing1.xml")
  mode?: "External";           // TargetMode — omitted for internal rels
}
```

Add the type/content-type constants (pure data — neutral to add now, centralises them for C/D):

```ts
REL_TYPE.drawing      = ".../officeDocument/2006/relationships/drawing"
REL_TYPE.comments     = ".../officeDocument/2006/relationships/comments"
REL_TYPE.vmlDrawing   = ".../officeDocument/2006/relationships/vmlDrawing"
REL_TYPE.image        = ".../officeDocument/2006/relationships/image"

CONTENT_TYPE.drawing    = "application/vnd.openxmlformats-officedocument.drawing+xml"
CONTENT_TYPE.comments   = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
CONTENT_TYPE.vmlDrawing = "application/vnd.openxmlformats-officedocument.vmlDrawing"

/** Image content-type by lowercase file extension, for the media `<Default>`s. */
export const IMAGE_CONTENT_TYPE: Record<string, string> = {
  png: "image/png", jpeg: "image/jpeg", jpg: "image/jpeg", gif: "image/gif", bmp: "image/bmp",
};
```

## 3. `worksheet-rels.ts` — render any rel type, TargetMode-optional

```ts
export function renderWorksheetRels(rels: WorksheetRel[]): string; // unchanged signature
```

Emit `Type: rel.type` (not the hardcoded hyperlink) and include `TargetMode: rel.mode` **only when
`rel.mode` is set** (`XmlWriter` already drops undefined attrs, so passing `TargetMode: rel.mode` directly
does the right thing). Nothing else changes.

## 4. `content-types.ts` — pluggable Defaults + Overrides

```ts
export interface ContentTypesInput {
  worksheetFiles: string[];
  hasSharedStrings: boolean;
  extraDefaults?: { extension: string; contentType: string }[];   // e.g. png, vml
  extraOverrides?: { partName: string; contentType: string }[];   // e.g. /xl/drawings/drawing1.xml
}
```

Emit the two base `<Default>`s, then the **deduped** `extraDefaults` (by lowercase extension — two PNG
images must yield one `<Default Extension="png">`), then all `<Override>`s (base set, then `extraOverrides`).
Schema requires every `<Default>` before any `<Override>`; keep that invariant. `partName` values are
absolute (`/xl/drawings/drawing1.xml`).

## 5. `worksheet.ts` — unify the sheet's rels + surface extra parts

Today `renderWorksheet` collects `HyperlinkRecord[]` and `renderHyperlinks` both emits `<hyperlinks>` **and**
returns the rendered `.rels` string. Refactor so the sheet has **one** rel accumulator and **one** rId
allocator that every feature shares:

- Introduce a local `SheetRelAllocator`: `next(): string` handing out `rId1, rId2, …` for this sheet, and a
  `rels: WorksheetRel[]` it appends to. Hyperlinks become the first consumer — each external link calls
  `alloc.next()` for its `r:id` (same value it already emits in `<hyperlink r:id>`) and pushes
  `{ rId, type: REL_TYPE.hyperlink, target, mode: "External" }`.
- Add an `extraParts: ExtraPart[]` accumulator (empty in F0 — C/D push to it):

  ```ts
  export interface ExtraPart {
    path: string;                                   // archive path, e.g. "xl/drawings/drawing1.xml" or "xl/media/image1.png"
    data: string | Uint8Array;                      // XML text (encoded by the orchestrator) or raw media bytes
    contentType?: string;                           // → an <Override> for this exact part
    defaultFor?: { extension: string; contentType: string }; // → a <Default> for media by extension
  }
  ```

- After the sheetData/features pass, call `renderWorksheetRels(alloc.rels)` **once** (when non-empty) for
  the sheet's `.rels`. Widen the return type:

  ```ts
  export interface RenderedWorksheet {
    xml: string;
    rels?: string;                 // the single unified .rels (undefined when the sheet has no rels)
    extraParts?: ExtraPart[];      // undefined/empty when the sheet owns no child parts
  }
  ```

- Thread the **1-based sheet number** into `WorksheetContext` (`sheetNumber: number`) so C/D can name their
  parts deterministically (`drawing${n}.xml`, `comments${n}.xml`, `vmlDrawing${n}.vml`, and workbook-unique
  media names). F0 only adds the field; the orchestrator sets it in the worksheet loop.

The hyperlink `<hyperlinks>` block, ordering, and output are otherwise unchanged — verify a hyperlink-only
workbook still round-trips identically.

## 6. `write-workbook.ts` — collect + assemble (the testable seam)

Factor the extra-part assembly into an **exported pure helper** so F0 can prove the path without a real
feature:

```ts
export interface CollectedParts {
  media: Map<string, Uint8Array>;                             // path -> bytes (already-encoded XML + raw media)
  extraDefaults: { extension: string; contentType: string }[];
  extraOverrides: { partName: string; contentType: string }[];
}
export function collectWorksheetParts(rendered: Iterable<RenderedWorksheet>, enc: TextEncoder): CollectedParts;
```

`collectWorksheetParts` walks each `RenderedWorksheet.extraParts`, encoding string `data` via `enc` and
passing `Uint8Array` through untouched into `media`, and gathers each part's `contentType` →
`extraOverrides` (partName = `/${path}`) and `defaultFor` → `extraDefaults`. The orchestrator then:

1. Renders worksheets (unchanged — this fills the pools) into `worksheetXml`.
2. `const collected = collectWorksheetParts(worksheetXml.values(), enc)`.
3. Passes `collected.extraDefaults` / `collected.extraOverrides` into `renderContentTypes`.
4. In the worksheet write loop, after `put`ing each sheet's `xml` and `.rels`, also `entries.set(path,
   bytes)` for every `collected.media` entry (do this once, not per sheet).
5. Sets `ctx.sheetNumber = i + 1` per sheet.

Everything else (rId discipline for workbook rels, pool decisions) is unchanged.

## 7. Tests (synthetic — no feature yet)

- **`worksheet-rels`:** a mixed `WorksheetRel[]` — one hyperlink (`type=hyperlink`, `TargetMode="External"`)
  + one drawing (`type=drawing`, target `../drawings/drawing1.xml`, **no** `TargetMode`) → assert both
  `<Relationship>`s render with the right `Type` and that only the hyperlink carries `TargetMode`.
- **`content-types`:** `extraDefaults` with a duplicate `png` collapses to one `<Default>`; an
  `extraOverride` for `/xl/drawings/drawing1.xml` appears; every `<Default>` precedes every `<Override>`.
- **`collectWorksheetParts`:** feed a hand-built `RenderedWorksheet[]` containing a synthetic XML part
  (`{ path: "xl/drawings/drawing1.xml", data: "<xml/>", contentType: CONTENT_TYPE.drawing }`) and a
  synthetic binary media part (`{ path: "xl/media/image1.png", data: Uint8Array([1,2,3]),
  defaultFor: { extension:"png", contentType: IMAGE_CONTENT_TYPE.png } }`) → assert the bytes land under
  the right paths, the override + default are collected, and (via a full `writeWorkbookBuffer` +
  `readZipArchive`) both parts are present in the actual archive with a matching `[Content_Types].xml`.
- **Regression:** a hyperlink workbook (existing fixture) serializes byte-identically to before, and the
  full suite (651) stays green.

## 8. Scope guard

F0 adds **no** model, no `<drawing>`/`<legacyDrawing>` emission, no reader changes, and no compat changes.
The `extraParts` accumulator stays empty in normal output, so `RenderedWorksheet.extraParts` is `undefined`
and the archive is unchanged. Comments (Track C) and images (Track D) are the first real consumers.

## 9. Sessions & models

| # | Session | Model | Depends | Parallel? |
|---|---------|-------|---------|-----------|
| F0 | Generalise rels + content-types + orchestrator seam + tests | **Opus 4.8** | 0.7.0 | ‖ Track E (streaming) |

After F0 merges to `main`, **Track C (comments) and Track D (images) can run in parallel** in separate
worktrees — they share only F0's now-stable seam plus one trailing line each in `renderWorksheet`
(`<legacyDrawing>` / `<drawing>`), so rebase D on C (or merge C first) to settle that line.

*Grounded in the v0.7.0 tree: `src/xlsx/parts/{worksheet-rels,content-types,constants,worksheet}.ts` and
`src/xlsx/write-workbook.ts` (the `put`/`ZipEntries` assembly, the `ctx` object, the worksheet render
loop). Clean-room OPC/relationship shapes only.*
