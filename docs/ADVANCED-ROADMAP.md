# @alosha/xlsx — Advanced Features Roadmap (post-0.5.0)

> Execution plan for the remaining "Stage 2 advanced" features: **freeze panes / auto-filter**,
> **conditional formatting**, **comments**, **images / drawings**, and the **streaming writer**.
> Same shape as [`STAGE-2-RUNBOOK.md`](./STAGE-2-RUNBOOK.md): ready-to-paste prompts, one per build
> session, each with its environment, model, dependencies, and whether it can run in parallel.
>
> **Current baseline: v0.5.0** — model, writer, reader, `@alosha/xlsx/compat` shim, ExcelJS parity,
> external hyperlinks (relationships), and cell data validations are all shipped and green.
> Conventions unchanged: TS strict, ESM, `Uint8Array`/`TextDecoder` (no Node `Buffer`), no `console.*`,
> pure functions where practical, **fflate stays the only runtime dep**, clean-room (no ExcelJS code
> copied — only OOXML shapes). Each feature ships as its own minor version behind a SPEC + tests, exactly
> like hyperlinks (0.4.0) and data validation (0.5.0).

---

## How to run these

- **Spec first.** Every feature gets a `docs/<FEATURE>-SPEC.md` authored in **Cowork (Opus 4.8)** before
  the code sessions — that's the "session 0" of each track. All five specs are independent docs and can be
  written up front in one Cowork sitting; the code sessions then just reference them.
- **All code sessions run in Claude Code** on your machine: they write code, run
  `pnpm typecheck && pnpm lint && pnpm test`, then `git commit` + `git push`. One session = one branch;
  parallel sessions = separate `git worktree`s that merge to `main` before the next wave.
- **Within a feature** the arc is always **model → write → read → compat + tests → release**, mirroring
  the hyperlink and data-validation tracks.
- **`pnpm` + Node ≥ 22.** After each merge, `pnpm install` if deps moved (none of these add runtime deps).

## The shared hotspot: `renderWorksheet` element order

Four of these features add a child element to `src/xlsx/parts/worksheet.ts` → `renderWorksheet`, and the
`CT_Worksheet` schema is **order-sensitive**. Every prompt below must keep this order:

```
sheetPr → dimension → sheetViews → sheetFormatPr → cols → sheetData → sheetProtection →
autoFilter → mergeCells → conditionalFormatting → dataValidations → hyperlinks →
pageMargins → pageSetup → drawing → legacyDrawing
```

So: **auto-filter** slots before `<mergeCells>`, **conditionalFormatting** between `<mergeCells>` and the
existing `<dataValidations>`, **drawing** (images) near the end, **legacyDrawing** (comments) last. Because
they all edit the same function, the tracks that touch it are ordered to minimise merge pain (see waves).

---

## Master overview

| Track | Feature | Target | New parts? | Sessions | Truly parallel with |
|-------|---------|--------|-----------|----------|---------------------|
| **A** | Freeze panes + auto-filter | 0.6.0 | no (in-sheet) | 3 | E (streaming) |
| **B** | Conditional formatting | 0.7.0 | no (in-sheet + dxf pool) | 3 | E, F0 |
| **F0**| Multi-part worksheet plumbing (shared) | — | yes (infra only) | 1 | B, E |
| **C** | Comments | 0.8.0 | yes (comments + VML) | 3 | **D**, E |
| **D** | Images / drawings | 0.9.0 | yes (drawing + media) | 3 | **C**, E |
| **E** | Streaming writer | 0.10.0 | no (new module) | 3 | **everything** |

### Execution waves (recommended)

- **Wave 0** — baseline: push/install/green + write all five SPECs in Cowork.
- **Wave 1** — **Track A** (sequential internally) ‖ **E1** (streaming core, own worktree). All SPECs done.
- **Wave 2** — **Track B** (after A merges — shares `renderWorksheet`) ‖ **F0** (shared plumbing, disjoint files).
- **Wave 3** — **Track C** ‖ **Track D** (both depend on F0; each adds one line to `renderWorksheet`, so
  rebase D on C or merge C first).
- **Wave 4** — **E2/E3** streaming finalisation (rebased on A–D so it serialises the new features) + release.

**Truly safe parallelism:** Track E lives entirely in `src/xlsx/stream/*`, so it can run start-to-finish
in its own worktree alongside any other track. C and D are 90% disjoint new files (they only collide on
one trailing line of `renderWorksheet` + a couple of Overrides in `content-types.ts`), so they parallelise
well once F0 is in. A and B both rewrite the middle of `renderWorksheet`, so they are **sequential**.

**Model rationale (mixed):** **Opus 4.8** for the fiddly / high-blast-radius / architectural sessions
(conditional-formatting model + dxf pool, the drawing/media plumbing, DrawingML anchors, the streaming
engine); **Sonnet 5** for the mechanical OOXML write/read sessions; **Fable 5** for the trivial
push / compat-passthrough / release sessions.

---

## Track A — Freeze panes + auto-filter → 0.6.0

Smallest track, no new parts. The model already has `WorksheetView` (`src/model/worksheet.ts:43`) as a
plain-data bag, and `renderWorksheet` already emits a bare `<sheetView>` (`parts/worksheet.ts:59`). This
just fills in `<pane>` + a new `autoFilter` model field and its `<autoFilter>` element.

| # | Session | Env | Model | Depends | Parallel? |
|---|---------|-----|-------|---------|-----------|
| A0 | SPEC (`PANES-AUTOFILTER-SPEC.md`) | Cowork | Opus 4.8 | — | ✅ all specs |
| A1 | Model + writer | Code | **Sonnet 5** | baseline | ‖ E |
| A2 | Reader | Code | **Sonnet 5** | A1 | ‖ E |
| A3 | Compat + tests + release 0.6.0 | Code | **Fable 5** | A2 | ‖ E |

**A1 — model + writer (Sonnet 5)**
> Implement Track A session 1 per `docs/PANES-AUTOFILTER-SPEC.md`. In `src/model/worksheet.ts`, extend
> `WorksheetView` with typed freeze/split fields (`xSplit`, `ySplit`, `topLeftCell`, `activePane`) via a
> `views` setter (`worksheet.views = [{ state: "frozen", xSplit: 1, ySplit: 1 }]`) and add an `autoFilter?:
> string` property (an A1 range like `"A1:D1"`). In `src/xlsx/parts/worksheet.ts`, upgrade the existing
> `<sheetView>` render to emit a `<pane>` child when `state` is `frozen`/`split` (xSplit/ySplit/topLeftCell/
> activePane/state), and add a `renderAutoFilter` that emits `<autoFilter ref="…">` **after `<sheetData>`/
> `<sheetProtection>` and before `<mergeCells>`** (respect the CT_Worksheet order in ADVANCED-ROADMAP.md).
> Add unit tests asserting the golden XML fragments. Run `pnpm typecheck && pnpm lint && pnpm test`, fix
> everything, commit, push.

**A2 — reader (Sonnet 5)**
> Implement Track A session 2 per `docs/PANES-AUTOFILTER-SPEC.md`. In `src/xlsx/read/worksheet.ts`, parse
> `<sheetView><pane>` back into `worksheet.views` and `<autoFilter ref>` into `worksheet.autoFilter`,
> inverting exactly what A1's writer emits. Add reader unit tests plus a write→read round-trip case in
> `tests/xlsx/read-roundtrip.test.ts` covering a frozen-pane + auto-filter sheet. Run
> `pnpm typecheck && pnpm lint && pnpm test`, commit, push.

**A3 — compat + tests + release (Fable 5)**
> Implement Track A session 3. Expose `worksheet.views` and `worksheet.autoFilter` through the
> `@alosha/xlsx/compat` shim (`src/compat/workbook.ts`) matching ExcelJS's `views`/`autoFilter` API, add a
> compat parity test, then run the release: changeset for **0.6.0**, update CHANGELOG/README feature list.
> Run `pnpm typecheck && pnpm lint && pnpm test`, commit, push. Do **not** `npm publish` (manual, by design).

---

## Track B — Conditional formatting → 0.7.0

In-sheet `<conditionalFormatting>` block, but the differential styles (`<dxf>`) live in a **new dedup pool
in `styles.xml`** — so this touches `src/xlsx/styles.ts` / `read/styles.ts`, the highest-blast-radius file
in the writer. Run **after Track A merges** (both rewrite `renderWorksheet`).

| # | Session | Env | Model | Depends | Parallel? |
|---|---------|-----|-------|---------|-----------|
| B0 | SPEC (`CONDITIONAL-FORMATTING-SPEC.md`) | Cowork | Opus 4.8 | — | ✅ all specs |
| B1 | Model + writer + dxf pool | Code | **Opus 4.8** | A merged | ‖ E, F0 |
| B2 | Reader | Code | **Sonnet 5** (Opus optional) | B1 | ‖ E, F0 |
| B3 | Compat + tests + release 0.7.0 | Code | **Fable 5** | B2 | ‖ E, F0 |

**B1 — model + writer + dxf pool (Opus 4.8)**
> Implement Track B session 1 per `docs/CONDITIONAL-FORMATTING-SPEC.md`. Add a `ConditionalFormatting`
> model: a per-worksheet list of `{ ref, rules[] }` where a rule is a discriminated union covering at least
> `cellIs` (operator + formulae), `expression`, `colorScale`, `dataBar`, `iconSet`, `top10`,
> `containsText`, and `duplicateValues`, with an optional differential `Style` (font/fill/border/numFmt).
> In `src/xlsx/styles.ts` add a **`<dxfs>` dedup pool** (reuse the existing font/fill/border builders; a dxf
> is a partial style) and return dxf indices. In `src/xlsx/parts/worksheet.ts` emit `<conditionalFormatting
> sqref>` blocks with `<cfRule>` children **between `<mergeCells>` and `<dataValidations>`** (CT_Worksheet
> order). Golden-fragment tests for each rule type + the dxf pool. Run
> `pnpm typecheck && pnpm lint && pnpm test`, fix everything, commit, push.

**B2 — reader (Sonnet 5)**
> Implement Track B session 2 per `docs/CONDITIONAL-FORMATTING-SPEC.md`. In `src/xlsx/read/styles.ts` parse
> `<dxfs>` into a dxf table; in `src/xlsx/read/worksheet.ts` parse `<conditionalFormatting>`/`<cfRule>` back
> into the model, resolving `dxfId` through the dxf table. Add reader unit tests + a round-trip case per
> rule type in `tests/xlsx/read-roundtrip.test.ts`. Run `pnpm typecheck && pnpm lint && pnpm test`, commit,
> push.

**B3 — compat + tests + release (Fable 5)**
> Implement Track B session 3. Expose `worksheet.addConditionalFormatting(...)` through
> `src/compat/workbook.ts` matching the ExcelJS signature; add a compat parity test and a foreign-file test
> (build a CF workbook with the `exceljs` devDep, load it, assert a rule + dxf survive). Release **0.7.0**
> (changeset, CHANGELOG, README). Run `pnpm typecheck && pnpm lint && pnpm test`, commit, push.

---

## F0 — Multi-part worksheet plumbing (shared foundation)

Comments and images both need a worksheet to own **extra child parts + binary media** in the zip and to
wire them through `content-types.ts` and a worksheet `.rels`. Today `worksheet-rels.ts` only knows about
hyperlinks and `write-workbook.ts` only `put()`s XML strings. Build this once; C and D then parallelise.

| # | Session | Env | Model | Depends | Parallel? |
|---|---------|-----|-------|---------|-----------|
| F0 | Plumbing infra + tests | Code | **Opus 4.8** | baseline | ‖ B, E |

**F0 — plumbing (Opus 4.8)**
> Implement the shared multi-part plumbing (foundation for comments + images), no user-facing feature yet.
> Generalise `src/xlsx/parts/worksheet-rels.ts` so a worksheet can carry **multiple** relationships of
> mixed types (not just hyperlinks) — accept a typed list `{ rId, type, target }`. Extend
> `src/xlsx/parts/content-types.ts` to accept extra `<Override>` part declarations and a `<Default
> Extension="…">` for binary media (e.g. `png`, `jpeg`, `vml`). Extend `src/xlsx/write-workbook.ts` so
> `renderWorksheet`'s `RenderedWorksheet` can return additional sibling parts — both XML strings and raw
> `Uint8Array` media — and have the orchestrator place them into the `ZipEntries` map with correct paths
> (`xl/drawings/…`, `xl/media/…`, `xl/comments…`). Add unit tests proving a synthetic extra XML part + a
> synthetic binary media entry land in the archive with a matching content-type + worksheet rel. Keep
> fflate the only runtime dep. Run `pnpm typecheck && pnpm lint && pnpm test`, commit, push.

---

## Track C — Comments → 0.8.0

Classic (ExcelJS-compatible) comments: `xl/comments1.xml` for the note text + author table, and a legacy
`xl/drawings/vmlDrawing1.vml` for the anchor shapes, referenced from the worksheet via `<legacyDrawing
r:id>` (**last** child in CT_Worksheet order). Depends on **F0**. Parallel with **Track D**.

| # | Session | Env | Model | Depends | Parallel? |
|---|---------|-----|-------|---------|-----------|
| C0 | SPEC (`COMMENTS-SPEC.md`) | Cowork | Opus 4.8 | — | ✅ all specs |
| C1 | Model + writer (comments + VML) | Code | **Sonnet 5** | F0 | ‖ **D**, E |
| C2 | Reader | Code | **Sonnet 5** | C1 | ‖ **D**, E |
| C3 | Compat + tests + release 0.8.0 | Code | **Fable 5** | C2 | ‖ **D**, E |

**C1 — model + writer (Sonnet 5)**
> Implement Track C session 1 per `docs/COMMENTS-SPEC.md`, on top of the F0 plumbing. Add a `comment` (aka
> `note`) to the cell model: `{ texts: RichText | string, author?, ref }` collected per worksheet. In a new
> `src/xlsx/parts/comments.ts` render `xl/comments{n}.xml` (`<authors>` + `<commentList>` with `<comment
> ref authorId><text>` runs) and a new `src/xlsx/parts/vml-drawing.ts` render the legacy
> `xl/drawings/vmlDrawing{n}.vml` anchor shapes. Wire both as worksheet sibling parts + rels (F0 API) and
> emit `<legacyDrawing r:id>` as the **last** child of `renderWorksheet`. Register the content-types.
> Golden-fragment tests for comments.xml, the VML, and the worksheet wiring. Run
> `pnpm typecheck && pnpm lint && pnpm test`, commit, push.

**C2 — reader (Sonnet 5)**
> Implement Track C session 2 per `docs/COMMENTS-SPEC.md`. Follow the worksheet's `<legacyDrawing>`/rels to
> locate and parse `comments{n}.xml` back into per-cell comments (author + rich text). VML is anchor-only —
> read text/author from comments.xml, ignore VML geometry beyond confirming presence. Add reader unit tests
> + a round-trip case in `tests/xlsx/read-roundtrip.test.ts`. Run `pnpm typecheck && pnpm lint && pnpm
> test`, commit, push.

**C3 — compat + tests + release (Fable 5)**
> Implement Track C session 3. Expose `cell.note` / comment access through `src/compat/workbook.ts` matching
> ExcelJS's `cell.note` API; add a compat parity test and a foreign-file test (ExcelJS-authored comment,
> loaded and asserted). Release **0.8.0**. Run `pnpm typecheck && pnpm lint && pnpm test`, commit, push.

---

## Track D — Images / drawings → 0.9.0

DrawingML images: `xl/drawings/drawing{n}.xml` (anchors), `xl/media/image{k}.{ext}` (the binary), a drawing
`.rels`, and `<drawing r:id>` in the worksheet (near the end, before `legacyDrawing`). Anchor math
(twoCellAnchor/oneCellAnchor EMU offsets) is the fiddly bit → **Opus** on the writer. Depends on **F0**.
Parallel with **Track C**.

| # | Session | Env | Model | Depends | Parallel? |
|---|---------|-----|-------|---------|-----------|
| D0 | SPEC (`IMAGES-DRAWINGS-SPEC.md`) | Cowork | Opus 4.8 | — | ✅ all specs |
| D1 | Model + writer (drawing + media) | Code | **Opus 4.8** | F0 | ‖ **C**, E |
| D2 | Reader | Code | **Sonnet 5** | D1 | ‖ **C**, E |
| D3 | Compat + tests + release 0.9.0 | Code | **Fable 5** | D2 | ‖ **C**, E |

**D1 — model + writer (Opus 4.8)**
> Implement Track D session 1 per `docs/IMAGES-DRAWINGS-SPEC.md`, on top of the F0 plumbing. Add a
> workbook-level image registry (`workbook.addImage({ buffer, extension })` → id) and
> `worksheet.addImage(imageId, anchor)` where anchor is a `tl/br` cell range or a one-cell anchor with EMU
> extents. In a new `src/xlsx/parts/drawing.ts` render `xl/drawings/drawing{n}.xml` (`<xdr:twoCellAnchor>` /
> `<xdr:oneCellAnchor>` with `<xdr:from>`/`<xdr:to>` col/row/offset in EMU, `<xdr:pic>` → `<a:blip
> r:embed>`), emit the media bytes as a `Uint8Array` sibling part (`xl/media/image{k}.{ext}`) and the
> drawing `.rels`, and emit `<drawing r:id>` in `renderWorksheet` (before `<legacyDrawing>`). Add a
> `<Default Extension="png"/>` etc. via F0's content-types API. Golden-fragment tests for the drawing XML +
> a byte-identity test that the media survives the zip. Run `pnpm typecheck && pnpm lint && pnpm test`, fix
> everything, commit, push.

**D2 — reader (Sonnet 5)**
> Implement Track D session 2 per `docs/IMAGES-DRAWINGS-SPEC.md`. Follow the worksheet's `<drawing>`/rels →
> `drawing{n}.xml` → blip `r:embed` → media entry; reconstruct the image registry + per-worksheet anchors,
> exposing the raw bytes. Add reader unit tests + a round-trip case asserting the media bytes and anchor are
> byte/structurally identical. Run `pnpm typecheck && pnpm lint && pnpm test`, commit, push.

**D3 — compat + tests + release (Fable 5)**
> Implement Track D session 3. Expose `workbook.addImage` / `worksheet.addImage` through
> `src/compat/workbook.ts` matching ExcelJS; add a compat parity test + a foreign-file test (ExcelJS-embedded
> image, loaded, bytes asserted). Release **0.9.0**. Run `pnpm typecheck && pnpm lint && pnpm test`, commit,
> push.

---

## Track E — Streaming writer → 0.10.0

Architectural and **fully independent** — a new `src/xlsx/stream/` module offering an incremental
`WorkbookWriter` (add sheet → add rows → commit sheet → finalise) that flushes worksheet XML and shared
strings progressively instead of holding the whole model in memory. Can run in its own worktree from
Wave 1; finalise its feature coverage after A–D land so it serialises panes/CF/comments/images too.

| # | Session | Env | Model | Depends | Parallel? |
|---|---------|-----|-------|---------|-----------|
| E0 | SPEC (`STREAMING-WRITER-SPEC.md`) | Cowork | Opus 4.8 | — | ✅ all specs |
| E1 | Core streaming engine | Code | **Opus 4.8** | baseline | ‖ **everything** |
| E2 | Feature coverage + parity/bench | Code | **Opus 4.8** | E1 (+ A–D merged) | ‖ **everything** |
| E3 | Docs + release 0.10.0 | Code | **Fable 5** | E2 | ‖ **everything** |

**E1 — core streaming engine (Opus 4.8)**
> Implement Track E session 1 per `docs/STREAMING-WRITER-SPEC.md`. Create `src/xlsx/stream/` with a
> `WorkbookWriter` that mirrors ExcelJS's streaming API surface: `addWorksheet()` →
> `worksheet.addRow(...).commit()` → `worksheet.commit()` → `workbook.commit()`, writing worksheet XML
> incrementally to an output sink (an async iterable of `Uint8Array` chunks, plus a Node `Writable`
> adapter isolated like `writeWorkbookFile`). Stream shared strings and styles, buffering only what OOXML
> forces (the sharedStrings table). Reuse the existing `parts/*` renderers where they're row-granular;
> factor a row-level renderer out of `renderSheetData` rather than duplicating it. Unit tests: streamed
> output for a small workbook is byte-equal (or structurally equal) to `writeWorkbookBuffer`. Keep fflate
> the only runtime dep. Run `pnpm typecheck && pnpm lint && pnpm test`, commit, push.

**E2 — feature coverage + parity + bench (Opus 4.8)**
> Implement Track E session 2. Extend the streaming writer to cover everything the buffered writer emits —
> styles, merges, panes + auto-filter (A), conditional formatting (B), comments (C), images/drawings (D) —
> reusing each feature's part renderer. Add a parity test matrix asserting streamed vs buffered output match
> across a fixture exercising all features, and a `bench/` case measuring peak memory / throughput for a
> large (e.g. 500k-row) sheet vs the buffered path. Run `pnpm typecheck && pnpm lint && pnpm test`, commit,
> push. (Rebase on `main` after A–D have merged so all features are available.)

**E3 — docs + release (Fable 5)**
> Implement Track E session 3. Document the streaming API in README + a `docs/STREAMING.md` usage guide,
> expose it from the barrel and (if it fits the ExcelJS shape) `@alosha/xlsx/compat`. Release **0.10.0**.
> Run `pnpm typecheck && pnpm lint && pnpm test`, commit, push.

---

## At-a-glance: every session + model

| Order | Session | Model | Wave |
|-------|---------|-------|------|
| 0 | Baseline push/install/green + all 5 SPECs (Cowork) | Fable 5 / Opus 4.8 | 0 |
| 1 | A1 model+writer (panes/autofilter) | Sonnet 5 | 1 |
| 2 | A2 reader | Sonnet 5 | 1 |
| 3 | A3 compat+release 0.6.0 | Fable 5 | 1 |
| ‖ | E1 streaming core | Opus 4.8 | 1 (parallel worktree) |
| 4 | B1 model+writer+dxf (conditional fmt) | Opus 4.8 | 2 |
| 5 | B2 reader | Sonnet 5 | 2 |
| 6 | B3 compat+release 0.7.0 | Fable 5 | 2 |
| ‖ | F0 multi-part plumbing | Opus 4.8 | 2 (parallel worktree) |
| 7 | C1 comments model+writer | Sonnet 5 | 3 |
| 8 | C2 comments reader | Sonnet 5 | 3 |
| 9 | C3 compat+release 0.8.0 | Fable 5 | 3 |
| ‖ | D1 images model+writer | Opus 4.8 | 3 (parallel worktree) |
| ‖ | D2 images reader | Sonnet 5 | 3 (parallel worktree) |
| ‖ | D3 compat+release 0.9.0 | Fable 5 | 3 (parallel worktree) |
| 10 | E2 streaming coverage+bench | Opus 4.8 | 4 |
| 11 | E3 streaming docs+release 0.10.0 | Fable 5 | 4 |

*Total: 5 spec sessions (Cowork) + 17 code sessions. Opus on 6, Sonnet on 7, Fable on 6 (incl. baseline).*

---

*Grounded in the v0.5.0 tree (`src/model/worksheet.ts`, `src/xlsx/parts/worksheet.ts`,
`src/xlsx/write-workbook.ts`, `src/xlsx/parts/{content-types,constants,worksheet-rels}.ts`) and the
hyperlink/data-validation precedents. Clean-room throughout — OOXML shapes only.*
