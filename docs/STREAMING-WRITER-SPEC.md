# Streaming Writer Build Spec (Track E → 0.10.0)

> An incremental `WorkbookWriter` (ExcelJS `stream.xlsx.WorkbookWriter`): add sheet → add rows → commit,
> flushing worksheet XML to an output sink as rows arrive instead of holding the whole model in memory.
> A **new, fully independent module** (`src/xlsx/stream/`) — it touches none of the buffered writer's
> files, so it runs in its own worktree in parallel with every other track. Companion to
> [`ADVANCED-ROADMAP.md`](./ADVANCED-ROADMAP.md) Track E. Baseline: **v0.7.0+** (main). Conventions
> unchanged: TS strict, ESM, `Uint8Array`/`TextDecoder`, no `console.*`, **fflate the only runtime dep**,
> clean-room OOXML.
>
> **Architectural → Opus** on the engine + coverage sessions. E2 rebases on A–D once merged so it can
> serialise their features too.

## 0. Scope

**In:** a write-only streaming API that produces a valid `.xlsx` for large workbooks with bounded memory:
```ts
const wb = new WorkbookWriter({ stream /* or */ filename, useSharedStrings=false, useStyles=true });
const ws = wb.addWorksheet("Sheet1");
ws.addRow([1, "a", new Date()]).commit();   // row XML flushed to the sink now
ws.commit();                                  // finalise this sheet's part
await wb.commit();                            // finalise sharedStrings/styles/workbook/content-types + close zip
```
Output sinks: an **async iterable of `Uint8Array` chunks** (environment-agnostic) plus a Node `Writable`
adapter isolated like `writeWorkbookFile` (so browser bundles stay `fs`-free). Cell value coverage matches
the buffered writer (string/number/bool/date/error/formula/richtext/hyperlink).

**Out (defer):** a streaming **reader** (`WorkbookReader`); random-access edits after a row commits;
CSV streaming. Styling/merges/panes/auto-filter/conditional-formatting/comments/images coverage is added in
**E2**, not E1.

**Definition of done:** (a) for a small workbook, streamed output is **structurally equal** (ideally
byte-equal) to `writeWorkbookBuffer` with the same options; (b) a large sheet (e.g. 500k rows) serialises
with roughly flat peak memory vs the buffered path (bench proves it); (c) the output opens in Excel without
repair; (d) parity holds across all in-scope features after E2.

## 1. Session E1 — core engine (Opus 4.8)

### The memory crux

OOXML forces exactly one thing to be buffered: the **sharedStrings** table (it's written after all sheets
but string cells reference it by index). So the streaming writer **defaults `useSharedStrings: false`**
(inline `<is><t>` strings) — the same choice ExcelJS's `WorkbookWriter` makes — which lets worksheet rows
flush immediately with nothing held. When `useSharedStrings: true`, only the string table is retained in
memory; rows still stream (they carry the interned index). Styles pool is small and stays in memory.

### Incremental zip

Use **fflate's streaming `Zip`** (`Zip`, `ZipPassThrough`/`AsyncZipDeflate`) rather than `zipSync`: open a
`ZipPassThrough` entry per part, push chunks as rows commit, and `.end()` each entry on `ws.commit()`.
`wb.commit()` writes the remaining parts (`[Content_Types].xml`, root/workbook rels, `styles.xml`,
`sharedStrings.xml` if used, `theme`, `workbook.xml`) and ends the zip. The `Zip` `ondata` callback feeds
the output sink (async iterable / Writable).

### Reuse, don't duplicate

Factor a **row-level renderer** out of the buffered writer so both paths share it: extract the per-row body
of `renderSheetData` (the `for (const cell of row.cells)` → `renderCell` loop) into an exported
`renderRowXml(row, ctx): string` (or `renderRow`), and have `renderSheetData` call it. The streaming
`ws.addRow(...).commit()` builds a transient `Row` and appends `renderRowXml` to the entry. Reuse
`createSharedStrings`, `createStyles`, `parts/worksheet.ts`'s sheetView/cols/`renderCell`, and the
`date.ts` helpers unchanged. The streaming worksheet emits the sheet prologue (`<worksheet>`, `<sheetViews>`,
`<sheetFormatPr>`, `<cols>`, `<sheetData>` open) on the first row (or `ws.commit()` if empty) and the
epilogue (`</sheetData>`, trailing elements, `</worksheet>`) on `ws.commit()`.

### Public surface

New `src/xlsx/stream/index.ts` exporting `WorkbookWriter`; barrel it from `src/xlsx/index.ts`. Keep it off
the core `Workbook` prototype (it's a separate constructor, not the `workbook.xlsx` accessor).

**Tests:** streamed output for a small multi-sheet, multi-type workbook is structurally equal to
`writeWorkbookBuffer({ useSharedStrings:false })` (parse both with `readZipArchive` + the reader, deep-diff);
inline vs shared-strings modes; empty sheet; the async-iterable sink collects a valid archive.

## 2. Session E2 — feature coverage + parity + bench (Opus 4.8)

Rebase on `main` after Tracks A–D have merged. Extend the streaming writer to emit everything the buffered
writer does, reusing each feature's renderer:

- **Row-commit-safe** features (fold into the streaming row/sheet epilogue): styles (`s` attrs already flow
  through `renderCell`), merges, panes + auto-filter (A), conditional formatting (B, in the epilogue),
  comments (C — collect during row commits, flush `comments{n}.xml` + VML at `ws.commit()`), images (D —
  registered on the workbook, drawing part flushed at `ws.commit()`).
- Anything that must know the full sheet before it can be written (e.g. `<dimension>`, merge/CF blocks that
  follow `</sheetData>`) is emitted in the **epilogue** at `ws.commit()`, which is fine — only the row body
  streams.

**Parity matrix + bench:** a fixture exercising every feature asserted streamed-vs-buffered equal; a
`bench/` case (extend `bench/index.ts` / `bench/workbook-spec.ts`) measuring peak RSS + throughput for a
~500k-row sheet, streaming vs buffered, written to `BENCHMARKS.md`.

**Tests:** the parity matrix; a bounded-memory smoke (large row count doesn't retain rows — assert via a
counting sink or a coarse memory check).

## 3. Session E3 — docs + release 0.10.0 (Fable 5)

Document the streaming API in README + a `docs/STREAMING.md` usage guide (async-iterable and Node stream
examples, the `useSharedStrings:false` default + why). If it fits the ExcelJS shape, expose it through
`@alosha/xlsx/compat` (`ExcelJS.stream.xlsx.WorkbookWriter` — currently a throw-stub in
`src/compat/workbook.ts`); otherwise leave that stub and document the free-function entry. `pnpm changeset`
minor → **0.10.0**; both entries in `npm pack`; `--dry-run` clean; full gate; commit; push. **No
publish/tag/release — owner-only.**

## 4. Sessions & models

| # | Session | Model | Depends | Parallel? |
|---|---------|-------|---------|-----------|
| E1 | Core streaming engine (`src/xlsx/stream/`) | **Opus 4.8** | 0.7.0 baseline | ‖ **everything** |
| E2 | Feature coverage + parity + bench | **Opus 4.8** | E1 (+ A–D merged) | ‖ everything |
| E3 | Docs + release 0.10.0 | **Fable 5** | E2 | ‖ everything |

Branch `feat/streaming` off main; because the module is self-contained, E1 can start immediately and only
E2 needs a rebase once A–D land. The one shared touch is extracting `renderRowXml` from
`parts/worksheet.ts` in E1 — a pure refactor with no output change (guard it with the existing
worksheet-writer tests).

*Grounded in the v0.7.0 tree: `renderSheetData`/`renderCell` in `src/xlsx/parts/worksheet.ts`,
`createSharedStrings`/`createStyles`, `WriteOptions`, `writeZipArchive`/`ZipEntries` (fflate) in
`zip-layer.ts`, the `workbook.xlsx` accessor + barrel in `src/index.ts`, and the existing `bench/`. fflate's
streaming `Zip` is already available in the sole runtime dep. Clean-room OOXML only.*
