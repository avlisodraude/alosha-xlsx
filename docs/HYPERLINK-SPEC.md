# Hyperlink Relationships Build Spec — real clickable links (write + read + compat)

> The first Stage 2 write feature. Persist cell hyperlinks as proper OOXML relationships instead of the
> current plain-text degradation, and read them back into the `{ text, hyperlink, tooltip? }` model value.
> Closes the "hyperlink cells round-trip as their display text" **Known limitation** (README /
> `docs/MIGRATION.md`). Ships as **0.4.0**.
>
> Prereqs (all published in 0.3.1): the writer (`src/xlsx/parts/*` + `write-workbook.ts`), the reader
> (`src/xlsx/read/*` + `read-workbook.ts`), the compat facade, and the `Hyperlink` model value
> (`{ text, hyperlink, tooltip? }`, `CellType.Hyperlink`). Conventions unchanged: TS strict, ESM+CJS,
> fflate the only runtime dep, `exceljs` a devDep used by the parity/foreign tests.

## 0. Scope

**In:** write and read **external** cell hyperlinks (http/https/mailto/file URLs). A hyperlink cell keeps
its visible string value AND gains a clickable target that survives a round-trip and opens correctly in
Excel/LibreOffice. Wire it through the compat facade so `cell.type` is `Hyperlink` and the value reads back
as `{ text, hyperlink, tooltip? }`, matching ExcelJS. Update the docs to drop the limitation.

**Out (defer):** internal in-workbook hyperlinks (the `location="Sheet2!A1"` form, which uses no external
relationship), hyperlinks anchored to drawings/images, and hyperlinks spanning a merged range as a single
`ref`. These can be follow-ons; note them in `MIGRATION.md` as still-deferred.

**Behaviour-change note (call out in the changelog):** reading a workbook that contains hyperlinks now
yields `{ text, hyperlink }` cell values (`ValueType.Hyperlink`) where 0.3.x returned a plain string. This
is the intended fix, but it changes read output for hyperlink-bearing files — hence the **minor** bump.

**Definition of done:** a hyperlink written by `@alosha/xlsx` (a) unzips to a `<hyperlinks>` block in the
sheet + a `.../hyperlink` external relationship in `xl/worksheets/_rels/sheetN.xml.rels`, (b) re-opens in
Excel/LibreOffice as a working link with the right display text and tooltip, (c) round-trips through our own
reader back to an equal `{ text, hyperlink, tooltip? }`, and (d) an ExcelJS-written hyperlink parses with
our reader (and vice-versa) — asserted in tests. Docs updated; typecheck/lint/tests green.

## 1. The OOXML shape (three coordinated pieces)

A hyperlink cell is **not** a distinct cell type on the wire — it is a normal string cell **plus** a sheet-
level hyperlink record **plus** an external relationship:

1. **`sheetData`** — the cell renders exactly as a string cell today (shared or inline), carrying the
   *display text*. No change to the `<c>` element.
2. **`<hyperlinks>`** — a sheet-level block (CT_Worksheet order: **after `<mergeCells>`**, before
   `<pageMargins>`):
   ```xml
   <hyperlinks>
     <hyperlink ref="A1" r:id="rId1" tooltip="Visit"/>
   </hyperlinks>
   ```
   `ref` is the cell (or range) address; `r:id` resolves in the **worksheet's own** rels; `tooltip` is
   optional; `display` is optional (Excel falls back to the cell value, so we omit it).
3. **`xl/worksheets/_rels/sheetN.xml.rels`** — one relationship per hyperlink:
   ```xml
   <Relationship Id="rId1"
     Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
     Target="https://example.com" TargetMode="External"/>
   ```
   The worksheet `.rels` file is covered by the existing `<Default Extension="rels" …>` in
   `[Content_Types].xml`, so **content-types needs no change**. rIds here live in the worksheet's own
   relationship namespace (rId1, rId2, … per sheet), independent of the workbook rels.

## 2. Write path

- **`src/xlsx/parts/constants.ts`** — add `REL_TYPE.hyperlink =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"`.
- **`src/xlsx/parts/worksheet.ts`** — the `CellType.Hyperlink` case still emits the display text as a
  string cell (unchanged visual), but ALSO records `{ ref, target, tooltip }` in a per-sheet collector.
  After `renderMergeCells`, emit a `<hyperlinks>` block, assigning each a local `rId${n}`. Return the
  collected hyperlink rels so the orchestrator can write the sheet's `.rels`. Change `renderWorksheet`'s
  return type from `string` to `{ xml: string; rels?: string }` (rels is the rendered
  `sheetN.xml.rels`, or `undefined` when the sheet has no hyperlinks).
- **`src/xlsx/parts/worksheet-rels.ts`** (new) — `renderWorksheetRels(rels: {rId, target, mode}[]):
  string` producing the `<Relationships>` doc (mirrors `workbook-rels.ts`).
- **`src/xlsx/write-workbook.ts`** — `worksheetXml` becomes `Map<file, { xml, rels? }>`; when `rels` is
  present, also `put("xl/worksheets/_rels/" + file + ".rels", rels)`.

rId discipline: hyperlink rIds are **local to each worksheet's `.rels`** — do not mix them with the
workbook-level rId space. Assign them in `worksheet.ts` in one pass so the `<hyperlink r:id>` and the
`.rels` `Id` always match.

## 3. Read path

- **`src/xlsx/read/package.ts`** — expose a helper to read+parse an arbitrary part's rels (reuse the
  existing `parseRels` + `relsPathFor`): `partRels(partPath): Map<rId, {type, target, mode}>`. For a
  worksheet at `xl/worksheets/sheetN.xml`, its rels live at `xl/worksheets/_rels/sheetN.xml.rels`.
- **`src/xlsx/read/worksheet.ts`** — parse the `<hyperlinks>` block into `{ ref, rId, tooltip }[]`, and
  accept the resolved worksheet-rels map in the read context. After cells are placed, for each hyperlink:
  resolve `rId → target`, take the cell's current string value as the display `text`, and set the cell to
  `{ text, hyperlink: target, tooltip? }` (`CellType.Hyperlink`). A hyperlink `ref` that is a range applies
  to its top-left cell (defer multi-cell).
- **`src/xlsx/read-workbook.ts`** — for each sheet, load its `partRels` and pass them into
  `parseWorksheet`'s context alongside styles/sharedStrings.

## 4. Compat + docs

- **Compat facade** — no code change needed: hyperlink cells already pass through as `{ text, hyperlink }`
  and `cell.type` maps to `ValueType.Hyperlink` (5). With real relationships persisted, the value now
  survives a round-trip, so the facade is automatically correct.
- **`docs/MIGRATION.md` / README** — remove the "hyperlink cells round-trip as their display text" Known
  limitation; if internal-location links remain deferred, note *that* narrower gap instead.
- **Parity corpus** (`tests/compat/parity.test.ts`) — the session-4 corpus currently **excludes** hyperlink
  values (they were a known divergence). Add a hyperlink snippet now that they round-trip.

## 5. Test & verification plan

Unit:
- **worksheet writer:** a hyperlink cell emits its display string in `sheetData` AND a `<hyperlinks>` entry
  with a matching `r:id`; `renderWorksheetRels` emits the external relationship; tooltip present/absent.
- **worksheet-rels:** golden string for the `.rels` doc.
- **worksheet reader:** a `<hyperlinks>` block + a rels map reconstructs `{ text, hyperlink, tooltip? }`.

Integration:
- **Round-trip** (`read-roundtrip`): a fixture with 2–3 hyperlinks (http, mailto, one with a tooltip)
  writes → the exact part set now includes `xl/worksheets/_rels/sheet1.xml.rels` → reads back to equal
  `{ text, hyperlink, tooltip? }` values with `CellType.Hyperlink`.
- **Foreign file** (`read-foreign`, exceljs devDep-guarded): build a hyperlink with ExcelJS → `writeBuffer`
  → our reader recovers `{ text, hyperlink }`; and our hyperlink buffer re-opens in ExcelJS with the link
  intact.
- **Excel/LibreOffice sanity:** the fixture opens as a working, clickable link (manual or the optional
  headless-LibreOffice CI check).

## 6. Deferred to later minors

Internal `location=` hyperlinks (no external rel), drawing/image-anchored hyperlinks, per-range hyperlink
`ref` spanning many cells, and `mailto:` header niceties (subject/body) beyond the bare URL.

## Appendix — session breakdown & model recommendations

1. **Write path** — `constants.ts` rel type, `worksheet.ts` `<hyperlinks>` + per-sheet rId collection,
   new `worksheet-rels.ts`, `write-workbook.ts` wiring (`renderWorksheet` return type change) + tests.
   Cross-part rId wiring is the fiddly bit. **Opus 4.8.**
2. **Read path** — `read/package.ts` `partRels`, `read/worksheet.ts` `<hyperlinks>` parse +
   reconstruction, `read-workbook.ts` per-sheet rels wiring + tests. **Sonnet 5 (Opus if desired).**
3. **Round-trip + compat + docs** — round-trip & foreign hyperlink tests, add the compat parity snippet,
   strip the Known-limitation from README/`MIGRATION.md`. **Sonnet 5.**
4. **Release 0.4.0** — `minor` changeset (note the read behaviour change), version bump, `npm pack` audit
   (both entries still ship), dry-run, publish (owner-only). Mirror `RELEASE-0.2-RUNBOOK.md`. **Opus 4.8.**

Parallelism: sessions 1 (write) and 2 (read) touch disjoint files and *could* run in parallel worktrees,
but the reader is far easier to build and test against real writer output — recommend **sequential** (1 → 2
→ 3 → 4).

---

*Grounded in the shipped `@alosha/xlsx@0.3.1` (writer emits hyperlinks as plain text at
`parts/worksheet.ts`; reader has no hyperlink handling; no `REL_TYPE.hyperlink` yet) and the OOXML
CT_Worksheet / hyperlink-relationship shape. Clean-room.*
