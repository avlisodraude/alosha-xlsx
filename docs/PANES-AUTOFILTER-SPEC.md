# Panes + Auto-filter Build Spec (Track A → 0.6.0)

> Freeze/split panes and a sheet auto-filter range. Small, in-sheet, **no new parts**. Companion to
> [`ADVANCED-ROADMAP.md`](./ADVANCED-ROADMAP.md) Track A. Session **A1 (model + writer) is already merged**
> (`703000c`); this spec pins the exact shapes so **A2 (reader)** is a clean inverse and **A3 (compat +
> release 0.6.0)** matches the ExcelJS surface.
>
> Conventions unchanged: TS strict, ESM, `Uint8Array`/`TextDecoder`, no `console.*`, fflate the only
> runtime dep, clean-room OOXML.

## 0. Scope

**In:** a single `<sheetView>` per sheet carrying a `<pane>` for `state: "frozen" | "split"`, and a
sheet-level `<autoFilter ref>`. Round-trips through the model on both write and read.

**Out (defer):** multiple views per sheet, `<selection>`, per-pane active-cell tracking, split-window
zoom, and auto-filter **filter criteria / sort state** (`<filterColumn>`, `<sortState>`) — we persist the
`ref` range only, matching where ExcelJS's basic `autoFilter` sits.

**Definition of done:** a write→read round-trip of a sheet with `views = [{ state: "frozen", xSplit, ySplit,
topLeftCell }]` and `autoFilter = "A1:D1"` reproduces both exactly; a foreign ExcelJS file with a frozen
pane + auto-filter parses with `views[0].state === "frozen"` and the `autoFilter` range intact; malformed
input never throws a bare `TypeError` (typed `AloshaXlsxError`).

## 1. Model (shipped in A1 — reference)

`src/model/worksheet.ts`:

```ts
export interface WorksheetView {
  state?: "normal" | "frozen" | "split";
  showGridLines?: boolean;
  zoomScale?: number;
  xSplit?: number;        // CT_Pane
  ySplit?: number;
  topLeftCell?: string;
  activePane?: "bottomRight" | "topRight" | "bottomLeft" | "topLeft";
}
// Worksheet also gains:  autoFilter?: string   // an A1 range, e.g. "A1:D1"
```

Both `views` and `autoFilter` are carried on `WorksheetModel` and survive `model` get/set (A1 wired the
snapshot round-trip).

## 2. Write output (shipped in A1 — the goldens the reader inverts)

`renderSheetViews` (only `views[0]` backs the single `<sheetView>`; `showGridLines="0"` emitted only when
the property is explicitly `false`):

```xml
<!-- default (no view / normal) -->
<sheetViews><sheetView workbookViewId="0"/></sheetViews>

<!-- frozen -->
<sheetViews><sheetView workbookViewId="0">
  <pane xSplit="1" ySplit="1" topLeftCell="B2" state="frozen"/>
</sheetView></sheetViews>

<!-- split, with activePane -->
<pane xSplit="2000" ySplit="3000" topLeftCell="C4" activePane="bottomRight" state="split"/>
```

`renderAutoFilter` — emitted **after `</sheetData>` and before `<mergeCells>`** (CT_Worksheet order):

```xml
<autoFilter ref="A1:D1"/>
```

Attribute emit order on `<pane>`: `xSplit, ySplit, topLeftCell, activePane, state` (omit-when-undefined).
`XmlWriter` drops undefined attrs, so a frozen pane with only `ySplit` emits `<pane ySplit="1"
state="frozen"/>`.

## 3. Read plan — session A2 (`src/xlsx/read/worksheet.ts`)

Invert §2 while parsing the sheet, using the existing `xml-reader` cursor already used for merges/cols/
hyperlinks:

- **`<sheetViews><sheetView>`**: read `showGridLines` (`"0"` → `view.showGridLines = false`). If it holds a
  `<pane>` child, build `views[0]` = `{ state: pane.state, xSplit?: +pane.xSplit, ySplit?: +pane.ySplit,
  topLeftCell?: pane.topLeftCell, activePane?: pane.activePane }` — coerce the numeric attrs, drop absent
  ones (don't materialise `xSplit: 0`). No `<pane>` and no non-default attrs → leave `views` empty (a bare
  `<sheetView workbookViewId="0"/>` carries no model state worth keeping).
- **`<autoFilter ref>`**: `ws.autoFilter = ref`. Ignore any `<filterColumn>`/`<sortState>` children (out of
  scope) but don't choke on them.
- Assign onto the same `Worksheet` the grid reader already populates; order in the file is guaranteed
  (pane inside sheetView near the top, autoFilter after sheetData) so no buffering is needed.

**Reader tests:** golden-fragment inversions of the three §2 pane cases + the auto-filter case; a
malformed `<pane state="frozen" xSplit="abc"/>` yields a typed error, not `NaN` silently. Add a round-trip
case to `tests/xlsx/read-roundtrip.test.ts` covering frozen pane + auto-filter together.

## 4. Compat plan — session A3 (`src/compat/workbook.ts`)

ExcelJS exposes `worksheet.views` (array) and `worksheet.autoFilter` (string or `{from,to}`). Surface both
through the shim as thin pass-throughs to the model:

- `worksheet.views` get/set ⇒ the model `views` array (ExcelJS's frozen-view fields — `state`, `xSplit`,
  `ySplit`, `topLeftCell`, `activePane`, `showGridLines` — are already our field names, so it's structural).
- `worksheet.autoFilter` get/set ⇒ the model `autoFilter` string; accept ExcelJS's `{ from, to }` object
  form and normalise to an `"A1:D1"` range on set.

**Compat tests:** parity assertion that setting `views`/`autoFilter` via the shim then `writeBuffer` +
reload reproduces them; a foreign-file test (build with the `exceljs` devDep, `writeBuffer`, load ours,
assert `views[0].state` + `autoFilter`).

## 5. Release — session A3

Changeset **minor → 0.6.0**; add "freeze/split panes + auto-filter" to CHANGELOG + the README feature
list. Run `pnpm typecheck && pnpm lint && pnpm test`, commit, push. No `npm publish` (manual, by design).

## Appendix — sessions & models

| # | Session | Model | Status |
|---|---------|-------|--------|
| A1 | Model + writer (`<pane>`, `<autoFilter>`) | Sonnet 5 | ✅ merged `703000c` |
| A2 | Reader (invert §3) | **Sonnet 5** | next |
| A3 | Compat + tests + release 0.6.0 | **Fable 5** | after A2 |

*Grounded in the merged A1 code (`src/model/worksheet.ts`, `src/xlsx/parts/worksheet.ts`,
`tests/xlsx/parts/panes-autofilter.test.ts`). Clean-room; OOXML CT_Pane/CT_AutoFilter shapes only.*
