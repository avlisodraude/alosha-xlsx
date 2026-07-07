# Data Validation Build Spec — dropdowns & cell constraints (model + write + read + compat)

> The next Stage 2 write feature. Add cell **data validations** — list dropdowns and whole/decimal/date/
> time/text-length/custom constraints — to the model, serialize them, read them back, and expose them
> through the compat facade (turning today's `UnsupportedFeatureError` stubs into a real API). Ships as
> **0.5.0**.
>
> Prereqs (published in 0.4.0): the model (`src/model/*`), writer (`src/xlsx/parts/*` +
> `write-workbook.ts`), reader (`src/xlsx/read/*`), and the compat facade (which currently throws for
> `cell.dataValidation` and `worksheet.dataValidations`). Conventions unchanged: TS strict, ESM+CJS, fflate
> the only runtime dep, `exceljs` a devDep for the parity/foreign tests.
>
> **Difference from the hyperlink feature:** data validations are a **new model concept**, not just a
> serialization detail — so this starts with a model session (add the type + worksheet storage + a
> `cell.dataValidation` accessor + `WorksheetModel` round-trip) before touching the writer.

## 0. Scope

**In:** the common ExcelJS-shaped data-validation surface —
- `type: "list"` (dropdowns), with `formulae[0]` being either an **inline** list (`'"Yes,No,Maybe"'`) or a
  **same-sheet** range (`"$E$1:$E$5"`);
- `type: "whole" | "decimal" | "date" | "time" | "textLength"` with an `operator`
  (`between`/`notBetween`/`equal`/`notEqual`/`greaterThan`/`lessThan`/`greaterThanOrEqual`/
  `lessThanOrEqual`) and `formulae`;
- `type: "custom"` (a boolean formula in `formulae[0]`);
- the flags/messages: `allowBlank`, `showInputMessage`, `showErrorMessage`, `errorStyle`
  (`stop`/`warning`/`information`), `errorTitle`, `error`, `promptTitle`, `prompt`, and the list
  `showDropDown` (with its OOXML inversion handled — see §4).

Applied per cell (`cell.dataValidation`) and per range (`worksheet.dataValidations.add(sqref, dv)`), where
`sqref` may be a single address, a range, or space-separated ranges. Write + read + compat + docs.

**Out (defer):** the **x14 / `<extLst>`** data-validation form Excel uses when a list references a range on
**another sheet** (cross-sheet list refs) — same-sheet ranges and inline lists use the plain
`<dataValidations>` form we implement here; note the cross-sheet case as still-deferred in `MIGRATION.md`.
Also defer: `imeMode`, and data validation attached to defined names.

**Definition of done:** a validation set via `cell.dataValidation` or `worksheet.dataValidations` (a) writes
a `<dataValidations>` block in the sheet (correct CT_Worksheet position), (b) re-opens in Excel/LibreOffice
as a working dropdown/constraint, (c) round-trips through our own reader to an equal model, and (d) an
ExcelJS-written validation parses with our reader and vice-versa. Compat stubs replaced by a real
pass-through. Typecheck/lint/tests green.

## 1. Model (`src/model/data-validation.ts` + worksheet/cell wiring)

```ts
export type DataValidationType =
  | "list" | "whole" | "decimal" | "date" | "time" | "textLength" | "custom";
export type DataValidationOperator =
  | "between" | "notBetween" | "equal" | "notEqual"
  | "greaterThan" | "lessThan" | "greaterThanOrEqual" | "lessThanOrEqual";
export type DataValidationErrorStyle = "stop" | "warning" | "information";

/** Mirrors ExcelJS's `cell.dataValidation` shape (so the compat facade is a pass-through). */
export interface DataValidation {
  type: DataValidationType;
  operator?: DataValidationOperator;
  allowBlank?: boolean;
  showInputMessage?: boolean;
  showErrorMessage?: boolean;
  /** `[formula1]` or `[formula1, formula2]`. For `list`: an inline `'"a,b,c"'` or a range `"$E$1:$E$5"`. */
  formulae?: (string | number)[];
  showDropDown?: boolean;         // list only; NOTE the OOXML attribute is inverted (see §4)
  errorStyle?: DataValidationErrorStyle;
  errorTitle?: string;
  error?: string;
  promptTitle?: string;
  prompt?: string;
}
```

Storage — **per worksheet, keyed by `sqref`** (a validation is inherently range-based). Add to `Worksheet`:
- a private `Map<string /* sqref */, DataValidation>` and a small manager exposed as
  `worksheet.dataValidations` with `add(sqref, dv)`, `find(address): DataValidation | undefined`,
  `remove(sqref)`, and a `model` getter (`{ sqref, ...dv }[]`), mirroring ExcelJS's `DataValidations`.
- `cell.dataValidation` get/set on `Cell`: the setter calls `add(cell.address, dv)`; the getter calls
  `find(cell.address)`. `Cell` reaches the manager through its `WorksheetHost` (extend `WorksheetHost` with
  the manager, like `getRow` was added for `Column#header`).
- `WorksheetModel` gains `dataValidations?: DataValidationModel[]` (`{ sqref } & DataValidation`); the
  worksheet `model` getter/setter serialize the map to/from it (so it round-trips like `merges`).

`find(address)` resolves a single address against each stored `sqref` (which may be a multi-range) — reuse
`Range`/`decodeAddress` from `src/address`.

## 2. Serialization position (all three: writer, reader, model order)

CT_Worksheet child order (ECMA-376): `… sheetData, …, mergeCells, phoneticPr, conditionalFormatting,`
**`dataValidations, hyperlinks,`** `printOptions, …`. So the writer emits `<dataValidations>` **after
`<mergeCells>` and before `<hyperlinks>`** — i.e. between the existing `renderMergeCells` and
`renderHyperlinks` calls in `parts/worksheet.ts`. (OOXML-NOTES §note lists these approximately; this is the
authoritative order.)

## 3. Write (`src/xlsx/parts/worksheet.ts`)

Add `renderDataValidations(w, validations)` emitting:
```xml
<dataValidations count="1">
  <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1:A10">
    <formula1>"Yes,No,Maybe"</formula1>
  </dataValidation>
</dataValidations>
```
- `type` omitted when `"none"`; `operator` only for the numeric/length types; `<formula1>`/`<formula2>`
  from `formulae[0]`/`[1]` (emit `formula2` only for `between`/`notBetween`). Numbers stringify; inline
  list strings keep their surrounding quotes as authored.
- Flags map straight through (`allowBlank`, `showInputMessage`, `showErrorMessage`, `errorStyle`,
  `errorTitle`, `error`, `promptTitle`, `prompt`) — omit when unset (booleans emit `1` when true).
- **`showDropDown` inversion (§4):** the OOXML attribute is `showDropDown="1"` to **hide** the dropdown.
  So emit `showDropDown="1"` only when the model's `showDropDown === false`; omit otherwise (default shows
  the dropdown). Document this clearly — it's the classic OOXML footgun.
- No relationships or extra parts — `<dataValidations>` lives entirely in the sheet.

`renderWorksheet` reads `ws.model.dataValidations` (or the live manager) and calls
`renderDataValidations` between `renderMergeCells` and `renderHyperlinks`. No change to `RenderedWorksheet`
(no `.rels`).

## 4. Read (`src/xlsx/read/worksheet.ts`)

Parse the `<dataValidations>` child: for each `<dataValidation>`, read the attributes + `<formula1>`/
`<formula2>` text into a `DataValidation`, invert `showDropDown` back to the model's sense, and register it
via the worksheet's manager keyed by `sqref`. Collected during the child walk (like merges/hyperlinks) and
applied after — order-independent. Malformed entries are skipped, never thrown.

## 5. Compat (`src/compat/workbook.ts`)

Replace the two `UnsupportedFeatureError` throwers with real pass-throughs:
- **`cell.dataValidation`** (get/set) — the core `Cell` accessor already matches ExcelJS's shape, so the
  proxy just stops intercepting it (delete the `set`-side throw and let it reflect through).
- **`worksheet.dataValidations`** — return the wrapped manager (ExcelJS's `DataValidations` has
  `add`/`find`/`remove`/`model`); our manager mirrors it, so it passes through.

Keep the deferred-x14 note: a cross-sheet list ref still isn't supported.

## 6. Docs

Remove "No data validation … yet" from the README Known-limitations and the compat matrix in
`docs/MIGRATION.md`; add a `dataValidation` row to the supported matrix with the deferred **cross-sheet
list ref (x14)** caveat and the `showDropDown` inversion note.

## 7. Test & verification plan

Unit:
- **model:** `cell.dataValidation` round-trips through `find`/`add`; `WorksheetModel.dataValidations`
  serializes and restores; `find(address)` resolves within a multi-range `sqref`.
- **writer:** golden strings for a list (inline + range), a `between` decimal (two formulae), a `custom`,
  the message/error attrs, and the `showDropDown` inversion (model `false` → attr `"1"`).
- **reader:** each of the above parses back to the model, `showDropDown` re-inverted.

Integration:
- **Round-trip** (`read-roundtrip`): a fixture with a list dropdown, a numeric `between`, and a custom rule
  writes → reads back to an equal model.
- **Foreign** (`read-foreign`, exceljs-guarded): an ExcelJS-authored dropdown → our reader recovers it; our
  validation buffer re-opens in ExcelJS with the dropdown intact.
- **Compat parity** (`tests/compat/parity.test.ts`): add a `cell.dataValidation` list snippet.

## 8. Deferred to later minors

Cross-sheet list references (x14 `<extLst>`), `imeMode`, data validation on defined names, and
per-validation UI niceties beyond the standard attributes.

## Appendix — session breakdown & model recommendations

1. **Model** — `src/model/data-validation.ts` (types), `Worksheet` manager + `Map<sqref,dv>` +
   `WorksheetModel.dataValidations` round-trip, `cell.dataValidation` accessor, `WorksheetHost` wiring +
   tests. New public model surface. **Opus 4.8.**
2. **Write** — `parts/worksheet.ts` `renderDataValidations` (between mergeCells & hyperlinks; the
   `showDropDown` inversion) + golden tests. **Opus 4.8** (the inversion + attr matrix are error-prone).
3. **Read** — `read/worksheet.ts` `<dataValidations>` parse → manager, re-invert `showDropDown` + tests.
   **Sonnet 5.**
4. **Compat + round-trip + docs** — un-stub the two facade throwers, round-trip & foreign tests, parity
   snippet, strip README/`MIGRATION.md` limitation. **Sonnet 5.**
5. **Release 0.5.0** — `minor` changeset, version bump, `npm pack` audit (both entries), dry-run, publish
   (owner-only). Mirror the prior release runbooks. **Opus 4.8.**

Parallelism: sequential — the model (1) gates everything; write (2) and read (3) build on it and are
easier verified in order.

---

*Grounded in `@alosha/xlsx@0.4.0` (no data-validation model yet; compat facade throws at
`workbook.ts:199` cell + `:276` worksheet; `WorksheetModel` round-trips sheet state) and the ECMA-376
CT_Worksheet / CT_DataValidation shape. Clean-room.*
