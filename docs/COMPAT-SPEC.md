# Compat Shim Build Spec — `@alosha/xlsx/compat` (the ExcelJS drop-in)

> The migration wedge. A subpath export that mirrors the ExcelJS public surface over the Stage 1.2 model
> + Stage 1.3/Stage 2 serializer, so an ExcelJS user changes **one import** and their code runs. This is
> the roadmap's "Stage 3 — migration shim"; it ships as the **0.3.0** minor. Grounded in
> [`API-AUDIT.md`](./API-AUDIT.md) (the ExcelJS surface) and the shipped `@alosha/xlsx@0.2.0` model.
>
> Prereqs (all published in 0.2.0): the model (`Workbook`/`Worksheet`/`Row`/`Column`/`Cell`), the
> `workbook.xlsx.writeBuffer/writeFile/load/readFile/read` accessor, the free-function core. Conventions
> unchanged: TS strict, ESM+CJS dual build, fflate the only runtime dep, `exceljs` stays a devDep (used
> only by the parity tests).

## 0. Scope

**In:** a `@alosha/xlsx/compat` entry whose **default export** is shaped like the ExcelJS namespace
(`{ Workbook, ... }` + `ValueType` / `Enums`), such that the common ExcelJS program — create workbook, add
worksheets/rows/cells, set values + styles + number formats + merges + column defs, then
`workbook.xlsx.writeBuffer/writeFile` and `load/readFile` — produces an equivalent `.xlsx` and reads back
equivalent values **without code changes beyond the import**. Emulate the two behaviours our model
deliberately diverged on (below). Everything the shim can't do yet throws a clear
`UnsupportedFeatureError` naming the feature.

**Out (defer):** the streaming `stream.xlsx.WorkbookWriter/Reader`, `workbook.csv.*`, images/media,
tables, pivot tables, conditional formatting, data validation, comments, defined names, worksheet
events, and protection — i.e. the same set Stages 1–2 deferred. These throw with a pointer to the tracking
issue rather than silently no-op'ing.

**Definition of done:** a **parity corpus** of real ExcelJS snippets (§8) runs against both `exceljs` and
`@alosha/xlsx/compat`; for each, the two outputs load back to the same sampled cell values, types, number
formats, merges, and sheet structure. `cell.type` returns ExcelJS `ValueType` numbers in compat mode.
Style precedence matches ExcelJS (whole-object per category). Typecheck/lint/tests green; the subpath
builds and type-resolves (`import ExcelJS from "@alosha/xlsx/compat"` and named imports both work).

## 1. Package & module layout

```
src/
  compat/
    index.ts        // default export = ExcelJS-shaped namespace; named re-exports; ValueType/Enums
    value-type.ts   // CellType <-> ExcelJS ValueType maps (the numbering differs — see §3)
    workbook.ts     // CompatWorkbook: the ExcelJS-faithful facade over the model (see §2/§4/§5)
    style-compat.ts // whole-object style precedence (ExcelJS semantics) — §4
tests/
  compat/
    parity.test.ts       // the real proof: run snippets vs exceljs, diff loaded output (devDep-guarded)
    value-type.test.ts
    style-compat.test.ts
```

`package.json` gains a subpath export; `tsup.config.ts` gains the second entry:

```jsonc
// package.json "exports"
"./compat": { "types": "./dist/compat.d.ts", "import": "./dist/compat.js", "require": "./dist/compat.cjs" }
```
```ts
// tsup.config.ts
entry: ["src/index.ts", "src/compat/index.ts"],
```

Keep the shim a **thin adapter**: the model was designed for ExcelJS parity (API-AUDIT §5 "keep"), so most
of the surface is re-export, not reimplementation. The shim only bridges the documented gaps.

## 2. The compatibility matrix

**Already compatible (re-export as-is):** `new Workbook()`, `addWorksheet(name, opts)`, `getWorksheet`,
`worksheets`, workbook metadata (`creator`/`created`/`modified`/`title`/…), `worksheet.getRow/getCell/
addRow/addRows/mergeCells/unMergeCells/columns/eachRow`, `row.getCell/values/eachCell`, `column.width/
key/header/hidden`, cell `value`/`text`/`numFmt`/`font`/`fill`/`border`/`alignment`, and the
`workbook.xlsx.writeBuffer/writeFile/load/readFile/read` accessor. Value shapes (`{formula,result}`,
`{richText}`, `{text,hyperlink}`, `{error}`) already match ExcelJS.

**Adapt (the shim's real work):**
- `cell.type` / `cell.effectiveType` → ExcelJS `ValueType` numbers (§3).
- `Cell` style **precedence** on assignment → ExcelJS whole-object-per-category (§4).
- `Enums.ValueType`, `Enums.FormulaType`, `Enums.RelationshipType` exports; `Cell.Types` alias.
- `row.values` 1-based sparse array with `[0]` empty — confirm our offset behaviour matches ExcelJS and
  normalise in the facade if not.

**Defer (throw `UnsupportedFeatureError`):** `stream.xlsx.*`, `workbook.csv.*`, `addImage`/`getImage`,
`addTable`, conditional formatting, data validation, `addComment`, defined names, events, protection.

## 3. `value-type.ts` — the enum bridge

Our `CellType` and ExcelJS `ValueType` share names but **not numbers**:

| Name | @alosha `CellType` | ExcelJS `ValueType` |
|------|-------------------:|--------------------:|
| Null | 0 | 0 |
| Number | 1 | 2 |
| String | 2 | 3 |
| Boolean | 3 | 9 |
| Date | 4 | 4 |
| Error | 5 | 10 |
| RichText | 6 | 8 |
| Hyperlink | 7 | 5 |
| Formula | 8 | 6 |
| Merge | 9 | 1 |
| SharedString | 10 | 7 |

Export the ExcelJS `ValueType` object (name→number as the right column) and `toExcelJSValueType(cellType)`
/ `fromExcelJSValueType(v)` maps. In compat mode `cell.type` returns `toExcelJSValueType(model.type)`.

## 4. `style-compat.ts` — the precedence divergence (subtle, must match)

API-AUDIT §4 flags this: our `resolveCellStyle` does **deep field-level merge** (a lone `cell.font.bold`
coexists with inherited `font.name/size`), while ExcelJS replaces the **whole** category (setting
`cell.font` discards inherited font fields). Migrated files must render identically, so the compat cell's
style setters apply **ExcelJS whole-object semantics**: assigning `cell.font`/`fill`/`border`/`alignment`/
`protection` **replaces** that category outright rather than merging. Provide `applyExcelJsStyle(cell,
patch)` used by the facade's setters; the underlying model store is unchanged (the serializer still dedups
final styles). Round-trip note: on **read**, styles already come back fully-resolved per cell, so no
divergence there — this only affects the authoring path.

## 5. `workbook.ts` — the facade contract

Define the observable contract the parity tests enforce; the implementer picks the mechanism (subclass +
per-cell facade, or a `Proxy`, or prototype augmentation scoped to compat-constructed workbooks — must NOT
leak ExcelJS semantics into the core `@alosha/xlsx` entry). The contract:
- `cell.type`/`effectiveType` → ExcelJS `ValueType`.
- style setters → §4 whole-object precedence.
- `worksheet.addRow(array | object)`, `row.values` 1-based with empty `[0]`, `worksheet.columns = [{header,
  key, width}]` all behave as ExcelJS documents.
- I/O is the existing accessor, unchanged.

> **Decision D1 — faithfulness of `cell.type`.** Full emulation (compat cells always report ExcelJS
> numbers) vs a lighter "export `ValueType` + document the difference". *Recommend full emulation* — it's
> the difference between "drop-in" and "mostly drop-in", and it's cheap via the value-type map. Lock before
> the facade session.
> **Decision D2 — style precedence** is NOT optional: the shim must emulate ExcelJS whole-object precedence
> (§4), else migrated spreadsheets restyle silently. No decision — just build it.

## 6. Unsupported surface

Every deferred method exists but throws `new UnsupportedFeatureError("<feature>", "not yet supported in
@alosha/xlsx/compat; see <tracking-issue>")` — so migrating code fails loud and precise at the exact call,
never silently produces a wrong file. List them in the README migration table with status (supported /
planned / won't-port).

## 7. Docs

A "Migrating from ExcelJS" README section (or `docs/MIGRATION.md`) with the one-line import change, the
supported/deferred matrix, the two behavioural notes (D1 `cell.type` numbers, D2 style precedence), and a
before/after snippet. Link it from the main README.

## 8. Test & verification plan

Unit:
- **value-type:** every row of the §3 table round-trips; `cell.type` on a compat cell of each kind returns
  the ExcelJS number.
- **style-compat:** setting `cell.font = {bold:true}` after an inherited font **replaces** (ExcelJS) rather
  than merges (core) — asserted against both behaviours.

Integration — the real proof (`parity.test.ts`, guarded on the `exceljs` devDep, `ctx.skip()` if absent):
1. A **corpus** of ~8–12 representative ExcelJS snippets (from the ExcelJS README / common Stack-Overflow
   patterns): values of every type, styled headers, number formats, a merge, multi-sheet, column keys +
   `addRow(object)`, a formula with cached result, `date1904`.
2. Run each snippet twice — once with `import ExcelJS from "exceljs"`, once with the same code and `import
   ExcelJS from "@alosha/xlsx/compat"` — producing two buffers.
3. Load **both** buffers with real `exceljs` and assert equality of sampled cell values, `cell.type`,
   `numFmt`, merged ranges, sheet names/order. (Cross-loading with exceljs normalises away trivially
   different-but-equivalent XML.)
4. Also load the compat buffer with our own `readWorkbookBuffer` and assert the same — proves the shim's
   output is valid for our reader too.

Keep snippets deterministic (fixed dates) so failures are legible.

## 9. Deferred to later minors

Streaming shim, CSV, images/tables/pivots/conditional-formatting/data-validation/comments/defined-names/
events/protection, and a `Cell.Types`-style deep legacy-alias surface. Each is a follow-on minor gated by
real migration demand.

## Appendix — session breakdown & model recommendations

1. **Enum bridge + subpath wiring** — `value-type.ts`, `compat/index.ts` skeleton, `package.json`
   `exports` + `tsup` entry, a smoke test that `import ExcelJS from "@alosha/xlsx/compat"` resolves in ESM
   and CJS. Mechanical. **Sonnet 5.**
2. **Style-precedence adapter** — `style-compat.ts` (whole-object semantics) + unit tests. Fiddly, must
   mirror ExcelJS exactly. **Opus 4.8.**
3. **Facade + surface** — `compat/workbook.ts` (cell.type emulation, style setters, row/column/value
   parity), unsupported-method stubs. Judgment on the facade mechanism (Decision D1). **Opus 4.8.**
4. **Parity corpus + migration docs** — `parity.test.ts` vs exceljs, `MIGRATION.md`/README section. Where
   real-world gaps surface. **Sonnet 5 (Opus if the corpus exposes surprises).**
5. **Release 0.3.0** — changeset (minor), version bump, `npm pack` audit (confirm `dist/compat.*` ships),
   dry-run, publish (owner-only). Mirror `RELEASE-0.2-RUNBOOK.md`. **Opus 4.8.**

---

*Grounded in `API-AUDIT.md` (ExcelJS 4.4.0 surface + keep/drop/redesign notes) and the published
`@alosha/xlsx@0.2.0`. Clean-room; the shim mirrors ExcelJS's public API shape, not its source.*
