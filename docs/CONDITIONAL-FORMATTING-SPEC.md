# Conditional Formatting Build Spec (Track B → 0.7.0)

> The largest in-sheet feature: a per-worksheet list of `<conditionalFormatting>` blocks, whose
> differential styles live in a **new `<dxfs>` dedup pool** in `styles.xml`. Companion to
> [`ADVANCED-ROADMAP.md`](./ADVANCED-ROADMAP.md) Track B. Baseline: **v0.6.0** (panes/auto-filter shipped,
> merge `ba288ec`). Conventions unchanged: TS strict, ESM, `Uint8Array`/`TextDecoder`, no `console.*`,
> fflate the only runtime dep, clean-room OOXML.
>
> **Session refinement:** the roadmap budgeted 3 sessions; this spec splits into **4** (model / writer+dxf
> pool / reader / compat+release) because the rule-type breadth plus the `styles.ts` dxf pool make a
> combined model+writer session too large to land cleanly. Still one feature, one minor release.

## 0. Scope

**In:** `worksheet.addConditionalFormatting({ ref, rules })` (ExcelJS shape) with these `cfRule` types —
- **dxf-bearing** (carry a differential `style` → a `<dxf>`): `cellIs`, `expression`, `containsText`
  family (`containsText`/`notContainsText`/`beginsWith`/`endsWith`), `duplicateValues`/`uniqueValues`,
  `top10`, `aboveAverage`/`belowAverage`.
- **self-visualising** (no dxf): `colorScale` (2–3 stops), `dataBar` (basic), `iconSet` (3/4/5-icon presets).

**Out (defer to a later minor):** `timePeriod` rules; the x14 `<extLst>` dataBar extensions (gradient fill,
axis, negative-value fill) — we emit the ECMA-376 baseline `<dataBar>` only; custom `<iconSet>` icons via
extLst; `stopIfTrue` beyond a boolean pass-through; and GUID-keyed extLst rules generally.

**Definition of done:** (a) a write→read round-trip reproduces every in-scope rule type — including its
`sqref`, `priority` order, operator/formulae/cfvo/color/icon fields, and the resolved differential
`style`; (b) an ExcelJS-authored file with a `cellIs` + a `colorScale` rule parses with both intact
(`exceljs` devDep-guarded); (c) malformed input (bad `sqref`, unknown `type`, non-numeric `priority`)
throws a typed `AloshaXlsxError`, never a bare `TypeError`.

## 1. Session B1 — model (`src/model/conditional-formatting.ts`)

Mirror ExcelJS's `cell.dataValidation`-style shape so the compat layer stays a pass-through. A
discriminated union keyed by `type`:

```ts
export type CfValueObject = {              // <cfvo>
  type: "num" | "percent" | "max" | "min" | "percentile" | "formula";
  value?: number | string;                // val="" (omit for min/max)
};
export type CfColor = string;              // ARGB "FFF8696B"

export interface CfRuleBase { priority?: number; stopIfTrue?: boolean; }
export interface CfCellIs   extends CfRuleBase { type: "cellIs"; operator: CfOperator; formulae: (string|number)[]; style?: Partial<Style>; }
export interface CfExpression extends CfRuleBase { type: "expression"; formulae: [string]; style?: Partial<Style>; }
export interface CfText      extends CfRuleBase { type: "containsText"|"notContainsText"|"beginsWith"|"endsWith"; text: string; style?: Partial<Style>; }
export interface CfTop10     extends CfRuleBase { type: "top10"; rank: number; percent?: boolean; bottom?: boolean; style?: Partial<Style>; }
export interface CfAverage   extends CfRuleBase { type: "aboveAverage"|"belowAverage"; style?: Partial<Style>; }
export interface CfDuplicate extends CfRuleBase { type: "duplicateValues"|"uniqueValues"; style?: Partial<Style>; }
export interface CfColorScale extends CfRuleBase { type: "colorScale"; cfvo: CfValueObject[]; color: CfColor[]; }
export interface CfDataBar   extends CfRuleBase { type: "dataBar"; cfvo: [CfValueObject, CfValueObject]; color: CfColor; }
export interface CfIconSet   extends CfRuleBase { type: "iconSet"; iconSet: string; cfvo: CfValueObject[]; }
export type ConditionalRule = CfCellIs | CfExpression | CfText | CfTop10 | CfAverage | CfDuplicate | CfColorScale | CfDataBar | CfIconSet;

export interface ConditionalFormatting { ref: string; rules: ConditionalRule[]; }
export type CfOperator = "lessThan"|"lessThanOrEqual"|"equal"|"notEqual"|"greaterThanOrEqual"|"greaterThan"|"between"|"notBetween"|"containsText"|"notContains"|"beginsWith"|"endsWith";
```

Worksheet manager (parallel to `DataValidations`, [[DATA-VALIDATION-SPEC]] §1 precedent): `Worksheet` owns
a private `ConditionalFormatting[]`; public `addConditionalFormatting(cf)` / `removeConditionalFormatting(ref)`
/ getter `conditionalFormattings`; round-trips via a new optional `WorksheetModel.conditionalFormattings`.
Validate `ref` through the existing `Range` decoder (typed error on garbage). No `WorksheetHost` change
needed (this is sheet-level, not cell-routed). Export from the model barrel + `src/index.ts`.

**Tests:** manager add/remove/get, model snapshot round-trip, `ref` validation error, union type-narrowing
smoke.

## 2. Session B2 — writer + dxf pool

### 2a. The dxf pool (`src/xlsx/styles.ts`)

Today `toXml` emits a placeholder `<dxfs count="0"/>` (line ~198). Replace with a real pool:

- Add `addDxf(style: Partial<Style> | undefined): number` to the `Styles` interface + a `dxfs` `InternPool`
  **with no baseline seed** (dxf ids start at 0, unlike fonts/fills/borders which reserve defaults).
- Render a `<dxf>` in **CT_Dxf child order: `font, numFmt, fill, alignment, border, protection`** (only the
  present sub-objects). Reuse `renderFont` as-is. **Footgun:** a dxf `<fill>` puts its colour in
  `<patternFill><bgColor rgb=…/>`, *not* `<fgColor>` like a cellXf solid fill — add a dxf-mode fill render
  (don't reuse `renderFill` verbatim). `renderBorder` may be reused (Excel tolerates a full border in a dxf).
- `toXml`: `<dxfs count="N">…</dxfs>` (omit entirely — or keep `count="0"` — when empty; match what the
  reader expects). Keep the CT_Stylesheet element order (`…cellXfs, cellStyles, dxfs, tableStyles`).

### 2b. `renderConditionalFormatting` (`src/xlsx/parts/worksheet.ts`)

Call it **between `renderMergeCells` and `renderDataValidations`** — CT_Worksheet order is
`…mergeCells, conditionalFormatting, dataValidations, hyperlinks…` (same rule the data-validation and
auto-filter sessions followed). The writer already has the shared pool via `ctx.styles`, so a dxf-bearing
rule does `const dxfId = ctx.styles.addDxf(rule.style)`.

Per worksheet, emit one `<conditionalFormatting sqref="REF">` per `ConditionalFormatting`, then its
`<cfRule>`s. **Priority**: assign a workbook-monotonic increasing integer when `rule.priority` is unset
(ExcelJS behaviour); required attribute, must be present on every rule. Shapes:

```xml
<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>5</formula></cfRule>
<!-- between/notBetween → two <formula> -->
<cfRule type="expression" dxfId="1" priority="2"><formula>$A1&gt;0</formula></cfRule>
<cfRule type="containsText" dxfId="2" priority="3" operator="containsText" text="foo">
  <formula>NOT(ISERROR(SEARCH("foo",A1)))</formula></cfRule>
<cfRule type="top10" dxfId="3" priority="4" rank="10" percent="1" bottom="1"/>
<cfRule type="aboveAverage" dxfId="4" priority="5"/>            <!-- belowAverage → aboveAverage="0" -->
<cfRule type="duplicateValues" dxfId="5" priority="6"/>          <!-- uniqueValues → type="uniqueValues" -->
<cfRule type="colorScale" priority="7"><colorScale>
  <cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule>
<cfRule type="dataBar" priority="8"><dataBar>
  <cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule>
<cfRule type="iconSet" priority="9"><iconSet iconSet="3TrafficLights1">
  <cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule>
```

Notes: `colorScale`/`dataBar`/`iconSet` carry **no `dxfId`**. `top10` percent/bottom emit `="1"` only when
true. `<cfvo val>` omitted for `min`/`max`. Formula text is XML-escaped by the existing `XmlWriter`.

**Tests:** a golden fragment per rule type, the dxf pool (a `cellIs` with a fill highlight → `<dxfs>` with
one bgColor dxf), the between-two-formulae case, and the `</mergeCells> … <conditionalFormatting> …
<dataValidations>` ordering (regex, like the auto-filter placement test).

## 3. Session B3 — reader

- **`src/xlsx/read/styles.ts`:** parse `<dxfs>` into a `Style[]` (invert §2a incl. `bgColor` → fill
  colour); expose `dxfFor(index): Style` on `StyleTable` (or a sibling `parseDxfs`). A missing `<dxfs>` →
  empty table.
- **`src/xlsx/read/worksheet.ts`:** collect `<conditionalFormatting>`/`<cfRule>` during the child walk
  (like `collectDataValidations`), applied after via `ws.addConditionalFormatting`. Invert every §2b shape;
  resolve `dxfId` through `dxfFor` into `rule.style`. Coerce numeric attrs (`priority`, `rank`) with the
  same typed-error guard the pane reader uses (no silent `NaN`). Unknown `type` / missing `sqref` → skip,
  never throw.

**Tests:** golden-fragment inversions per type; a round-trip case per family in
`tests/xlsx/read-roundtrip.test.ts`; a malformed `priority="x"` typed-error test.

## 4. Session B4 — compat + release 0.7.0

- **Compat (`src/compat/workbook.ts`):** remove `"addConditionalFormatting"` / `"removeConditionalFormatting"`
  from the UnsupportedFeature throw list (~lines 87–88) and let the worksheet proxy pass through to core
  (the model shape mirrors ExcelJS, so no bridge beyond that). Verify `rule.style` (font/fill/border/numFmt)
  round-trips through the existing `applyExcelJsStyle` semantics where relevant.
- **Tests:** a compat parity snippet (`cellIs` + `colorScale`, built via real `exceljs` and the facade,
  both reloaded with `exceljs` and diffed on rule type/operator/priority + the highlight fill); an
  `exceljs`-guarded foreign-file test.
- **Docs + release:** move conditional formatting to SUPPORTED in README + `MIGRATION.md` (note the deferred
  x14 dataBar/timePeriod bits); `pnpm changeset` minor → confirm **0.7.0**; `npm pack` ships both
  `dist/index.*` + `dist/compat.*`; `--dry-run` clean. Run `pnpm typecheck && pnpm lint && pnpm test`,
  commit, push. **No publish/tag/release — owner-only.**

## 5. Sessions & models

| # | Session | Model | Depends | Parallel? |
|---|---------|-------|---------|-----------|
| B1 | Model (`conditional-formatting.ts` + manager) | **Opus 4.8** | 0.6.0 | ‖ E (streaming), F0 |
| B2 | Writer + `<dxfs>` pool (`styles.ts`, `parts/worksheet.ts`) | **Opus 4.8** | B1 | ‖ E, F0 |
| B3 | Reader (`read/styles.ts`, `read/worksheet.ts`) | **Sonnet 5** (Opus optional) | B2 | ‖ E, F0 |
| B4 | Compat un-stub + tests + release 0.7.0 | **Sonnet 5** | B3 | ‖ E, F0 |

Sequential within the track. Track B shares no files with Track E (streaming) or F0 (plumbing), so those
can run in parallel worktrees. **Do not** run a second `renderWorksheet`-touching track (C/D) concurrently
with B2 — merge B first.

*Grounded in the v0.6.0 tree: `src/xlsx/styles.ts` (`InternPool`, `addStyle`, the `<dxfs count="0"/>`
placeholder, exported `renderFont/renderFill/renderBorder/renderNumFmt`), `WorksheetContext.styles`,
`src/xlsx/read/styles.ts` `StyleTable`, and `src/compat/workbook.ts` throw-list. Clean-room; CT_Dxf /
CT_ConditionalFormatting / CT_CfRule shapes only.*
