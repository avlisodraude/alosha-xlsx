# Data Validation Runbook — session by session

> Execution companion to [`DATA-VALIDATION-SPEC.md`](./DATA-VALIDATION-SPEC.md). Ready-to-paste prompts for
> adding cell data validations (dropdowns + constraints) across model → write → read → compat, shipping as
> **0.5.0**. All coding sessions run in **Claude Code** (write code, run
> `pnpm typecheck && pnpm lint && pnpm test && pnpm build`, commit+push).

## Overview

| # | Session | Env | Model | Depends on | Parallel? |
|---|---------|-----|-------|-----------|-----------|
| 0 | Baseline + branch | Code | Fable | 0.4.0 on main | — |
| 1 | Model (type + storage + accessors) | Code | **Opus 4.8** | 0 | no (gates all) |
| 2 | Write path | Code | **Opus 4.8** | 1 | no |
| 3 | Read path | Code | **Sonnet 5** | 1, 2 | no |
| 4 | Compat + round-trip + docs | Code | **Sonnet 5** | 3 | no |
| 5 | Release 0.5.0 | Code | **Opus 4.8** | 4 | no |

**Parallelism:** sequential — the model gates everything, and write/read are easier verified in order.

**Two footguns to watch (spec §2/§4):** `<dataValidations>` goes **between `<mergeCells>` and
`<hyperlinks>`** in CT_Worksheet order; and `showDropDown="1"` in OOXML **hides** the dropdown (inverted vs
the model's `showDropDown`).

---

## Prompts

### Session 0 — baseline + branch (Fable, Claude Code)

> In `~/WebstormProjects/alosha-xlsx` on `main`: confirm `@alosha/xlsx@0.4.0` is the published baseline,
> `git status` clean, `pnpm install`, then `pnpm typecheck && pnpm lint && pnpm test && pnpm build` all
> green. Create branch `feat/data-validation`. Report.

### Session 1 — model (Opus 4.8, Claude Code)

> Implement data-validation session 1 per `docs/DATA-VALIDATION-SPEC.md` §1, on `feat/data-validation`.
> Create `src/model/data-validation.ts` with the `DataValidation` type + enums (mirroring ExcelJS's
> `cell.dataValidation` shape). Add to `Worksheet`: a private `Map<sqref, DataValidation>` and a
> `dataValidations` manager (`add(sqref, dv)`, `find(address)`, `remove(sqref)`, `model` getter). Add
> `cell.dataValidation` get/set on `Cell` (set → `add(cell.address, dv)`, get → `find(cell.address)`),
> wiring the manager through `WorksheetHost` (as `getRow` was added for `Column#header`). Add
> `dataValidations?: ({ sqref } & DataValidation)[]` to `WorksheetModel` and serialize it in the worksheet
> `model` getter/setter (round-trips like `merges`). Export the new types from the model barrel + `src/index.ts`.
> `find(address)` resolves an address within a multi-range `sqref` (reuse `Range`/`decodeAddress`). Add unit
> tests. Run `pnpm typecheck && pnpm lint && pnpm test && pnpm build`, commit, push.

### Session 2 — write path (Opus 4.8, Claude Code)

> After session 1, implement data-validation session 2 per `docs/DATA-VALIDATION-SPEC.md` §2–§3, on
> `feat/data-validation`. In `src/xlsx/parts/worksheet.ts` add `renderDataValidations(w, validations)` and
> call it **between `renderMergeCells` and `renderHyperlinks`** (CT_Worksheet order). Emit
> `<dataValidations count><dataValidation type operator? allowBlank? showInputMessage? showErrorMessage?
> errorStyle? errorTitle? error? promptTitle? prompt? sqref><formula1>…</formula1><formula2>…</formula2>?
> </dataValidation></dataValidations>`; `formula2` only for between/notBetween; `operator` only for the
> numeric/length types; type omitted for `none`. **Handle the `showDropDown` inversion:** emit
> `showDropDown="1"` only when the model's `showDropDown === false`. No rels/extra parts. Add golden-string
> tests (list inline + range, decimal `between`, custom, messages, the inversion). Run tsc/lint/test/build,
> commit, push.

### Session 3 — read path (Sonnet 5, Claude Code)

> After sessions 1–2, implement data-validation session 3 per `docs/DATA-VALIDATION-SPEC.md` §4, on
> `feat/data-validation`. In `src/xlsx/read/worksheet.ts` parse the `<dataValidations>` child: for each
> `<dataValidation>`, read attrs + `<formula1>`/`<formula2>` into a `DataValidation`, **re-invert
> `showDropDown`** back to the model sense, and register it on the worksheet's manager keyed by `sqref`
> (collect during the child walk, apply after — order-independent; skip malformed entries, never throw). Add
> unit tests. Run tsc/lint/test/build, commit, push.

### Session 4 — compat + round-trip + docs (Sonnet 5, Claude Code)

> After sessions 1–3, implement data-validation session 4 per `docs/DATA-VALIDATION-SPEC.md` §5–§7, on
> `feat/data-validation`. In `src/compat/workbook.ts` replace the two `UnsupportedFeatureError` throwers
> (`cell.dataValidation` set at ~L199, `worksheet.dataValidations` get at ~L276) with real pass-throughs
> (the core accessor/manager already match ExcelJS's shape — stop intercepting and let them reflect
> through). Add a write→read round-trip test (list dropdown + numeric `between` + custom → equal model), an
> exceljs-guarded foreign test (exceljs dropdown → our reader; our validation buffer → exceljs intact), and
> a `cell.dataValidation` list snippet to `tests/compat/parity.test.ts`. Remove "No data validation" from
> the README Known-limitations + `docs/MIGRATION.md` matrix; add a `dataValidation` supported row with the
> deferred **cross-sheet list ref (x14)** caveat + the `showDropDown` inversion note. Run tsc/lint/test/build,
> commit, push.

### Session 5 — release 0.5.0 (Opus 4.8, Claude Code; publish owner-only)

> After sessions 1–4 are merged to `main` and CI is green, cut 0.5.0 mirroring the prior release runbooks:
> `pnpm changeset` → a **`minor`** changeset describing cell data validations (list dropdowns + whole/
> decimal/date/time/textLength/custom constraints, write+read+compat). `pnpm changeset version` (0.4.0 →
> 0.5.0), `pnpm build`, `npm pack` and confirm both `dist/index.*` and `dist/compat.*` still ship,
> `npm publish --dry-run`. STOP before the real publish; commit the bump + changelog on `main`, push, and
> hand me the `npm publish --otp=<OTP>` / `git tag v0.5.0 && git push origin v0.5.0` /
> `gh release create v0.5.0 --title "v0.5.0" --notes "$(sed -n '/## 0.5.0/,/## 0.4.0/p' CHANGELOG.md | sed '$d')"`
> commands. (Do NOT create or publish from a throwaway smoke-test npm project — mark any such dir
> `"private": true`.)
