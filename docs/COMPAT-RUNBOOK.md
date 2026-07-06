# Compat Shim Runbook — session by session

> Execution companion to [`COMPAT-SPEC.md`](./COMPAT-SPEC.md). Ready-to-paste prompts for the
> `@alosha/xlsx/compat` ExcelJS drop-in (ships as **0.3.0**), with environment, model, dependencies, and
> parallelism. All coding sessions run in **Claude Code** (they write code, run
> `pnpm typecheck && pnpm lint && pnpm test && pnpm build`, and commit+push).

## Overview

| # | Session | Env | Model | Depends on | Parallel? |
|---|---------|-----|-------|-----------|-----------|
| 0 | Housekeeping + baseline | Code | Fable | 0.2.0 on main | — |
| 1 | Enum bridge + subpath wiring | Code | **Sonnet 5** | 0 | no (scaffolds compat) |
| 2 | Style-precedence adapter | Code | **Opus 4.8** | 1 | — |
| 3 | Facade + surface | Code | **Opus 4.8** | 1, 2 | no |
| 4 | Parity corpus + migration docs | Code | **Sonnet 5** | 3 | no |
| 5 | Release 0.3.0 | Code | **Opus 4.8** | 4 | no |

**Parallelism:** mostly sequential — each layer builds on the last (enum + scaffold → style adapter →
facade → parity → release). The one option: once session 1 has created `src/compat/`, sessions 2
(style-compat) and the value-type unit tests are independent enough to run in a second worktree, but the
facade (3) needs both, so the payoff is small. Recommend straight sequential.

**Two decisions to lock (see spec §5):** **D1** — emulate `cell.type` as ExcelJS `ValueType` numbers
(recommended: yes, full emulation). **D2** — emulate ExcelJS whole-object style precedence (not optional).
Confirm D1 before session 3.

---

## Prompts

### Session 0 — housekeeping + baseline (Fable, Claude Code)

> In `~/WebstormProjects/alosha-xlsx` on `main`: confirm `@alosha/xlsx@0.2.0` is the published/tagged
> baseline and `git status` is clean. Delete the merged feature/release branches locally and on origin:
> `feat/read-sst-props`, `feat/read-styles`, `feat/read-worksheet`, `feat/read-orchestrator`,
> `chore/release-ci`, `docs/release-readme`, `chore/release-bench`. Run `pnpm install` and
> `pnpm typecheck && pnpm lint && pnpm test && pnpm build` — confirm all green. Report.

### Session 1 — enum bridge + subpath wiring (Sonnet 5, Claude Code)

> Implement compat session 1 per `docs/COMPAT-SPEC.md` §1–§3. Branch `feat/compat-scaffold`. Create
> `src/compat/value-type.ts` (the ExcelJS `ValueType` object with ExcelJS's numbers, `FormulaType`, and
> `toExcelJSValueType`/`fromExcelJSValueType` maps per the §3 table) and a `src/compat/index.ts` skeleton
> (default export shaped like the ExcelJS namespace + named re-exports of the model classes + `ValueType`/
> `Enums`). Add the `./compat` subpath to `package.json` `exports` and `src/compat/index.ts` to the
> `tsup.config.ts` `entry`. Add a smoke test asserting `import ExcelJS from "@alosha/xlsx/compat"` and
> `import { ValueType } from "@alosha/xlsx/compat"` resolve, and that the ValueType numbers match §3. Run
> `pnpm typecheck && pnpm lint && pnpm test && pnpm build` (confirm `dist/compat.*` emits), commit, push.

### Session 2 — style-precedence adapter (Opus 4.8, Claude Code)

> Implement compat session 2 per `docs/COMPAT-SPEC.md` §4. Branch `feat/compat-style`. Create
> `src/compat/style-compat.ts`: `applyExcelJsStyle(cell, patch)` applying **ExcelJS whole-object per-category
> precedence** — assigning `font`/`fill`/`border`/`alignment`/`protection` replaces that category outright
> (unlike the core's deep field-merge in `resolveCellStyle`). Add unit tests asserting the divergence
> explicitly (a lone `cell.font = {bold:true}` after an inherited font REPLACES, not merges). Do not change
> the core `resolveCellStyle`. Run tsc/lint/test/build, commit, push.

### Session 3 — facade + surface (Opus 4.8, Claude Code)

> After sessions 1–2 are merged, implement compat session 3 per `docs/COMPAT-SPEC.md` §5 (and §2 matrix).
> Branch `feat/compat-facade`. Build `src/compat/workbook.ts` so compat-constructed workbooks satisfy the
> §5 contract: `cell.type`/`effectiveType` return ExcelJS `ValueType` numbers (Decision D1 = full
> emulation), style setters use `applyExcelJsStyle` (§4), and `addRow(array|object)` / `row.values`
> (1-based, empty `[0]`) / `worksheet.columns = [{header,key,width}]` behave as ExcelJS documents. The
> mechanism (facade/Proxy/scoped subclass) is yours, but it MUST NOT leak ExcelJS semantics into the core
> `@alosha/xlsx` entry. Add `UnsupportedFeatureError`-throwing stubs for the deferred surface (§6: stream/
> csv/images/tables/conditional-formatting/data-validation/comments/defined-names/protection). Wire it all
> through `compat/index.ts`. Run tsc/lint/test/build, commit, push.

### Session 4 — parity corpus + migration docs (Sonnet 5, Claude Code)

> Implement compat session 4 per `docs/COMPAT-SPEC.md` §7–§8. Branch `feat/compat-parity`. Add
> `tests/compat/parity.test.ts` (guarded on the `exceljs` devDep, `ctx.skip()` if absent): a corpus of
> ~8–12 representative ExcelJS snippets (every value type, styled headers, number formats, a merge,
> multi-sheet, column keys + `addRow(object)`, a formula with cached result, `date1904`); run each with
> real `exceljs` and with `@alosha/xlsx/compat`, load BOTH buffers with real exceljs, and assert equal
> sampled values / `cell.type` / `numFmt` / merges / sheet structure; also load the compat buffer with our
> `readWorkbookBuffer` and assert the same. Write `docs/MIGRATION.md` (one-line import change, supported/
> deferred matrix, the D1/D2 behaviour notes, a before/after snippet) and link it from README. Run
> tsc/lint/test/build, commit, push.

### Session 5 — release 0.3.0 (Opus 4.8, Claude Code; publish owner-only)

> After sessions 1–4 are merged and CI is green, cut 0.3.0 mirroring `docs/RELEASE-0.2-RUNBOOK.md`
> session 4: add a `minor` changeset describing the `@alosha/xlsx/compat` ExcelJS shim, `pnpm changeset
> version` (0.2.0 → 0.3.0), `pnpm build`, `npm pack` and confirm the tarball now also ships
> `dist/compat.js`/`.cjs`/`.d.ts`, `npm publish --dry-run`. STOP before the real publish and hand me
> (Eduardo) the `npm publish --otp=…` / `git tag v0.3.0 && git push --tags` / `gh release create` commands.
> Commit the version bump + changelog on `main`, push.
