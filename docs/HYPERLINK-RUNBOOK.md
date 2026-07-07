# Hyperlink Relationships Runbook — session by session

> Execution companion to [`HYPERLINK-SPEC.md`](./HYPERLINK-SPEC.md). Ready-to-paste prompts for making cell
> hyperlinks real OOXML relationships (write + read + compat), shipping as **0.4.0**. All coding sessions
> run in **Claude Code** (they write code, run `pnpm typecheck && pnpm lint && pnpm test && pnpm build`, and
> commit+push).

## Overview

| # | Session | Env | Model | Depends on | Parallel? |
|---|---------|-----|-------|-----------|-----------|
| 0 | Baseline + branch | Code | Fable | 0.3.1 on main | — |
| 1 | Write path | Code | **Opus 4.8** | 0 | no |
| 2 | Read path | Code | **Sonnet 5** | 1 | no |
| 3 | Round-trip + compat + docs | Code | **Sonnet 5** | 2 | no |
| 4 | Release 0.4.0 | Code | **Opus 4.8** | 3 | no |

**Parallelism:** sessions 1 (write) and 2 (read) touch disjoint files and *could* parallelize, but the
reader is far easier to build and verify against real writer output — recommend **sequential**.

**Model rationale:** Opus on 1 (cross-part rId wiring + the `renderWorksheet` return-type change ripple)
and 4 (release); Sonnet on the read parse and the docs/round-trip session.

---

## Prompts

### Session 0 — baseline + branch (Fable, Claude Code)

> In `~/WebstormProjects/alosha-xlsx` on `main`: confirm `@alosha/xlsx@0.3.1` is the published baseline,
> `git status` clean, `pnpm install`, then `pnpm typecheck && pnpm lint && pnpm test && pnpm build` all
> green. Create branch `feat/hyperlinks`. Report.

### Session 1 — write path (Opus 4.8, Claude Code)

> Implement hyperlink session 1 per `docs/HYPERLINK-SPEC.md` §1–§2, on branch `feat/hyperlinks`. Add
> `REL_TYPE.hyperlink` to `src/xlsx/parts/constants.ts`. In `src/xlsx/parts/worksheet.ts`, keep the
> `CellType.Hyperlink` cell emitting its display text as a string cell, but also collect
> `{ ref, target, tooltip }` per sheet and emit a `<hyperlinks>` block **after `<mergeCells>`** with local
> `rId${n}` ids; change `renderWorksheet` to return `{ xml, rels? }` (rels = the rendered
> `sheetN.xml.rels`). Add `src/xlsx/parts/worksheet-rels.ts` (`renderWorksheetRels`, external
> `TargetMode="External"` relationships, mirroring `workbook-rels.ts`). Update `src/xlsx/write-workbook.ts`
> so `worksheetXml` carries `{ xml, rels? }` and writes `xl/worksheets/_rels/<file>.rels` when present
> (content-types needs NO change — the `Default Extension="rels"` covers it). Add unit tests (hyperlink cell
> → string value + `<hyperlinks>` entry with matching r:id; `renderWorksheetRels` golden; tooltip present/
> absent). Run `pnpm typecheck && pnpm lint && pnpm test && pnpm build`, commit, push.

### Session 2 — read path (Sonnet 5, Claude Code)

> After session 1 is committed, implement hyperlink session 2 per `docs/HYPERLINK-SPEC.md` §3, on
> `feat/hyperlinks`. In `src/xlsx/read/package.ts` add a `partRels(partPath)` helper (reuse `parseRels` +
> `relsPathFor`) returning `rId → { type, target, mode }` for any part (worksheet rels live at
> `xl/worksheets/_rels/sheetN.xml.rels`). In `src/xlsx/read/worksheet.ts` parse the `<hyperlinks>` block and,
> given the resolved worksheet-rels map in the read context, reconstruct each hyperlink cell as
> `{ text: <current cell string value>, hyperlink: target, tooltip? }` (`CellType.Hyperlink`); a range `ref`
> applies to its top-left cell. Wire the per-sheet rels through `src/xlsx/read-workbook.ts`. Add unit tests
> (a `<hyperlinks>` block + rels map reconstructs the model value). Run tsc/lint/test/build, commit, push.

### Session 3 — round-trip + compat + docs (Sonnet 5, Claude Code)

> After sessions 1–2, implement hyperlink session 3 per `docs/HYPERLINK-SPEC.md` §4–§5, on
> `feat/hyperlinks`. Add a write→read round-trip test (fixture with 2–3 hyperlinks incl. a tooltip and a
> `mailto:`; assert the part set now includes `xl/worksheets/_rels/sheet1.xml.rels` and the values read back
> equal `{ text, hyperlink, tooltip? }` with `CellType.Hyperlink`). Add an exceljs-guarded foreign-file test
> (ExcelJS-written hyperlink → our reader recovers it; our hyperlink buffer re-opens in ExcelJS intact). Add
> a hyperlink snippet to `tests/compat/parity.test.ts` (previously excluded). Remove the "hyperlink cells
> round-trip as their display text" Known limitation from `README.md` and `docs/MIGRATION.md` (note
> internal-`location=` links as still-deferred if you keep a caveat). Run tsc/lint/test/build, commit, push.

### Session 4 — release 0.4.0 (Opus 4.8, Claude Code; publish owner-only)

> After sessions 1–3 are merged to `main` and CI is green, cut 0.4.0 mirroring
> `docs/COMPAT-RUNBOOK.md` session 5 / `RELEASE-0.2-RUNBOOK.md`: `pnpm changeset` → a **`minor`** changeset
> describing real hyperlink relationships **and the read behaviour change** (hyperlink cells now read back as
> `{ text, hyperlink }` / `ValueType.Hyperlink`, not a string). `pnpm changeset version` (0.3.1 → 0.4.0),
> `pnpm build`, `npm pack` and confirm both `dist/index.*` and `dist/compat.*` still ship, `npm publish
> --dry-run`. STOP before the real publish; commit the bump + changelog on `main`, push, and hand me
> (Eduardo) the `npm publish --otp=<OTP>` / `git tag v0.4.0 && git push origin v0.4.0` /
> `gh release create v0.4.0 --title "v0.4.0" --notes "$(sed -n '/## 0.4.0/,/## 0.3.1/p' CHANGELOG.md | sed '$d')"`
> commands.
