# Stage 2 Runbook — the Reader, session by session

> Execution companion to [`STAGE-2-SPEC.md`](./STAGE-2-SPEC.md). Ready-to-paste prompts for each build
> session, with the environment, model, dependencies, and what can run in parallel. All coding sessions
> run in **Claude Code** (they write code, run `pnpm typecheck && pnpm lint && pnpm test`, and
> commit+push on your machine). Cowork's part — the spec — is done.

## Overview

| # | Session | Env | Model | Depends on | Parallel? |
|---|---------|-----|-------|-----------|-----------|
| 0 | Push + install | Code | Fable (trivial) | — | — |
| 1 | Parser + package discovery | Code | **Sonnet 5** | 0 | no (blocks all) |
| 2 | sharedStrings + props | Code | **Sonnet 5** | 1 | ✅ with 3 |
| 3 | Styles reader | Code | **Opus 4.8** | 1 | ✅ with 2 |
| 4 | Worksheet grid reader | Code | **Sonnet 5** (Opus optional) | 1, 2, 3 | no |
| 5 | Orchestrator + API + verify | Code | **Opus 4.8** | 4 | no |

**Parallelism:** once session 1 is on `main`, sessions **2 and 3 are independent** — run them in two
`git worktree`s (branches `feat/read-sst-props` and `feat/read-styles`) and merge both before session 4.
Everything else is sequential.

**Model rationale:** Opus on 3 and 5 (inverse-dedup styles logic and the integration/verification wiring —
where symmetry bugs hide), Sonnet on the mechanical sessions, Fable only for the trivial session 0.

---

## Prompts

### Session 0 — prereq (Fable)

> In `~/WebstormProjects/alosha-xlsx`: `git push` (origin/main is behind by 650f98a, e019214, 558c091,
> 6f6b504), then `pnpm install` and run `pnpm test` — confirm the exceljs round-trip test now runs green
> in-repo. Report the result.

### Session 1 — parser + package discovery (Sonnet 5)

> Implement Stage 2 session 1 per `docs/STAGE-2-SPEC.md` §2–§3. Create `src/xml/xml-reader.ts`
> (non-validating pull parser: `parseXml` generator + `cursor` helper per the §2 signatures — skip
> `<?xml?>`/comments, decode named + numeric entities, handle CDATA and self-closing tags, preserve
> significant whitespace, keep namespace prefixes verbatim) and `src/xlsx/read/package.ts` (`openPackage`:
> unzip via the existing `zip-layer` `readZipArchive`, parse `_rels/.rels` + `xl/_rels/workbook.xml.rels`,
> resolve targets relative to the workbook part, sanity-check `[Content_Types].xml`). Add
> `tests/xml/xml-reader.test.ts` + `tests/xlsx/read/package.test.ts` per §9. fflate stays the only runtime
> dep. Run `pnpm typecheck && pnpm lint && pnpm test`, fix everything, commit, push.

### Session 2 — sharedStrings + props (Sonnet 5, parallel with 3)

> Branch `feat/read-sst-props` from `main`. Implement Stage 2 session 2 per `docs/STAGE-2-SPEC.md` §4 and
> the props part of §7. Create `src/xlsx/read/shared-strings.ts` (`parseSharedStrings` →
> `(string | RichText)[]`, inverting the writer's `renderRunProps` for `<rPr>`) and
> `src/xlsx/read/props.ts` (`applyCoreProps`/`applyAppProps`, W3CDTF → `Date`), both using
> `src/xml/xml-reader.ts`. Tests are the inverse of the Stage 1.3 writer goldens. Run
> `pnpm typecheck && pnpm lint && pnpm test`, commit, push the branch. Runs in parallel with the styles
> session — stay on your branch to avoid conflicts.

### Session 3 — styles reader (Opus 4.8, parallel with 2)

> Branch `feat/read-styles` from `main`. Implement Stage 2 session 3 per `docs/STAGE-2-SPEC.md` §5. Create
> `src/xlsx/read/styles.ts`: `parseStyles` → `StyleTable { styleFor(index), isDateStyle(index) }`. Invert
> the Stage 1.3 styles builder **exactly** — numFmt as the reverse of `BUILTIN_NUMFMT_BY_CODE` plus custom
> ids ≥164; fonts/fills/borders pools → `Font`/`Fill`/`Borders`; each `<xf>` → `Style`, emitting a
> sub-object only when its `applyX` flag or non-default id says so, so a write→read→write cycle yields
> structurally identical styles. Add the `isDateFormatCode` helper (built-in ids 14–22/45–47 + token-sniff
> outside quotes). Tests per §9. Run `pnpm typecheck && pnpm lint && pnpm test`, commit, push the branch.

### Session 4 — worksheet grid reader (Sonnet 5; Opus if you prefer)

> After sessions 1–3 are merged to `main`, implement Stage 2 session 4 per `docs/STAGE-2-SPEC.md` §6.
> Create `src/xlsx/read/worksheet.ts`: `parseWorksheet(ws, xml, ctx)` populating `<cols>`, `<sheetData>`
> rows/cells, and `<mergeCells>`; decode the `t`-mapping table (s/inlineStr/str/b/e/n) with date-vs-number
> resolved via `ctx.styles.isDateStyle` + `serialToDate`, and formulas as `<f>` + typed cached `<v>`. Apply
> cell/row/column styles via the StyleTable. Tests per §9. Run `pnpm typecheck && pnpm lint && pnpm test`,
> commit, push.

### Session 5 — orchestrator + public API + verification (Opus 4.8)

> Implement Stage 2 session 5 per `docs/STAGE-2-SPEC.md` §8–§9. Create `src/xlsx/read-workbook.ts`
> (`readWorkbookBuffer`/`readWorkbookFile`), add the read exports to `src/xlsx/index.ts`, and extend the
> `workbook.xlsx` accessor in `src/index.ts` with `load`/`readFile`/`read`. Add
> `tests/xlsx/read-roundtrip.test.ts` (write→read equality of the Stage 1.3 fixture, plus inline-strings
> and styles-off variants) and `tests/xlsx/read-foreign.test.ts` (guarded on the exceljs devDep: build with
> ExcelJS → `writeBuffer` → read with ours → assert values, a merge, a numFmt). Run
> `pnpm typecheck && pnpm lint && pnpm test`, commit, push. This closes the Stage 2 read path.
