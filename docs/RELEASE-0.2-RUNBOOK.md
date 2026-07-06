# Release 0.2 Runbook — session by session

> Execution companion to [`RELEASE-0.2-SPEC.md`](./RELEASE-0.2-SPEC.md). Ready-to-paste prompts for the
> first public npm release, with environment, model, dependencies, and parallelism. Runs after the Stage 2
> reader is merged to `main`.

## Overview

| # | Session | Env | Model | Depends on | Parallel? |
|---|---------|-----|-------|-----------|-----------|
| 0 | Prereq: merge, push, green | Code | Fable (trivial) | Stage 2 done | — |
| 1 | CI build+matrix + package metadata | Code | **Sonnet 5** | 0 | ✅ with 2, 3 |
| 2 | README quickstart + docs | Cowork | **Opus 4.8** | 0 | ✅ with 1, 3 |
| 3 | Benchmarks vs exceljs | Code | **Sonnet 5** | 0 | ✅ with 1, 2 |
| 4 | Release mechanics + publish | Code | **Opus 4.8** | 1, 2, 3 | no |

**Parallelism:** sessions **1, 2, 3 touch disjoint files** (`.github/` + `package.json` / `README.md` +
`CHANGELOG.md` / `bench/` + `BENCHMARKS.md`) — run them together in separate `git worktree`s and merge all
three before session 4. Session 2 links to `BENCHMARKS.md` rather than inlining numbers, so it needn't wait
on session 3.

**Model rationale:** Opus on 2 (public-facing docs, migration framing) and 4 (release/publish — where a
wrong `files`/version/tag bites); Sonnet on the mechanical CI and bench work; Fable on the trivial prereq.
Session 2 is the one **Cowork** job (doc-writing); the rest are **Claude Code** (they run build/test/git).

**Owner-only (Eduardo, in session 4):** the actual `npm publish` (scoped public + OTP), the npm automation
token, `git tag`/GitHub release, and enabling branch protection to require the CI check.

---

## Prompts

### Session 0 — prereq (Fable, Claude Code)

> In `~/WebstormProjects/alosha-xlsx`: confirm the Stage 2 reader is merged to `main`, `git push`, run
> `pnpm install`, then `pnpm typecheck && pnpm lint && pnpm test && pnpm build` — all must be green locally
> (the build matters here; it's what CI and the release will pack). Report any failure before proceeding.

### Session 1 — CI build+matrix + package metadata (Sonnet 5, Claude Code) — parallel

> Per `docs/RELEASE-0.2-SPEC.md` workstream 1. Branch `chore/release-ci`. Edit `.github/workflows/test.yml`
> to add a `- run: pnpm build` step and run the job across `node-version: [20.x, 22.x]`. In `package.json`
> add `repository` (`git+https://github.com/avlisodraude/alosha-xlsx.git`), `homepage`
> (`https://github.com/avlisodraude/alosha-xlsx#readme`), and `bugs`
> (`https://github.com/avlisodraude/alosha-xlsx/issues`); confirm `publishConfig.access: "public"` and that
> `files` is `["dist","README.md","LICENSE"]`. Do NOT touch README (session 2 owns it). Run
> `pnpm typecheck && pnpm lint && pnpm test && pnpm build`, commit, push the branch.

### Session 2 — README quickstart + docs (Opus 4.8, Cowork) — parallel

> Per `docs/RELEASE-0.2-SPEC.md` workstream 2. Rewrite `README.md` for the shipped read+write API: install;
> a write quickstart (`new Workbook()`, add a sheet/cells, `workbook.xlsx.writeBuffer` and `writeFile`); a
> read quickstart (`workbook.xlsx.load(buffer)` / `readFile(path)`); the free-function core
> (`writeWorkbookBuffer` / `readWorkbookBuffer`); a feature matrix; a short "Migrating from ExcelJS"
> section; a "Known limitations" list (hyperlink cells round-trip as display text; streaming and
> drawings/images deferred); a CI status badge; and a link to `BENCHMARKS.md`. Keep the existing dev-notes
> section. Add a `CHANGELOG.md` stub (Changesets takes it over in session 4). Verify every code sample
> compiles against the current public types in `src/index.ts`. This is a docs-only session — no `src/`
> changes. Commit on branch `docs/release-readme`, push.

### Session 3 — benchmarks vs exceljs (Sonnet 5, Claude Code) — parallel

> Per `docs/RELEASE-0.2-SPEC.md` workstream 3. Branch `chore/release-bench`. Add a `bench/` script that
> builds a mid-size workbook (~50k cells, mixed value types + styles + a few merges), times
> write→`Uint8Array` and read-back for `@alosha/xlsx` vs the `exceljs` devDep, and prints output size + wall
> time (median of a few runs). Make it node-runnable (via tsx or a small compiled step), NOT part of
> `pnpm test`; add a `bench` script to `package.json`. Write the results to `BENCHMARKS.md` (machine/Node
> version noted). Run `pnpm typecheck && pnpm lint`, commit, push the branch.

### Session 4 — release mechanics + publish (Opus 4.8, Claude Code; publish is owner-only)

> After sessions 1–3 are merged to `main` and CI is green, per `docs/RELEASE-0.2-SPEC.md` workstream 4:
> `pnpm changeset init`, add a `minor` changeset describing the 0.2.0 read+write release, run
> `pnpm changeset version` (bumps `0.0.0`→`0.2.0` and writes CHANGELOG), then `pnpm build`, `npm pack`, and
> inspect the tarball — it must contain only `dist/**`, `README.md`, `LICENSE`. Run `npm publish --dry-run`
> and fix anything it flags. STOP before the real publish and hand me (Eduardo) the exact commands to run:
> `npm publish` (scoped public, with OTP), `git tag v0.2.0 && git push --tags`, and the GitHub release notes
> drawn from the changelog. Commit the version bump + changelog + changeset on `main`, push.
