# Release 0.2 Spec — first public npm release (read + write)

> The milestone after the Stage 2 reader: cut the first published `@alosha/xlsx` on npm. With write
> (Stage 1.3) and read (Stage 2) both done, the library is a real read+write OOXML tool — this release
> makes it installable and gives the portfolio a credible, npm-native anchor while SheetJS is off-npm and
> ExcelJS is stalled. Execution companion: [`RELEASE-0.2-RUNBOOK.md`](./RELEASE-0.2-RUNBOOK.md).

## 0. Scope

**In:** finish CI (add the missing build step + a Node 20/22 matrix to the existing `test.yml`); fill the
npm-facing `package.json` metadata (`repository`, `homepage`, `bugs`); a real README quickstart covering
**both** read and write plus an exceljs-migration note and a known-limitations list; a small benchmark vs
exceljs with published numbers; Changesets initialised with a `0.2.0` changeset; a verified package tarball
(dist + README + LICENSE only); and the publish itself.

**Out (later stages):** the `@alosha/xlsx/compat` shim, streaming, the remaining Stage 2 write features
(data validation, conditional formatting, images, comments, real hyperlink relationships), the
`xlsx.alosha.dev` product subdomain (Stage 4 — adjacent, not part of this npm cut), and any API changes
(this release ships the current surface as-is).

**Version decision:** publish the first release directly as **0.2.0** (read + write), skipping a separate
0.1.0 — nothing was ever published, and 0.2 matches the roadmap's "reading = v0.2." (Publishing 0.1.0 first
is a valid alternative if you'd rather mark the write-only milestone; pick one before the release session.)

**Definition of done:** CI green on `main` (lint + typecheck + test + **build**) across Node 20 and 22;
`npm pack` contains only `dist/**`, `README.md`, `LICENSE`; `npm publish --dry-run` clean; the README
renders correct read+write examples; a git tag `v0.2.0` + GitHub release exist; and `npm view @alosha/xlsx`
shows the published package. The actual `npm publish`, the npm token/OTP, and GitHub branch protection are
**owner-only** (Eduardo) — the sessions prepare and dry-run everything up to that line.

## 1. Workstreams

1. **CI + package metadata.** Extend `.github/workflows/test.yml`: add `- run: pnpm build`, run the job on
   a `node-version: [20.x, 22.x]` matrix. Add `repository` (`git+https://github.com/avlisodraude/alosha-xlsx.git`),
   `homepage`, and `bugs` to `package.json` (needed for the npm page and publish provenance). Confirm
   `publishConfig.access: public` and the `files` allowlist are correct. Owner action: enable branch
   protection requiring the CI check.
2. **Docs.** Rewrite `README.md` around the shipped API: install, a write quickstart
   (`new Workbook()` → `workbook.xlsx.writeBuffer/writeFile`), a read quickstart
   (`workbook.xlsx.load/readFile`), the free-function core (`writeWorkbookBuffer`/`readWorkbookBuffer`), a
   feature matrix, a short "migrating from exceljs" section, a **known limitations** list (hyperlink cells
   currently round-trip as display text; streaming/read-of-drawings deferred), a CI badge, and a link to
   `BENCHMARKS.md`. Add a `CHANGELOG.md` (Changesets will own it going forward).
3. **Benchmarks.** A `bench/` script that builds a mid-size workbook (e.g. 50k cells, mixed
   types/styles/merges), times write→bytes and read-back for `@alosha/xlsx` vs `exceljs`, and reports
   output size + wall time. Write results to `BENCHMARKS.md`. Uses the existing `exceljs` devDep. Keep it
   deterministic and node-run (`node bench/…` via tsx or a compiled step), not part of `pnpm test`.
4. **Release mechanics + publish.** `pnpm changeset init`, add a `minor` changeset describing 0.2.0,
   `pnpm changeset version` (bumps `0.0.0`→`0.2.0`, generates CHANGELOG). Locally: `pnpm build`, then
   `npm pack` and inspect the tarball contents; `npm publish --dry-run` (add `--provenance` only if
   publishing from CI OIDC). **Owner runs** `npm publish` (scoped public, OTP), `git tag v0.2.0`, push the
   tag, and cuts the GitHub release from the changelog.

## 2. Risks / notes

- **Build in the sandbox:** `tsup` may hit the esbuild host/binary version mismatch in the Linux sandbox;
  it builds fine on macOS. The release must confirm `pnpm build` locally before packing.
- **Provenance:** simplest path is a manual `npm publish` with OTP. Provenance (`--provenance`) needs the
  `repository` field (added in workstream 1) and ideally a CI OIDC publish; treat it as a nice-to-have, not
  a blocker for 0.2.0.
- **No API churn:** freeze the surface for this cut. Anything that would change public types belongs to a
  later minor.
- **Subdomain:** `xlsx.alosha.dev` wiring (Vercel + Cloudflare CNAME + landing in alosha-web) is a separate,
  optional follow-on in the alosha-web repo — see the other packages' subdomain pattern; not required to
  publish.

---

*Grounded in the current repo state (v0.0.0, existing `test.yml`, no Changesets, `files` = dist/README/
LICENSE) and the Alosha stack decisions (pnpm, tsup, Changesets, MIT).*
