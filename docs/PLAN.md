# @alosha/xlsx — Project Plan

> A modern TypeScript rewrite inspired by [exceljs](https://github.com/exceljs/exceljs),
> focused on a clean, strictly-typed, ESM-first API for reading and writing XLSX files.

## Goals

- **Modern toolchain:** ESM + CJS dual output, full type declarations, Node >= 20.
- **Strict TypeScript:** no implicit `any`, checked indexed access, tree-shakeable.
- **Great DX:** small, composable public API; helpful types; thorough tests.
- **Standards-based:** read/write the OOXML SpreadsheetML (`.xlsx`) format.

## Current status

**Stage 1.2 — ZIP layer**

Stage 0 (foundation, build tooling, test harness) is complete:
- Repo scaffold: ✅ package.json, tsconfig.json (strict), tsup/vitest/biome configs
- Dual build: ✅ ESM (index.js) + CJS (index.cjs) + type declarations generated
- Test harness: ✅ vitest with 3 passing placeholder tests
- Lint/format: ✅ Biome configured (check + format scripts)
- Git: ✅ chore/scaffold branch with initial commit

Stage 1.1 added:
- CI: ✅ GitHub Actions workflow (lint, typecheck, test) on push/PR + manual dispatch
- fflate: ✅ added as dependency; `src/zip-layer.ts` stub in place for Stage 1.2

Stage 1.2 adds:
- ZIP layer: ✅ `readZipArchive`/`writeZipArchive` in `src/zip-layer.ts`, backed by fflate's
  `unzipSync`/`zipSync`, operating on `Map<string, Uint8Array>` entries. Round-trip tested.

## Roadmap (high level)

1. **Foundation** ✅ — repo, tooling, CI-ready scripts, dual build.
2. **Core model** — Workbook / Worksheet / Row / Cell data structures.
3. **Reader** — parse `.xlsx` (zip + XML) into the core model.
4. **Writer** — serialize the core model back to `.xlsx`.
5. **Styles & formats** — number formats, fonts, fills, borders, alignment.
6. **Advanced** — formulas, shared strings, streaming read/write, merged cells.

_This document is a living plan and will evolve as the library takes shape._
