# @alosha/xlsx — Project Plan

> A modern TypeScript rewrite inspired by [exceljs](https://github.com/exceljs/exceljs),
> focused on a clean, strictly-typed, ESM-first API for reading and writing XLSX files.

## Goals

- **Modern toolchain:** ESM + CJS dual output, full type declarations, Node >= 20.
- **Strict TypeScript:** no implicit `any`, checked indexed access, tree-shakeable.
- **Great DX:** small, composable public API; helpful types; thorough tests.
- **Standards-based:** read/write the OOXML SpreadsheetML (`.xlsx`) format.

## Current status

Foundation scaffolding only. No feature code yet. The build, lint, type-check, and
test pipelines are green with a placeholder export.

## Roadmap (high level)

1. **Foundation** ✅ — repo, tooling, CI-ready scripts, dual build.
2. **Core model** — Workbook / Worksheet / Row / Cell data structures.
3. **Reader** — parse `.xlsx` (zip + XML) into the core model.
4. **Writer** — serialize the core model back to `.xlsx`.
5. **Styles & formats** — number formats, fonts, fills, borders, alignment.
6. **Advanced** — formulas, shared strings, streaming read/write, merged cells.

_This document is a living plan and will evolve as the library takes shape._
