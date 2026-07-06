# @alosha/xlsx

> A modern TypeScript library for reading and writing XLSX spreadsheets — a fresh rewrite inspired by [exceljs](https://github.com/exceljs/exceljs).

[![Test](https://github.com/avlisodraude/alosha-xlsx/actions/workflows/test.yml/badge.svg)](https://github.com/avlisodraude/alosha-xlsx/actions/workflows/test.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](./LICENSE)

`@alosha/xlsx` reads and writes real `.xlsx` (OOXML) workbooks from a single, ESM-first
document model. The public surface is intentionally ExcelJS-shaped, so most code ports with
minimal changes — but the internals are a clean rewrite with a discriminated-union value model,
no per-cell class allocation, and a single runtime dependency ([`fflate`](https://github.com/101arrowz/fflate)).

## Install

```sh
pnpm add @alosha/xlsx
# or
npm install @alosha/xlsx
# or
yarn add @alosha/xlsx
```

Requires Node.js **>= 20**. The package ships ESM + CJS builds and full type declarations.

## Write a workbook

```ts
import { Workbook } from "@alosha/xlsx";

const workbook = new Workbook();
workbook.creator = "Ada Lovelace";

const sheet = workbook.addWorksheet("Report");

// Address strings or (row, col) numbers both work.
sheet.getCell("A1").value = "Item";
sheet.getCell("B1").value = "Qty";
sheet.getCell("A2").value = "Widgets";
sheet.getCell("B2").value = 42;

// Number formats and formulas.
sheet.getCell("B2").numFmt = "#,##0";
sheet.getCell("B3").value = { formula: "SUM(B2:B2)", result: 42 };

// Column widths and merged cells.
sheet.getColumn(1).width = 24;
sheet.getCell("A4").value = "Total";
sheet.mergeCells("A4:B4");

// Serialize to bytes (browser-safe) …
const bytes: Uint8Array = await workbook.xlsx.writeBuffer();

// … or straight to disk (Node only).
await workbook.xlsx.writeFile("report.xlsx");
```

`writeBuffer()` returns a `Uint8Array` and never touches `node:fs`, so it runs in the browser,
a worker, or an edge runtime. `writeFile()` is the Node convenience wrapper.

## Read a workbook

`load` / `readFile` mutate the receiving workbook in place and return it (matching ExcelJS):

```ts
import { Workbook } from "@alosha/xlsx";

// From bytes you already have (fetch, upload, buffer) …
const workbook = new Workbook();
await workbook.xlsx.load(bytes); // bytes: Uint8Array

// … or from a path (Node only).
await workbook.xlsx.readFile("report.xlsx");

const sheet = workbook.getWorksheet("Report");
console.log(sheet?.getCell("B2").value); // 42
console.log(sheet?.getCell("A1").text); // "Item"

sheet?.eachRow((row, rowNumber) => {
  console.log(rowNumber, row.values);
});
```

## The free-function core

The `workbook.xlsx.*` accessor is a thin ExcelJS-style wrapper over an environment-agnostic,
**synchronous** core. Reach for these when you want no magic accessor, tree-shakeable imports,
or to keep the model layer free of the serializer:

```ts
import { Workbook, writeWorkbookBuffer, readWorkbookBuffer } from "@alosha/xlsx";

const wb = new Workbook();
wb.addWorksheet("Sheet1").getCell("A1").value = "hi";

const bytes: Uint8Array = writeWorkbookBuffer(wb); // sync
const roundTripped: Workbook = readWorkbookBuffer(bytes); // sync
```

Node-only file variants (`writeWorkbookFile`, `readWorkbookFile`) are also exported; they lazily
import `node:fs/promises` so browser bundles that only use the buffer functions never pull it in.

`writeWorkbookBuffer` accepts a `WriteOptions` (`useSharedStrings`, `useStyles`); the read
functions accept a `ReadOptions` (reserved for future use).

## Feature matrix

| Area | Read | Write | Notes |
| --- | :---: | :---: | --- |
| Cell values — string, number, boolean, date | ✅ | ✅ | 1900 and 1904 date systems |
| Formulas (with cached result) | ✅ | ✅ | shared/array formulas round-trip |
| Rich text runs | ✅ | ✅ | per-run fonts |
| Error values (`#REF!`, `#DIV/0!`, …) | ✅ | ✅ | |
| Shared strings | ✅ | ✅ | `useSharedStrings` toggle on write |
| Styles — fonts, fills, borders, alignment, number formats, protection | ✅ | ✅ | `useStyles` toggle on write |
| Merged cells | ✅ | ✅ | |
| Column width / styles / keys | ✅ | ✅ | |
| Row height / styles / outline levels | ✅ | ✅ | |
| Multiple worksheets, sheet state (visible / hidden / veryHidden) | ✅ | ✅ | |
| Workbook & sheet views (freeze/split, zoom, active tab) | ✅ | ✅ | round-trip as data |
| Workbook metadata (creator, dates, title, …) | ✅ | ✅ | |
| Hyperlink cells | ⚠️ | ⚠️ | round-trip as display **text** — see limitations |
| Streaming read/write | ❌ | ❌ | deferred |
| Drawings / images / charts | ❌ | ❌ | deferred |
| Data validation, conditional formatting, comments | ❌ | ❌ | deferred |

## Migrating from ExcelJS

The API is deliberately familiar — `new Workbook()`, `addWorksheet`, `getCell`, `getRow`,
`mergeCells`, and the `workbook.xlsx.writeBuffer / writeFile / load / readFile` accessor all
behave as you'd expect. Common differences to watch for:

- **Bytes, not Node `Buffer`.** `writeBuffer()` resolves to a `Uint8Array`. In Node a
  `Uint8Array` is accepted anywhere a `Buffer` is (`fs.writeFile`, `res.end`, …); pass it
  through, or wrap with `Buffer.from(bytes)` if an API insists on a `Buffer`.
- **`worksheets` is a dense, 0-based array.** No leading `undefined` / 1-based sparse hole —
  `workbook.worksheets[0]` is the first sheet. Use `getWorksheet(idOrName)` for lookups by id or
  name.
- **Clearer row inheritance.** Instead of ExcelJS's `'i' | 'o' | 'i+' | 'o+'` DSL, `addRow` /
  `insertRow` take `{ inheritFrom: "above" | "below", includeEmpty?: boolean }`.
- **ESM-first.** The package is ESM with a CJS fallback; prefer `import` over `require`.
- **Hyperlinks round-trip as text** (see below).

Not everything ExcelJS does is implemented yet — check the feature matrix and limitations before
porting a workbook that leans on drawings, streaming, or data validation.

## Known limitations

- **Hyperlink cells round-trip as their display text.** Setting a `{ text, hyperlink }` value
  writes the visible text; the clickable link relationship is not yet persisted, and reading a
  workbook with hyperlinks yields the text only. Real hyperlink relationships are planned.
- **No streaming.** Reading and writing are buffer-based (the whole workbook lives in memory). A
  streaming API for very large sheets is deferred.
- **No drawings, images, or charts.** These parts are ignored on read and not emitted on write.
- **No data validation, conditional formatting, or comments** yet.

See [`BENCHMARKS.md`](./BENCHMARKS.md) for write/read throughput and output-size numbers against
ExcelJS, and [`CHANGELOG.md`](./CHANGELOG.md) for release notes.

## Development

This project uses [pnpm](https://pnpm.io/), [tsup](https://tsup.egoist.dev/),
[Biome](https://biomejs.dev/), and [Vitest](https://vitest.dev/).

```sh
pnpm install      # install dependencies
pnpm build        # build ESM + CJS + type declarations
pnpm dev          # rebuild on change
pnpm test         # run the test suite once
pnpm test:watch   # run tests in watch mode
pnpm lint         # lint + format check with Biome
pnpm format       # auto-format with Biome
pnpm typecheck    # type-check with tsc
```

### A note on `pnpm install` warnings

Installing dev dependencies prints a handful of `npm warn deprecated` lines
(`rimraf`, `inflight`, `lodash.isequal`, `glob`, `uuid`, `fstream`). These are
**transitive dependencies of `exceljs`**, which we pull in only as a
`devDependency` to verify our output: the round-trip test in
`tests/xlsx/write-workbook.test.ts` re-opens a generated workbook with ExcelJS
to confirm it loads without a repair prompt.

None of these ship in `@alosha/xlsx` — the published package has a single
runtime dependency, `fflate`. The warnings are ExcelJS's baggage, not ours;
in fact they're a tidy illustration of why this library exists. If ExcelJS is
not installed, that one round-trip test simply skips.

## License

[MIT](./LICENSE) © E.Y. Silva Navarrete
