# @alosha/xlsx

> A modern TypeScript library for reading and writing XLSX spreadsheets — a fresh rewrite inspired by [exceljs](https://github.com/exceljs/exceljs).

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](./LICENSE)

## Status

🚧 **Early scaffolding.** The public API is not yet implemented — this repo currently
contains only the build, test, and tooling foundation.

## Install

```sh
pnpm add @alosha/xlsx
# or
npm install @alosha/xlsx
# or
yarn add @alosha/xlsx
```

Requires Node.js **>= 20**.

## Quick example

> ⚠️ Placeholder — the real spreadsheet API is coming soon.

```ts
import { hello } from "@alosha/xlsx";

console.log(hello("world"));
// => "Hello, world, from @alosha/xlsx!"
```

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
