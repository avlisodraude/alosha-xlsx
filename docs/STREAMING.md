# Streaming writer (`WorkbookWriter`)

> Write large workbooks with bounded memory: rows are rendered and deflated into the output
> `.xlsx` as you commit them, instead of building the whole document model in memory first.

The buffered writer (`workbook.xlsx.writeBuffer()` / `writeWorkbookBuffer`) must hold every
worksheet, every row, and the entire output archive in memory at once — fine for reports, painful
for exports with hundreds of thousands of rows. The streaming `WorkbookWriter` is an incremental,
ExcelJS-shaped alternative (`ExcelJS.stream.xlsx.WorkbookWriter`): add a sheet, commit rows one at
a time, and each committed row's XML is deflated straight into an open zip entry and handed to your
output sink. Memory stays roughly flat no matter how many rows you write.

At 500,000 rows × 5 columns, streaming holds **~15.6× less peak memory** than the buffered writer
while sustaining **~1.9× the write throughput** (see [Benchmarks](#benchmarks)).

## Quick start

```ts
import { WorkbookWriter } from "@alosha/xlsx";

const wb = new WorkbookWriter({ filename: "big-report.xlsx" });
const ws = wb.addWorksheet("Data");

for (const record of hugeDataset) {
  ws.addRow([record.id, record.name, record.amount, record.when]).commit();
  // ^ this row's XML is rendered and deflated to disk now — nothing is retained
}

ws.commit();       // finalise the worksheet part
await wb.commit(); // write styles/workbook/content-types, end the zip, flush the file
```

The lifecycle is always the same three commits:

1. **`row.commit()`** — flushes that row's XML to the sink. The row is gone from memory afterwards;
   you cannot revisit or edit it.
2. **`ws.commit()`** — closes `<sheetData>`, emits the sheet's trailing features (merges,
   auto-filter, conditional formatting, hyperlink/data-validation blocks, comments, images) and its
   child parts, and closes the worksheet zip entry.
3. **`await wb.commit()`** — writes the remaining workbook-level parts (`[Content_Types].xml`,
   rels, `styles.xml`, `sharedStrings.xml` if enabled, `workbook.xml`, theme, doc properties), ends
   the archive, and — for a `filename`/`stream` destination — resolves once the destination has
   fully flushed. Uncommitted sheets are committed for you, but committing each sheet explicitly
   (as ExcelJS requires) is the clearer style.

One sheet streams at a time: once a sheet's first row commits, its zip entry is open, and starting
rows on a *different* sheet before calling `ws.commit()` on the first throws. Finish a sheet, then
start the next.

## Constructing: options and output sinks

```ts
new WorkbookWriter(options?: WorkbookWriterOptions)
```

| Option | Default | Meaning |
| --- | --- | --- |
| `filename` | — | Write the archive to this path (Node only; `node:fs` is imported lazily). |
| `stream` | — | Pipe the archive into a Node `Writable` (e.g. an HTTP response). Backpressure is honoured. |
| `useSharedStrings` | **`false`** | Emit a shared-string table instead of inline strings. See [below](#the-usesharedstrings-false-default) for why the default is the opposite of the buffered writer's. |
| `useStyles` | `true` | Intern cell styles into `styles.xml` (same as the buffered `WriteOptions.useStyles`). |

`filename` and `stream` are mutually exclusive. When **neither** is given, the writer itself is the
sink: it implements `AsyncIterable<Uint8Array>` and yields archive chunks as they are produced.

### Sink 1 — async iterable of `Uint8Array` (environment-agnostic)

No `node:*` imports are touched on this path, so it works in the browser, a worker, or an edge
runtime as well as Node. Consume the chunks concurrently with producing them (start iterating
before/while you commit, or right after — chunks queue until read):

```ts
import { WorkbookWriter } from "@alosha/xlsx";

const wb = new WorkbookWriter(); // no destination → wb is an AsyncIterable<Uint8Array>
const ws = wb.addWorksheet("Data");

// Drain concurrently — e.g. into a web ReadableStream, an upload, or a collector.
const drained = (async () => {
  const chunks: Uint8Array[] = [];
  for await (const chunk of wb) chunks.push(chunk);
  return chunks;
})();

for (let i = 1; i <= 1_000_000; i++) {
  ws.addRow([i, `row ${i}`, i * 0.01]).commit();
}
ws.commit();
await wb.commit();

const chunks = await drained; // the complete .xlsx, as a sequence of byte chunks
```

### Sink 2 — Node `Writable` or file path

Pass `stream` (any Node `Writable` — a file stream, an HTTP response, a pipeline stage) or the
`filename` shorthand. Both are drained by a background pump that honours backpressure; `await
wb.commit()` resolves only after the destination has fully flushed.

```ts
import { createWriteStream } from "node:fs";
import { WorkbookWriter } from "@alosha/xlsx";

// Explicit Writable (identical to the filename shorthand below):
const wb = new WorkbookWriter({ stream: createWriteStream("export.xlsx") });

// …or straight into an HTTP response:
// const wb = new WorkbookWriter({ stream: res });

// …or just:
// const wb = new WorkbookWriter({ filename: "export.xlsx" });
```

The `Writable` is duck-typed (`write`/`end`/`once`/`destroy`), so anything structurally compatible
works and browser bundles that never pass one stay free of `node:*` imports.

## The `useSharedStrings: false` default

OOXML forces exactly one thing to be buffered when strings are *shared*: the sharedStrings table.
Sheets reference strings by index, but the table itself is written **after** every sheet — so every
distinct string must be retained in memory until the very end. That's why the streaming writer
defaults to `useSharedStrings: false` (each string cell carries its text inline as `<is><t>`), the
same choice ExcelJS's `WorkbookWriter` makes: with inline strings, a committed row leaves
**nothing** behind, and memory stays flat regardless of row count.

Set `useSharedStrings: true` if you want the deduplicated table (smaller output when the same
strings repeat a lot): rows still stream — they carry the interned index — but the string table
grows with the number of *distinct* strings and is held until `wb.commit()`.

Note the buffered writer defaults the other way (`useSharedStrings: true`), since it holds
everything in memory anyway.

## What a streamed row can carry

`addRow(values)` accepts the same row shapes as the buffered writer (an array of cell values, or a
sparse object) and returns a real model `Row` plus `.commit()`. Everything you can put on a
buffered row works — set it **before** committing the row:

```ts
const row = ws.addRow(["Total", 1234.5, new Date()]);
row.height = 22;
row.getCell(1).style = { font: { bold: true } };
row.getCell(2).numFmt = "#,##0.00";
row.getCell(3).note = "generated";           // comments are collected, flushed at ws.commit()
row.commit();
```

Full cell-value coverage matches the buffered writer: strings, numbers, booleans, dates (1900 and
1904 systems), error values, formulas (with cached results, shared/array forms), rich text, and
hyperlink cells.

Sheet-level features are configured on the `WorksheetWriter` and emitted in the epilogue at
`ws.commit()`: `mergeCells(...)`, `autoFilter`, `addConditionalFormatting(...)`,
`dataValidations`, images (`wb.addImage(...)` + `ws.addImage(id, anchor)`), and comments via
`cell.note`. Two of them — `views` (frozen/split panes, gridline visibility) and `properties`
(default row height, …) — are written into the sheet *prologue*, so set them **before the first
row commits**.

Workbook metadata (`creator`, `created`, `title`, …, `properties.date1904`, `calcProperties`,
`views`) lives directly on the `WorkbookWriter`, mirroring ExcelJS.

## Parity with the buffered writer

For every in-scope feature, streamed output is asserted **structurally equal** to
`writeWorkbookBuffer` with the same options (see `tests/xlsx/stream/parity.test.ts`), and the
archive opens in Excel without repair. Two deliberate differences:

- `<dimension>` is omitted from streamed sheets (the sheet's extent isn't known until the last
  row); readers — including this library's — recompute it.
- The default `useSharedStrings` differs, as described above.

## Benchmarks

From [`BENCHMARKS.md`](../BENCHMARKS.md) (`pnpm bench:stream`, Apple M3 Pro, Node 22): one sheet of
500,000 rows × 5 columns (2.5 M cells, mixed types, inline strings, written to a file):

| Mode | Write time | Throughput | Peak RSS |
| --- | --- | --- | --- |
| streaming (`WorkbookWriter`) | 6.32 s | 79,074 rows/s | 188.6 MiB |
| buffered (`writeWorkbookBuffer`) | 12.05 s | 41,500 rows/s | 2939.6 MiB |

**~15.6× less peak memory at ~1.9× the throughput.** The buffered writer must materialise the full
model plus the whole output buffer, so its memory grows with the sheet; the streaming writer's
stays roughly flat as rows grow.

## Known limitations

- **Write-only.** There is no streaming *reader* (`WorkbookReader`) yet; reading is buffer-based
  (`workbook.xlsx.load` / `readWorkbookBuffer`).
- **Column widths / `<cols>` are not yet supported in streaming mode.** `getColumn` / `columns`
  are intentionally not exposed on the streaming `WorksheetWriter`, and no `<cols>` block is
  emitted — a planned follow-up. (Column *content* is unaffected; only widths/column styles/keys
  are missing.) Use the buffered writer if you need column widths today.
- **No random access after commit.** A committed row is flushed and gone; committed sheets can't
  be reopened. This is inherent to streaming, not a gap.
- **`@alosha/xlsx/compat`'s `ExcelJS.stream.xlsx.WorkbookWriter` is a real bridge onto this writer**
  (docs/MIGRATION.md D6) — the same lifecycle, with rows/cells proxied the same way the buffered
  compat facade proxies them. A couple of real-ExcelJS constructor/`addWorksheet` options with no
  equivalent here (`zip`, `lastPrinted`, `pageSetup`, `headerFooter`, a per-sheet `useSharedStrings`
  override) throw `UnsupportedFeatureError` instead of being silently dropped. `stream.xlsx.
  WorkbookReader` still throws — no streaming reader exists (previous point).
- **No CSV streaming** (matching the absence of CSV support generally).

## API reference (summary)

```ts
import { WorkbookWriter } from "@alosha/xlsx";
// also exported: WorksheetWriter, and the types WorkbookWriterOptions, StreamedRow,
// ByteSink, NodeWritableLike

class WorkbookWriter implements AsyncIterable<Uint8Array> {
  constructor(options?: WorkbookWriterOptions);
  addWorksheet(name?: string, options?: { state?: "visible" | "hidden" | "veryHidden" }): WorksheetWriter;
  addImage(options: AddImageOptions): number;   // register once, place per-sheet
  commit(): Promise<void>;                      // finalise the archive
  // + doc metadata: creator, created, modified, title, …, properties, calcProperties, views
}

class WorksheetWriter {
  addRow(values: RowValues): StreamedRow;       // StreamedRow = Row & { commit(): void }
  commit(): void;                               // finalise this sheet
  mergeCells(...args): void;
  addConditionalFormatting(cf): ConditionalFormatting;
  addImage(imageId: number, anchor): WorksheetImage;
  autoFilter?: string;                          // e.g. "A1:D1"
  views: WorksheetView[];                       // set before the first row commits
  properties;                                   // set before the first row commits
  dataValidations;
}
```
