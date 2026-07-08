---
"@alosha/xlsx": minor
---

Streaming writer: add `flush()` and `writeRows()` for flat-memory large exports, and fix a drain
listener leak.

A fully synchronous `addRow().commit()` loop never yields the event loop, so with a `filename`/`stream`
destination the whole compressed archive accumulated in the sink's buffer until `commit()`. Two
additions give callers an await point that lets the background pump drain mid-write:

- `await workbookWriter.flush()` — resolves once every chunk produced so far has reached the
  destination. Call it periodically in a large write loop. No-op on the async-iterable path (the
  consumer drives draining) or once committed.
- `await worksheetWriter.writeRows(rows, { flushEvery })` — adds and commits every row from a sync or
  async iterable, flushing every `flushEvery` rows (default 10,000). On a 400k-row poorly-compressible
  export, peak RSS dropped ~23% versus the un-flushed loop.

Also fixes a latent listener leak in the internal `Writable` drain pump (a per-backpressure `"error"`
listener was never removed, tripping Node's `MaxListenersExceededWarning` on long writes), and pins
fflate's per-chunk fresh-buffer behavior with a regression test — the streaming zip hands retained
chunks to the sink without copying, which is only safe as long as fflate does not reuse its output
buffer. No output change.
