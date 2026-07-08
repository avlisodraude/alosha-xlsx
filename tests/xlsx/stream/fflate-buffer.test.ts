// Regression guard for a latent correctness assumption in the streaming zip (src/xlsx/stream/zip.ts):
// the `ondata` callback hands its chunk straight to the ByteSink, and QueueSink *retains* that chunk
// in a buffer until a consumer/pump pulls it. That is only safe if fflate hands back a fresh buffer
// per callback rather than reusing (and later overwriting) one shared output buffer. If a future
// fflate version reuses its output buffer, retained chunks would be silently corrupted and zip.ts must
// copy defensively. This test pins the current (safe) behavior so that regression is caught here.

import { Zip, ZipDeflate } from "fflate";
import { describe, expect, it } from "vitest";

describe("fflate output-buffer aliasing (streaming-zip safety invariant)", () => {
  it("emits a fresh, independently-backed buffer per ondata chunk (no reuse/mutation)", () => {
    // Retain each view without copying (exactly what QueueSink does) alongside a head-snapshot copied
    // the moment it arrived, so a later mutation of the shared buffer would show up as a mismatch.
    const chunks: { view: Uint8Array; headAtArrival: string }[] = [];

    const zip = new Zip((err, data, _final) => {
      if (err) throw err;
      if (data.length > 0) {
        chunks.push({ view: data, headAtArrival: Array.from(data.slice(0, 16)).join(",") });
      }
    });

    // Several entries with compressible payloads to force multiple ondata chunks mid-stream.
    for (let e = 0; e < 5; e++) {
      const entry = new ZipDeflate(`f${e}.txt`, { level: 6 });
      zip.add(entry);
      entry.push(new TextEncoder().encode(`row-${e}-`.repeat(20_000)), true);
    }
    zip.end();

    expect(chunks.length).toBeGreaterThan(1);

    // 1. No retained chunk was mutated after it was handed off.
    for (const { view, headAtArrival } of chunks) {
      expect(Array.from(view.slice(0, 16)).join(",")).toBe(headAtArrival);
    }

    // 2. Chunks don't alias one shared backing ArrayBuffer (which reuse would imply).
    const distinctBuffers = new Set(chunks.map((c) => c.view.buffer));
    expect(distinctBuffers.size).toBe(chunks.length);
  });
});
