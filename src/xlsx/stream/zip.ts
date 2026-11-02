// Incremental ZIP over fflate's streaming `Zip` (SPEC §1's "incremental zip"). Unlike `zipSync`, which
// needs every part's bytes up front, this opens one `ZipDeflate` entry at a time and deflates chunks as
// they are pushed — so a worksheet part grows on the wire as rows commit, with nothing buffered but the
// deflate window. The `Zip` `ondata` callback feeds a {@link ByteSink}.
//
// Entries are strictly sequential: `startEntry` → `push(..)*` → `push(.., true)`, then the next entry.
// fflate can interleave concurrent entries, but sequencing keeps peak memory to a single open part and
// mirrors how the writer commits sheets one after another.

import { Zip, ZipDeflate } from "fflate";
import type { ByteSink } from "./output.js";

export class StreamingZip {
  private readonly zip: Zip;
  private current: ZipDeflate | undefined;
  private ended = false;

  constructor(sink: ByteSink) {
    this.zip = new Zip((err, data, final) => {
      if (err) {
        sink.abort(err);
        return;
      }
      // The sink may retain this chunk (QueueSink buffers it until a consumer/pump pulls it), so we
      // rely on fflate handing back a fresh buffer per `ondata` call rather than reusing one — pinned
      // by a regression test (tests/xlsx/stream/fflate-buffer.test.ts). If a future fflate reuses its
      // output buffer, copy here (`data.slice()`) before handing it off. `final` closes the whole
      // archive (not an individual entry's last chunk).
      if (data.length > 0) sink.write(data);
      if (final) sink.close();
    });
  }

  /** Open a new deflate entry at `path`. The previous entry must already be finalised. */
  startEntry(path: string): void {
    if (this.current) throw new Error("streaming zip: previous entry not finalised");
    const entry = new ZipDeflate(path, { level: 6 });
    this.current = entry;
    this.zip.add(entry);
  }

  /** Push a chunk into the open entry; `final` closes the entry (but not the archive). */
  push(bytes: Uint8Array, final = false): void {
    if (!this.current) throw new Error("streaming zip: no open entry");
    this.current.push(bytes, final);
    if (final) this.current = undefined;
  }

  /** Whether an entry is currently open for pushing. */
  get hasOpenEntry(): boolean {
    return this.current !== undefined;
  }

  /** Write a complete part in a single deflate entry. */
  addFile(path: string, bytes: Uint8Array): void {
    this.startEntry(path);
    this.push(bytes, true);
  }

  /** Finalise the archive: emit the central directory and drive the sink's final `close`. */
  end(): void {
    if (this.ended) return;
    if (this.current) throw new Error("streaming zip: entry still open at end()");
    this.ended = true;
    this.zip.end();
  }

  abort(): void {
    this.ended = true;
    this.current = undefined;
    this.zip.terminate();
  }
}
