// Output sinks for the streaming writer (SPEC §1). The streaming zip pushes its byte chunks into a
// `ByteSink`; the writer wires one of three destinations to it:
//
//   • an in-memory async-iterable queue (the environment-agnostic default) — the `WorkbookWriter` is
//     itself `AsyncIterable<Uint8Array>` and yields from it;
//   • a Node `Writable` (e.g. `fs.createWriteStream(...)`), duck-typed so this file needs no `node:*`
//     import and browser bundles that bring their own stream stay clean;
//   • a file path, whose `node:fs` write stream is imported lazily inside {@link drainToFile} — the one
//     place `fs` is touched, isolated exactly like `writeWorkbookFile`.
//
// The zip driver is synchronous (fflate's `ZipDeflate`), so by the time `zip.end()` returns every chunk
// has already been handed to the sink; a `filename`/`stream` destination is drained asynchronously by a
// background pump the writer awaits in `commit()`.

/** Terminal destination for the streaming zip's byte chunks. Written to synchronously from the zip's
 *  `ondata` callback; `close`/`abort` fire exactly once at the end of the archive. */
export interface ByteSink {
  write(chunk: Uint8Array): void;
  /** No more chunks — the archive is complete. */
  close(): void;
  /** A fatal error occurred; surface it to any consumer. */
  abort(error: unknown): void;
}

/** The minimal slice of a Node `Writable` the writer relies on — duck-typed so this module imports no
 *  `node:*` and stays browser-safe. A real `fs.WriteStream`/`Writable` satisfies it structurally. */
export interface NodeWritableLike {
  write(chunk: Uint8Array): boolean;
  end(cb?: () => void): void;
  once(event: string, listener: (arg?: unknown) => void): void;
  destroy(error?: Error): void;
}

/** An in-memory {@link ByteSink} that is also an `AsyncIterable<Uint8Array>`: chunks written before a
 *  consumer asks are buffered; a pending `next()` is resolved as soon as a chunk arrives. Single
 *  consumer (one iterator) — matching how the writer is drained. */
export class QueueSink implements ByteSink, AsyncIterable<Uint8Array> {
  private readonly queue: Uint8Array[] = [];
  private pending?: {
    resolve: (r: IteratorResult<Uint8Array>) => void;
    reject: (e: unknown) => void;
  };
  private done = false;
  private failure?: unknown;

  write(chunk: Uint8Array): void {
    if (this.pending) {
      const p = this.pending;
      this.pending = undefined;
      p.resolve({ value: chunk, done: false });
    } else {
      this.queue.push(chunk);
    }
  }

  close(): void {
    this.done = true;
    if (this.pending) {
      const p = this.pending;
      this.pending = undefined;
      p.resolve({ value: undefined, done: true });
    }
  }

  abort(error: unknown): void {
    this.failure = error;
    if (this.pending) {
      const p = this.pending;
      this.pending = undefined;
      p.reject(error);
    }
  }

  [Symbol.asyncIterator](): AsyncIterator<Uint8Array> {
    return {
      next: (): Promise<IteratorResult<Uint8Array>> => {
        if (this.queue.length > 0) {
          return Promise.resolve({ value: this.queue.shift() as Uint8Array, done: false });
        }
        if (this.failure !== undefined) return Promise.reject(this.failure);
        if (this.done) return Promise.resolve({ value: undefined, done: true });
        return new Promise((resolve, reject) => {
          this.pending = { resolve, reject };
        });
      },
    };
  }
}

/** Pump an async byte stream into a Node `Writable`, honouring backpressure (awaiting `"drain"` when
 *  `write` returns `false`) and resolving once the destination has fully flushed (`end`'s callback). */
export async function drainToWritable(
  source: AsyncIterable<Uint8Array>,
  writable: NodeWritableLike,
): Promise<void> {
  for await (const chunk of source) {
    if (!writable.write(chunk)) {
      await new Promise<void>((resolve, reject) => {
        writable.once("drain", () => resolve());
        writable.once("error", (e) => reject(e));
      });
    }
  }
  await new Promise<void>((resolve, reject) => {
    writable.once("error", (e) => reject(e));
    writable.end(() => resolve());
  });
}

/** Pump an async byte stream to a file. The sole `node:fs` touch in the streaming writer — imported
 *  lazily so browser bundles that only use the async-iterable or `Writable` sinks never pull it in. */
export async function drainToFile(source: AsyncIterable<Uint8Array>, path: string): Promise<void> {
  const { createWriteStream } = await import("node:fs");
  const writable = createWriteStream(path) as unknown as NodeWritableLike;
  await drainToWritable(source, writable);
}
