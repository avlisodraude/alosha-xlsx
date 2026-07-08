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
  removeListener(event: string, listener: (arg?: unknown) => void): void;
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
  /** Resolvers waiting on {@link whenDrained}; settled once the buffer empties (or the stream ends). */
  private drainWaiters: Array<() => void> = [];

  write(chunk: Uint8Array): void {
    if (this.pending) {
      // Handed straight to a waiting consumer — never enters the queue, so the buffer stays empty.
      const p = this.pending;
      this.pending = undefined;
      p.resolve({ value: chunk, done: false });
    } else {
      this.queue.push(chunk);
    }
  }

  close(): void {
    this.done = true;
    this.settleDrainWaiters();
    if (this.pending) {
      const p = this.pending;
      this.pending = undefined;
      p.resolve({ value: undefined, done: true });
    }
  }

  abort(error: unknown): void {
    this.failure = error;
    this.settleDrainWaiters();
    if (this.pending) {
      const p = this.pending;
      this.pending = undefined;
      p.reject(error);
    }
  }

  /** Resolve once the buffered queue has drained to the consumer (or the stream has ended/failed).
   *  The writer's `flush()` awaits this so a caller can yield the event loop mid-write and let a
   *  background pump empty the queue, keeping peak memory flat instead of buffering the whole archive. */
  whenDrained(): Promise<void> {
    if (this.queue.length === 0 || this.done || this.failure !== undefined) {
      return Promise.resolve();
    }
    return new Promise<void>((resolve) => {
      this.drainWaiters.push(resolve);
    });
  }

  private settleDrainWaiters(): void {
    if (this.drainWaiters.length === 0) return;
    const waiters = this.drainWaiters;
    this.drainWaiters = [];
    for (const resolve of waiters) resolve();
  }

  [Symbol.asyncIterator](): AsyncIterator<Uint8Array> {
    return {
      next: (): Promise<IteratorResult<Uint8Array>> => {
        if (this.queue.length > 0) {
          const chunk = this.queue.shift() as Uint8Array;
          if (this.queue.length === 0) this.settleDrainWaiters();
          return Promise.resolve({ value: chunk, done: false });
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
      // Await "drain", removing BOTH listeners on either outcome. A backpressured write that resolves
      // via "drain" must not leave its "error" listener attached, or they accumulate one per chunk and
      // trip Node's MaxListenersExceededWarning on a long, frequently-draining write.
      await new Promise<void>((resolve, reject) => {
        const onDrain = (): void => {
          cleanup();
          resolve();
        };
        const onError = (e?: unknown): void => {
          cleanup();
          reject(e);
        };
        const cleanup = (): void => {
          writable.removeListener("drain", onDrain);
          writable.removeListener("error", onError);
        };
        writable.once("drain", onDrain);
        writable.once("error", onError);
      });
    }
  }
  await new Promise<void>((resolve, reject) => {
    const onError = (e?: unknown): void => reject(e);
    writable.once("error", onError);
    writable.end(() => {
      writable.removeListener("error", onError);
      resolve();
    });
  });
}

/** Pump an async byte stream to a file. The sole `node:fs` touch in the streaming writer — imported
 *  lazily so browser bundles that only use the async-iterable or `Writable` sinks never pull it in. */
export async function drainToFile(source: AsyncIterable<Uint8Array>, path: string): Promise<void> {
  const { createWriteStream } = await import("node:fs");
  const writable = createWriteStream(path) as unknown as NodeWritableLike;
  await drainToWritable(source, writable);
}
