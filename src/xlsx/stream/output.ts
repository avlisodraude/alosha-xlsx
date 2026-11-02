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
  end(cb?: (error?: Error | null) => void): void;
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
  private failed = false;
  private failure?: unknown;
  private outstanding = 0;
  /** Resolvers waiting on {@link whenDrained}; settled once the buffer empties (or the stream ends). */
  private drainWaiters: Array<{ resolve: () => void; reject: (error: unknown) => void }> = [];

  throwIfFailed(): void {
    if (this.failed) throw this.failure;
  }

  write(chunk: Uint8Array): void {
    this.throwIfFailed();
    this.outstanding++;
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
    if (this.failed) return;
    this.failed = true;
    this.failure = error;
    this.queue.length = 0;
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
    if (this.failed) return Promise.reject(this.failure);
    if (this.outstanding === 0) {
      return Promise.resolve();
    }
    return new Promise<void>((resolve, reject) => {
      this.drainWaiters.push({ resolve, reject });
    });
  }

  /** Called by the destination pump after a chunk has been accepted, including backpressure. */
  acknowledge(): void {
    this.outstanding--;
    if (this.outstanding === 0) this.settleDrainWaiters();
  }

  private settleDrainWaiters(): void {
    if (!this.failed && this.outstanding !== 0) return;
    if (this.drainWaiters.length === 0) return;
    const waiters = this.drainWaiters;
    this.drainWaiters = [];
    for (const waiter of waiters) {
      if (this.failed) waiter.reject(this.failure);
      else waiter.resolve();
    }
  }

  [Symbol.asyncIterator](): AsyncIterator<Uint8Array> {
    return {
      next: (): Promise<IteratorResult<Uint8Array>> => {
        if (this.failed) return Promise.reject(this.failure);
        if (this.queue.length > 0) {
          const chunk = this.queue.shift() as Uint8Array;
          return Promise.resolve({ value: chunk, done: false });
        }
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
  let rejectFailure!: (error: unknown) => void;
  let finished = false;
  const failure = new Promise<never>((_resolve, reject) => {
    rejectFailure = reject;
  });
  const onError = (error?: unknown): void => {
    rejectFailure(error);
  };
  const onClose = (): void => {
    if (!finished) rejectFailure(new Error("export destination closed before completion"));
  };
  // Watch the whole lifecycle, including periods when the producer has supplied no bytes yet.
  writable.once("error", onError);
  writable.once("close", onClose);
  const iterator = source[Symbol.asyncIterator]();
  try {
    for (;;) {
      const item = await Promise.race([iterator.next(), failure]);
      if (item.done) break;
      if (!writable.write(item.value)) {
        let onDrain!: () => void;
        try {
          await Promise.race([
            new Promise<void>((resolve) => {
              onDrain = resolve;
              writable.once("drain", onDrain);
            }),
            failure,
          ]);
        } finally {
          writable.removeListener("drain", onDrain);
        }
      }
      if (source instanceof QueueSink) source.acknowledge();
    }
    await Promise.race([
      new Promise<void>((resolve) => {
        writable.end((error) => {
          if (error) {
            rejectFailure(error);
            return;
          }
          finished = true;
          resolve();
        });
      }),
      failure,
    ]);
  } catch (error) {
    if (source instanceof QueueSink) source.abort(error);
    writable.destroy();
    throw error;
  } finally {
    writable.removeListener("error", onError);
    writable.removeListener("close", onClose);
  }
}

/** Pump an async byte stream to a file. The sole `node:fs` touch in the streaming writer — imported
 *  lazily so browser bundles that only use the async-iterable or `Writable` sinks never pull it in. */
export async function drainToFile(source: AsyncIterable<Uint8Array>, path: string): Promise<void> {
  const { createWriteStream } = await import("node:fs");
  const writable = createWriteStream(path) as unknown as NodeWritableLike;
  await drainToWritable(source, writable);
}
