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

import { ExportAbortedError, StreamBufferLimitError } from "../../errors.js";
import { openFileStream } from "../../platform/io.js";

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
  private queue: Array<Uint8Array | undefined> = [];
  private head = 0;
  private pending?: {
    resolve: (r: IteratorResult<Uint8Array>) => void;
    reject: (e: unknown) => void;
  };
  private done = false;
  private failed = false;
  private failure?: unknown;
  private bytes = 0;
  private claimed = false;
  private waiters: Array<{
    capacity: boolean;
    resolve: () => void;
    reject: (error: unknown) => void;
  }> = [];
  private readonly failures = new Set<(error: unknown) => void>();
  private completion?: () => void;

  constructor(
    readonly highWaterMarkBytes = 1024 * 1024,
    readonly maxBufferedBytes = 16 * 1024 * 1024,
  ) {
    for (const value of [highWaterMarkBytes, maxBufferedBytes]) {
      if (!Number.isSafeInteger(value) || value < 1)
        throw new RangeError("stream byte limits must be positive safe integers");
    }
    if (highWaterMarkBytes > maxBufferedBytes)
      throw new RangeError("highWaterMarkBytes must not exceed maxBufferedBytes");
  }

  get bufferedBytes(): number {
    return this.bytes;
  }
  throwIfFailed(): void {
    if (this.failed) throw this.failure;
  }
  onFailure(listener: (error: unknown) => void): () => void {
    if (this.failed) listener(this.failure);
    else this.failures.add(listener);
    return () => {
      this.failures.delete(listener);
    };
  }
  onComplete(listener: () => void): void {
    this.completion = listener;
    this.settle();
  }

  write(chunk: Uint8Array): void {
    this.throwIfFailed();
    if (this.done) throw new Error("stream output is already closed");
    const size = this.bytes + chunk.byteLength;
    if (size > this.maxBufferedBytes) {
      const error = new StreamBufferLimitError(this.maxBufferedBytes, size);
      this.abort(error);
      throw error;
    }
    this.bytes = size;
    if (this.pending) {
      const pending = this.pending;
      this.pending = undefined;
      pending.resolve({ value: chunk, done: false });
    } else this.queue.push(chunk);
  }

  close(): void {
    this.done = true;
    if (this.pending) {
      const pending = this.pending;
      this.pending = undefined;
      pending.resolve({ value: undefined, done: true });
    }
    this.settle();
  }

  abort(error: unknown): void {
    if (this.failed) return;
    this.failed = true;
    this.failure = error;
    this.queue = [];
    this.head = 0;
    this.bytes = 0;
    if (this.pending) {
      const pending = this.pending;
      this.pending = undefined;
      pending.reject(error);
    }
    this.settle();
    for (const listener of this.failures) listener(error);
    this.failures.clear();
    this.completion = undefined;
  }

  acknowledge(size: number): void {
    if (!this.failed) this.bytes -= size;
    this.settle();
  }
  whenDrained(): Promise<void> {
    return this.wait(false);
  }
  whenReady(): Promise<void> {
    return this.wait(true);
  }
  private wait(capacity: boolean): Promise<void> {
    if (this.failed) return Promise.reject(this.failure);
    if (capacity ? this.bytes < this.highWaterMarkBytes : this.bytes === 0)
      return Promise.resolve();
    return new Promise((resolve, reject) => {
      this.waiters.push({ capacity, resolve, reject });
    });
  }
  private settle(): void {
    const remaining: typeof this.waiters = [];
    for (const waiter of this.waiters) {
      if (this.failed) waiter.reject(this.failure);
      else if (waiter.capacity ? this.bytes < this.highWaterMarkBytes : this.bytes === 0)
        waiter.resolve();
      else remaining.push(waiter);
    }
    this.waiters = remaining;
    if (!this.failed && this.done && this.bytes === 0 && this.completion) {
      const complete = this.completion;
      this.completion = undefined;
      complete();
    }
  }

  /** Await a single operation while observing cancellation, without retaining one promise reaction
   * per chunk on a long-lived rejection promise. */
  withCancellation<T>(operation: Promise<T>): Promise<T> {
    return new Promise<T>((resolve, reject) => {
      const unsubscribe = this.onFailure(reject);
      operation.then(
        (value) => {
          unsubscribe();
          if (this.failed) reject(this.failure);
          else resolve(value);
        },
        (error: unknown) => {
          unsubscribe();
          reject(error);
        },
      );
    });
  }

  iterator(manualAcknowledgement = false): AsyncIterator<Uint8Array> {
    if (this.claimed) throw new Error("stream output supports one consumer");
    this.claimed = true;
    let delivered = 0;
    let reading = false;
    let finished = false;
    return {
      next: async () => {
        if (reading) throw new Error("concurrent stream next() calls are unsupported");
        if (finished) return { value: undefined, done: true };
        this.throwIfFailed();
        if (!manualAcknowledgement && delivered) {
          this.acknowledge(delivered);
          delivered = 0;
        }
        reading = true;
        try {
          let result: IteratorResult<Uint8Array>;
          if (this.head < this.queue.length) {
            const value = this.queue[this.head] as Uint8Array;
            this.queue[this.head++] = undefined;
            if (this.head === this.queue.length) {
              this.queue = [];
              this.head = 0;
            }
            result = { value, done: false };
          } else if (this.done) result = { value: undefined, done: true };
          else
            result = await new Promise<IteratorResult<Uint8Array>>((resolve, reject) => {
              this.pending = { resolve, reject };
            });
          if (result.done) finished = true;
          else delivered = result.value.byteLength;
          return result;
        } finally {
          reading = false;
        }
      },
      return: async () => {
        if (!finished)
          this.abort(new ExportAbortedError("output consumer stopped before completion"));
        finished = true;
        return { value: undefined, done: true };
      },
      throw: async (error) => {
        this.abort(error);
        finished = true;
        throw error;
      },
    };
  }
  [Symbol.asyncIterator](): AsyncIterator<Uint8Array> {
    return this.iterator();
  }
}

/** Pump an async byte stream into a Node `Writable`, honouring backpressure (awaiting `"drain"` when
 *  `write` returns `false`) and resolving once the destination has fully flushed (`end`'s callback). */
export async function drainToWritable(
  source: AsyncIterable<Uint8Array>,
  writable: NodeWritableLike,
): Promise<void> {
  let rejectPending: ((error: unknown) => void) | undefined;
  let failed = false;
  let failure: unknown;
  let finished = false;
  const fail = (error: unknown): void => {
    if (failed) return;
    failed = true;
    failure = error;
    if (source instanceof QueueSink) source.abort(error);
    rejectPending?.(error);
  };
  // Subscribe only the current operation. Racing every chunk against one never-settled promise
  // would accumulate a promise reaction per chunk for the entire lifetime of a large export.
  const waitFor = <T>(operation: Promise<T>): Promise<T> =>
    new Promise<T>((resolve, reject) => {
      rejectPending = reject;
      operation.then((value) => {
        rejectPending = undefined;
        if (failed) reject(failure);
        else resolve(value);
      }, reject);
      if (failed) reject(failure);
    });
  const onError = (error?: unknown): void => {
    fail(error);
  };
  const onClose = (): void => {
    if (!finished) fail(new Error("export destination closed before completion"));
  };
  // Watch the whole lifecycle, including periods when the producer has supplied no bytes yet.
  writable.once("error", onError);
  writable.once("close", onClose);
  const iterator =
    source instanceof QueueSink ? source.iterator(true) : source[Symbol.asyncIterator]();
  const unsubscribe = source instanceof QueueSink ? source.onFailure(fail) : undefined;
  try {
    for (;;) {
      const item = await waitFor(iterator.next());
      if (item.done) break;
      if (!writable.write(item.value)) {
        let onDrain!: () => void;
        try {
          await waitFor(
            new Promise<void>((resolve) => {
              onDrain = resolve;
              writable.once("drain", onDrain);
            }),
          );
        } finally {
          writable.removeListener("drain", onDrain);
        }
      }
      if (source instanceof QueueSink) source.acknowledge(item.value.byteLength);
    }
    await waitFor(
      new Promise<void>((resolve) => {
        writable.end((error) => {
          if (error) {
            fail(error);
            return;
          }
          finished = true;
          resolve();
        });
      }),
    );
  } catch (error) {
    if (source instanceof QueueSink) source.abort(error);
    writable.destroy();
    throw error;
  } finally {
    writable.removeListener("error", onError);
    writable.removeListener("close", onClose);
    unsubscribe?.();
  }
}

/** Pump an async byte stream to a file. The sole `node:fs` touch in the streaming writer — imported
 *  lazily so browser bundles that only use the async-iterable or `Writable` sinks never pull it in. */
export async function drainToFile(source: AsyncIterable<Uint8Array>, path: string): Promise<void> {
  const writable = await openFileStream(path);
  await drainToWritable(source, writable);
}
