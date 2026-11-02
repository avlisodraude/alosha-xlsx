import { checkReadLimit } from "../read-limits.js";
import type { NodeWritableLike } from "../xlsx/stream/output.js";

/** Node file operations. Browser builds replace this module with explicit unsupported stubs. */
export async function readFileBytes(path: string, maxInputBytes = Infinity): Promise<Uint8Array> {
  const { createReadStream } = await import("node:fs");
  const chunks: Uint8Array[] = [];
  let total = 0;
  // At most one byte beyond the budget is needed to distinguish an exact fit from overflow.
  const stream = createReadStream(path, {
    end: Number.isFinite(maxInputBytes) ? maxInputBytes : undefined,
  });
  for await (const chunk of stream) {
    const bytes = chunk as Uint8Array;
    total += bytes.byteLength;
    checkReadLimit("maxInputBytes", maxInputBytes, total);
    chunks.push(bytes);
  }
  const bytes = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return bytes;
}

export async function writeFileBytes(path: string, bytes: Uint8Array): Promise<void> {
  const { writeFile } = await import("node:fs/promises");
  await writeFile(path, bytes);
}

export async function openFileStream(path: string): Promise<NodeWritableLike> {
  const { createWriteStream } = await import("node:fs");
  return createWriteStream(path);
}
