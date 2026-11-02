import { UnsupportedFeatureError } from "../errors.js";
import type { NodeWritableLike } from "../xlsx/stream/output.js";

const unsupported = (): never => {
  throw new UnsupportedFeatureError(
    "filesystem access is unavailable in browser builds; use bytes or streams",
  );
};

export async function readFileBytes(_path: string, _maxInputBytes?: number): Promise<Uint8Array> {
  return unsupported();
}
export async function writeFileBytes(_path: string, _bytes: Uint8Array): Promise<void> {
  unsupported();
}
export async function openFileStream(_path: string): Promise<NodeWritableLike> {
  return unsupported();
}
