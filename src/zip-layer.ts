// ZIP abstraction layer. Uses fflate under the hood for reading/writing OOXML archives.

import { unzipSync, zipSync } from "fflate";
import { InvalidPackageError } from "./errors.js";
import { checkReadLimit, type ReadLimits, resolveReadLimits } from "./read-limits.js";

/** Entries of a ZIP archive, keyed by path (e.g. "xl/worksheets/sheet1.xml"). */
export type ZipEntries = Map<string, Uint8Array>;

/** Reads a ZIP archive buffer into a map of entry path -> raw bytes. */
function readZipArchive(data: Uint8Array, options?: ReadLimits): ZipEntries {
  const limits = resolveReadLimits(options);
  checkReadLimit("maxInputBytes", limits.maxInputBytes, data.byteLength);
  let count = 0;
  let total = 0;
  const unzipped = unzipSync(data, {
    filter(file) {
      checkReadLimit("maxEntries", limits.maxEntries, ++count);
      // Stored entries copy size bytes, whereas DEFLATE allocates originalSize bytes. Validate both
      // fields for stored entries so a forged, smaller originalSize cannot bypass the budget.
      const size =
        file.compression === 0 ? Math.max(file.size, file.originalSize) : file.originalSize;
      if (!Number.isSafeInteger(size) || size < 0)
        throw new InvalidPackageError("invalid ZIP entry size");
      checkReadLimit("maxEntryBytes", limits.maxEntryBytes, size);
      total += size;
      checkReadLimit("maxTotalBytes", limits.maxTotalBytes, total);
      return true;
    },
  });
  return new Map(Object.entries(unzipped));
}

/** Writes a map of entry path -> raw bytes into a ZIP archive buffer. */
function writeZipArchive(entries: ZipEntries): Uint8Array {
  return zipSync(Object.fromEntries(entries));
}

export { readZipArchive, writeZipArchive };
