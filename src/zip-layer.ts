// ZIP abstraction layer. Uses fflate under the hood for reading/writing OOXML archives.

import { unzipSync, zipSync } from "fflate";

/** Entries of a ZIP archive, keyed by path (e.g. "xl/worksheets/sheet1.xml"). */
export type ZipEntries = Map<string, Uint8Array>;

/** Reads a ZIP archive buffer into a map of entry path -> raw bytes. */
function readZipArchive(data: Uint8Array): ZipEntries {
  const unzipped = unzipSync(data);
  return new Map(Object.entries(unzipped));
}

/** Writes a map of entry path -> raw bytes into a ZIP archive buffer. */
function writeZipArchive(entries: ZipEntries): Uint8Array {
  return zipSync(Object.fromEntries(entries));
}

export { readZipArchive, writeZipArchive };
