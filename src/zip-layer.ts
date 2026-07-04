// ZIP abstraction layer. Uses fflate under the hood for reading/writing OOXML archives.

function readZipArchive(): never {
  throw new Error("not implemented");
}

function writeZipArchive(): never {
  throw new Error("not implemented");
}

export { readZipArchive, writeZipArchive };
