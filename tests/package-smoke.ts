// Exercise the actual packed package, including conditional exports and downstream type resolution.
import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import {
  mkdir,
  mkdtemp,
  readdir,
  readFile,
  realpath,
  rm,
  symlink,
  writeFile,
} from "node:fs/promises";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { runInNewContext } from "node:vm";
import { build } from "esbuild";

const require = createRequire(import.meta.url);
const temp = await mkdtemp(join(tmpdir(), "alosha-package-"));
try {
  execFileSync("pnpm", ["pack", "--pack-destination", temp], { stdio: "pipe" });
  const archive = (await readdir(temp)).find((name) => name.endsWith(".tgz"));
  assert(archive);
  execFileSync("tar", ["-xzf", join(temp, archive), "-C", temp]);
  await mkdir(join(temp, "node_modules/@alosha"), { recursive: true });
  await symlink(join(temp, "package"), join(temp, "node_modules/@alosha/xlsx"), "dir");
  const fflate = dirname(await realpath(require.resolve("fflate/package.json")));
  await symlink(fflate, join(temp, "node_modules/fflate"), "dir");

  const consumer = `
import assert from "node:assert/strict";
import { Workbook, version } from "@alosha/xlsx";
import ExcelJS from "@alosha/xlsx/compat";
const wb = new Workbook();
wb.addWorksheet("S").getCell("A1").value = "packed";
const loaded = new Workbook();
await loaded.xlsx.load(await wb.xlsx.writeBuffer());
assert.equal(loaded.getWorksheet("S")?.getCell("A1").value, "packed");
assert.equal(typeof version, "string");
assert.equal(new ExcelJS.Workbook().addWorksheet("C").name, "C");
`;
  await writeFile(join(temp, "consumer.mjs"), consumer);
  execFileSync(process.execPath, [join(temp, "consumer.mjs")], { stdio: "pipe" });
  await writeFile(
    join(temp, "consumer.cjs"),
    `
const assert = require("node:assert/strict");
const { Workbook } = require("@alosha/xlsx");
const { default: ExcelJS } = require("@alosha/xlsx/compat");
assert.equal(new Workbook().addWorksheet("S").name, "S");
assert.equal(new ExcelJS.Workbook().addWorksheet("C").name, "C");
`,
  );
  execFileSync(process.execPath, [join(temp, "consumer.cjs")], { stdio: "pipe" });

  const typed = `import { Workbook, type ReadOptions } from "@alosha/xlsx";
import ExcelJS from "@alosha/xlsx/compat";
const wb: Workbook = new Workbook();
const limits: ReadOptions = { maxCells: 100, maxRows: 100, maxMergeCells: 100, maxColumnDefinitions: 100 };
void wb.xlsx.load(new Uint8Array(), limits).catch(() => undefined);
wb.addWorksheet("S").getCell("A1").value = "typed";
new ExcelJS.Workbook().addWorksheet("C");
`;
  await writeFile(join(temp, "consumer.mts"), typed);
  await writeFile(join(temp, "consumer.cts"), typed);
  await writeFile(
    join(temp, "tsconfig.json"),
    JSON.stringify({
      compilerOptions: {
        target: "ES2022",
        module: "NodeNext",
        moduleResolution: "NodeNext",
        strict: true,
        noEmit: true,
        types: [],
      },
      files: ["consumer.mts", "consumer.cts"],
    }),
  );
  execFileSync(
    process.execPath,
    [require.resolve("typescript/bin/tsc"), "-p", join(temp, "tsconfig.json")],
    { stdio: "pipe" },
  );
  execFileSync(
    process.execPath,
    [
      require.resolve("typescript/bin/tsc"),
      "-p",
      join(temp, "tsconfig.json"),
      "--customConditions",
      "browser",
    ],
    { stdio: "pipe" },
  );

  const browser = await build({
    stdin: {
      resolveDir: temp,
      contents: `
import { Workbook, WorkbookWriter } from "@alosha/xlsx";
import ExcelJS from "@alosha/xlsx/compat";
globalThis.smoke = async () => {
  const wb = new Workbook(); wb.addWorksheet("S").getCell("A1").value = "browser";
  const loaded = new Workbook(); await loaded.xlsx.load(await wb.xlsx.writeBuffer(), { maxCells: 1, maxRows: 1 });
  if (loaded.getWorksheet("S").getCell("A1").value !== "browser") throw Error("buffer roundtrip");
  if (new ExcelJS.Workbook().addWorksheet("C").name !== "C") throw Error("compat export");
  const writer = new WorkbookWriter(); writer.addWorksheet("Stream").addRow([7]).commit();
  await writer.commit(); const chunks = []; for await (const c of writer) chunks.push(c);
  const bytes = new Uint8Array(chunks.reduce((n,c) => n+c.length,0));
  let offset = 0; for (const c of chunks) { bytes.set(c,offset); offset += c.length; }
  const parsed = new Workbook(); await parsed.xlsx.load(bytes);
  if (parsed.getWorksheet("Stream").getCell("A1").value !== 7) throw Error("stream roundtrip");
  try { await wb.xlsx.readFile("unavailable.xlsx"); } catch(e) {
    if (e.code === "UNSUPPORTED_FEATURE") return;
    throw e;
  }
  throw Error("browser file access should reject");
};`,
    },
    bundle: true,
    platform: "browser",
    format: "iife",
    write: false,
    metafile: true,
  });
  assert(Object.keys(browser.metafile.inputs).some((path) => path.includes("dist-browser")));
  const context: { smoke?: () => Promise<void> } = {};
  Object.assign(context, { TextEncoder, TextDecoder, Uint8Array });
  runInNewContext(browser.outputFiles[0]?.text ?? "", context);
  assert(context.smoke);
  await context.smoke();
  const packed = JSON.parse(await readFile(join(temp, "package/package.json"), "utf8")) as {
    version: string;
  };
  process.stdout.write(
    `Package ${packed.version}: ESM, CommonJS, NodeNext types, and browser consumers passed.\n`,
  );
} finally {
  await rm(temp, { recursive: true, force: true });
}
