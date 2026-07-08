import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { defineConfig } from "tsup";

const pkg = JSON.parse(
  readFileSync(fileURLToPath(new URL("./package.json", import.meta.url)), "utf8"),
) as { version: string };

export default defineConfig((options) => ({
  entry: { index: "src/index.ts", compat: "src/compat/index.ts" },
  format: ["esm", "cjs"],
  dts: true,
  // Sourcemaps roughly triple the published tarball (~1.7MB of `.map`); emit them only for
  // `pnpm dev` (watch) builds, never for the published `pnpm build`.
  sourcemap: options.watch === true,
  clean: true,
  treeshake: true,
  target: "node22",
  // Inject the real version so `export const version` matches package.json in the built output.
  define: {
    __PKG_VERSION__: JSON.stringify(pkg.version),
  },
  outExtension({ format }) {
    return { js: format === "cjs" ? ".cjs" : ".js" };
  },
}));
