import { defineConfig } from "tsup";

export default defineConfig({
  entry: { index: "src/index.ts", compat: "src/compat/index.ts" },
  format: ["esm", "cjs"],
  dts: true,
  sourcemap: true,
  clean: true,
  treeshake: true,
  target: "node22",
  outExtension({ format }) {
    return { js: format === "cjs" ? ".cjs" : ".js" };
  },
});
