import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { defineConfig } from "vitest/config";

const pkg = JSON.parse(
  readFileSync(fileURLToPath(new URL("./package.json", import.meta.url)), "utf8"),
) as { version: string };

export default defineConfig({
  // Mirror tsup's build-time `define` so `version` resolves to the real package.json version under
  // test (the build injects the same constant), instead of the "0.0.0" ReferenceError sentinel.
  define: {
    __PKG_VERSION__: JSON.stringify(pkg.version),
  },
  test: {
    include: ["tests/**/*.test.ts"],
    environment: "node",
    globals: false,
  },
  resolve: {
    alias: {
      // Resolve the published subpath to source so tests don't depend on a prior `pnpm build`;
      // real consumers still go through package.json's `exports` -> `dist/compat.*`.
      "@alosha/xlsx/compat": fileURLToPath(new URL("./src/compat/index.ts", import.meta.url)),
    },
  },
});
