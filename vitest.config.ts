import { fileURLToPath } from "node:url";
import { defineConfig } from "vitest/config";

export default defineConfig({
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
