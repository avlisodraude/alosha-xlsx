import { describe, expect, it } from "vitest";
import { hello, version } from "../src/index.js";

describe("@alosha/xlsx", () => {
  it("exposes a version string", () => {
    expect(version).toBe("0.0.0");
  });

  it("greets the world by default", () => {
    expect(hello()).toBe("Hello, world, from @alosha/xlsx!");
  });

  it("greets a given name", () => {
    expect(hello("Alosha")).toBe("Hello, Alosha, from @alosha/xlsx!");
  });
});
