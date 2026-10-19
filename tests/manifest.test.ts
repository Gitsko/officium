import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

const manifest = JSON.parse(readFileSync(new URL("../manifest.json", import.meta.url), "utf8"));

describe("manifest", () => {
  it("carries the Officium identity instead of the template defaults", () => {
    expect(manifest.name.short).toBe("Officium");
    expect(manifest.id).not.toBe("f2b103f1-1ab1-4e1b-8f0b-072aa3d4e19d");
    expect(JSON.stringify(manifest)).not.toMatch(/contoso/i);
  });
});
