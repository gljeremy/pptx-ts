// The image header reader (python-pptx uses Pillow) against what Pillow reports for each
// image; regenerate the expectations with scripts/image_golden.py.
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { readImageProps } from "../../src/image/headers.js";

const golden: Record<string, { format?: string; size?: [number, number]; dpi?: number | number[] | null; error?: string }> = JSON.parse(
  readFileSync(new URL("../fixtures/images_pillow.json", import.meta.url), "utf8"),
);
const root = new URL("../../", import.meta.url);

describe("readImageProps matches Pillow", () => {
  for (const [path, want] of Object.entries(golden)) {
    it(path, () => {
      const blob = new Uint8Array(readFileSync(new URL(path, root)));
      if (want.error) {
        expect(() => readImageProps(blob)).toThrow();
        return;
      }
      const [format, size, dpi] = readImageProps(blob);
      expect(format).toBe(want.format);
      if (want.format !== "EPS") expect(size).toEqual(want.size);
      if (Array.isArray(want.dpi)) {
        expect(Array.isArray(dpi)).toBe(true);
        (dpi as number[]).forEach((d, i) => expect(d).toBeCloseTo((want.dpi as number[])[i], 9));
      } else expect(dpi).toEqual(want.dpi);
    });
  }
});
