// fitText() and its text measurement against python-pptx and Pillow on the upstream test
// font; regenerate the expectations with scripts/text_golden.py.
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { Inches, Presentation } from "../../src/index.js";
import { TrueTypeFont } from "../../src/text/ttf.js";

const fontPath = new URL("../test_files/calibriz.ttf", import.meta.url);
const golden: { bboxes: [number, string, number[]][]; fits: [string, number, number, number | "error"][] } = JSON.parse(
  readFileSync(new URL("../fixtures/text_golden.json", import.meta.url), "utf8"),
);

describe("fitText", () => {
  it("chooses the same font size as python-pptx", () => {
    const fontFile = new Uint8Array(readFileSync(fontPath));
    const prs = Presentation();
    const slide = prs.slides.addSlide(prs.slideLayouts.at(6));
    const got = golden.fits.map(([text, w, h]) => {
      const tb = slide.shapes.addTextbox(0, 0, Inches(w), Inches(h));
      tb.textFrame.text = text;
      try {
        tb.textFrame.fitText("Calibri", 40, true, true, fontFile);
        return (tb.textFrame.paragraphs.at(0).runs.at(0).font.size as number) / 12700;
      } catch (e) {
        // like python-pptx, a word wider than the box can't be fitted
        if (e instanceof TypeError) return "error";
        throw e;
      }
    });
    expect(got).toEqual(golden.fits.map((f) => f[3]));
  });
});

describe("text measurement vs Pillow", () => {
  const font = new TrueTypeFont(new Uint8Array(readFileSync(fontPath)));
  it("matches Pillow's widths and heights closely (FreeType hinting aside)", () => {
    let exactW = 0;
    let exactH = 0;
    for (const [size, text, b] of golden.bboxes) {
      const g = font.getbbox(text, size);
      const dw = g[2] - g[0] - (b[2] - b[0]);
      const dh = g[3] - g[1] - (b[3] - b[1]);
      expect(Math.abs(dw)).toBeLessThanOrEqual(5);
      expect(Math.abs(dh)).toBeLessThanOrEqual(2);
      if (dw === 0) exactW++;
      if (dh === 0) exactH++;
    }
    // today: widths exact in 104/110 cases, heights in 80/110
    expect(exactW).toBeGreaterThanOrEqual(100);
    expect(exactH).toBeGreaterThanOrEqual(75);
  });
});
