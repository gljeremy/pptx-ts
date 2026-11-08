// The README's example, end to end; `scripts/crosscheck` and the README note that the same
// steps in python-pptx produce byte-identical parts.
import { expect, it } from "vitest";
import { BytesIO, Presentation } from "../src/index.js";

it("runs the README example", () => {
  const prs = Presentation();
  const slide = prs.slides.addSlide(prs.slideLayouts.at(0));
  slide.shapes.title!.textFrame.text = "Hello, pptx-ts";
  slide.placeholders.get(1).textFrame.text = "built in TypeScript";
  prs.coreProperties.title = "Hello";
  expect([...prs.slides].map((s) => [s.slideId, s.slideLayout.name])).toEqual([[256, "Title Slide"]]);

  const out = new BytesIO();
  prs.save(out);
  const reopened = Presentation(out.getvalue());
  expect([...reopened.slides.at(0).shapes].map((sh: any) => sh.textFrame.text)).toEqual(["Hello, pptx-ts", "built in TypeScript"]);
  expect(reopened.coreProperties.title).toBe("Hello");
});
