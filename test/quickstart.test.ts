// The "Getting started" examples in README.md (python-pptx's docs/user/quickstart.rst in
// TypeScript), checked part by part against python-pptx's output for the same steps
// (test/fixtures/quickstart, written by scripts/quickstart_golden.py).
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { expect, it } from "vitest";
import { unzipSync } from "fflate";
import { BytesIO, Inches, MSO_SHAPE, Presentation, Pt } from "../src/index.js";

const imgPath = fileURLToPath(new URL("./features/test_files/monty-truth.png", import.meta.url));

const examples: Record<string, () => ReturnType<typeof Presentation>> = {
  hello_world() {
    const prs = Presentation();
    const titleSlideLayout = prs.slideLayouts.at(0);
    const slide = prs.slides.addSlide(titleSlideLayout);
    const title = slide.shapes.title!;
    const subtitle = slide.placeholders.get(1);
    title.text = "Hello, World!";
    subtitle.text = "python-pptx was here!";
    return prs;
  },

  bullet_slide() {
    const prs = Presentation();
    const bulletSlideLayout = prs.slideLayouts.at(1);
    const slide = prs.slides.addSlide(bulletSlideLayout);
    const shapes = slide.shapes;
    const titleShape = shapes.title!;
    const bodyShape = shapes.placeholders.get(1);
    titleShape.text = "Adding a Bullet Slide";
    const tf = bodyShape.textFrame;
    tf.text = "Find the bullet slide layout";
    let p = tf.addParagraph();
    p.text = "Use _TextFrame.text for first bullet";
    p.level = 1;
    p = tf.addParagraph();
    p.text = "Use _TextFrame.add_paragraph() for subsequent bullets";
    p.level = 2;
    return prs;
  },

  add_textbox() {
    const prs = Presentation();
    const blankSlideLayout = prs.slideLayouts.at(6);
    const slide = prs.slides.addSlide(blankSlideLayout);
    const left = Inches(1), top = Inches(1), width = Inches(1), height = Inches(1);
    const txBox = slide.shapes.addTextbox(left, top, width, height);
    const tf = txBox.textFrame;
    tf.text = "This is text inside a textbox";
    let p = tf.addParagraph();
    p.text = "This is a second paragraph that's bold";
    p.font.bold = true;
    p = tf.addParagraph();
    p.text = "This is a third paragraph that's big";
    p.font.size = Pt(40);
    return prs;
  },

  add_picture() {
    const prs = Presentation();
    const blankSlideLayout = prs.slideLayouts.at(6);
    const slide = prs.slides.addSlide(blankSlideLayout);
    let left = Inches(1);
    const top = Inches(1);
    slide.shapes.addPicture(imgPath, left, top);
    left = Inches(5);
    const height = Inches(5.5);
    slide.shapes.addPicture(imgPath, left, top, null, height);
    return prs;
  },

  add_shape() {
    const prs = Presentation();
    const titleOnlySlideLayout = prs.slideLayouts.at(5);
    const slide = prs.slides.addSlide(titleOnlySlideLayout);
    const shapes = slide.shapes;
    shapes.title!.text = "Adding an AutoShape";
    let left = Inches(0.93); // 0.93" centers this overall set of shapes
    const top = Inches(3.0);
    let width = Inches(1.75);
    const height = Inches(1.0);
    let shape = shapes.addShape(MSO_SHAPE.PENTAGON, left, top, width, height);
    shape.text = "Step 1";
    left = left + width - Inches(0.4);
    width = Inches(2.0); // chevrons need more width for visual balance
    for (let n = 2; n < 6; n++) {
      shape = shapes.addShape(MSO_SHAPE.CHEVRON, left, top, width, height);
      shape.text = `Step ${n}`;
      left = left + width - Inches(0.4);
    }
    return prs;
  },

  add_table() {
    const prs = Presentation();
    const titleOnlySlideLayout = prs.slideLayouts.at(5);
    const slide = prs.slides.addSlide(titleOnlySlideLayout);
    const shapes = slide.shapes;
    shapes.title!.text = "Adding a Table";
    const rows = 2, cols = 2;
    const left = Inches(2.0), top = Inches(2.0);
    const width = Inches(6.0);
    const height = Inches(0.8);
    const table = shapes.addTable(rows, cols, left, top, width, height).table;
    // set column widths
    table.columns.at(0).width = Inches(2.0);
    table.columns.at(1).width = Inches(4.0);
    // write column headings
    table.cell(0, 0).text = "Foo";
    table.cell(0, 1).text = "Bar";
    // write body cells
    table.cell(1, 0).text = "Baz";
    table.cell(1, 1).text = "Qux";
    return prs;
  },
};

for (const [name, build] of Object.entries(examples)) {
  it(`quickstart ${name} matches python-pptx part by part`, () => {
    const out = new BytesIO();
    build().save(out);
    const ts = unzipSync(out.getvalue());
    const py = unzipSync(new Uint8Array(readFileSync(new URL(`./fixtures/quickstart/${name}.pptx`, import.meta.url))));
    expect(Object.keys(ts).sort()).toEqual(Object.keys(py).sort());
    for (const part of Object.keys(py)) {
      expect(ts[part], part).toEqual(py[part]);
    }
  });
}

it("extracts all text from slides in a presentation", () => {
  const prs = Presentation(fileURLToPath(new URL("./features/test_files/txt-text.pptx", import.meta.url)));
  // textRuns will be populated with a list of strings, one for each text run in the presentation
  const textRuns: string[] = [];
  for (const slide of prs.slides) {
    for (const shape of slide.shapes) {
      if (!shape.hasTextFrame) continue;
      for (const paragraph of shape.textFrame.paragraphs) {
        for (const run of paragraph.runs) textRuns.push(run.text);
      }
    }
  }
  expect(textRuns.length).toBeGreaterThan(0);
});
