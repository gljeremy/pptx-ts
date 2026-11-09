// pptx-ts stage-2 workloads (Presentation API), for comparison with bench_prs.py.
// Usage: npm run build && node bench/bench_prs.mjs <deck.pptx> [acceptance test_files dir]
import fs from "node:fs";
import path from "node:path";
import { Presentation, BytesIO, CategoryChartData, XL_CHART_TYPE, MSO_SHAPE, RGBColor } from "../dist/index.js";

const data = new Uint8Array(fs.readFileSync(process.argv[2]));
const filesDir = process.argv[3] ?? "test/features/test_files";
const decks = fs.readdirSync(filesDir).filter((f) => f.endsWith(".pptx")).sort().map((f) => new Uint8Array(fs.readFileSync(path.join(filesDir, f))));
const image = new Uint8Array(fs.readFileSync(path.join(filesDir, "monty-truth.png")));
const best = (fn, n = 5) => {
  let m = Infinity;
  for (let i = 0; i < n; i++) {
    const s = performance.now();
    fn();
    m = Math.min(m, performance.now() - s);
  }
  return m;
};
const openSave = () => Presentation(data).save(new BytesIO());
const build = () => {
  const prs = Presentation();
  for (let i = 0; i < 100; i++) prs.slides.addSlide(prs.slideLayouts.at(i % 2));
  prs.save(new BytesIO());
};
const shapes = () => {
  const prs = Presentation();
  const slide = prs.slides.addSlide(prs.slideLayouts.at(6));
  slide.shapes.turboAddEnabled = true;
  for (let i = 0; i < 1000; i++) {
    const tb = slide.shapes.addTextbox(i * 1000, i * 1000, 914400, 457200);
    tb.textFrame.text = `shape ${i}`;
  }
  prs.save(new BytesIO());
};
const charts = () => {
  const prs = Presentation();
  for (let i = 0; i < 50; i++) {
    const slide = prs.slides.addSlide(prs.slideLayouts.at(6));
    const cd = new CategoryChartData();
    cd.categories = Array.from({ length: 12 }, (_, m) => `Month ${m + 1}`);
    for (let s = 0; s < 3; s++) cd.addSeries(`Series ${s}`, Array.from({ length: 12 }, (_, m) => ((m * 7 + s * 3 + i) % 50) + 0.5));
    slide.shapes.addChart(XL_CHART_TYPE.COLUMN_CLUSTERED, 914400, 914400, 6 * 914400, 4 * 914400, cd);
  }
  prs.save(new BytesIO());
};
const openSaveAll = () => {
  for (const d of decks) Presentation(d).save(new BytesIO());
};
const build200 = () => {
  const prs = Presentation();
  for (let i = 0; i < 200; i++) {
    const slide = prs.slides.addSlide(prs.slideLayouts.at(1));
    slide.shapes.title.textFrame.text = `Slide ${i + 1}`;
    const body = slide.placeholders.get(1).textFrame;
    body.text = "First point";
    for (let j = 0; j < 3; j++) {
      const p = body.addParagraph();
      p.text = `Sub-point ${j + 1}`;
      p.level = 1;
    }
    const shape = slide.shapes.addShape(MSO_SHAPE.ROUNDED_RECTANGLE, 457200, 5486400, 1828800, 457200);
    shape.fill.solid();
    shape.fill.foreColor.rgb = new RGBColor(0x1f, 0x77, 0xb4);
    shape.textFrame.text = `Box ${i + 1}`;
    slide.shapes.addPicture(new BytesIO(image), 6400800, 5029200, 1828800);
  }
  prs.save(new BytesIO());
};
const shapes10k = () => {
  const prs = Presentation();
  const slide = prs.slides.addSlide(prs.slideLayouts.at(6));
  slide.shapes.turboAddEnabled = true;
  for (let i = 0; i < 10000; i++) slide.shapes.addShape(MSO_SHAPE.RECTANGLE, (i % 100) * 91440, Math.floor(i / 100) * 68580, 91440, 68580);
  prs.save(new BytesIO());
};
console.log(`open + save presentation  ${best(openSave).toFixed(1).padStart(8)} ms`);
console.log(`build 100-slide deck      ${best(build).toFixed(1).padStart(8)} ms`);
console.log(`1000 text boxes on a slide ${best(shapes, 3).toFixed(1).padStart(8)} ms`);
console.log(`50 charts with workbooks  ${best(charts, 3).toFixed(1).padStart(8)} ms`);
console.log(`open + save ${decks.length} test decks     ${best(openSaveAll, 3).toFixed(1).padStart(8)} ms`);
console.log(`200 slides, text+shape+picture ${best(build200, 3).toFixed(1).padStart(8)} ms`);
console.log(`10000 shapes on a slide   ${best(shapes10k, 3).toFixed(1).padStart(8)} ms`);
