// pptx-ts stage-2 workloads (Presentation API), for comparison with bench_prs.py.
// Usage: npm run build && node bench/bench_prs.mjs <deck.pptx>
import fs from "node:fs";
import { Presentation, BytesIO, CategoryChartData, XL_CHART_TYPE } from "../dist/index.js";

const data = new Uint8Array(fs.readFileSync(process.argv[2]));
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
console.log(`open + save presentation  ${best(openSave).toFixed(1).padStart(8)} ms`);
console.log(`build 100-slide deck      ${best(build).toFixed(1).padStart(8)} ms`);
console.log(`1000 text boxes on a slide ${best(shapes, 3).toFixed(1).padStart(8)} ms`);
console.log(`50 charts with workbooks  ${best(charts, 3).toFixed(1).padStart(8)} ms`);
