// pptx-ts stage-2 workloads (Presentation API), for comparison with bench_prs.py.
// Usage: npm run build && node bench/bench_prs.mjs <deck.pptx>
import fs from "node:fs";
import { Presentation, BytesIO } from "../dist/index.js";

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
console.log(`open + save presentation  ${best(openSave).toFixed(1).padStart(8)} ms`);
console.log(`build 100-slide deck      ${best(build).toFixed(1).padStart(8)} ms`);
