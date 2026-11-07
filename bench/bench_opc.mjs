// pptx-ts stage-1 workloads (OPC package + XML), for comparison with bench_opc.py.
// Usage: npm run build && node bench/bench_opc.mjs <deck.pptx>
import fs from "node:fs";
import "../dist/index.js";
import { OpcPackage } from "../dist/opc/package.js";
import { BytesIO } from "../dist/fileio.js";
import { parseXml } from "../dist/oxml/parser.js";
import { tostring } from "../dist/oxml/xml.js";

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
const openSave = () => {
  const pkg = OpcPackage.open(data);
  pkg.save(new BytesIO());
};
const xmlParts = [...OpcPackage.open(data).iterParts()].filter((p) => String(p.partname).endsWith(".xml")).map((p) => p.blob);
const parseSerialize = () => {
  for (let i = 0; i < 20; i++) for (const b of xmlParts) tostring(parseXml(b), { xmlDeclaration: true, standalone: true });
};
console.log(`open + save package     ${best(openSave).toFixed(1).padStart(8)} ms`);
console.log(`parse+serialize x20     ${best(parseSerialize).toFixed(1).padStart(8)} ms  (${((xmlParts.reduce((a, b) => a + b.length, 0) * 20) / 1e6).toFixed(1)} MB of XML)`);
