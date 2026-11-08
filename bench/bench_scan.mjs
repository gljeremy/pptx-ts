// How much of XML parsing is scanning (what a WASM parser could replace) versus building the
// JS element tree (which stays in JS either way)?
import fs from "node:fs";
import "../dist/index.js";
import { OpcPackage } from "../dist/opc/package.js";
import { parseXml } from "../dist/oxml/parser.js";
const data = new Uint8Array(fs.readFileSync(process.argv[2]));
const parts = [...OpcPackage.open(data).iterParts()].filter((p) => String(p.partname).endsWith(".xml")).map((p) => p.blob);
const strs = parts.map((b) => new TextDecoder().decode(b));
const best = (fn, n = 7) => { let m = Infinity; for (let i = 0; i < n; i++) { const s = performance.now(); fn(); m = Math.min(m, performance.now() - s); } return m; };
// scan only: find every tag and attribute boundary, create no objects
const scan = (s) => {
  let i = 0, n = 0;
  for (;;) {
    const lt = s.indexOf("<", i);
    if (lt < 0) break;
    const gt = s.indexOf(">", lt);
    for (let q = s.indexOf('"', lt); q >= 0 && q < gt; q = s.indexOf('"', s.indexOf('"', q + 1) + 1)) n++;
    i = gt + 1;
    n++;
  }
  return n;
};
const full = best(() => { for (let i = 0; i < 20; i++) for (const b of parts) parseXml(b); });
const decode = best(() => { const d = new TextDecoder(); for (let i = 0; i < 20; i++) for (const b of parts) d.decode(b); });
const scanOnly = best(() => { for (let i = 0; i < 20; i++) for (const s of strs) scan(s); });
console.log(`full parse x20        ${full.toFixed(1)} ms`);
console.log(`utf-8 decode x20      ${decode.toFixed(1)} ms`);
console.log(`scan only x20         ${scanOnly.toFixed(1)} ms  (${((scanOnly + decode) / full * 100).toFixed(0)}% of parse is decode + scan)`);
console.log(`tree building         ${(full - scanOnly - decode).toFixed(1)} ms  (stays in JS with a WASM parser)`);
