// XML engine split: parse vs serialize vs UTF-8 decode (pptx-ts only).
// Usage: npm run build && node bench/bench_xml.mjs <deck.pptx>
import fs from "node:fs";
import "../dist/index.js";
import { OpcPackage } from "../dist/opc/package.js";
import { parseXml } from "../dist/oxml/parser.js";
import { tostring } from "../dist/oxml/xml.js";
const data = new Uint8Array(fs.readFileSync(process.argv[2]));
const parts = [...OpcPackage.open(data).iterParts()].filter((p) => String(p.partname).endsWith(".xml")).map((p) => p.blob);
const best = (fn, n = 7) => { let m = Infinity; for (let i = 0; i < n; i++) { const s = performance.now(); fn(); m = Math.min(m, performance.now() - s); } return m; };
const trees = parts.map((b) => parseXml(b));
console.log("parse x20", best(() => { for (let i = 0; i < 20; i++) for (const b of parts) parseXml(b); }).toFixed(1));
console.log("serialize x20", best(() => { for (let i = 0; i < 20; i++) for (const t of trees) tostring(t, { xmlDeclaration: true, standalone: true }); }).toFixed(1));
console.log("decode only x20", best(() => { const d = new TextDecoder(); for (let i = 0; i < 20; i++) for (const b of parts) d.decode(b); }).toFixed(1));
