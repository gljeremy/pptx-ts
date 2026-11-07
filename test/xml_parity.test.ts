// Checks the XML engine against lxml on real presentations: for every XML part of the decks
// in test/fixtures/lxml_golden.json.gz (made by scripts/xml_golden.py) parsing + serializing
// must give lxml's exact bytes, as must pretty printing, subtree serialization, and XPath.
import { expect, it } from "vitest";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { gunzipSync } from "fflate";
import { parseXml } from "../src/oxml/parser.js";
import { tostring, XmlElement } from "../src/oxml/xml.js";

const golden = JSON.parse(
  new TextDecoder().decode(gunzipSync(fs.readFileSync(path.join(path.dirname(fileURLToPath(import.meta.url)), "fixtures", "lxml_golden.json.gz")))),
);

it(`matches lxml on ${golden.length} XML parts`, () => {
  const bad: string[] = [];
  for (const m of golden) {
    const root = parseXml(m.raw);
    const first = root._children.length && root._children[0] instanceof XmlElement ? root._children[0] : root;
    const got = {
      part: tostring(root, { xmlDeclaration: true, encoding: "UTF-8", standalone: true }),
      pretty: tostring(root, { prettyPrint: true }),
      sub: tostring(first, { prettyPrint: true }),
      ids: JSON.stringify(root.xpath("//@r:id")),
    };
    const exp = { part: m.part, pretty: m.pretty, sub: m.sub, ids: JSON.stringify(m.ids) };
    for (const k of Object.keys(got) as (keyof typeof got)[]) if (got[k] !== exp[k]) bad.push(`${m.file}:${m.name} ${k}`);
  }
  expect(bad).toEqual([]);
});
