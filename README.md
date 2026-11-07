# pptx-ts

A TypeScript port of [python-pptx](https://github.com/scanny/python-pptx) for creating,
reading and updating PowerPoint (.pptx) files in browsers and Node, with no server and no
Python.

**Status: stage 1 of 6.** The foundation is done: a fast XML engine that matches lxml's
output byte for byte, python-pptx's declarative element layer (xmlchemy), every XML simple
type, and the complete OPC package layer, so packages open, save and round-trip losslessly.
The presentation object model (slides, shapes, text, tables, charts) arrives in the next
stages; see [PARITY.md](PARITY.md) and the stage issues on GitHub.

```ts
import { BytesIO, OpcPackage } from "pptx-ts";

const pkg = OpcPackage.open(bytes);                  // Uint8Array, ArrayBuffer, path (Node) or file-like
for (const part of pkg.iterParts()) console.log(String(part.partname), part.contentType);
const out = new BytesIO();
pkg.save(out);                                       // or pkg.save("out.pptx") in Node
```

## Design

- **XML engine** (`src/oxml/xml.ts`, `src/oxml/xpath.ts`): parser, lxml-style element tree
  (text/tail model, namespace declarations, custom element classes per tag) and serializer.
  `test/xml_parity.test.ts` checks it against lxml on 307 real XML parts: serialization,
  pretty printing, subtree output and XPath results are identical.
- **xmlchemy** (`src/oxml/xmlchemy.ts`): `defineChildren(this, { solidFill: new ZeroOrOne(...) })`
  generates the same members python-pptx's metaclass does, in camelCase
  (`get_or_add_spPr` → `getOrAddSpPr`, `p_lst` → `pLst`).
- **Zip** via [fflate](https://github.com/101arrowz/fflate) (pinned), synchronous in both
  browsers and Node.
- **Python → TypeScript naming**: methods and properties are camelCase; class names and
  constants are unchanged. Lengths are plain numbers in EMU; `Inches(1)`, `Pt(12)` etc.
  return numbers, and `new Length(emu).inches` converts.

## Tests

Upstream's unit tests are machine-translated from python-pptx by `scripts/py2ts_tests.py`
(pytest classes, fixtures with params, `unittest.mock` helpers) and run under vitest with
Python semantics from `test/pyshim.ts` and `test/mock.ts`. Post-translation edits and skips
are recorded in `test/fixups.json`, so the files can be regenerated:

```bash
scripts/regen_ported_tests.sh ../python-pptx python   # files listed in test/PORTED
```

## Performance

Stage-1 workloads on `cht-charts.pptx` (best of 5, Node 22 vs CPython 3.13 + lxml 5.3,
`bench/bench_opc.mjs` / `bench/bench_opc.py`):

| Workload | python-pptx | pptx-ts |
| --- | ---: | ---: |
| open + save package | 10.8 ms | 10.6 ms |
| parse + serialize 12 MB of part XML | 24.4 ms | 50.8 ms |

python-pptx also parses the XML parts it has classes for on open; pptx-ts will too from
stage 2. The pure-JS XML engine runs at about half of lxml's C speed.

## Development

```bash
npm install
npm test           # vitest
npm run typecheck
npm run build      # tsc -> dist/
node scripts/gen_templates.mjs                                  # after changing templates/
py -I scripts/gen_constants.py ../python-pptx src/opc/constants.ts
```
