# pptx-ts

A TypeScript port of [python-pptx](https://github.com/scanny/python-pptx) for creating,
reading and updating PowerPoint (.pptx) files in browsers and Node, with no server and no
Python.

**Status: stage 2 of 6.** You can open any deck, or start from the built-in template, add
slides from its layouts, read slides, layouts, masters, notes and core properties, and save.
Decks built with pptx-ts are byte-identical to the same steps in python-pptx. The shape, text,
table and chart APIs are ported with this stage but are covered by upstream's tests from
stages 3–5; see [PARITY.md](PARITY.md) and the stage issues on GitHub.

```ts
import { Presentation, BytesIO } from "pptx-ts";

const prs = Presentation();                        // or Presentation(bytes | path | file-like)
const slide = prs.slides.addSlide(prs.slideLayouts.at(0));
slide.shapes.title!.textFrame.text = "Hello, pptx-ts";
slide.placeholders.get(1).textFrame.text = "built in TypeScript";
prs.coreProperties.title = "Hello";

for (const s of prs.slides) console.log(s.slideId, s.slideLayout.name);

const out = new BytesIO();
prs.save(out);                                     // or prs.save("hello.pptx") in Node
const bytes = out.getvalue();                      // Uint8Array, e.g. for a download Blob
```

### From python-pptx

| python-pptx | pptx-ts |
| --- | --- |
| `from pptx import Presentation` | `import { Presentation } from "pptx-ts"` |
| `prs.slide_layouts[6]`, `prs.slides[-1]` | `prs.slideLayouts.at(6)`, `prs.slides.at(-1)` |
| `len(prs.slides)`, `for s in prs.slides` | `prs.slides.length`, `for (const s of prs.slides)` |
| `slide.placeholders[1]` (looked up by idx) | `slide.placeholders.get(1)` |
| `shape.text_frame`, `get_or_add_xfrm()` | `shape.textFrame`, `getOrAddXfrm()` (camelCase) |
| `Inches(1)`, `Pt(12)` (`Length` ints) | `Inches(1)`, `Pt(12)` (plain numbers in EMU) |
| `datetime.datetime(2024, 1, 2)` | `new Date(Date.UTC(2024, 0, 2))` (naive times held as UTC) |
| `MSO_SHAPE.RECTANGLE == 1` | `MSO_SHAPE.RECTANGLE` is a `Number`; `+MSO_SHAPE.RECTANGLE === 1` |

## Design

- **XML engine** (`src/oxml/xml.ts`, `src/oxml/xpath.ts`): parser, lxml-style element tree
  (text/tail model, namespace declarations, custom element classes per tag) and serializer.
  `test/xml_parity.test.ts` checks it against lxml on 307 real XML parts: serialization,
  pretty printing, subtree output and XPath results are identical.
- **xmlchemy** (`src/oxml/xmlchemy.ts`): `defineChildren(this, { solidFill: new ZeroOrOne(...) })`
  generates the same members python-pptx's metaclass does, in camelCase
  (`get_or_add_spPr` → `getOrAddSpPr`, `p_lst` → `pLst`).
- **Enums** (`src/enum/*`) are generated from python-pptx by `scripts/gen_enums.py`: members
  are `Number` objects with `.name`, `.value` and (for XML enums) `.xmlValue`; enums are
  callable (`PP_ACTION_TYPE(10)`) and have `fromXml()`/`toXml()`.
- **Images**: python-pptx reads image size and DPI with Pillow; `src/image/headers.ts` reads
  them from PNG, JPEG (JFIF and EXIF), GIF, BMP, TIFF and WMF/EMF headers and is checked
  against Pillow's values (`test/image/headers.test.ts`).
- **Zip** via [fflate](https://github.com/101arrowz/fflate) (pinned), synchronous in both
  browsers and Node. SHA-1 for image/media de-duplication is a small synchronous
  implementation (`src/sha1.ts`).

## Tests

Upstream's unit tests are machine-translated from python-pptx by `scripts/py2ts_tests.py`
(pytest classes, fixtures with params, `unittest.mock` helpers) and run under vitest with
Python semantics from `test/pyshim.ts` and `test/mock.ts`. Post-translation edits and skips
are recorded in `test/fixups.json`, so the files can be regenerated:

```bash
scripts/regen_ported_tests.sh ../python-pptx python   # files listed in test/PORTED
```

`scripts/crosscheck/` compares pptx-ts with python-pptx on every upstream test deck (what the
API reports for slides, layouts, masters, shapes, placeholders, notes and core properties,
and the bytes of every part after re-saving):

```bash
python scripts/crosscheck/py_side.py ../python-pptx py.json
npx tsx scripts/crosscheck/ts_side.ts py.json ../python-pptx    # 65/66 decks identical
```

The one difference is a deck without core properties, where both libraries stamp the
current time.

Source modules start from drafts written by `scripts/py2ts_src.py` (a porting aid that keeps
upstream's structure and docstrings) and are then reviewed and corrected by hand.

## Performance

Best of 5 on `cht-charts.pptx`, Node 22 vs CPython 3.13 + lxml 5.3
(`bench/bench_prs.*`, `bench/bench_opc.*`):

| Workload | python-pptx | pptx-ts |
| --- | ---: | ---: |
| open + save presentation (XML parts parsed) | 10.8 ms | 15.4 ms |
| build a 100-slide deck from the default template | 297.5 ms | 120.4 ms |
| open + save package (OPC only) | 10.8 ms | 10.6 ms |
| parse + serialize 12 MB of part XML | 24.4 ms | 50.8 ms |

The pure-JS XML parser runs at about half of lxml's C speed (tracked in the performance
issue); everything above the XML layer is faster than python-pptx.

## Development

```bash
npm install
npm test           # vitest
npm run typecheck
npm run build      # tsc -> dist/
node scripts/gen_templates.mjs                                  # after changing templates/
py -I scripts/gen_constants.py ../python-pptx src/opc/constants.ts
python scripts/gen_enums.py src/enum                            # python-pptx importable
python scripts/gen_spec.py src/spec.ts
python scripts/image_golden.py test/fixtures/images_pillow.json test/fixtures/images test/test_files
```
