# pptx-ts

A TypeScript port of [python-pptx](https://github.com/scanny/python-pptx) for creating,
reading and updating PowerPoint (.pptx) files in browsers and Node, with no server and no
Python.

**Status: stages 1–5 of 6: all of python-pptx's API.** Open any deck or start from the
built-in template; slides, layouts, masters, notes and core properties; shapes, pictures,
movies, connectors, freeforms, groups, placeholders, fills, lines, colors, hyperlinks and click
actions; text frames with fonts and paragraphs, `fitText()`; tables; charts of every type
python-pptx creates, with their embedded Excel workbooks. It runs in browsers as well as
Node ([demo](demo/index.html)). See [PARITY.md](PARITY.md) for cross-check results and known
output differences. Stage 6 ports upstream's acceptance tests; see the stage issues on GitHub.

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

Charts take the same chart data objects as python-pptx:

```ts
import { CategoryChartData, XL_CHART_TYPE, XL_LEGEND_POSITION, Inches } from "pptx-ts";

const data = new CategoryChartData();
data.categories = ["East", "West", "Midwest"];
data.addSeries("Q1 Sales", [19.2, 21.4, 16.7]);
data.addSeries("Q2 Sales", [22.3, 28.6, 15.2]);

const slide2 = prs.slides.addSlide(prs.slideLayouts.at(6));
const chart = slide2.shapes.addChart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1), Inches(8), Inches(5), data).chart;
chart.hasLegend = true;
chart.legend.position = XL_LEGEND_POSITION.BOTTOM;
chart.plots.at(0).hasDataLabels = true;
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
| `CategoryChartData()`, `chart.plots[0]` | `new CategoryChartData()`, `chart.plots.at(0)` |
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
- **Chart workbooks**: `src/xlsx/workbook.ts` replaces XlsxWriter; see
  [PARITY.md, item 14](PARITY.md#behavioral-differences) for its compatibility scope.
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
python scripts/crosscheck/charts_py.py ../python-pptx charts.json
npx tsx scripts/crosscheck/charts_ts.ts charts.json ../python-pptx
python scripts/crosscheck/chart_build_py.py charts.pptx
npx tsx scripts/crosscheck/chart_build_ts.ts charts.pptx
```

The one deck difference is a deck without core properties, where both libraries stamp the
current time. Chart cross-check results and the remaining difference are recorded in
[PARITY.md](PARITY.md#ported-upstream-tests).

Source modules start from drafts written by `scripts/py2ts_src.py` (a porting aid that keeps
upstream's structure and docstrings) and are then reviewed and corrected by hand.

## Performance

Best of 5 (best of 3 for the 1000-shape case), Node 22 vs CPython 3.13 + lxml 5.3, on
`cht-charts.pptx` from python-pptx's test files (`bench/bench_prs.*`, `bench/bench_opc.*`):

| Workload | python-pptx | pptx-ts |
| --- | ---: | ---: |
| build a 100-slide deck from the default template | 297.5 ms | 120.9 ms |
| add 1000 text boxes with text to a slide, save | 309.9 ms | 150.5 ms |
| add 50 column charts (12 categories, 3 series) with workbooks, save | 364.0 ms | 320.0 ms |
| open + save a presentation (all XML parts parsed) | 10.9 ms | 15.2 ms |
| parse + serialize the deck's XML parts 20 times | 24.3 ms | 46.2 ms |

Everything above the XML layer is about twice as fast as python-pptx. Charts gain less: about
40% of that workload is deflating the 50 embedded workbooks and then the deck, which python
does in C zlib. The XML parser itself is
pure JS and runs at about half the speed of lxml's C parser, so opening a deck is a few
milliseconds slower.

### Why not WebAssembly?

`bench/bench_scan.mjs` splits XML parsing into the part a WebAssembly parser could take over
(decoding and scanning the text: 22% of parse time) and building the element tree (78%). The
tree is JS objects with per-tag classes that the rest of the library works on, so it has to be
built in JS either way, and a WASM parser would add the cost of passing every tag and
attribute string across the boundary. At best it would save about 1 ms per large deck, while
adding an asynchronous load step and a binary to ship. The JS engine is the better trade-off.

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
python scripts/text_golden.py test/test_files/calibriz.ttf test/fixtures/text_golden.json
python scripts/gen_unitdata.py ../python-pptx test/oxml/unitdata
npm run demo:build                                              # demo/pptx.js for demo/index.html
```

CI (`.github/workflows/ci.yml`) runs typecheck, tests, build and a browser bundle on Node 20
and 22, Linux and Windows.
