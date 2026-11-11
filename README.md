# pptx-ts

A TypeScript port of [python-pptx](https://github.com/scanny/python-pptx) for creating,
reading and updating PowerPoint (.pptx) files in browsers and Node, with no server and no
Python.

**Status: complete.** All of python-pptx's API: open any deck or start from the built-in
template; slides, layouts, masters, notes and core properties; shapes, pictures, movies,
connectors, freeforms, groups, placeholders, fills, lines, colors, hyperlinks and click actions;
text frames with fonts and paragraphs, `fitText()`; tables; charts of every type python-pptx
creates, with their embedded Excel workbooks. Upstream's unit tests and all 973 of its
acceptance scenarios pass. It runs in browsers as well as Node ([demo](demo/index.html)). See
[PARITY.md](PARITY.md) for cross-check results and known output differences.

```sh
npm install @gljeremy/pptx-ts
```

```ts
import { Presentation, BytesIO } from "@gljeremy/pptx-ts";

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
import { CategoryChartData, XL_CHART_TYPE, XL_LEGEND_POSITION, Inches } from "@gljeremy/pptx-ts";

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

## Getting started

These are python-pptx's [Getting Started](https://python-pptx.readthedocs.io/en/latest/user/quickstart.html)
examples in TypeScript. Each deck is identical, part by part, to python-pptx's output for the
same steps (`test/quickstart.test.ts`).

### Hello World!

```ts
import { Presentation } from "@gljeremy/pptx-ts";

const prs = Presentation();
const titleSlideLayout = prs.slideLayouts.at(0);
const slide = prs.slides.addSlide(titleSlideLayout);
const title = slide.shapes.title!;
const subtitle = slide.placeholders.get(1);

title.text = "Hello, World!";
subtitle.text = "python-pptx was here!";

prs.save("test.pptx");
```

### Bullet slide

```ts
const prs = Presentation();
const slide = prs.slides.addSlide(prs.slideLayouts.at(1));
const shapes = slide.shapes;
shapes.title!.text = "Adding a Bullet Slide";

const tf = shapes.placeholders.get(1).textFrame;
tf.text = "Find the bullet slide layout";

let p = tf.addParagraph();
p.text = "Use _TextFrame.text for first bullet";
p.level = 1;

p = tf.addParagraph();
p.text = "Use _TextFrame.add_paragraph() for subsequent bullets";
p.level = 2;

prs.save("test.pptx");
```

Not every shape can hold text; `shape.hasTextFrame` says whether it can. A text frame always
has at least one paragraph. Assigning `shape.text` or `textFrame.text` replaces all existing
contents, creating a paragraph for each newline (`\n`) and a soft line break for each
vertical tab (`\v`) in the assigned string.

### `addTextbox()`

```ts
import { Presentation, Inches, Pt } from "@gljeremy/pptx-ts";

const prs = Presentation();
const slide = prs.slides.addSlide(prs.slideLayouts.at(6)); // blank layout
const txBox = slide.shapes.addTextbox(Inches(1), Inches(1), Inches(1), Inches(1));
const tf = txBox.textFrame;
tf.text = "This is text inside a textbox";

let p = tf.addParagraph();
p.text = "This is a second paragraph that's bold";
p.font.bold = true;

p = tf.addParagraph();
p.text = "This is a third paragraph that's big";
p.font.size = Pt(40);

prs.save("test.pptx");
```

### `addPicture()`

```ts
const imgPath = "monty-truth.png";                  // or a Uint8Array, or a BytesIO

const prs = Presentation();
const slide = prs.slides.addSlide(prs.slideLayouts.at(6));
slide.shapes.addPicture(imgPath, Inches(1), Inches(1));                       // native size
slide.shapes.addPicture(imgPath, Inches(5), Inches(1), null, Inches(5.5));    // height=, aspect kept

prs.save("test.pptx");
```

### `addShape()`

```ts
import { Presentation, Inches, MSO_SHAPE } from "@gljeremy/pptx-ts";

const prs = Presentation();
const slide = prs.slides.addSlide(prs.slideLayouts.at(5)); // title only
const shapes = slide.shapes;
shapes.title!.text = "Adding an AutoShape";

let left = Inches(0.93); // 0.93" centers this overall set of shapes
const top = Inches(3.0);
let width = Inches(1.75);
const height = Inches(1.0);

let shape = shapes.addShape(MSO_SHAPE.PENTAGON, left, top, width, height);
shape.text = "Step 1";

left = left + width - Inches(0.4); // lengths are numbers, so arithmetic just works
width = Inches(2.0); // chevrons need more width for visual balance
for (let n = 2; n < 6; n++) {
  shape = shapes.addShape(MSO_SHAPE.CHEVRON, left, top, width, height);
  shape.text = `Step ${n}`;
  left = left + width - Inches(0.4);
}

prs.save("test.pptx");
```

`MSO_SHAPE` has every auto shape type python-pptx has (`MSO_SHAPE.ROUNDED_RECTANGLE`,
`MSO_SHAPE.CHEVRON`, ...).

### `addTable()`

```ts
const prs = Presentation();
const slide = prs.slides.addSlide(prs.slideLayouts.at(5));
slide.shapes.title!.text = "Adding a Table";

const table = slide.shapes.addTable(2, 2, Inches(2.0), Inches(2.0), Inches(6.0), Inches(0.8)).table;

// set column widths
table.columns.at(0).width = Inches(2.0);
table.columns.at(1).width = Inches(4.0);

// write column headings
table.cell(0, 0).text = "Foo";
table.cell(0, 1).text = "Bar";

// write body cells
table.cell(1, 0).text = "Baz";
table.cell(1, 1).text = "Qux";

prs.save("test.pptx");
```

### Extract all text from the slides in a presentation

```ts
const prs = Presentation(pathToPresentation);

// textRuns will be populated with a list of strings, one for each text run in the presentation
const textRuns: string[] = [];
for (const slide of prs.slides) {
  for (const shape of slide.shapes) {
    if (!shape.hasTextFrame) continue;
    for (const paragraph of shape.textFrame.paragraphs) {
      for (const run of paragraph.runs) textRuns.push(run.text);
    }
  }
}
```

### In the browser

`Presentation()` takes a `Uint8Array` (for example from a file input) and `save()` writes to a
`BytesIO`, so no file system is needed:

```ts
const prs = Presentation(new Uint8Array(await file.arrayBuffer()));
// ... edit ...
const out = new BytesIO();
prs.save(out);
const blob = new Blob([out.getvalue()], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
```

The package is ES modules with types, and its only dependency is
[fflate](https://github.com/101arrowz/fflate). Bundled for the browser, it is about 400 KB
minified, 130 KB gzipped. `addOleObject()` without an icon file uses python-pptx's built-in
icons, which are kept out of that bundle (about 260 KB). In the browser, load them first:

```ts
import "@gljeremy/pptx-ts/ole-icons";                 // or: await import("@gljeremy/pptx-ts/ole-icons") when needed
slide.shapes.addOleObject(xlsxBytes, PROG_ID.XLSX, Inches(1), Inches(1));
```

Without this import or a custom icon file, browser `addOleObject()` raises `KeyError`
before embedding the attachment. Node reads the icons from the package's files, so it needs
no import.

### From python-pptx

| python-pptx | pptx-ts |
| --- | --- |
| `from pptx import Presentation` | `import { Presentation } from "@gljeremy/pptx-ts"` |
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

Upstream's behave acceptance tests run too: `test/features/` holds its 54 `.feature` files and
test decks, `scripts/py2ts_steps.py` translates its step definitions, and
`test/features/runner.ts` is a small Gherkin runner for vitest. All 973 scenarios pass, the
same number behave runs for python-pptx:

```bash
scripts/regen_features.sh ../python-pptx python       # test/features/steps/*.ts
npx vitest run test/features
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

Node 22 vs CPython 3.13 + lxml 5.3 (`bench/bench_prs.*`, `bench/bench_opc.*`).
The single deck is `cht-charts.pptx` from python-pptx's test files; the 60 decks are its
acceptance test decks. Trial counts for each workload are defined in the benchmark scripts:

| Workload | python-pptx | pptx-ts |
| --- | ---: | ---: |
| build a 100-slide deck from the default template | 297.9 ms | 121.4 ms |
| build a 200-slide deck, each slide with title, 4 bullets, a filled shape and a picture | 1216.9 ms | 491.1 ms |
| add 1000 text boxes with text to a slide, save | 311.3 ms | 146.2 ms |
| add 10,000 shapes to a slide, save | 4094.3 ms | 2233.0 ms |
| add 50 column charts (12 categories, 3 series) with workbooks, save | 365.2 ms | 329.7 ms |
| open + save a presentation (all XML parts parsed) | 10.8 ms | 15.1 ms |
| open + save 60 different decks | 942.2 ms | 902.4 ms |
| parse + serialize the deck's XML parts 20 times | 24.3 ms | 46.2 ms |

Building decks is two to two and a half times as fast as python-pptx. Charts gain less: about
40% of that workload is deflating the 50 embedded workbooks and then the deck, which python
does in C zlib. Opening and saving decks is about even across the 60 test decks, and a few
milliseconds slower on the chart deck: the XML parser is pure JS and runs at about half the
speed of lxml's C parser.

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
npm run check:bundle                                            # browser bundle: size, no OLE icons, runs
npm run check:pack                                              # npm pack -> fresh Vite project
python scripts/quickstart_golden.py ../python-pptx test/fixtures/quickstart
```

CI runs typecheck, tests, build, browser bundle checks (including the gzip size budget and
OLE icon registration), and the packed-tarball check in a fresh Vite project. See
[the workflow](.github/workflows/ci.yml) for the Node and OS matrix.
