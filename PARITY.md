# pptx-ts ↔ python-pptx parity

Port of [python-pptx](https://github.com/scanny/python-pptx) at commit `278b47b`
(v1.0.2, 2024-08-06).

## Stages

| Stage | Scope | Status |
| --- | --- | --- |
| 1 | XML engine, `oxml/xmlchemy`, `oxml/simpletypes`, `oxml/ns`, `opc/*`, `util`, `shared`, `exc` | **done** |
| 2 | `enum/*`, `package.py`, `presentation.py`, `parts/*` (presentation, slide, coreprops, image, media, embeddedpackage), `slide.py`, `oxml/presentation`, `oxml/slide`, `oxml/theme`, `oxml/coreprops`, image header parsing (replaces Pillow) — `Presentation()` works | open |
| 3 | `shapes/*`, `oxml/shapes/*`, `dml/*`, `oxml/dml/*`, `action.py`, `media.py` | open |
| 4 | `text/*`, `oxml/text`, `table.py`, `oxml/table`, font metrics (`text/fonts.py`, `text/layout.py`) | open |
| 5 | `chart/*`, `oxml/chart/*`, `parts/chart`, embedded workbook writer (replaces XlsxWriter) | open |
| 6 | the 54 behave feature files (371 scenarios) as vitest acceptance tests, docs, benchmarks | open |

## Ported upstream tests (stage 1)

| Upstream | pptx-ts |
| --- | --- |
| `tests/opc/test_packuri.py`, `test_oxml.py`, `test_serialized.py`, `test_package.py` | `test/opc/*.test.ts` |
| `tests/oxml/test_xmlchemy.py`, `test_simpletypes.py`, `test_ns.py` | `test/oxml/*.test.ts` |
| `tests/test_util.py`, `tests/test_shared.py` | `test/util.test.ts`, `test/shared.test.ts` |
| `tests/unitdata.py`, `tests/unitutil/*` | `test/unitdata.ts`, `test/unitutil/*`, `test/mock.ts` (hand-ported) |

Additional pptx-ts tests: `test/xml_parity.test.ts` (byte-for-byte lxml comparison on 307
XML parts), `test/opc/roundtrip.test.ts` (open → save → reopen of every test deck),
`test/opc/zip_writer.test.ts`.

### Skipped or adjusted (see `test/fixups.json`)

| Test | Why |
| --- | --- |
| `DescribeCustomElementClass.it_has_the_MetaOxmlElement_metaclass` | python metaclass; classes call `defineChildren()` |
| `DescribeOptionalAttribute.it_adds_a_docstring_for_the_property` and `__doc__` asserts in xmlchemy tests | no docstrings in JS |
| `Describe_ZipPkgWriter` `__enter__` / `__exit__` / `it_can_write_a_blob` / `_zipf` | python `zipfile` and context managers; replaced by `test/opc/zip_writer.test.ts` |
| `DescribeLength.it_can_construct_from_convenient_units` (`isinstance(..., Length)` dropped) | lengths are numbers |
| `Describe_Relationships.it_collects_relationships_by_reltype_to_help` (CHART lookup) | python defaultdict; callers default to `[]` |

## Behavioral differences

1. **Lengths are numbers** (EMU). `Inches()`, `Pt()`, `Cm()`, `Mm()`, `Emu()`,
   `Centipoints()` return numbers; use `new Length(x).inches` etc. to convert.
2. **One number type**: integer validation is `Number.isInteger`, so `1.0` counts as an int,
   and booleans are not numbers.
3. **Files**: anything that takes a file accepts bytes (`Uint8Array`/`ArrayBuffer`), a
   file-like object with `read()`/`write()` (e.g. the exported `BytesIO`), or — in Node — a
   path. Browsers have no paths.
4. **Mappings**: python-style collections (`_Relationships`, readers) expose `get(k)`
   (raising `KeyError` like `obj[k]`), `has(k)` for `in`, `length` for `len()`, and iterate
   their keys.
5. **Exceptions**: `ValueError`, `KeyError`, `IndexError`, `AttributeError`,
   `NotImplementedError` are exported classes; JS `TypeError` stands in for python's.
6. **Zip output** is written by fflate (deflate level 6, like python's zipfile); archives
   differ byte-wise from python's but contain identical members.
