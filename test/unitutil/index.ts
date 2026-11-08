// Port of python-pptx tests/unitutil/__init__.py: helper objects for unit testing.

/** Local implementation of `itertools.count()` (upstream keeps it for py2.6 compatibility). */
export function* count(start = 0, step = 1): Generator<number> {
  let n = start;
  for (;;) {
    yield n;
    n += step;
  }
}
