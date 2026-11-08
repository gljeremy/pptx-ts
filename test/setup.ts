// Loads the element-class registry, as importing any pptx module does in python.
import "../src/index.js";

// JS's built-in `escape()`/`unescape()` (URL-style %XX encoding) share their names with
// python's `xml.sax.saxutils.escape`, so a port that forgets to switch to `xmlEscape` still
// compiles. Make any call from code under test fail loudly instead of writing %XX into XML.
for (const name of ["escape", "unescape"] as const) {
  (globalThis as any)[name] = () => {
    throw new Error(`global ${name}() called: use xmlEscape() from src/util.ts`);
  };
}
