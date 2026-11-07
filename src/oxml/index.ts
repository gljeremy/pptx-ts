// Registers every custom element class with the parser, mirroring python-pptx's
// `pptx/oxml/__init__.py`. Importing the package (src/index.ts) loads this.
export { parseFromTemplate, parseXml, registerElementCls } from "./parser.js";
import "./xmlchemy.js";
import "../opc/oxml.js";
