// Hand-ported from python-pptx tests/oxml/test___init__.py (listed in test/HAND_EDITED).
// Upstream tests lxml's parser objects (`etree.fromstring(xml, oxml_parser)`); these check the
// same behaviors through pptx-ts's parser. `it_uses_oxml_configured_parser_to_parse_xml` only
// checks that parse_xml delegates to lxml with the configured parser, so it has no equivalent.
import { describe, expect, it } from "vitest";
import { ValueError } from "../../src/errors.js";
import { qn } from "../../src/oxml/ns.js";
import { parseXml, registerElementCls } from "../../src/oxml/parser.js";
import { tostring } from "../../src/oxml/xml.js";
import { BaseOxmlElement } from "../../src/oxml/xmlchemy.js";

const xmlText =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<a:foo xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">\n' +
  "  <a:bar>foobar</a:bar>\n" +
  "</a:foo>\n";
const xmlBytes = new TextEncoder().encode(xmlText);

class CustElmCls extends BaseOxmlElement {}

describe("DescribeOxmlParser", () => {
  it("it_strips_whitespace_between_elements", () => {
    const foo = parseXml(xmlBytes);
    expect(tostring(foo)).toBe('<a:foo xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bar>foobar</a:bar></a:foo>');
  });
});

describe("DescribeParseXml", () => {
  it("it_prefers_to_parse_bytes", () => {
    parseXml(xmlBytes);
  });

  it("but_accepts_unicode_providing_there_is_no_encoding_declaration", () => {
    const nonEncDecl = '<?xml version="1.0" standalone="yes"?>';
    const encDecl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    const xmlBody = "<foo><bar>føøbår</bar></foo>";
    // unicode body by itself doesn't raise
    parseXml(xmlBody);
    // adding XML decl without encoding attr doesn't raise either
    parseXml(`${nonEncDecl}\n${xmlBody}`);
    // but adding encoding in the declaration raises ValueError
    expect(() => parseXml(`${encDecl}\n${xmlBody}`)).toThrow(ValueError);
  });
});

describe("DescribeRegisterCustomElementClass", () => {
  it("it_determines_cust_elm_class_constructed_for_specified_tag", () => {
    registerElementCls("a:foo", CustElmCls);
    const foo = parseXml(xmlBytes);
    expect(foo.constructor).toBe(CustElmCls);
    // unregistered tags get the default element class
    expect(foo.find(qn("a:bar")).constructor).toBe(BaseOxmlElement);
  });
});
