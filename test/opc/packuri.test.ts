// Machine-translated from python-pptx tests/opc/test_packuri.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assertEqual, $call, $raises, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { PackURI } from "../../src/opc/packuri.js";
void $py;


const DescribePackURI: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_can_construct_from_relative_ref", args: [], fn: (self: any): any => {
        let pack_uri: any;
        pack_uri = PackURI.fromRelRef("/ppt/slides", "../slideLayouts/slideLayout1.xml");
        $assertEqual(pack_uri, "/ppt/slideLayouts/slideLayout1.xml");
      } },
    { name: "it_should_raise_on_construct_with_bad_pack_uri_str", args: [], fn: (self: any): any => {
        $raises(ValueError, () => {
          $call(PackURI, "foobar");
        });
      } },
    { name: "it_knows_its_base_URI", args: ["uri", "expected_value"], parametrize: [{ names: ["uri", "expected_value"], values: [["/", "/"], ["/ppt/presentation.xml", "/ppt"], ["/ppt/slides/slide1.xml", "/ppt/slides"]] }], fn: (self: any, uri: any, expected_value: any): any => {
        $assertEqual($call(PackURI, uri).baseURI, expected_value);
      } },
    { name: "it_knows_its_extension", args: ["uri", "expected_value"], parametrize: [{ names: ["uri", "expected_value"], values: [["/", ""], ["/ppt/presentation.xml", "xml"], ["/ppt/media/image.PnG", "PnG"]] }], fn: (self: any, uri: any, expected_value: any): any => {
        $assertEqual($call(PackURI, uri).ext, expected_value);
      } },
    { name: "it_knows_its_filename", args: ["uri", "expected_value"], parametrize: [{ names: ["uri", "expected_value"], values: [["/", ""], ["/ppt/presentation.xml", "presentation.xml"], ["/ppt/media/image.png", "image.png"]] }], fn: (self: any, uri: any, expected_value: any): any => {
        $assertEqual($call(PackURI, uri).filename, expected_value);
      } },
    { name: "it_knows_the_filename_index", args: ["uri", "expected_value"], parametrize: [{ names: ["uri", "expected_value"], values: [["/", null], ["/ppt/presentation.xml", null], ["/ppt/,foo,grob!.xml", null], ["/ppt/media/image42.png", 42]] }], fn: (self: any, uri: any, expected_value: any): any => {
        $assertEqual($call(PackURI, uri).idx, expected_value);
      } },
    { name: "it_can_compute_its_relative_reference", args: ["uri", "base_uri", "expected_value"], parametrize: [{ names: ["uri", "base_uri", "expected_value"], values: [["/ppt/presentation.xml", "/", "ppt/presentation.xml"], ["/ppt/slideMasters/slideMaster1.xml", "/ppt", "slideMasters/slideMaster1.xml"], ["/ppt/slideLayouts/slideLayout1.xml", "/ppt/slides", "../slideLayouts/slideLayout1.xml"]] }], fn: (self: any, uri: any, base_uri: any, expected_value: any): any => {
        $assertEqual($call(PackURI, uri).relativeRef(base_uri), expected_value);
      } },
    { name: "it_knows_the_uri_of_its_rels_part", args: ["uri", "expected_value"], parametrize: [{ names: ["uri", "expected_value"], values: [["/", "/_rels/.rels"], ["/ppt/presentation.xml", "/ppt/_rels/presentation.xml.rels"], ["/ppt/slides/slide42.xml", "/ppt/slides/_rels/slide42.xml.rels"]] }], fn: (self: any, uri: any, expected_value: any): any => {
        $assertEqual($call(PackURI, uri).relsUri, expected_value);
      } },
  ],
};
pytestClass("DescribePackURI", DescribePackURI);

