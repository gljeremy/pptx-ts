// Machine-translated from python-pptx tests/dml/test_effect.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $call, $is, $list, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { ShadowFormat as ShadowFormat$imported } from "../../src/dml/effect.js";
import { element, xml } from "../unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const ShadowFormat = ShadowFormat$imported;
void $py;


const DescribeShadowFormat: any = {
  fixtures: {
    inherit_get_fixture: { args: ["request"], params: [["p:spPr", true], ["p:spPr/a:effectLst", false], ["p:grpSpPr", true], ["p:grpSpPr/a:effectLst", false]], fn: (self: any, request: any): any => {
        let cxml: any, expected_value: any, shadow: any;
        [cxml, expected_value] = $list(request.param);
        shadow = $call(ShadowFormat, element(cxml));
        return [shadow, expected_value];
      } },
    inherit_set_fixture: { args: ["request"], params: [["p:spPr{a:b=c}", false, "p:spPr{a:b=c}/a:effectLst"], ["p:grpSpPr{a:b=c}", false, "p:grpSpPr{a:b=c}/a:effectLst"], ["p:spPr{a:b=c}/a:effectLst", true, "p:spPr{a:b=c}"], ["p:grpSpPr{a:b=c}/a:effectLst", true, "p:grpSpPr{a:b=c}"], ["p:spPr", true, "p:spPr"], ["p:grpSpPr", true, "p:grpSpPr"], ["p:spPr/a:effectLst", false, "p:spPr/a:effectLst"], ["p:grpSpPr/a:effectLst", false, "p:grpSpPr/a:effectLst"]], fn: (self: any, request: any): any => {
        let cxml: any, expected_cxml: any, expected_value: any, shadow: any, value: any;
        [cxml, value, expected_cxml] = $list(request.param);
        shadow = $call(ShadowFormat, element(cxml));
        expected_value = xml(expected_cxml);
        return [shadow, value, expected_value];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_whether_it_inherits", args: ["inherit_get_fixture"], fn: (self: any, inherit_get_fixture: any): any => {
        let expected_value: any, inherit: any, shadow: any;
        [shadow, expected_value] = $list(inherit_get_fixture);
        inherit = shadow.inherit;
        $assert($is(inherit, expected_value));
      } },
    { name: "it_can_change_whether_it_inherits", args: ["inherit_set_fixture"], fn: (self: any, inherit_set_fixture: any): any => {
        let expected_xml: any, shadow: any, value: any;
        [shadow, value, expected_xml] = $list(inherit_set_fixture);
        shadow.inherit = value;
        $assertEqual(shadow._element.xml, expected_xml);
      } },
  ],
};
pytestClass("DescribeShadowFormat", DescribeShadowFormat);

