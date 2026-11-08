// Machine-translated from python-pptx tests/oxml/test_presentation.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $cast, $cmp, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { CT_SlideIdList as CT_SlideIdList$imported } from "../../src/oxml/presentation.js";
import { element, xml } from "../unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const CT_SlideIdList = CT_SlideIdList$imported;
void $py;


const DescribeCT_SlideIdList: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_can_add_a_sldId_element_as_a_child", args: [], fn: (self: any): any => {
        let sldIdLst: any;
        sldIdLst = $cast(CT_SlideIdList, element("p:sldIdLst/p:sldId{r:id=rId4,id=256}"));
        sldIdLst.addSldId("rId1");
        $assertEqual(sldIdLst.xml, xml("p:sldIdLst/(p:sldId{r:id=rId4,id=256},p:sldId{r:id=rId1,id=257})"));
      } },
    { name: "it_knows_the_next_available_slide_id", args: ["sldIdLst_cxml", "expected_value"], parametrize: [{ names: ["sldIdLst_cxml", "expected_value"], values: [["p:sldIdLst", 256], ["p:sldIdLst/p:sldId{id=42}", 256], ["p:sldIdLst/p:sldId{id=256}", 257], ["p:sldIdLst/(p:sldId{id=256},p:sldId{id=712})", 713], ["p:sldIdLst/(p:sldId{id=280},p:sldId{id=257})", 281]] }], fn: (self: any, sldIdLst_cxml: any, expected_value: any): any => {
        let sldIdLst: any;
        sldIdLst = $cast(CT_SlideIdList, element(sldIdLst_cxml));
        $assertEqual(sldIdLst._nextId, expected_value);
      } },
    { name: "and_it_chooses_a_valid_slide_id_when_max_slide_id_is_used_for_a_slide", args: ["sldIdLst_cxml", "expected_value"], parametrize: [{ names: ["sldIdLst_cxml", "expected_value"], values: [["p:sldIdLst/p:sldId{id=2147483646}", 2147483647], ["p:sldIdLst/p:sldId{id=2147483647}", 256], ["p:sldIdLst/p:sldId{id=2147483648}", 256], ["p:sldIdLst/(p:sldId{id=256},p:sldId{id=2147483647})", 257], ["p:sldIdLst/(p:sldId{id=256},p:sldId{id=2147483647},p:sldId{id=257})", 258], ["p:sldIdLst/(p:sldId{id=245},p:sldId{id=2147483647},p:sldId{id=256})", 257]] }], fn: (self: any, sldIdLst_cxml: any, expected_value: any): any => {
        let sldIdLst: any, slide_id: any;
        sldIdLst = $cast(CT_SlideIdList, element(sldIdLst_cxml));
        slide_id = sldIdLst._nextId;
        $assert($cmp(256, "<=", slide_id) && $cmp(slide_id, "<=", 2147483647));
        $assertEqual(slide_id, expected_value);
      } },
  ],
};
pytestClass("DescribeCT_SlideIdList", DescribeCT_SlideIdList);

