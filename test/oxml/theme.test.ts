// Machine-translated from python-pptx tests/oxml/test_theme.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assertEqual, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { CT_OfficeStyleSheet as CT_OfficeStyleSheet$imported } from "../../src/oxml/theme.js";
import { snippetText } from "../unitutil/file.js";
// python binds imported names at import time; patches made later don't change them
const CT_OfficeStyleSheet = CT_OfficeStyleSheet$imported;
void $py;


const DescribeCT_OfficeStyleSheet: any = {
  fixtures: {
    new_fixture: { args: [], fn: (self: any): any => {
        let expected_xml: any;
        expected_xml = snippetText("default-theme");
        return expected_xml;
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_create_a_default_theme_element", args: ["new_fixture"], fn: (self: any, new_fixture: any): any => {
        let expected_xml: any, theme: any;
        expected_xml = new_fixture;
        theme = CT_OfficeStyleSheet.newDefault();
        $assertEqual(theme.xml, expected_xml);
      } },
  ],
};
pytestClass("DescribeCT_OfficeStyleSheet", DescribeCT_OfficeStyleSheet);

