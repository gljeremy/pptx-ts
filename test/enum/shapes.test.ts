// Machine-translated from python-pptx tests/enum/test_shapes.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $is, $isinstance, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { PROG_ID as PROG_ID$imported } from "../../src/enum/shapes.js";
// python binds imported names at import time; patches made later don't change them
const PROG_ID = PROG_ID$imported;
void $py;


const DescribeProgId: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_has_members_for_the_OLE_embeddings_known_to_work_on_Windows", args: [], fn: (self: any): any => {
        $assert(PROG_ID.DOCX);
        $assert(PROG_ID.PPTX);
        $assert(PROG_ID.XLSX);
      } },
    { name: "it_knows_its_height", args: ["member", "expected_value"], parametrize: [{ names: ["member", "expected_value"], values: [[PROG_ID.DOCX, 609600], [PROG_ID.PPTX, 609600], [PROG_ID.XLSX, 609600]] }], fn: (self: any, member: any, expected_value: any): any => {
        $assertEqual(member.height, expected_value);
      } },
    { name: "it_knows_its_icon_filename", args: [], fn: (self: any): any => {
        $assertEqual(PROG_ID.DOCX.iconFilename, "docx-icon.emf");
      } },
    { name: "it_knows_its_progId", args: [], fn: (self: any): any => {
        $assertEqual(PROG_ID.PPTX.progId, "PowerPoint.Show.12");
      } },
    { name: "it_knows_its_width", args: [], fn: (self: any): any => {
        $assertEqual(PROG_ID.XLSX.width, 965200);
      } },
    { name: "it_knows_each_of_its_members_is_an_instance", args: ["value", "expected_value"], parametrize: [{ names: ["value", "expected_value"], values: [[PROG_ID.DOCX, true], [PROG_ID.PPTX, true], [PROG_ID.XLSX, true], [17, false], ["XLSX", false]] }], fn: (self: any, value: any, expected_value: any): any => {
        $assert($is($isinstance(value, PROG_ID), expected_value));
      } },
  ],
};
pytestClass("DescribeProgId", DescribeProgId);

