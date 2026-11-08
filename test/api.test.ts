// Machine-translated from python-pptx tests/test_api.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "./pyshim.js";
import { $assert, $assertEqual, $call, $get, $is, $list, $path, $pptx, ANY, classMock, instanceMock, pytestClass } from "./pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../src/exc.js";
import { defineChildren } from "../src/oxml/xmlchemy.js";
import { Presentation as Presentation$imported } from "../src/api.js";
import { CONTENT_TYPE as CT$imported } from "../src/opc/constants.js";
import { PresentationPart as PresentationPart$imported } from "../src/parts/presentation.js";
// python binds imported names at import time; patches made later don't change them
const Presentation = Presentation$imported;
const CT = CT$imported;
const PresentationPart = PresentationPart$imported;
void $py;


const DescribePresentation: any = {
  fixtures: {
    call_fixture: { args: ["Package_", "prs_", "prs_part_"], fn: (self: any, Package_: any, prs_: any, prs_part_: any): any => {
        let path: any;
        path = $path.abspath($path.join($get($path.split($pptx.__file__), 0), "templates", "default.pptx"));
        Package_.open.returnValue.mainDocumentPart = prs_part_;
        prs_part_.contentType = CT.PML_PRESENTATION_MAIN;
        prs_part_.presentation = prs_;
        return [Package_, path, prs_];
      } },
    Package_: { args: ["request"], fn: (self: any, request: any): any => {
        return classMock(request, "pptx.api.Package");
      } },
    prs_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Presentation);
      } },
    prs_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, PresentationPart);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_opens_default_template_on_no_path_provided", args: ["call_fixture"], fn: (self: any, call_fixture: any): any => {
        let Package_: any, path: any, prs: any, prs_: any;
        [Package_, path, prs_] = $list(call_fixture);
        prs = $call(Presentation);
        // the default template is opened from the bytes bundled in src/templates/data.ts
        Package_.open.assertCalledOnceWith(ANY);
        $assertEqual(Package_.open.callArgs.args[0].getvalue(), new Uint8Array(require("node:fs").readFileSync(path)));
        $assert($is(prs, prs_));
      } },
  ],
};
pytestClass("DescribePresentation", DescribePresentation);

