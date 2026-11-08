// Machine-translated from python-pptx tests/dml/test_chtfmt.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $call, $get, $is, $kw, $list, classMock, instanceMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { ChartFormat as ChartFormat$imported } from "../../src/dml/chtfmt.js";
import { FillFormat as FillFormat$imported } from "../../src/dml/fill.js";
import { LineFormat as LineFormat$imported } from "../../src/dml/line.js";
import { element, xml } from "../unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const ChartFormat = ChartFormat$imported;
const FillFormat = FillFormat$imported;
const LineFormat = LineFormat$imported;
void $py;


const DescribeChartFormat: any = {
  fixtures: {
    fill_fixture: { args: ["request", "FillFormat_", "fill_"], params: [["c:catAx", "c:catAx/c:spPr"], ["c:catAx/c:spPr", "c:catAx/c:spPr"], ["c:dPt", "c:dPt/c:spPr"], ["c:dPt/c:spPr", "c:dPt/c:spPr"], ["c:majorGridlines", "c:majorGridlines/c:spPr"], ["c:majorGridlines/c:spPr", "c:majorGridlines/c:spPr"], ["c:valAx", "c:valAx/c:spPr"], ["c:valAx/c:spPr", "c:valAx/c:spPr"]], fn: (self: any, request: any, FillFormat_: any, fill_: any): any => {
        let chart_format: any, dPt_cxml: any, expected_cxml: any, expected_xml: any;
        [dPt_cxml, expected_cxml] = $list(request.param);
        chart_format = $call(ChartFormat, element(dPt_cxml));
        FillFormat_.fromFillParent.returnValue = fill_;
        expected_xml = xml(expected_cxml);
        return [chart_format, FillFormat_, fill_, expected_xml];
      } },
    line_fixture: { args: ["request", "LineFormat_", "line_"], params: [["c:catAx", "c:catAx/c:spPr"], ["c:catAx/c:spPr", "c:catAx/c:spPr"], ["c:dPt", "c:dPt/c:spPr"], ["c:dPt/c:spPr", "c:dPt/c:spPr"], ["c:majorGridlines", "c:majorGridlines/c:spPr"], ["c:majorGridlines/c:spPr", "c:majorGridlines/c:spPr"], ["c:valAx", "c:valAx/c:spPr"], ["c:valAx/c:spPr", "c:valAx/c:spPr"]], fn: (self: any, request: any, LineFormat_: any, line_: any): any => {
        let chart_format: any, cxml: any, expected_cxml: any, expected_xml: any;
        [cxml, expected_cxml] = $list(request.param);
        chart_format = $call(ChartFormat, element(cxml));
        expected_xml = xml(expected_cxml);
        return [chart_format, LineFormat_, line_, expected_xml];
      } },
    FillFormat_: { args: ["request"], fn: (self: any, request: any): any => {
        return classMock(request, "pptx.dml.chtfmt.FillFormat");
      } },
    fill_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, FillFormat);
      } },
    LineFormat_: { args: ["request", "line_"], fn: (self: any, request: any, line_: any): any => {
        return $call(classMock, request, "pptx.dml.chtfmt.LineFormat", $kw({return_value: line_}));
      } },
    line_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, LineFormat);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_access_to_its_fill", args: ["fill_fixture"], fn: (self: any, fill_fixture: any): any => {
        let FillFormat_: any, chart_format: any, expected_xml: any, fill: any, fill_: any;
        [chart_format, FillFormat_, fill_, expected_xml] = $list(fill_fixture);
        fill = chart_format.fill;
        FillFormat_.fromFillParent.assertCalledOnceWith($get(chart_format._element.xpath("c:spPr"), 0));
        $assert($is(fill, fill_));
        $assertEqual(chart_format._element.xml, expected_xml);
      } },
    { name: "it_provides_access_to_its_line", args: ["line_fixture"], fn: (self: any, line_fixture: any): any => {
        let LineFormat_: any, chart_format: any, expected_xml: any, line: any, line_: any;
        [chart_format, LineFormat_, line_, expected_xml] = $list(line_fixture);
        line = chart_format.line;
        LineFormat_.assertCalledOnceWith($get(chart_format._element.xpath("c:spPr"), 0));
        $assert($is(line, line_));
        $assertEqual(chart_format._element.xml, expected_xml);
      } },
  ],
};
pytestClass("DescribeChartFormat", DescribeChartFormat);

