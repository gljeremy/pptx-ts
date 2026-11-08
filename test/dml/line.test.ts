// Machine-translated from python-pptx tests/dml/test_line.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $call, $is, $kw, $list, $t1, $truthy, call, classMock, instanceMock, propertyMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { ColorFormat as ColorFormat$imported } from "../../src/dml/color.js";
import { FillFormat as FillFormat$imported } from "../../src/dml/fill.js";
import { LineFormat as LineFormat$imported } from "../../src/dml/line.js";
import { MSO_FILL as MSO_FILL$imported, MSO_LINE as MSO_LINE$imported } from "../../src/enum/dml.js";
import { CT_LineProperties as CT_LineProperties$imported } from "../../src/oxml/shapes/shared.js";
import { Shape as Shape$imported } from "../../src/shapes/autoshape.js";
import { anLn } from "../oxml/unitdata/dml.js";
import { element, xml } from "../unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const ColorFormat = ColorFormat$imported;
const FillFormat = FillFormat$imported;
const LineFormat = LineFormat$imported;
const MSO_FILL = MSO_FILL$imported;
const MSO_LINE = MSO_LINE$imported;
const CT_LineProperties = CT_LineProperties$imported;
const Shape = Shape$imported;
void $py;


const DescribeLineFormat: any = {
  fixtures: {
    color_fixture: { args: ["request", "line", "fill_prop_", "fill_", "color_"], params: [[MSO_FILL.SOLID, false], [MSO_FILL.BACKGROUND, true], [null, true]], fn: (self: any, request: any, line: any, fill_prop_: any, fill_: any, color_: any): any => {
        let expected_solid_calls: any, pre_call_fill_type: any, solid_call_expected: any;
        [pre_call_fill_type, solid_call_expected] = $list(request.param);
        fill_.type = pre_call_fill_type;
        expected_solid_calls = ($truthy(solid_call_expected) ? [call()] : []);
        return [line, fill_, expected_solid_calls, color_];
      } },
    dash_style_get_fixture: { args: ["request"], params: [["p:spPr", null], ["p:spPr/a:ln", null], ["p:spPr/a:ln/a:prstDash", null], ["p:spPr/a:ln/a:prstDash{val=dash}", MSO_LINE.DASH], ["p:spPr/a:ln/a:prstDash{val=solid}", MSO_LINE.SOLID]], fn: (self: any, request: any): any => {
        let expected_value: any, line: any, spPr: any, spPr_cxml: any;
        [spPr_cxml, expected_value] = $list(request.param);
        spPr = element(spPr_cxml);
        line = $call(LineFormat, spPr);
        return [line, expected_value];
      } },
    fill_fixture: { args: ["line", "FillFormat_", "ln_", "fill_"], fn: (self: any, line: any, FillFormat_: any, ln_: any, fill_: any): any => {
        return [line, FillFormat_, ln_, fill_];
      } },
    width_get_fixture: { args: ["request", "shape_"], params: [[null, 0], [12700, 12700]], fn: (self: any, request: any, shape_: any): any => {
        let expected_line_width: any, line: any, w: any;
        [w, expected_line_width] = $list(request.param);
        shape_.ln = self.lnBldr(w).element;
        line = $call(LineFormat, shape_);
        return [line, expected_line_width];
      } },
    width_set_fixture: { args: ["request", "shape_"], params: [[null, null], [null, 12700], [12700, 12700], [12700, 25400], [25400, null]], fn: (self: any, request: any, shape_: any): any => {
        let expected_xml: any, initial_width: any, line: any, width: any;
        [initial_width, width] = $list(request.param);
        const $t1 = self.lnBldr(initial_width).element;
        shape_.ln = $t1;
        shape_.getOrAddLn.returnValue = $t1;
        line = $call(LineFormat, shape_);
        expected_xml = self.lnBldr(width).xml();
        return [line, width, expected_xml];
      } },
    color_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, ColorFormat);
      } },
    fill_: { args: ["request", "color_"], fn: (self: any, request: any, color_: any): any => {
        return $call(instanceMock, request, FillFormat, $kw({fore_color: color_}));
      } },
    fill_prop_: { args: ["request", "fill_"], fn: (self: any, request: any, fill_: any): any => {
        return $call(propertyMock, request, LineFormat, "fill", $kw({return_value: fill_}));
      } },
    FillFormat_: { args: ["request", "fill_"], fn: (self: any, request: any, fill_: any): any => {
        let FillFormat_: any;
        FillFormat_ = classMock(request, "pptx.dml.line.FillFormat");
        FillFormat_.fromFillParent.returnValue = fill_;
        return FillFormat_;
      } },
    line: { args: ["shape_"], fn: (self: any, shape_: any): any => {
        return $call(LineFormat, shape_);
      } },
    ln_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, CT_LineProperties);
      } },
    shape_: { args: ["request", "ln_"], fn: (self: any, request: any, ln_: any): any => {
        let shape_: any;
        shape_ = instanceMock(request, Shape);
        shape_.getOrAddLn.returnValue = ln_;
        return shape_;
      } },
  },
  helpers: {
    lnBldr: (self: any, w: any): any => {
        let ln_bldr: any;
        ln_bldr = anLn().withNsdecls();
        if ($truthy(!$is(w, null))) {
          ln_bldr.withW(w);
        }
        return ln_bldr;
      },
  },
  tests: [
    { name: "it_knows_its_dash_style", args: ["dash_style_get_fixture"], fn: (self: any, dash_style_get_fixture: any): any => {
        let expected_value: any, line: any;
        [line, expected_value] = $list(dash_style_get_fixture);
        $assertEqual(line.dashStyle, expected_value);
      } },
    { name: "it_can_change_its_dash_style", args: ["spPr_cxml", "dash_style", "expected_cxml"], parametrize: [{ names: ["spPr_cxml", "dash_style", "expected_cxml"], values: [["p:spPr{a:b=c}", MSO_LINE.DASH, "p:spPr{a:b=c}/a:ln/a:prstDash{val=dash}"], ["p:spPr/a:ln", MSO_LINE.ROUND_DOT, "p:spPr/a:ln/a:prstDash{val=sysDot}"], ["p:spPr/a:ln/a:prstDash", MSO_LINE.SOLID, "p:spPr/a:ln/a:prstDash{val=solid}"], ["p:spPr/a:ln/a:custDash", MSO_LINE.DASH_DOT, "p:spPr/a:ln/a:prstDash{val=dashDot}"], ["p:spPr/a:ln/a:prstDash{val=dash}", MSO_LINE.LONG_DASH, "p:spPr/a:ln/a:prstDash{val=lgDash}"], ["p:spPr/a:ln/a:prstDash{val=dash}", null, "p:spPr/a:ln"], ["p:spPr/a:ln/a:custDash", null, "p:spPr/a:ln"]] }], fn: (self: any, spPr_cxml: any, dash_style: any, expected_cxml: any): any => {
        let line: any, spPr: any;
        spPr = element(spPr_cxml);
        line = $call(LineFormat, spPr);
        line.dashStyle = dash_style;
        $assertEqual(spPr.xml, xml(expected_cxml));
      } },
    { name: "it_knows_its_width", args: ["width_get_fixture"], fn: (self: any, width_get_fixture: any): any => {
        let expected_line_width: any, line: any;
        [line, expected_line_width] = $list(width_get_fixture);
        $assertEqual(line.width, expected_line_width);
      } },
    { name: "it_can_change_its_width", args: ["width_set_fixture"], fn: (self: any, width_set_fixture: any): any => {
        let expected_xml: any, line: any, width: any;
        [line, width, expected_xml] = $list(width_set_fixture);
        line.width = width;
        $assertEqual(line._ln.xml, expected_xml);
      } },
    { name: "it_has_a_fill", args: ["fill_fixture"], fn: (self: any, fill_fixture: any): any => {
        let FillFormat_: any, fill: any, fill_: any, line: any, ln_: any;
        [line, FillFormat_, ln_, fill_] = $list(fill_fixture);
        fill = line.fill;
        FillFormat_.fromFillParent.assertCalledOnceWith(ln_);
        $assert($is(fill, fill_));
      } },
    { name: "it_has_a_color", args: ["color_fixture"], fn: (self: any, color_fixture: any): any => {
        let color: any, color_: any, expected_solid_calls: any, fill_: any, line: any;
        [line, fill_, expected_solid_calls, color_] = $list(color_fixture);
        color = line.color;
        $assertEqual(fill_.solid.mockCalls, expected_solid_calls);
        $assert($is(color, color_));
      } },
  ],
};
pytestClass("DescribeLineFormat", DescribeLineFormat);

