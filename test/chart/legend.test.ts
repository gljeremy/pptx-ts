// Machine-translated from python-pptx tests/chart/test_legend.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $call, $get, $getattr, $isinstance, $list, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { Legend as Legend$imported } from "../../src/chart/legend.js";
import { XL_LEGEND_POSITION as XL_LEGEND_POSITION$imported } from "../../src/enum/chart.js";
import { Font as Font$imported } from "../../src/text/text.js";
import { element, xml } from "../unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const Legend = Legend$imported;
const XL_LEGEND_POSITION = XL_LEGEND_POSITION$imported;
const Font = Font$imported;
void $py;


const DescribeLegend: any = {
  fixtures: {
    font_fixture: { args: ["request"], params: [["c:legend{a:b=c}", "c:legend{a:b=c}/c:txPr/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr)"]], fn: (self: any, request: any): any => {
        let expected_cxml: any, expected_xml: any, legend: any, legend_cxml: any;
        [legend_cxml, expected_cxml] = $list(request.param);
        legend = $call(Legend, element(legend_cxml));
        expected_xml = xml(expected_cxml);
        return [legend, expected_xml];
      } },
    horz_offset_get_fixture: { args: ["request"], params: [["c:legend", 0.0], ["c:legend/c:layout", 0.0], ["c:legend/c:layout/c:manualLayout", 0.0], ["c:legend/c:layout/c:manualLayout/c:xMode", 0.0], ["c:legend/c:layout/c:manualLayout/c:xMode{val=edge}", 0.0], ["c:legend/c:layout/c:manualLayout/c:xMode{val=factor}", 0.0], ["c:legend/c:layout/c:manualLayout/(c:xMode,c:x{val=0.42})", 0.42], ["c:legend/c:layout/c:manualLayout/(c:xMode{val=factor},c:x{val=0.42})", 0.42]], fn: (self: any, request: any): any => {
        let expected_value: any, legend: any, legend_cxml: any;
        [legend_cxml, expected_value] = $list(request.param);
        legend = $call(Legend, element(legend_cxml));
        return [legend, expected_value];
      } },
    horz_offset_set_fixture: { args: ["request"], params: [["c:legend", 0.42, "c:legend/c:layout/c:manualLayout/(c:xMode,c:x{val=0.42})"], ["c:legend/c:layout", 0.42, "c:legend/c:layout/c:manualLayout/(c:xMode,c:x{val=0.42})"], ["c:legend/c:layout/c:manualLayout", -0.42, "c:legend/c:layout/c:manualLayout/(c:xMode,c:x{val=-0.42})"], ["c:legend/c:layout/c:manualLayout/c:xMode{val=edge}", -0.1, "c:legend/c:layout/c:manualLayout/(c:xMode,c:x{val=-0.1})"], ["c:legend/c:layout/c:manualLayout/c:xMode{val=factor}", 0.2, "c:legend/c:layout/c:manualLayout/(c:xMode,c:x{val=0.2})"], ["c:legend/c:layout/c:manualLayout", 0, "c:legend/c:layout"]], fn: (self: any, request: any): any => {
        let expected_legend_cxml: any, expected_xml: any, legend: any, legend_cxml: any, new_value: any;
        [legend_cxml, new_value, expected_legend_cxml] = $list(request.param);
        legend = $call(Legend, element(legend_cxml));
        expected_xml = xml(expected_legend_cxml);
        return [legend, new_value, expected_xml];
      } },
    include_in_layout_get_fixture: { args: ["request"], params: [["c:legend", true], ["c:legend/c:overlay", true], ["c:legend/c:overlay{val=1}", true], ["c:legend/c:overlay{val=0}", false]], fn: (self: any, request: any): any => {
        let expected_value: any, legend: any, legend_cxml: any;
        [legend_cxml, expected_value] = $list(request.param);
        legend = $call(Legend, element(legend_cxml));
        return [legend, expected_value];
      } },
    include_in_layout_set_fixture: { args: ["request"], params: [["c:legend", true, "c:legend/c:overlay{val=1}"], ["c:legend", false, "c:legend/c:overlay{val=0}"], ["c:legend/c:overlay{val=0}", true, "c:legend/c:overlay{val=1}"], ["c:legend/c:overlay{val=1}", false, "c:legend/c:overlay{val=0}"], ["c:legend/c:overlay{val=1}", null, "c:legend"]], fn: (self: any, request: any): any => {
        let expected_legend_cxml: any, expected_xml: any, legend: any, legend_cxml: any, new_value: any;
        [legend_cxml, new_value, expected_legend_cxml] = $list(request.param);
        legend = $call(Legend, element(legend_cxml));
        expected_xml = xml(expected_legend_cxml);
        return [legend, new_value, expected_xml];
      } },
    position_get_fixture: { args: ["request"], params: [["c:legend", "RIGHT"], ["c:legend/c:legendPos", "RIGHT"], ["c:legend/c:legendPos{val=r}", "RIGHT"], ["c:legend/c:legendPos{val=b}", "BOTTOM"]], fn: (self: any, request: any): any => {
        let expected_enum_member: any, expected_value: any, legend: any, legend_cxml: any;
        [legend_cxml, expected_enum_member] = $list(request.param);
        legend = $call(Legend, element(legend_cxml));
        expected_value = $getattr(XL_LEGEND_POSITION, expected_enum_member);
        return [legend, expected_value];
      } },
    position_set_fixture: { args: ["request"], params: [["c:legend/c:legendPos{val=r}", "BOTTOM", "c:legend/c:legendPos{val=b}"], ["c:legend/c:legendPos{val=b}", "RIGHT", "c:legend/c:legendPos"], ["c:legend", "TOP", "c:legend/c:legendPos{val=t}"], ["c:legend/c:legendPos", "CORNER", "c:legend/c:legendPos{val=tr}"]], fn: (self: any, request: any): any => {
        let expected_legend_cxml: any, expected_xml: any, legend: any, legend_cxml: any, new_enum_member: any, new_value: any;
        [legend_cxml, new_enum_member, expected_legend_cxml] = $list(request.param);
        legend = $call(Legend, element(legend_cxml));
        new_value = $getattr(XL_LEGEND_POSITION, new_enum_member);
        expected_xml = xml(expected_legend_cxml);
        return [legend, new_value, expected_xml];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_access_to_its_font", args: ["font_fixture"], fn: (self: any, font_fixture: any): any => {
        let expected_xml: any, font: any, legend: any;
        [legend, expected_xml] = $list(font_fixture);
        font = legend.font;
        $assertEqual(legend._element.xml, expected_xml);
        $assert($isinstance(font, Font));
        $assertEqual(font._element, $get(legend._element.xpath(".//a:defRPr"), 0));
      } },
    { name: "it_knows_its_horizontal_offset", args: ["horz_offset_get_fixture"], fn: (self: any, horz_offset_get_fixture: any): any => {
        let expected_value: any, legend: any;
        [legend, expected_value] = $list(horz_offset_get_fixture);
        $assertEqual(legend.horzOffset, expected_value);
      } },
    { name: "it_can_change_its_horizontal_offset", args: ["horz_offset_set_fixture"], fn: (self: any, horz_offset_set_fixture: any): any => {
        let expected_xml: any, legend: any, new_value: any;
        [legend, new_value, expected_xml] = $list(horz_offset_set_fixture);
        legend.horzOffset = new_value;
        $assertEqual(legend._element.xml, expected_xml);
      } },
    { name: "it_knows_whether_it_should_overlap_the_chart", args: ["include_in_layout_get_fixture"], fn: (self: any, include_in_layout_get_fixture: any): any => {
        let expected_value: any, legend: any;
        [legend, expected_value] = $list(include_in_layout_get_fixture);
        $assertEqual(legend.includeInLayout, expected_value);
      } },
    { name: "it_can_change_whether_it_overlaps_the_chart", args: ["include_in_layout_set_fixture"], fn: (self: any, include_in_layout_set_fixture: any): any => {
        let expected_xml: any, legend: any, new_value: any;
        [legend, new_value, expected_xml] = $list(include_in_layout_set_fixture);
        legend.includeInLayout = new_value;
        $assertEqual(legend._element.xml, expected_xml);
      } },
    { name: "it_knows_its_position_with_respect_to_the_chart", args: ["position_get_fixture"], fn: (self: any, position_get_fixture: any): any => {
        let expected_value: any, legend: any;
        [legend, expected_value] = $list(position_get_fixture);
        $assertEqual(legend.position, expected_value);
      } },
    { name: "it_can_change_its_position", args: ["position_set_fixture"], fn: (self: any, position_set_fixture: any): any => {
        let expected_xml: any, legend: any, new_value: any;
        [legend, new_value, expected_xml] = $list(position_set_fixture);
        legend.position = new_value;
        $assertEqual(legend._element.xml, expected_xml);
      } },
  ],
};
pytestClass("DescribeLegend", DescribeLegend);

