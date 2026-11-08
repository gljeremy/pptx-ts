// Machine-translated from python-pptx tests/chart/test_marker.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $call, $get, $is, $kw, $list, classMock, instanceMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { Marker as Marker$imported } from "../../src/chart/marker.js";
import { ChartFormat as ChartFormat$imported } from "../../src/dml/chtfmt.js";
import { XL_MARKER_STYLE as XL_MARKER_STYLE$imported } from "../../src/enum/chart.js";
import { element, xml } from "../unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const Marker = Marker$imported;
const ChartFormat = ChartFormat$imported;
const XL_MARKER_STYLE = XL_MARKER_STYLE$imported;
void $py;


const DescribeMarker: any = {
  fixtures: {
    format_fixture: { args: ["request", "ChartFormat_", "chart_format_"], params: [["c:ser", "c:ser/c:marker"], ["c:ser/c:marker", "c:ser/c:marker"], ["c:dPt", "c:dPt/c:marker"], ["c:dPt/c:marker", "c:dPt/c:marker"]], fn: (self: any, request: any, ChartFormat_: any, chart_format_: any): any => {
        let _element: any, cxml: any, expected_cxml: any, expected_xml: any, marker: any;
        [cxml, expected_cxml] = $list(request.param);
        _element = element(cxml);
        marker = $call(Marker, _element);
        expected_xml = xml(expected_cxml);
        return [marker, ChartFormat_, _element, chart_format_, expected_xml];
      } },
    size_get_fixture: { args: ["request"], params: [["c:ser", null], ["c:ser/c:marker", null], ["c:ser/c:marker/c:size{val=24}", 24], ["c:dPt", null], ["c:dPt/c:marker", null], ["c:dPt/c:marker/c:size{val=36}", 36]], fn: (self: any, request: any): any => {
        let cxml: any, expected_value: any, marker: any;
        [cxml, expected_value] = $list(request.param);
        marker = $call(Marker, element(cxml));
        return [marker, expected_value];
      } },
    size_set_fixture: { args: ["request"], params: [["c:ser", 42, "c:ser/c:marker/c:size{val=42}"], ["c:ser/c:marker", 42, "c:ser/c:marker/c:size{val=42}"], ["c:ser/c:marker/c:size{val=42}", 24, "c:ser/c:marker/c:size{val=24}"], ["c:ser/c:marker/c:size{val=24}", null, "c:ser/c:marker"], ["c:ser", null, "c:ser/c:marker"]], fn: (self: any, request: any): any => {
        let cxml: any, expected_cxml: any, expected_xml: any, marker: any, value: any;
        [cxml, value, expected_cxml] = $list(request.param);
        marker = $call(Marker, element(cxml));
        expected_xml = xml(expected_cxml);
        return [marker, value, expected_xml];
      } },
    style_get_fixture: { args: ["request"], params: [["c:ser", null], ["c:ser/c:marker", null], ["c:ser/c:marker/c:symbol{val=circle}", XL_MARKER_STYLE.CIRCLE], ["c:dPt", null], ["c:dPt/c:marker", null], ["c:dPt/c:marker/c:symbol{val=square}", XL_MARKER_STYLE.SQUARE]], fn: (self: any, request: any): any => {
        let cxml: any, expected_value: any, marker: any;
        [cxml, expected_value] = $list(request.param);
        marker = $call(Marker, element(cxml));
        return [marker, expected_value];
      } },
    style_set_fixture: { args: ["request"], params: [["c:ser", XL_MARKER_STYLE.CIRCLE, "c:ser/c:marker/c:symbol{val=circle}"], ["c:ser/c:marker", XL_MARKER_STYLE.SQUARE, "c:ser/c:marker/c:symbol{val=square}"], ["c:ser/c:marker/c:symbol{val=square}", XL_MARKER_STYLE.AUTOMATIC, "c:ser/c:marker/c:symbol{val=auto}"], ["c:ser/c:marker/c:symbol{val=auto}", null, "c:ser/c:marker"], ["c:ser", null, "c:ser/c:marker"]], fn: (self: any, request: any): any => {
        let cxml: any, expected_cxml: any, expected_xml: any, marker: any, value: any;
        [cxml, value, expected_cxml] = $list(request.param);
        marker = $call(Marker, element(cxml));
        expected_xml = xml(expected_cxml);
        return [marker, value, expected_xml];
      } },
    ChartFormat_: { args: ["request", "chart_format_"], fn: (self: any, request: any, chart_format_: any): any => {
        return $call(classMock, request, "pptx.chart.marker.ChartFormat", $kw({return_value: chart_format_}));
      } },
    chart_format_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, ChartFormat);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_size", args: ["size_get_fixture"], fn: (self: any, size_get_fixture: any): any => {
        let expected_value: any, marker: any;
        [marker, expected_value] = $list(size_get_fixture);
        $assertEqual(marker.size, expected_value);
      } },
    { name: "it_can_change_its_size", args: ["size_set_fixture"], fn: (self: any, size_set_fixture: any): any => {
        let expected_xml: any, marker: any, value: any;
        [marker, value, expected_xml] = $list(size_set_fixture);
        marker.size = value;
        $assertEqual(marker._element.xml, expected_xml);
      } },
    { name: "it_knows_its_style", args: ["style_get_fixture"], fn: (self: any, style_get_fixture: any): any => {
        let expected_value: any, marker: any;
        [marker, expected_value] = $list(style_get_fixture);
        $assertEqual(marker.style, expected_value);
      } },
    { name: "it_can_change_its_style", args: ["style_set_fixture"], fn: (self: any, style_set_fixture: any): any => {
        let expected_xml: any, marker: any, value: any;
        [marker, value, expected_xml] = $list(style_set_fixture);
        marker.style = value;
        $assertEqual(marker._element.xml, expected_xml);
      } },
    { name: "it_provides_access_to_its_format", args: ["format_fixture"], fn: (self: any, format_fixture: any): any => {
        let ChartFormat_: any, _element: any, chart_format: any, chart_format_: any, expected_xml: any, marker: any;
        [marker, ChartFormat_, _element, chart_format_, expected_xml] = $list(format_fixture);
        chart_format = marker.format;
        ChartFormat_.assertCalledOnceWith($get(_element.xpath("c:marker"), 0));
        $assert($is(chart_format, chart_format_));
        $assertEqual(marker._element.xml, expected_xml);
      } },
  ],
};
pytestClass("DescribeMarker", DescribeMarker);

