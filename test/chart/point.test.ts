// Machine-translated from python-pptx tests/chart/test_point.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $call, $get, $is, $kw, $len, $list, $raises, classMock, instanceMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { DataLabel as DataLabel$imported } from "../../src/chart/datalabel.js";
import { Marker as Marker$imported } from "../../src/chart/marker.js";
import { BubblePoints as BubblePoints$imported, CategoryPoints as CategoryPoints$imported, Point as Point$imported, XyPoints as XyPoints$imported } from "../../src/chart/point.js";
import { ChartFormat as ChartFormat$imported } from "../../src/dml/chtfmt.js";
import { element, xml } from "../unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const DataLabel = DataLabel$imported;
const Marker = Marker$imported;
const BubblePoints = BubblePoints$imported;
const CategoryPoints = CategoryPoints$imported;
const Point = Point$imported;
const XyPoints = XyPoints$imported;
const ChartFormat = ChartFormat$imported;
void $py;


const Describe_BasePoints: any = {
  fixtures: {
    getitem_fixture: { args: ["request", "Point_", "point_"], fn: (self: any, request: any, Point_: any, point_: any): any => {
        let idx: any, points: any, ser: any;
        ser = element("c:ser/(c:xVal/c:numRef/c:numCache/c:ptCount{val=3},c:yVal/c:numRef/c:numCache/c:ptCount{val=3})");
        points = $call(XyPoints, ser);
        idx = 2;
        return [points, idx, Point_, ser, point_];
      } },
    Point_: { args: ["request", "point_"], fn: (self: any, request: any, point_: any): any => {
        return $call(classMock, request, "pptx.chart.point.Point", $kw({return_value: point_}));
      } },
    point_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Point);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_supports_indexed_access", args: ["getitem_fixture"], fn: (self: any, getitem_fixture: any): any => {
        let Point_: any, idx: any, point: any, point_: any, points: any, ser: any;
        [points, idx, Point_, ser, point_] = $list(getitem_fixture);
        point = $get(points, idx);
        Point_.assertCalledOnceWith(ser, idx);
        $assert($is(point, point_));
      } },
    { name: "it_raises_on_indexed_access_out_of_range", args: [], fn: (self: any): any => {
        let points: any;
        points = $call(XyPoints, element("c:ser/(c:xVal/c:numRef/c:numCache/c:ptCount{val=3},c:yVal/c:numRef/c:numCache/c:ptCount{val=3})"));
        $raises(IndexError, () => {
          $get(points, -1);
        });
        $raises(IndexError, () => {
          $get(points, 3);
        });
      } },
  ],
};
pytestClass("Describe_BasePoints", Describe_BasePoints);

const DescribeBubblePoints: any = {
  fixtures: {
    len_fixture: { args: ["request"], params: [["c:ser", 0], ["c:ser/c:bubbleSize/c:numRef/c:numCache/c:ptCount{val=3}", 0], ["c:ser/(c:xVal/c:numRef/c:numCache/c:ptCount{val=1},c:yVal/c:numRef/c:numCache/c:ptCount{val=2},c:bubbleSize/c:numRef/c:numCache/c:ptCount{val=3})", 1], ["c:ser/(c:xVal/c:numRef/c:numCache/c:ptCount{val=3},c:yVal/c:numRef/c:numCache/c:ptCount{val=3},c:bubbleSize/c:numRef/c:numCache/c:ptCount{val=3})", 3]], fn: (self: any, request: any): any => {
        let expected_value: any, points: any, ser_cxml: any;
        [ser_cxml, expected_value] = $list(request.param);
        points = $call(BubblePoints, element(ser_cxml));
        return [points, expected_value];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_supports_len", args: ["len_fixture"], fn: (self: any, len_fixture: any): any => {
        let expected_value: any, points: any;
        [points, expected_value] = $list(len_fixture);
        $assertEqual($len(points), expected_value);
      } },
  ],
};
pytestClass("DescribeBubblePoints", DescribeBubblePoints);

const DescribeCategoryPoints: any = {
  fixtures: {
    len_fixture: { args: ["request"], params: [["c:ser", 0], ["c:ser/c:cat/c:numRef/c:numCache/c:ptCount{val=42}", 42], ["c:ser/c:cat/c:numRef/c:numCache/c:ptCount{val=24}", 24]], fn: (self: any, request: any): any => {
        let expected_value: any, points: any, ser_cxml: any;
        [ser_cxml, expected_value] = $list(request.param);
        points = $call(CategoryPoints, element(ser_cxml));
        return [points, expected_value];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_supports_len", args: ["len_fixture"], fn: (self: any, len_fixture: any): any => {
        let expected_value: any, points: any;
        [points, expected_value] = $list(len_fixture);
        $assertEqual($len(points), expected_value);
      } },
  ],
};
pytestClass("DescribeCategoryPoints", DescribeCategoryPoints);

const DescribePoint: any = {
  fixtures: {
    data_label_fixture: { args: ["DataLabel_", "data_label_"], fn: (self: any, DataLabel_: any, data_label_: any): any => {
        let idx: any, point: any, ser: any;
        [ser, idx] = $list([element("c:ser"), 42]);
        point = $call(Point, ser, idx);
        return [point, DataLabel_, ser, idx, data_label_];
      } },
    format_fixture: { args: ["request", "ChartFormat_", "chart_format_"], params: [["c:ser", "c:ser/c:dPt/c:idx{val=42}"], ["c:ser/c:dPt/c:idx{val=42}", "c:ser/c:dPt/c:idx{val=42}"], ["c:ser/c:dPt/c:idx{val=45}", "c:ser/(c:dPt/c:idx{val=45},c:dPt/c:idx{val=42})"]], fn: (self: any, request: any, ChartFormat_: any, chart_format_: any): any => {
        let expected_cxml: any, expected_xml: any, point: any, ser: any, ser_cxml: any;
        [ser_cxml, expected_cxml] = $list(request.param);
        ser = element(ser_cxml);
        point = $call(Point, ser, 42);
        expected_xml = xml(expected_cxml);
        return [point, ChartFormat_, ser, chart_format_, expected_xml];
      } },
    marker_fixture: { args: ["Marker_", "marker_"], fn: (self: any, Marker_: any, marker_: any): any => {
        let dPt: any, point: any, ser: any;
        ser = element("c:ser/c:dPt/c:idx{val=42}");
        point = $call(Point, ser, 42);
        dPt = $get(ser, 0);
        return [point, Marker_, dPt, marker_];
      } },
    ChartFormat_: { args: ["request", "chart_format_"], fn: (self: any, request: any, chart_format_: any): any => {
        return $call(classMock, request, "pptx.chart.point.ChartFormat", $kw({return_value: chart_format_}));
      } },
    chart_format_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, ChartFormat);
      } },
    DataLabel_: { args: ["request", "data_label_"], fn: (self: any, request: any, data_label_: any): any => {
        return $call(classMock, request, "pptx.chart.point.DataLabel", $kw({return_value: data_label_}));
      } },
    data_label_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, DataLabel);
      } },
    Marker_: { args: ["request", "marker_"], fn: (self: any, request: any, marker_: any): any => {
        return $call(classMock, request, "pptx.chart.point.Marker", $kw({return_value: marker_}));
      } },
    marker_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Marker);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_access_to_its_data_label", args: ["data_label_fixture"], fn: (self: any, data_label_fixture: any): any => {
        let DataLabel_: any, data_label: any, data_label_: any, idx: any, point: any, ser: any;
        [point, DataLabel_, ser, idx, data_label_] = $list(data_label_fixture);
        data_label = point.dataLabel;
        DataLabel_.assertCalledOnceWith(ser, idx);
        $assert($is(data_label, data_label_));
      } },
    { name: "it_provides_access_to_its_format", args: ["format_fixture"], fn: (self: any, format_fixture: any): any => {
        let ChartFormat_: any, chart_format: any, chart_format_: any, expected_xml: any, point: any, ser: any;
        [point, ChartFormat_, ser, chart_format_, expected_xml] = $list(format_fixture);
        chart_format = point.format;
        ChartFormat_.assertCalledOnceWith($get(ser.xpath("c:dPt[c:idx/@val=\"42\"]"), 0));
        $assert($is(chart_format, chart_format_));
        $assertEqual(point._element.xml, expected_xml);
      } },
    { name: "it_provides_access_to_its_marker", args: ["marker_fixture"], fn: (self: any, marker_fixture: any): any => {
        let Marker_: any, dPt: any, marker: any, marker_: any, point: any;
        [point, Marker_, dPt, marker_] = $list(marker_fixture);
        marker = point.marker;
        Marker_.assertCalledOnceWith(dPt);
        $assert($is(marker, marker_));
      } },
  ],
};
pytestClass("DescribePoint", DescribePoint);

const DescribeXyPoints: any = {
  fixtures: {
    len_fixture: { args: ["request"], params: [["c:ser", 0], ["c:ser/c:xVal", 0], ["c:ser/c:xVal/c:numRef/c:numCache/c:ptCount{val=3}", 0], ["c:ser/c:yVal/c:numRef/c:numCache/c:ptCount{val=3}", 0], ["c:ser/(c:xVal/c:numRef/c:numCache/c:ptCount{val=1},c:yVal/c:numRef/c:numCache/c:ptCount{val=3})", 1], ["c:ser/(c:xVal/c:numRef/c:numCache/c:ptCount{val=3},c:yVal/c:numRef/c:numCache/c:ptCount{val=1})", 1], ["c:ser/(c:xVal/c:numRef/c:numCache/c:ptCount{val=3},c:yVal/c:numRef/c:numCache/c:ptCount{val=3})", 3]], fn: (self: any, request: any): any => {
        let expected_value: any, points: any, ser_cxml: any;
        [ser_cxml, expected_value] = $list(request.param);
        points = $call(XyPoints, element(ser_cxml));
        return [points, expected_value];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_supports_len", args: ["len_fixture"], fn: (self: any, len_fixture: any): any => {
        let expected_value: any, points: any;
        [points, expected_value] = $list(len_fixture);
        $assertEqual($len(points), expected_value);
      } },
  ],
};
pytestClass("DescribeXyPoints", DescribeXyPoints);

