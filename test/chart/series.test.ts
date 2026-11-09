// Machine-translated from python-pptx tests/chart/test_series.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $call, $get, $is, $isinstance, $kw, $len, $list, $mod, classMock, functionMock, instanceMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { DataLabels as DataLabels$imported } from "../../src/chart/datalabel.js";
import { Marker as Marker$imported } from "../../src/chart/marker.js";
import { BubblePoints as BubblePoints$imported, CategoryPoints as CategoryPoints$imported, XyPoints as XyPoints$imported } from "../../src/chart/point.js";
import { AreaSeries as AreaSeries$imported, BarSeries as BarSeries$imported, BubbleSeries as BubbleSeries$imported, LineSeries as LineSeries$imported, PieSeries as PieSeries$imported, RadarSeries as RadarSeries$imported, SeriesCollection as SeriesCollection$imported, XySeries as XySeries$imported, _BaseCategorySeries as _BaseCategorySeries$imported, _BaseSeries as _BaseSeries$imported, _MarkerMixin as _MarkerMixin$imported, _SeriesFactory as _SeriesFactory$imported } from "../../src/chart/series.js";
import { ChartFormat as ChartFormat$imported } from "../../src/dml/chtfmt.js";
import { element, xml } from "../unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const DataLabels = DataLabels$imported;
const Marker = Marker$imported;
const BubblePoints = BubblePoints$imported;
const CategoryPoints = CategoryPoints$imported;
const XyPoints = XyPoints$imported;
const AreaSeries = AreaSeries$imported;
const BarSeries = BarSeries$imported;
const BubbleSeries = BubbleSeries$imported;
const LineSeries = LineSeries$imported;
const PieSeries = PieSeries$imported;
const RadarSeries = RadarSeries$imported;
const SeriesCollection = SeriesCollection$imported;
const XySeries = XySeries$imported;
const _BaseCategorySeries = _BaseCategorySeries$imported;
const _BaseSeries = _BaseSeries$imported;
const _MarkerMixin = _MarkerMixin$imported;
const _SeriesFactory = _SeriesFactory$imported;
const ChartFormat = ChartFormat$imported;
void $py;


const Describe_BaseSeries: any = {
  fixtures: {
    format_fixture: { args: ["ChartFormat_", "chart_format_"], fn: (self: any, ChartFormat_: any, chart_format_: any): any => {
        let ser: any, series: any;
        ser = element("c:ser");
        series = $call(_BaseSeries, ser);
        return [series, ChartFormat_, ser, chart_format_];
      } },
    index_fixture: { args: [], fn: (self: any): any => {
        let expected_value: any, ser_cxml: any, series: any;
        [ser_cxml, expected_value] = $list(["c:ser/c:idx{val=42}", 42]);
        series = $call(_BaseSeries, element(ser_cxml));
        return [series, expected_value];
      } },
    name_fixture: { args: ["request"], params: [["c:ser", ""], ["c:ser/c:tx/c:strRef/c:strCache/c:pt/c:v\"foobar\"", "foobar"]], fn: (self: any, request: any): any => {
        let expected_value: any, ser_cxml: any, series: any;
        [ser_cxml, expected_value] = $list(request.param);
        series = $call(_BaseSeries, element(ser_cxml));
        return [series, expected_value];
      } },
    ChartFormat_: { args: ["request", "chart_format_"], fn: (self: any, request: any, chart_format_: any): any => {
        return $call(classMock, request, "pptx.chart.series.ChartFormat", $kw({return_value: chart_format_}));
      } },
    chart_format_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, ChartFormat);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_name", args: ["name_fixture"], fn: (self: any, name_fixture: any): any => {
        let expected_value: any, series: any;
        [series, expected_value] = $list(name_fixture);
        $assertEqual(series.name, expected_value);
      } },
    { name: "it_knows_its_position_in_the_series_sequence", args: ["index_fixture"], fn: (self: any, index_fixture: any): any => {
        let expected_value: any, series: any;
        [series, expected_value] = $list(index_fixture);
        $assertEqual(series.index, expected_value);
      } },
    { name: "it_provides_access_to_its_format", args: ["format_fixture"], fn: (self: any, format_fixture: any): any => {
        let ChartFormat_: any, format: any, format_: any, ser: any, series: any;
        [series, ChartFormat_, ser, format_] = $list(format_fixture);
        format = series.format;
        ChartFormat_.assertCalledOnceWith(ser);
        $assert($is(format, format_));
      } },
  ],
};
pytestClass("Describe_BaseSeries", Describe_BaseSeries);

const Describe_BaseCategorySeries: any = {
  fixtures: {
    data_labels_fixture: { args: ["request"], params: [["c:ser", "c:dLbls/(c:showLegendKey{val=0},c:showVal{val=0},c:showCatName{val=0},c:showSerName{val=0},c:showPercent{val=0},c:showBubbleSize{val=0},c:showLeaderLines{val=1})"], ["c:ser/c:dLbls", "c:dLbls"]], fn: (self: any, request: any): any => {
        let expected_dLbls_cxml: any, expected_dLbls_xml: any, ser: any, ser_cxml: any;
        [ser_cxml, expected_dLbls_cxml] = $list(request.param);
        ser = element(ser_cxml);
        expected_dLbls_xml = xml(expected_dLbls_cxml);
        return [ser, expected_dLbls_xml];
      } },
    points_fixture: { args: ["CategoryPoints_", "points_"], fn: (self: any, CategoryPoints_: any, points_: any): any => {
        let ser: any, series: any;
        ser = element("c:ser");
        series = $call(_BaseCategorySeries, ser);
        return [series, CategoryPoints_, ser, points_];
      } },
    subclass_fixture: { args: [], fn: (self: any): any => {
        return $call(_BaseCategorySeries, null);
      } },
    values_get_fixture: { args: ["request"], params: [["c:ser", []], ["c:ser/c:val/c:numRef", []], ["c:ser/c:val/c:numLit", []], ["c:ser/c:val/c:numRef/c:numCache", []], ["c:ser/c:val/c:numRef/c:numCache/c:ptCount{val=0}", []], ["c:ser/c:val/c:numRef/c:numCache/(c:ptCount{val=1},c:pt{idx=0}/c:v\"1.1\")", [1.1]], ["c:ser/c:val/c:numRef/c:numCache/(c:ptCount{val=3},c:pt{idx=0}/c:v\"1.1\",c:pt{idx=2}/c:v\"3.3\")", [1.1, null, 3.3]], ["c:ser/c:val/c:numLit/(c:ptCount{val=3},c:pt{idx=0}/c:v\"1.1\",c:pt{idx=2}/c:v\"3.3\")", [1.1, null, 3.3]], ["c:ser/c:val/c:numRef/c:numCache/(c:ptCount{val=3},c:pt{idx=2}/c:v\"3.3\",c:pt{idx=0}/c:v\"1.1\")", [1.1, null, 3.3]]], fn: (self: any, request: any): any => {
        let expected_value: any, ser_cxml: any, series: any;
        [ser_cxml, expected_value] = $list(request.param);
        series = $call(_BaseCategorySeries, element(ser_cxml));
        return [series, expected_value];
      } },
    CategoryPoints_: { args: ["request", "points_"], fn: (self: any, request: any, points_: any): any => {
        return $call(classMock, request, "pptx.chart.series.CategoryPoints", $kw({return_value: points_}));
      } },
    DataLabels_: { args: ["request"], fn: (self: any, request: any): any => {
        return classMock(request, "pptx.chart.series.DataLabels");
      } },
    data_labels_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, DataLabels);
      } },
    points_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, CategoryPoints);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_is_a_BaseSeries_subclass", args: ["subclass_fixture"], fn: (self: any, subclass_fixture: any): any => {
        let base_category_series: any;
        base_category_series = subclass_fixture;
        $assert($isinstance(base_category_series, _BaseSeries));
      } },
    { name: "it_provides_access_to_its_data_labels", args: ["data_labels_fixture", "DataLabels_", "data_labels_"], fn: (self: any, data_labels_fixture: any, DataLabels_: any, data_labels_: any): any => {
        let dLbls: any, data_labels: any, expected_dLbls_xml: any, ser: any, series: any;
        [ser, expected_dLbls_xml] = $list(data_labels_fixture);
        DataLabels_.returnValue = data_labels_;
        series = $call(_BaseCategorySeries, ser);
        data_labels = series.dataLabels;
        dLbls = $get(ser.xpath("c:dLbls"), 0);
        $assertEqual(dLbls.xml, expected_dLbls_xml);
        DataLabels_.assertCalledOnceWith(dLbls);
        $assert($is(data_labels, data_labels_));
      } },
    { name: "it_provides_access_to_its_points", args: ["points_fixture"], fn: (self: any, points_fixture: any): any => {
        let CategoryPoints_: any, points: any, points_: any, ser: any, series: any;
        [series, CategoryPoints_, ser, points_] = $list(points_fixture);
        points = series.points;
        CategoryPoints_.assertCalledOnceWith(ser);
        $assert($is(points, points_));
      } },
    { name: "it_knows_its_values", args: ["values_get_fixture"], fn: (self: any, values_get_fixture: any): any => {
        let expected_value: any, series: any;
        [series, expected_value] = $list(values_get_fixture);
        $assertEqual(series.values, expected_value);
      } },
  ],
};
pytestClass("Describe_BaseCategorySeries", Describe_BaseCategorySeries);

const Describe_MarkerMixin: any = {
  fixtures: {
    marker_fixture: { args: ["Marker_", "marker_"], fn: (self: any, Marker_: any, marker_: any): any => {
        let ser: any, series: any;
        ser = element("c:ser");
        series = $call(LineSeries, ser);
        return [series, Marker_, ser, marker_];
      } },
    Marker_: { args: ["request", "marker_"], fn: (self: any, request: any, marker_: any): any => {
        return $call(classMock, request, "pptx.chart.series.Marker", $kw({return_value: marker_}));
      } },
    marker_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Marker);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_access_to_the_series_marker", args: ["marker_fixture"], fn: (self: any, marker_fixture: any): any => {
        let Marker_: any, marker: any, marker_: any, ser: any, series: any;
        [series, Marker_, ser, marker_] = $list(marker_fixture);
        marker = series.marker;
        Marker_.assertCalledOnceWith(ser);
        $assert($is(marker, marker_));
      } },
  ],
};
pytestClass("Describe_MarkerMixin", Describe_MarkerMixin);

const DescribeAreaSeries: any = {
  fixtures: {
    subclass_fixture: { args: [], fn: (self: any): any => {
        return $call(AreaSeries, null);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_is_a_BaseCategorySeries_subclass", args: ["subclass_fixture"], fn: (self: any, subclass_fixture: any): any => {
        let area_series: any;
        area_series = subclass_fixture;
        $assert($isinstance(area_series, _BaseCategorySeries));
      } },
  ],
};
pytestClass("DescribeAreaSeries", DescribeAreaSeries);

const DescribeBarSeries: any = {
  fixtures: {
    invert_if_negative_get_fixture: { args: ["request"], params: [["c:ser", true], ["c:ser/c:invertIfNegative", true], ["c:ser/c:invertIfNegative{val=1}", true], ["c:ser/c:invertIfNegative{val=0}", false]], fn: (self: any, request: any): any => {
        let bar_series: any, expected_value: any, ser_cxml: any;
        [ser_cxml, expected_value] = $list(request.param);
        bar_series = $call(BarSeries, element(ser_cxml));
        return [bar_series, expected_value];
      } },
    invert_if_negative_set_fixture: { args: ["request"], params: [["c:ser/c:order", true, "c:ser/(c:order,c:invertIfNegative{val=1})"], ["c:ser/c:order", false, "c:ser/(c:order,c:invertIfNegative{val=0})"], ["c:ser/(c:spPr,c:invertIfNegative{val=0})", true, "c:ser/(c:spPr,c:invertIfNegative{val=1})"], ["c:ser/(c:tx,c:invertIfNegative{val=1})", false, "c:ser/(c:tx,c:invertIfNegative{val=0})"]], fn: (self: any, request: any): any => {
        let bar_series: any, expected_ser_cxml: any, expected_xml: any, new_value: any, ser_cxml: any;
        [ser_cxml, new_value, expected_ser_cxml] = $list(request.param);
        bar_series = $call(BarSeries, element(ser_cxml));
        expected_xml = xml(expected_ser_cxml);
        return [bar_series, new_value, expected_xml];
      } },
    subclass_fixture: { args: [], fn: (self: any): any => {
        return $call(BarSeries, null);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_is_a_BaseCategorySeries_subclass", args: ["subclass_fixture"], fn: (self: any, subclass_fixture: any): any => {
        let bar_series: any;
        bar_series = subclass_fixture;
        $assert($isinstance(bar_series, _BaseCategorySeries));
      } },
    { name: "it_knows_whether_it_should_invert_if_negative", args: ["invert_if_negative_get_fixture"], fn: (self: any, invert_if_negative_get_fixture: any): any => {
        let bar_series: any, expected_value: any;
        [bar_series, expected_value] = $list(invert_if_negative_get_fixture);
        $assertEqual(bar_series.invertIfNegative, expected_value);
      } },
    { name: "it_can_change_whether_it_inverts_if_negative", args: ["invert_if_negative_set_fixture"], fn: (self: any, invert_if_negative_set_fixture: any): any => {
        let bar_series: any, expected_xml: any, new_value: any;
        [bar_series, new_value, expected_xml] = $list(invert_if_negative_set_fixture);
        bar_series.invertIfNegative = new_value;
        $assertEqual(bar_series._element.xml, expected_xml);
      } },
  ],
};
pytestClass("DescribeBarSeries", DescribeBarSeries);

const Describe_BubbleSeries: any = {
  fixtures: {
    points_fixture: { args: ["BubblePoints_", "points_"], fn: (self: any, BubblePoints_: any, points_: any): any => {
        let ser: any, series: any;
        ser = element("c:ser");
        series = $call(BubbleSeries, ser);
        return [series, BubblePoints_, ser, points_];
      } },
    BubblePoints_: { args: ["request", "points_"], fn: (self: any, request: any, points_: any): any => {
        return $call(classMock, request, "pptx.chart.series.BubblePoints", $kw({return_value: points_}));
      } },
    points_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, BubblePoints);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_access_to_its_points", args: ["points_fixture"], fn: (self: any, points_fixture: any): any => {
        let BubblePoints_: any, points: any, points_: any, ser: any, series: any;
        [series, BubblePoints_, ser, points_] = $list(points_fixture);
        points = series.points;
        BubblePoints_.assertCalledOnceWith(ser);
        $assert($is(points, points_));
      } },
  ],
};
pytestClass("Describe_BubbleSeries", Describe_BubbleSeries);

const DescribeLineSeries: any = {
  fixtures: {
    smooth_get_fixture: { args: ["request"], params: [["c:ser", true], ["c:ser/c:smooth", true], ["c:ser/c:smooth{val=1}", true], ["c:ser/c:smooth{val=0}", false]], fn: (self: any, request: any): any => {
        let expected_value: any, ser_cxml: any, series: any;
        [ser_cxml, expected_value] = $list(request.param);
        series = $call(LineSeries, element(ser_cxml));
        return [series, expected_value];
      } },
    smooth_set_fixture: { args: ["request"], params: [["c:ser", true, "c:ser/c:smooth"], ["c:ser/c:smooth", false, "c:ser/c:smooth{val=0}"], ["c:ser/c:smooth{val=0}", true, "c:ser/c:smooth"]], fn: (self: any, request: any): any => {
        let expected_ser_cxml: any, expected_xml: any, new_value: any, ser_cxml: any, series: any;
        [ser_cxml, new_value, expected_ser_cxml] = $list(request.param);
        series = $call(LineSeries, element(ser_cxml));
        expected_xml = xml(expected_ser_cxml);
        return [series, new_value, expected_xml];
      } },
    subclass_fixture: { args: [], fn: (self: any): any => {
        return $call(LineSeries, null);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_is_a_BaseCategorySeries_subclass", args: ["subclass_fixture"], fn: (self: any, subclass_fixture: any): any => {
        let line_series: any;
        line_series = subclass_fixture;
        $assert($isinstance(line_series, _BaseCategorySeries));
      } },
    { name: "it_uses__MarkerMixin", args: ["subclass_fixture"], fn: (self: any, subclass_fixture: any): any => {
        let line_series: any;
        line_series = subclass_fixture;
        $assert($isinstance(line_series, _MarkerMixin));
      } },
    { name: "it_knows_whether_it_should_use_curve_smoothing", args: ["smooth_get_fixture"], fn: (self: any, smooth_get_fixture: any): any => {
        let expected_value: any, series: any;
        [series, expected_value] = $list(smooth_get_fixture);
        $assertEqual(series.smooth, expected_value);
      } },
    { name: "it_can_change_whether_it_uses_curve_smoothing", args: ["smooth_set_fixture"], fn: (self: any, smooth_set_fixture: any): any => {
        let expected_xml: any, new_value: any, series: any;
        [series, new_value, expected_xml] = $list(smooth_set_fixture);
        series.smooth = new_value;
        $assertEqual(series._element.xml, expected_xml);
      } },
  ],
};
pytestClass("DescribeLineSeries", DescribeLineSeries);

const DescribePieSeries: any = {
  fixtures: {
    subclass_fixture: { args: [], fn: (self: any): any => {
        return $call(PieSeries, null);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_is_a_BaseCategorySeries_subclass", args: ["subclass_fixture"], fn: (self: any, subclass_fixture: any): any => {
        let pie_series: any;
        pie_series = subclass_fixture;
        $assert($isinstance(pie_series, _BaseCategorySeries));
      } },
  ],
};
pytestClass("DescribePieSeries", DescribePieSeries);

const DescribeRadarSeries: any = {
  fixtures: {
    subclass_fixture: { args: [], fn: (self: any): any => {
        return $call(RadarSeries, null);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_is_a_BaseCategorySeries_subclass", args: ["subclass_fixture"], fn: (self: any, subclass_fixture: any): any => {
        let radar_series: any;
        radar_series = subclass_fixture;
        $assert($isinstance(radar_series, _BaseCategorySeries));
      } },
    { name: "it_uses__MarkerMixin", args: ["subclass_fixture"], fn: (self: any, subclass_fixture: any): any => {
        let line_series: any;
        line_series = subclass_fixture;
        $assert($isinstance(line_series, _MarkerMixin));
      } },
  ],
};
pytestClass("DescribeRadarSeries", DescribeRadarSeries);

const Describe_XySeries: any = {
  fixtures: {
    points_fixture: { args: ["XyPoints_", "points_"], fn: (self: any, XyPoints_: any, points_: any): any => {
        let ser: any, series: any;
        ser = element("c:ser");
        series = $call(XySeries, ser);
        return [series, XyPoints_, ser, points_];
      } },
    subclass_fixture: { args: [], fn: (self: any): any => {
        return $call(XySeries, null);
      } },
    values_get_fixture: { args: ["request"], params: [["c:ser", []], ["c:ser/c:yVal/c:numRef", []], ["c:ser/c:val/c:numRef/c:numCache", []], ["c:ser/c:yVal/c:numRef/c:numCache/(c:ptCount{val=1},c:pt{idx=0}/c:v\"1.1\")", [1.1]], ["c:ser/c:yVal/c:numRef/c:numCache/(c:ptCount{val=3},c:pt{idx=0}/c:v\"1.1\",c:pt{idx=2}/c:v\"3.3\")", [1.1, null, 3.3]], ["c:ser/c:val/c:numLit", []], ["c:ser/c:yVal/c:numLit/(c:ptCount{val=3},c:pt{idx=0}/c:v\"1.1\",c:pt{idx=2}/c:v\"3.3\")", [1.1, null, 3.3]]], fn: (self: any, request: any): any => {
        let expected_values: any, ser_cxml: any, series: any;
        [ser_cxml, expected_values] = $list(request.param);
        series = $call(XySeries, element(ser_cxml));
        return [series, expected_values];
      } },
    XyPoints_: { args: ["request", "points_"], fn: (self: any, request: any, points_: any): any => {
        return $call(classMock, request, "pptx.chart.series.XyPoints", $kw({return_value: points_}));
      } },
    points_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, XyPoints);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_uses__MarkerMixin", args: ["subclass_fixture"], fn: (self: any, subclass_fixture: any): any => {
        let line_series: any;
        line_series = subclass_fixture;
        $assert($isinstance(line_series, _MarkerMixin));
      } },
    { name: "it_provides_access_to_its_points", args: ["points_fixture"], fn: (self: any, points_fixture: any): any => {
        let XyPoints_: any, points: any, points_: any, ser: any, series: any;
        [series, XyPoints_, ser, points_] = $list(points_fixture);
        points = series.points;
        XyPoints_.assertCalledOnceWith(ser);
        $assert($is(points, points_));
      } },
    { name: "it_knows_its_values", args: ["values_get_fixture"], fn: (self: any, values_get_fixture: any): any => {
        let expected_values: any, series: any;
        [series, expected_values] = $list(values_get_fixture);
        $assertEqual(series.values, expected_values);
      } },
  ],
};
pytestClass("Describe_XySeries", Describe_XySeries);

const DescribeSeriesCollection: any = {
  fixtures: {
    getitem_fixture: { args: ["request", "_SeriesFactory_", "series_"], params: [["c:barChart/c:ser/c:order{val=42}", 0, 0], ["c:barChart/(c:ser/c:order{val=9},c:ser/c:order{val=6},c:ser/c:order{val=3})", 2, 0], ["c:plotArea/(c:barChart/(c:ser/(c:idx{val=1},c:order{val=3}),c:ser/(c:idx{val=0},c:order{val=0})),c:lineChart/c:ser/(c:idx{val=2},c:order{val=1}))", 1, 0]], fn: (self: any, request: any, _SeriesFactory_: any, series_: any): any => {
        let _offset: any, cxml: any, index: any, parent_elm: any, ser: any, series_collection: any;
        [cxml, index, _offset] = $list(request.param);
        parent_elm = element(cxml);
        ser = $get(parent_elm.xpath(".//c:ser"), _offset);
        series_collection = $call(SeriesCollection, parent_elm);
        return [series_collection, index, _SeriesFactory_, ser, series_];
      } },
    len_fixture: { args: ["request"], params: [["c:barChart", 0], ["c:barChart/c:ser/c:order{val=4}", 1], ["c:barChart/(c:ser/c:order{val=4},c:ser/c:order{val=1},c:ser/c:order{val=6})", 3], ["c:plotArea/c:barChart", 0], ["c:plotArea/c:barChart/c:ser/c:order{val=4}", 1], ["c:plotArea/c:barChart/(c:ser/c:order{val=4},c:ser/c:order{val=1},c:ser/c:order{val=6})", 3]], fn: (self: any, request: any): any => {
        let cxml: any, expected_len: any, series_collection: any;
        [cxml, expected_len] = $list(request.param);
        series_collection = $call(SeriesCollection, element(cxml));
        return [series_collection, expected_len];
      } },
    _SeriesFactory_: { args: ["request", "series_"], fn: (self: any, request: any, series_: any): any => {
        return $call(functionMock, request, "pptx.chart.series._SeriesFactory", $kw({return_value: series_}));
      } },
    series_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _BaseSeries);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_supports_indexed_access", args: ["getitem_fixture"], fn: (self: any, getitem_fixture: any): any => {
        let _SeriesFactory_: any, index: any, ser: any, series: any, series_: any, series_collection: any;
        [series_collection, index, _SeriesFactory_, ser, series_] = $list(getitem_fixture);
        series = $get(series_collection, index);
        _SeriesFactory_.assertCalledOnceWith(ser);
        $assert($is(series, series_));
      } },
    { name: "it_supports_len", args: ["len_fixture"], fn: (self: any, len_fixture: any): any => {
        let expected_len: any, series_collection: any;
        [series_collection, expected_len] = $list(len_fixture);
        $assertEqual($len(series_collection), expected_len);
      } },
  ],
};
pytestClass("DescribeSeriesCollection", DescribeSeriesCollection);

const Describe_SeriesFactory: any = {
  fixtures: {
    call_fixture: { args: ["request"], params: [["c:areaChart/c:ser", "AreaSeries"], ["c:barChart/c:ser", "BarSeries"], ["c:bubbleChart/c:ser", "BubbleSeries"], ["c:doughnutChart/c:ser", "PieSeries"], ["c:lineChart/c:ser", "LineSeries"], ["c:pieChart/c:ser", "PieSeries"], ["c:radarChart/c:ser", "RadarSeries"], ["c:scatterChart/c:ser", "XySeries"]], fn: (self: any, request: any): any => {
        let SeriesCls_: any, cls_name: any, ser: any, series_: any, xChart_cxml: any;
        [xChart_cxml, cls_name] = $list(request.param);
        ser = $get(element(xChart_cxml).serLst, 0);
        SeriesCls_ = classMock(request, $mod("pptx.chart.series.%s", cls_name));
        series_ = SeriesCls_.returnValue;
        return [ser, SeriesCls_, series_];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_contructs_a_series_object_from_a_plot_element", args: ["call_fixture"], fn: (self: any, call_fixture: any): any => {
        let SeriesCls_: any, ser: any, series: any, series_: any;
        [ser, SeriesCls_, series_] = $list(call_fixture);
        series = $call(_SeriesFactory, ser);
        SeriesCls_.assertCalledOnceWith(ser);
        $assert($is(series, series_));
      } },
  ],
};
pytestClass("Describe_SeriesFactory", Describe_SeriesFactory);

