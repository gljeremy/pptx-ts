// Machine-translated from python-pptx features/steps/series.py by
// scripts/py2ts_steps.py. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $assert, $assertEqual, $call, $d, $endswith, $enumerate, $eq, $get, $getattr, $int, $is, $kw, $len, $list, $m, $mod, $pyEval, $set, $truthy, $typename } from "../../pyshim.js";
import { AttributeError, IndexError, KeyError, NotImplementedError, ValueError } from "../../../src/errors.js";
import { step } from "../runner.js";
import { test_pptx } from "./helpers.js";
import { RGBColor } from "../../../src/dml/color.js";
import { XL_MARKER_STYLE } from "../../../src/enum/chart.js";
import { MSO_FILL_TYPE, MSO_THEME_COLOR } from "../../../src/enum/dml.js";
import { Presentation } from "../../../src/index.js";
void $py;


function given_a_BarSeries_object_having_fill_type_as_series(context: any, fill_type: any): any {
  let plot: any, prs: any, series_idx: any;
  series_idx = $get($d({Automatic: 0, "No Fill": 1, Orange: 2, "Accent 1": 3}), fill_type);
  prs = $call(Presentation, test_pptx("cht-series"));
  plot = $get($get($get(prs.slides, 2).shapes, 0).chart.plots, 0);
  context.series = $get(plot.series, series_idx);
}
step("given", "a BarSeries object having {fill_type} fill as series", ["fill_type"], given_a_BarSeries_object_having_fill_type_as_series);

function given_a_bar_series_having_invert_if_negative_setting(context: any, setting: any): any {
  let plot: any, prs: any, series_idx: any;
  series_idx = $get($d({"no explicit setting": 0, True: 1, False: 2}), setting);
  prs = $call(Presentation, test_pptx("cht-series"));
  plot = $get($get($get(prs.slides, 2).shapes, 0).chart.plots, 0);
  context.series = $get(plot.series, series_idx);
}
step("given", "a BarSeries object having invert_if_negative of {setting} as series", ["setting"], given_a_bar_series_having_invert_if_negative_setting);

function given_a_bar_series_having_values_as_series(context: any, values: any): any {
  let prs: any, series_idx: any;
  prs = $call(Presentation, test_pptx("cht-series"));
  series_idx = $get($d({"1.2, 2.3, 3.4": 0, "4.5, None, 6.7": 1}), values);
  context.series = $get($get($get($get(prs.slides, 3).shapes, 0).chart.plots, 0).series, series_idx);
}
step("given", "a BarSeries object having values {values} as series", ["values"], given_a_bar_series_having_values_as_series);

function given_a_bar_series_having_width_line_as_series(context: any, width: any): any {
  let plot: any, prs: any, series_idx: any;
  series_idx = $get($d({no: 0, "1 point": 1}), width);
  prs = $call(Presentation, test_pptx("cht-series"));
  plot = $get($get($get(prs.slides, 2).shapes, 0).chart.plots, 0);
  context.series = $get(plot.series, series_idx);
}
step("given", "a BarSeries object having {width} line as series", ["width"], given_a_bar_series_having_width_line_as_series);

function given_a_marker(context: any): any {
  let prs: any, series: any;
  prs = $call(Presentation, test_pptx("cht-marker-props"));
  series = $get($get($get(prs.slides, 0).shapes, 0).chart.series, 0);
  context.marker = series.marker;
}
step("given", "a marker", [], given_a_marker);

function given_a_marker_having_size_of_case(context: any, case$: any): any {
  let prs: any, series: any, series_idx: any;
  series_idx = $get($d({"no explicit value": 0, "24 points": 1, "36 points": 2}), case$);
  prs = $call(Presentation, test_pptx("cht-marker-props"));
  series = $get($get($get(prs.slides, 0).shapes, 0).chart.series, series_idx);
  context.marker = series.marker;
}
step("given", "a marker having size of {case}", ["case"], given_a_marker_having_size_of_case);

function given_a_marker_having_style_of_case(context: any, case$: any): any {
  let prs: any, series: any, series_idx: any;
  series_idx = $get($d({"no explicit value": 0, circle: 1, triangle: 2}), case$);
  prs = $call(Presentation, test_pptx("cht-marker-props"));
  series = $get($get($get(prs.slides, 0).shapes, 0).chart.series, series_idx);
  context.marker = series.marker;
}
step("given", "a marker having style of {case}", ["case"], given_a_marker_having_style_of_case);

function given_a_point(context: any): any {
  let chart: any, prs: any;
  prs = $call(Presentation, test_pptx("cht-point-props"));
  chart = $get($get(prs.slides, 0).shapes, 0).chart;
  context.point = $get($get($get(chart.plots, 0).series, 0).points, 0);
}
step("given", "a point", [], given_a_point);

function given_a_points_type_object_containing_3_points(context: any, points_type: any): any {
  let prs: any, series: any, slide_idx: any;
  slide_idx = $get($d({XyPoints: 0, BubblePoints: 1, CategoryPoints: 2}), points_type);
  prs = $call(Presentation, test_pptx("cht-point-access"));
  series = $get($get($get($get(prs.slides, slide_idx).shapes, 0).chart.plots, 0).series, 0);
  context.points = series.points;
}
step("given", "a {points_type} object containing 3 points", ["points_type"], given_a_points_type_object_containing_3_points);

function given_a_series(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("cht-series"));
  context.series = $get($get($get($get(prs.slides, 3).shapes, 0).chart.plots, 0).series, 0);
}
step("given", "a series", [], given_a_series);

function given_a_series_of_type_series_type(context: any, prefix: any): any {
  let prs: any, slide_idx: any;
  slide_idx = $get($d({Area: 8, Bar: 3, Bubble: 5, Category: 3, Doughnut: 9, Line: 6, Pie: 10, Radar: 7, Xy: 4}), prefix);
  prs = $call(Presentation, test_pptx("cht-series"));
  context.series = $get($get($get($get(prs.slides, slide_idx).shapes, 0).chart.plots, 0).series, 0);
}
step("given", "a {prefix}Series object as series", ["prefix"], given_a_series_of_type_series_type);

function given_a_SeriesCollection_object_for_a_plot_having_n_series(context: any, n: any): any {
  let plot: any, prs: any;
  prs = $call(Presentation, test_pptx("cht-series"));
  plot = $get($get($get(prs.slides, 0).shapes, 0).chart.plots, 0);
  context.seriesCollection = plot.series;
  context.seriesCount = $int(n);
}
step("given", "a SeriesCollection object for a plot having {n} series", ["n"], given_a_SeriesCollection_object_for_a_plot_having_n_series);

function given_a_SeriesCollection_for_chart_having_n_series(context: any, type_: any, n: any): any {
  let prs: any, slide_idx: any;
  slide_idx = $get($d({"single-plot": 0, "multi-plot": 1}), type_);
  prs = $call(Presentation, test_pptx("cht-series"));
  context.seriesCollection = $get($get(prs.slides, slide_idx).shapes, 0).chart.series;
  context.seriesCount = $int(n);
}
step("given", "a SeriesCollection object for a {type_} chart having {n} series", ["type_", "n"], given_a_SeriesCollection_for_chart_having_n_series);

function when_I_add_a_series_with_number_format(context: any, strval: any): any {
  let chart_data: any, params: any;
  chart_data = context.chartData;
  params = $d({name: "Series Foo"});
  if ($truthy(!$eq(strval, "None"))) {
    $set(params, "number_format", $int(strval));
  }
  context.seriesData = $m(chart_data, "addSeries", $kw({...params}));
}
step("when", "I add a series with number format {strval}", ["strval"], when_I_add_a_series_with_number_format);

function when_I_assign_value_to_marker_size(context: any, value: any): any {
  let new_value: any;
  new_value = ($truthy($eq(value, "None")) ? null : $int(value));
  context.marker.size = new_value;
}
step("when", "I assign {value} to marker.size", ["value"], when_I_assign_value_to_marker_size);

function when_I_assign_value_to_marker_style(context: any, value: any): any {
  let new_value: any;
  new_value = ($truthy($eq(value, "None")) ? null : $getattr(XL_MARKER_STYLE, value));
  context.marker.style = new_value;
}
step("when", "I assign {value} to marker.style", ["value"], when_I_assign_value_to_marker_style);

function when_I_assign_value_to_series_invert_if_negative(context: any, value: any): any {
  let new_value: any;
  new_value = $get($d({True: true, False: false}), value);
  context.series.invertIfNegative = new_value;
}
step("when", "I assign {value} to series.invert_if_negative", ["value"], when_I_assign_value_to_series_invert_if_negative);

function then_data_point_number_format_is(context: any, value_str: any): any {
  let data_point: any, number_format: any;
  data_point = context.dataPoint;
  number_format = ($truthy($eq(value_str, "General")) ? value_str : $int(value_str));
  $assertEqual(data_point.numberFormat, number_format);
}
step("then", "data_point.number_format is {value_str}", ["value_str"], then_data_point_number_format_is);

function then_iterating_points_produces_3_point_objects(context: any): any {
  let idx: any, points: any;
  points = context.points;
  idx = -1;
  let point: any;
  for ([idx, point] of $list($enumerate(points))) {
    $assertEqual($typename(point), "Point");
  }
  $assertEqual(idx, 2, $mod("got %s", idx));
}
step("then", "iterating points produces 3 Point objects", [], then_iterating_points_produces_3_point_objects);

function then_iterating_series_collection_produces_count_series(context: any, count: any): any {
  let expected_idx: any, idx: any, type_name: any;
  expected_idx = ($int(count) - 1);
  idx = -1;
  let series: any;
  for ([idx, series] of $list($enumerate(context.seriesCollection))) {
    type_name = $typename(series);
    $assert($endswith(type_name, "Series"), $mod("got %s", type_name));
  }
  $assertEqual(idx, expected_idx, $mod("got %s", idx));
}
step("then", "iterating series_collection produces {count} Series objects", ["count"], then_iterating_series_collection_produces_count_series);

function then_len_points_is_3(context: any): any {
  let points: any;
  points = context.points;
  $assertEqual($len(points), 3);
}
step("then", "len(points) is 3", [], then_len_points_is_3);

function then_len_series_collection_is_count(context: any, count: any): any {
  let actual_len: any, expected_len: any;
  expected_len = $int(count);
  actual_len = $len(context.seriesCollection);
  $assertEqual(actual_len, expected_len, $mod("got %s", actual_len));
}
step("then", "len(series_collection) is {count}", ["count"], then_len_series_collection_is_count);

function then_len_series_values_is_count_for_each_series(context: any, count: any): any {
  let expected_count: any;
  expected_count = $int(count);
  for (const series of $list($get(context.chart.plots, 0).series)) {
    $assertEqual($len(series.values), expected_count);
  }
}
step("then", "len(series.values) is {count} for each series", ["count"], then_len_series_values_is_count_for_each_series);

function then_marker_format_is_a_ChartFormat_object(context: any): any {
  let marker: any;
  marker = context.marker;
  $assertEqual($typename(marker.format), "ChartFormat");
}
step("then", "marker.format is a ChartFormat object", [], then_marker_format_is_a_ChartFormat_object);

function then_marker_format_fill_is_a_FillFormat_object(context: any): any {
  let marker: any;
  marker = context.marker;
  $assertEqual($typename(marker.format.fill), "FillFormat");
}
step("then", "marker.format.fill is a FillFormat object", [], then_marker_format_fill_is_a_FillFormat_object);

function then_marker_format_line_is_a_LineFormat_object(context: any): any {
  let marker: any;
  marker = context.marker;
  $assertEqual($typename(marker.format.line), "LineFormat");
}
step("then", "marker.format.line is a LineFormat object", [], then_marker_format_line_is_a_LineFormat_object);

function then_marker_size_is_case(context: any, case$: any): any {
  let expected_value: any, marker: any;
  expected_value = ($truthy($eq(case$, "None")) ? null : $int(case$));
  marker = context.marker;
  $assertEqual(marker.size, expected_value, $mod("got %s", marker.size));
}
step("then", "marker.size is {case}", ["case"], then_marker_size_is_case);

function then_marker_style_is_case(context: any, case$: any): any {
  let expected_value: any, marker: any;
  expected_value = ($truthy($eq(case$, "None")) ? null : $getattr(XL_MARKER_STYLE, case$));
  marker = context.marker;
  $assertEqual(marker.style, expected_value, $mod("got %s", marker.style));
}
step("then", "marker.style is {case}", ["case"], then_marker_style_is_case);

function then_point_data_label_is_a_DataLabel_object(context: any): any {
  let point: any;
  point = context.point;
  $assertEqual($typename(point.dataLabel), "DataLabel");
}
step("then", "point.data_label is a DataLabel object", [], then_point_data_label_is_a_DataLabel_object);

function then_point_format_is_a_ChartFormat_object(context: any): any {
  let point: any;
  point = context.point;
  $assertEqual($typename(point.format), "ChartFormat");
}
step("then", "point.format is a ChartFormat object", [], then_point_format_is_a_ChartFormat_object);

function then_point_format_fill_is_a_FillFormat_object(context: any): any {
  let point: any;
  point = context.point;
  $assertEqual($typename(point.format.fill), "FillFormat");
}
step("then", "point.format.fill is a FillFormat object", [], then_point_format_fill_is_a_FillFormat_object);

function then_point_format_line_is_a_LineFormat_object(context: any): any {
  let point: any;
  point = context.point;
  $assertEqual($typename(point.format.line), "LineFormat");
}
step("then", "point.format.line is a LineFormat object", [], then_point_format_line_is_a_LineFormat_object);

function then_point_marker_is_a_Marker_object(context: any): any {
  let point: any;
  point = context.point;
  $assertEqual($typename(point.marker), "Marker");
}
step("then", "point.marker is a Marker object", [], then_point_marker_is_a_Marker_object);

function then_points_2_is_a_Point_object(context: any): any {
  let actual: any;
  actual = $typename($get(context.points, 2));
  $assertEqual(actual, "Point", $mod("points[2] is a %s object", actual));
}
step("then", "points[2] is a Point object", [], then_points_2_is_a_Point_object);

function then_series_collection_2_is_a_Series_object(context: any): any {
  let type_name: any;
  type_name = $typename($get(context.seriesCollection, 2));
  $assert($endswith(type_name, "Series"), $mod("got %s", type_name));
}
step("then", "series_collection[2] is a Series object", [], then_series_collection_2_is_a_Series_object);

function then_series_data_labels_is_a_DataLabels_object(context: any): any {
  let actual: any;
  actual = $typename(context.series.dataLabels);
  $assertEqual(actual, "DataLabels", $mod("series.data_labels is a %s object", actual));
}
step("then", "series.data_labels is a DataLabels object", [], then_series_data_labels_is_a_DataLabels_object);

function then_series_format_fill_fore_color_rgb_is_FF6600(context: any): any {
  let rgb_color: any;
  rgb_color = context.series.format.fill.foreColor.rgb;
  $assertEqual(rgb_color, $call(RGBColor, 255, 102, 0), $mod("got %s", rgb_color));
}
step("then", "series.format.fill.fore_color.rgb is FF6600", [], then_series_format_fill_fore_color_rgb_is_FF6600);

function then_series_format_fill_fore_color_theme_color_is_Accent_1(context: any): any {
  let theme_color: any;
  theme_color = context.series.format.fill.foreColor.themeColor;
  $assertEqual(theme_color, MSO_THEME_COLOR.ACCENT_1, $mod("got %s", theme_color));
}
step("then", "series.format.fill.fore_color.theme_color is Accent 1", [], then_series_format_fill_fore_color_theme_color_is_Accent_1);

function then_series_format_fill_type_is_type(context: any, fill_type: any): any {
  let expected_fill_type: any;
  expected_fill_type = $get($d({None: null, "MSO_FILL_TYPE.BACKGROUND": MSO_FILL_TYPE.BACKGROUND, "MSO_FILL_TYPE.SOLID": MSO_FILL_TYPE.SOLID}), fill_type);
  fill_type = context.series.format.fill.type;
  $assertEqual(fill_type, expected_fill_type, $mod("got %s", fill_type));
}
step("then", "series.format.fill.type is {fill_type}", ["fill_type"], then_series_format_fill_type_is_type);

function then_series_invert_if_negative_is_value(context: any, value: any): any {
  let expected_value: any, series: any;
  expected_value = $get($d({True: true, False: false}), value);
  series = context.series;
  $assert($is(series.invertIfNegative, expected_value));
}
step("then", "series.invert_if_negative is {value}", ["value"], then_series_invert_if_negative_is_value);

function then_series_format_line_width_is_width(context: any, width: any): any {
  let expected_width: any, line_width: any;
  expected_width = $int(width);
  line_width = context.series.format.line.width;
  $assertEqual(line_width, expected_width, $mod("got %s", line_width));
}
step("then", "series.format.line.width is {width}", ["width"], then_series_format_line_width_is_width);

function then_series_format_is_a_ChartFormat_object(context: any): any {
  let actual: any;
  actual = $typename(context.series.format);
  $assertEqual(actual, "ChartFormat", $mod("series.format is a %s object", actual));
}
step("then", "series.format is a ChartFormat object", [], then_series_format_is_a_ChartFormat_object);

function then_series_marker_is_a_Marker_object(context: any): any {
  let actual: any;
  actual = $typename(context.series.marker);
  $assertEqual(actual, "Marker", $mod("series.marker is a %s object", actual));
}
step("then", "series.marker is a Marker object", [], then_series_marker_is_a_Marker_object);

function then_series_points_is_a_type_name_object(context: any, type_name: any): any {
  let actual: any, expected: any;
  actual = $typename(context.series.points);
  expected = type_name;
  $assertEqual(actual, expected, $mod("series.points is a %s object", actual));
}
step("then", "series.points is a {type_name} object", ["type_name"], then_series_points_is_a_type_name_object);

function then_series_values_is_values(context: any, values: any): any {
  let expected_values: any, series: any;
  series = context.series;
  expected_values = $pyEval(values, {});
  $assertEqual(series.values, expected_values, $mod("got %s", [series.values]));
}
step("then", "series.values is {values}", ["values"], then_series_values_is_values);

// the names python's eval() sees in this module
const $scope = { MSO_FILL_TYPE, MSO_THEME_COLOR, Presentation, RGBColor, XL_MARKER_STYLE };
