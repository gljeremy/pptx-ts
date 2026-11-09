// Machine-translated from python-pptx features/steps/chart.py by
// scripts/py2ts_steps.py. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $add, $assert, $assertEqual, $call, $d, $eq, $get, $getattr, $hashlib, $int, $is, $isinstance, $len, $list, $mod, $mul, $range, $slice, $startswith, $t1, $typename } from "../../pyshim.js";
import { AttributeError, IndexError, KeyError, NotImplementedError, ValueError } from "../../../src/errors.js";
import { step } from "../runner.js";
import { count, test_pptx } from "./helpers.js";
import { Legend } from "../../../src/chart/legend.js";
import { BubbleChartData, CategoryChartData, ChartData, XyChartData } from "../../../src/chart/data.js";
import { XL_CHART_TYPE } from "../../../src/enum/chart.js";
import { Presentation } from "../../../src/index.js";
import { EmbeddedXlsxPart } from "../../../src/parts/embeddedpackage.js";
import { Inches } from "../../../src/util.js";
const islice = $py.$itertools.islice;
void $py;


function given_a_Chart_object_as_chart(context: any): any {
  let slide: any;
  slide = $get($call(Presentation, test_pptx("shp-common-props")).slides, 0);
  context.chart = $get(slide.shapes, 6).chart;
}
step("given", "a Chart object as chart", [], given_a_Chart_object_as_chart);

function given_a_chart_having_a_or_no_title(context: any, a_or_no: any): any {
  let prs: any, shape_idx: any;
  shape_idx = $get($d({no: 0, a: 1}), a_or_no);
  prs = $call(Presentation, test_pptx("cht-chart-props"));
  context.chart = $get($get(prs.slides, 0).shapes, shape_idx).chart;
}
step("given", "a chart having {a_or_no} title", ["a_or_no"], given_a_chart_having_a_or_no_title);

function given_a_chart_having_or_not_a_legend(context: any, having_or_not: any): any {
  let prs: any, slide_idx: any;
  slide_idx = $get($d({having: 0, "not having": 1}), having_or_not);
  prs = $call(Presentation, test_pptx("cht-legend"));
  context.chart = $get($get(prs.slides, slide_idx).shapes, 0).chart;
}
step("given", "a chart {having_or_not} a legend", ["having_or_not"], given_a_chart_having_or_not_a_legend);

function given_a_chart_of_size_and_type_spec(context: any, spec: any): any {
  let chart: any, prs: any, slide_idx: any;
  slide_idx = $get($d({"2x2 Clustered Bar": 0, "2x2 100% Stacked Bar": 1, "2x2 Clustered Column": 2, "4x3 Line": 3, "3x1 Pie": 4, "3x2 XY": 5, "3x2 Bubble": 6}), spec);
  prs = $call(Presentation, test_pptx("cht-replace-data"));
  chart = $get($get(prs.slides, slide_idx).shapes, 0).chart;
  context.chart = chart;
  context.xlsxSha1 = $hashlib.sha1(chart._workbook.xlsxPart.blob).hexdigest();
}
step("given", "a chart of size and type {spec}", ["spec"], given_a_chart_of_size_and_type_spec);

function given_a_chart_of_type_chart_type(context: any, chart_type: any): any {
  let prs: any, shape_idx: any, slide_idx: any;
  [slide_idx, shape_idx] = $list($get($d({Area: [0, 0], "Stacked Area": [0, 1], "100% Stacked Area": [0, 2], "3-D Area": [0, 3], "3-D Stacked Area": [0, 4], "3-D 100% Stacked Area": [0, 5], "Clustered Bar": [1, 0], "Stacked Bar": [1, 1], "100% Stacked Bar": [1, 2], "Clustered Column": [1, 3], "Stacked Column": [1, 4], "100% Stacked Column": [1, 5], Line: [2, 0], "Stacked Line": [2, 1], "100% Stacked Line": [2, 2], "Marked Line": [2, 3], "Stacked Marked Line": [2, 4], "100% Stacked Marked Line": [2, 5], Pie: [3, 0], "Exploded Pie": [3, 1], "XY (Scatter)": [4, 0], "XY Lines": [4, 1], "XY Lines No Markers": [4, 2], "XY Smooth Lines": [4, 3], "XY Smooth No Markers": [4, 4], Bubble: [5, 0], "3D-Bubble": [5, 1], Radar: [6, 0], "Marked Radar": [6, 1], "Filled Radar": [6, 2], "Line (with date categories)": [7, 0]}), chart_type));
  prs = $call(Presentation, test_pptx("cht-chart-type"));
  context.chart = $get($get(prs.slides, slide_idx).shapes, shape_idx).chart;
}
step("given", "a chart of type {chart_type}", ["chart_type"], given_a_chart_of_type_chart_type);

function given_a_chart_title(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("cht-chart-props"));
  context.chartTitle = $get($get(prs.slides, 0).shapes, 1).chart.chartTitle;
}
step("given", "a chart title", [], given_a_chart_title);

function given_a_chart_title_having_a_or_no_text_frame(context: any, a_or_no: any): any {
  let prs: any, shape_idx: any;
  prs = $call(Presentation, test_pptx("cht-chart-props"));
  shape_idx = $get($d({no: 0, a: 1}), a_or_no);
  context.chartTitle = $get($get(prs.slides, 1).shapes, shape_idx).chart.chartTitle;
}
step("given", "a chart title having {a_or_no} text frame", ["a_or_no"], given_a_chart_title_having_a_or_no_text_frame);

function when_I_add_a_clustered_bar_chart_with_multi_level_categories(context: any): any {
  let EAST: any, WEST: any, chart_data: any, chart_type: any;
  chart_type = XL_CHART_TYPE.BAR_CLUSTERED;
  chart_data = $call(CategoryChartData);
  WEST = chart_data.addCategory("WEST");
  WEST.addSubCategory("SF");
  WEST.addSubCategory("LA");
  EAST = chart_data.addCategory("EAST");
  EAST.addSubCategory("NY");
  EAST.addSubCategory("NJ");
  chart_data.addSeries("Series 1", [1, 2, null, 4]);
  chart_data.addSeries("Series 2", [5, null, 7, 8]);
  context.chart = context.slide.shapes.addChart(chart_type, $call(Inches, 1), $call(Inches, 1), $call(Inches, 8), $call(Inches, 5), chart_data).chart;
}
step("when", "I add a Clustered bar chart with multi-level categories", [], when_I_add_a_clustered_bar_chart_with_multi_level_categories);

function when_I_add_a_chart_with_categories_and_series(context: any, kind: any, cats: any, sers: any): any {
  let category_count: any, category_source: any, chart_data: any, chart_type: any, series_count: any, series_title: any, series_value_source: any, series_values: any;
  chart_type = $get($d({Area: XL_CHART_TYPE.AREA, "Stacked Area": XL_CHART_TYPE.AREA_STACKED, "100% Stacked Area": XL_CHART_TYPE.AREA_STACKED_100, "Clustered Bar": XL_CHART_TYPE.BAR_CLUSTERED, "Stacked Bar": XL_CHART_TYPE.BAR_STACKED, "100% Stacked Bar": XL_CHART_TYPE.BAR_STACKED_100, "Clustered Column": XL_CHART_TYPE.COLUMN_CLUSTERED, "Stacked Column": XL_CHART_TYPE.COLUMN_STACKED, "100% Stacked Column": XL_CHART_TYPE.COLUMN_STACKED_100, Doughnut: XL_CHART_TYPE.DOUGHNUT, "Exploded Doughnut": XL_CHART_TYPE.DOUGHNUT_EXPLODED, Line: XL_CHART_TYPE.LINE, "Line with Markers": XL_CHART_TYPE.LINE_MARKERS, "Line Markers Stacked": XL_CHART_TYPE.LINE_MARKERS_STACKED, "100% Line Markers Stacked": XL_CHART_TYPE.LINE_MARKERS_STACKED_100, "Line Stacked": XL_CHART_TYPE.LINE_STACKED, "100% Line Stacked": XL_CHART_TYPE.LINE_STACKED_100, Pie: XL_CHART_TYPE.PIE, "Exploded Pie": XL_CHART_TYPE.PIE_EXPLODED, Radar: XL_CHART_TYPE.RADAR, "Filled Radar": XL_CHART_TYPE.RADAR_FILLED, "Radar with markers": XL_CHART_TYPE.RADAR_MARKERS}), kind);
  [category_count, series_count] = $list([$int(cats), $int(sers)]);
  category_source = ["Foo", "Bar", "Baz", "Boo", "Far", "Faz"];
  series_value_source = count(1.1, 1.1);
  chart_data = $call(CategoryChartData);
  chart_data.categories = $slice(category_source, null, category_count, null);
  for (const idx of $list($range(series_count))) {
    series_title = $mod("Series %d", (idx + 1));
    series_values = $list(islice(series_value_source, category_count));
    chart_data.addSeries(series_title, series_values);
  }
  context.chart = context.slide.shapes.addChart(chart_type, $call(Inches, 1), $call(Inches, 1), $call(Inches, 8), $call(Inches, 5), chart_data).chart;
}
step("when", "I add a {kind} chart with {cats} categories and {sers} series", ["kind", "cats", "sers"], when_I_add_a_chart_with_categories_and_series);

function when_I_add_a_bubble_chart_having_2_series_of_3_pts(context: any, bubble_type: any): any {
  let chart_data: any, chart_type: any, data: any, points: any, series: any, series_label: any, size: any, x: any, y: any;
  chart_type = $getattr(XL_CHART_TYPE, bubble_type);
  data = [["Series 1", [[-0.1, 0.5, 1.0], [16.2, 0.0, 2.0], [8.0, -0.2, 3.0]]], ["Series 2", [[12.4, 0.8, 4.0], [-7.5, 0.5, 5.0], [5.1, -0.5, 6.0]]]];
  chart_data = $call(BubbleChartData);
  for (const series_data of $list(data)) {
    [series_label, points] = $list(series_data);
    series = chart_data.addSeries(series_label);
    for (const point of $list(points)) {
      [x, y, size] = $list(point);
      series.addDataPoint(x, y, size);
    }
  }
  context.chart = context.slide.shapes.addChart(chart_type, $call(Inches, 1), $call(Inches, 1), $call(Inches, 8), $call(Inches, 5), chart_data).chart;
}
step("when", "I add a {bubble_type} chart having 2 series of 3 points each", ["bubble_type"], when_I_add_a_bubble_chart_having_2_series_of_3_pts);

function when_I_assign_value_to_chart_has_legend(context: any, value: any): any {
  let new_value: any;
  new_value = $get($d({True: true, False: false}), value);
  context.chart.hasLegend = new_value;
}
step("when", "I assign {value} to chart.has_legend", ["value"], when_I_assign_value_to_chart_has_legend);

function when_I_assign_value_to_chart_has_title(context: any, value: any): any {
  context.chart.hasTitle = $get($d({True: true, False: false}), value);
}
step("when", "I assign {value} to chart.has_title", ["value"], when_I_assign_value_to_chart_has_title);

function when_I_assign_value_to_chart_title_has_text_frame(context: any, value: any): any {
  context.chartTitle.hasTextFrame = $get($d({True: true, False: false}), value);
}
step("when", "I assign {value} to chart_title.has_text_frame", ["value"], when_I_assign_value_to_chart_title_has_text_frame);

function when_I_replace_its_data_with_categories_and_series(context: any, cats: any, sers: any): any {
  let category_count: any, category_source: any, chart_data: any, series_count: any, series_title: any, series_value_source: any, series_values: any;
  [category_count, series_count] = $list([$int(cats), $int(sers)]);
  category_source = ["Foo", "Bar", "Baz", "Boo", "Far", "Faz"];
  series_value_source = count(1.1, 1.1);
  chart_data = $call(ChartData);
  chart_data.categories = $slice(category_source, null, category_count, null);
  for (const idx of $list($range(series_count))) {
    series_title = $mod("New Series %d", (idx + 1));
    series_values = $list(islice(series_value_source, category_count));
    chart_data.addSeries(series_title, series_values);
  }
  context.chart.replaceData(chart_data);
}
step("when", "I replace its data with {cats} categories and {sers} series", ["cats", "sers"], when_I_replace_its_data_with_categories_and_series);

function when_I_replace_its_data_with_3_series_of_three_bubble_pts_each(context: any): any {
  let chart_data: any, series: any, series_title: any, size: any, x: any, y: any;
  chart_data = $call(BubbleChartData);
  for (const idx of $list($range(3))) {
    series_title = $mod("New Series %d", (idx + 1));
    series = chart_data.addSeries(series_title);
    for (const jdx of $list($range(3))) {
      [x, y, size] = $list([$add($mul(idx, 3), jdx), $add($mul(idx, 2), jdx), $add(idx, jdx)]);
      series.addDataPoint(x, y, size);
    }
  }
  context.chart.replaceData(chart_data);
}
step("when", "I replace its data with 3 series of 3 bubble points each", [], when_I_replace_its_data_with_3_series_of_three_bubble_pts_each);

function when_I_replace_its_data_with_3_series_of_three_points_each(context: any): any {
  let chart_data: any, series: any, series_title: any, x: any, y: any;
  chart_data = $call(XyChartData);
  const $t1 = 0;
  x = $t1;
  y = $t1;
  for (const idx of $list($range(3))) {
    series_title = $mod("New Series %d", (idx + 1));
    series = chart_data.addSeries(series_title);
    for (const jdx of $list($range(3))) {
      [x, y] = $list([$add($mul(idx, 3), jdx), $add($mul(idx, 2), jdx)]);
      series.addDataPoint(x, y);
    }
  }
  context.chart.replaceData(chart_data);
}
step("when", "I replace its data with 3 series of 3 points each", [], when_I_replace_its_data_with_3_series_of_three_points_each);

function then_chart_category_axis_is_a_cls_name_object(context: any, cls_name: any): any {
  let category_axis: any, type_name: any;
  category_axis = context.chart.categoryAxis;
  type_name = $typename(category_axis);
  $assertEqual(type_name, cls_name, $mod("got %s", type_name));
}
step("then", "chart.category_axis is a {cls_name} object", ["cls_name"], then_chart_category_axis_is_a_cls_name_object);

function then_chart_chart_title_is_a_ChartTitle_object(context: any): any {
  let class_name: any;
  class_name = $typename(context.chart.chartTitle);
  $assertEqual(class_name, "ChartTitle", $mod("got %s", class_name));
}
step("then", "chart.chart_title is a ChartTitle object", [], then_chart_chart_title_is_a_ChartTitle_object);

function then_chart_chart_type_is_value(context: any, enum_member: any): any {
  let chart: any, expected_value: any;
  expected_value = $getattr(XL_CHART_TYPE, enum_member);
  chart = context.chart;
  $assert($is(chart.chartType, expected_value), $mod("got %s", chart.chartType));
}
step("then", "chart.chart_type is {enum_member}", ["enum_member"], then_chart_chart_type_is_value);

function then_chart_font_is_a_Font_object(context: any): any {
  let actual: any, expected: any;
  actual = $typename(context.chart.font);
  expected = "Font";
  $assertEqual(actual, expected, $mod("chart.font is a %s object", actual));
}
step("then", "chart.font is a Font object", [], then_chart_font_is_a_Font_object);

function then_chart_has_legend_is_value(context: any, value: any): any {
  let chart: any, expected_value: any;
  expected_value = $get($d({True: true, False: false}), value);
  chart = context.chart;
  $assert($is(chart.hasLegend, expected_value));
}
step("then", "chart.has_legend is {value}", ["value"], then_chart_has_legend_is_value);

function then_chart_has_title_is_value(context: any, value: any): any {
  let actual_value: any, chart: any, expected_value: any;
  chart = context.chart;
  actual_value = chart.hasTitle;
  expected_value = $get($d({True: true, False: false}), value);
  $assert($is(actual_value, expected_value), $mod("got %s", actual_value));
}
step("then", "chart.has_title is {value}", ["value"], then_chart_has_title_is_value);

function then_chart_legend_is_a_legend_object(context: any): any {
  let chart: any;
  chart = context.chart;
  $assert($isinstance(chart.legend, Legend));
}
step("then", "chart.legend is a legend object", [], then_chart_legend_is_a_legend_object);

function then_chart_series_is_a_SeriesCollection_object(context: any): any {
  let type_name: any;
  type_name = $typename(context.chart.series);
  $assertEqual(type_name, "SeriesCollection", $mod("got %s", type_name));
}
step("then", "chart.series is a SeriesCollection object", [], then_chart_series_is_a_SeriesCollection_object);

function then_chart_value_axis_is_a_ValueAxis_object(context: any): any {
  let value_axis: any;
  value_axis = context.chart.valueAxis;
  $assertEqual($typename(value_axis), "ValueAxis");
}
step("then", "chart.value_axis is a ValueAxis object", [], then_chart_value_axis_is_a_ValueAxis_object);

function then_chart_title_format_is_a_ChartFormat_object(context: any): any {
  let class_name: any;
  class_name = $typename(context.chartTitle.format);
  $assertEqual(class_name, "ChartFormat", $mod("got %s", class_name));
}
step("then", "chart_title.format is a ChartFormat object", [], then_chart_title_format_is_a_ChartFormat_object);

function then_chart_title_format_fill_is_a_FillFormat_object(context: any): any {
  let class_name: any;
  class_name = $typename(context.chartTitle.format.fill);
  $assertEqual(class_name, "FillFormat", $mod("got %s", class_name));
}
step("then", "chart_title.format.fill is a FillFormat object", [], then_chart_title_format_fill_is_a_FillFormat_object);

function then_chart_title_format_line_is_a_LineFormat_object(context: any): any {
  let class_name: any;
  class_name = $typename(context.chartTitle.format.line);
  $assertEqual(class_name, "LineFormat", $mod("got %s", class_name));
}
step("then", "chart_title.format.line is a LineFormat object", [], then_chart_title_format_line_is_a_LineFormat_object);

function then_chart_title_has_text_frame_is_value(context: any, value: any): any {
  let actual_value: any, expected_value: any;
  actual_value = context.chartTitle.hasTextFrame;
  expected_value = $get($d({True: true, False: false}), value);
  $assert($is(actual_value, expected_value), $mod("got %s", actual_value));
}
step("then", "chart_title.has_text_frame is {value}", ["value"], then_chart_title_has_text_frame_is_value);

function then_chart_title_text_frame_is_a_TextFrame_object(context: any): any {
  let class_name: any;
  class_name = $typename(context.chartTitle.textFrame);
  $assertEqual(class_name, "TextFrame", $mod("got %s", class_name));
}
step("then", "chart_title.text_frame is a TextFrame object", [], then_chart_title_text_frame_is_a_TextFrame_object);

function then_each_series_has_a_new_name(context: any): any {
  for (const series of $list($get(context.chart.plots, 0).series)) {
    $assert($startswith(series.name, "New "));
  }
}
step("then", "each series has a new name", [], then_each_series_has_a_new_name);

function then_each_series_has_count_values(context: any, count: any): any {
  let actual_value_count: any, expected_count: any;
  expected_count = $int(count);
  for (const series of $list($get(context.chart.plots, 0).series)) {
    actual_value_count = $len(series.values);
    $assertEqual(actual_value_count, expected_count);
  }
}
step("then", "each series has {count} values", ["count"], then_each_series_has_count_values);

function then_len_chart_series_is_count(context: any, count: any): any {
  let expected_count: any;
  expected_count = $int(count);
  $assertEqual($len(context.chart.series), expected_count);
}
step("then", "len(chart.series) is {count}", ["count"], then_len_chart_series_is_count);

function then_the_chart_has_an_Excel_data_worksheet(context: any): any {
  let xlsx_part: any;
  xlsx_part = context.chart._workbook.xlsxPart;
  $assert($isinstance(xlsx_part, EmbeddedXlsxPart));
}
step("then", "the chart has an Excel data worksheet", [], then_the_chart_has_an_Excel_data_worksheet);

function then_the_chart_has_new_chart_data(context: any): any {
  let new_xlsx_sha1: any, orig_xlsx_sha1: any;
  orig_xlsx_sha1 = context.xlsxSha1;
  new_xlsx_sha1 = $hashlib.sha1(context.chart._workbook.xlsxPart.blob).hexdigest();
  $assert(!$eq(new_xlsx_sha1, orig_xlsx_sha1));
}
step("then", "the chart has new chart data", [], then_the_chart_has_new_chart_data);

// the names python's eval() sees in this module
const $scope = { BubbleChartData, CategoryChartData, ChartData, EmbeddedXlsxPart, Inches, Legend, Presentation, XL_CHART_TYPE, XyChartData };
