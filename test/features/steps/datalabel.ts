// Machine-translated from python-pptx features/steps/datalabel.py by
// scripts/py2ts_steps.py. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $assert, $assertEqual, $call, $d, $eq, $get, $getattr, $is, $list, $mod, $pyEval, $truthy, $typename } from "../../pyshim.js";
import { AttributeError, IndexError, KeyError, NotImplementedError, ValueError } from "../../../src/errors.js";
import { step } from "../runner.js";
import { test_pptx } from "./helpers.js";
import { XL_DATA_LABEL_POSITION } from "../../../src/enum/chart.js";
import { Presentation } from "../../../src/index.js";
void $py;


function given_a_DataLabels_object_showing_or_not_cat_name(context: any, showing_or_not: any): any {
  let chart: any, prs: any, series_idx: any;
  series_idx = $get($d({"not showing": 0, showing: 1}), showing_or_not);
  prs = $call(Presentation, test_pptx("cht-datalabels"));
  chart = $get($get(prs.slides, 2).shapes, 0).chart;
  context.dataLabels = $get($get(chart.plots, 0).series, series_idx).dataLabels;
}
step("given", "a DataLabels object {showing_or_not} category-name as data_labels", ["showing_or_not"], given_a_DataLabels_object_showing_or_not_cat_name);

function given_a_DataLabels_object_showing_or_not_leg_key(context: any, showing_or_not: any): any {
  let chart: any, prs: any, series_idx: any;
  series_idx = $get($d({"not showing": 0, showing: 1}), showing_or_not);
  prs = $call(Presentation, test_pptx("cht-datalabels"));
  chart = $get($get(prs.slides, 2).shapes, 0).chart;
  context.dataLabels = $get($get(chart.plots, 0).series, series_idx).dataLabels;
}
step("given", "a DataLabels object {showing_or_not} legend-key as data_labels", ["showing_or_not"], given_a_DataLabels_object_showing_or_not_leg_key);

function given_a_DataLabels_object_showing_or_not_percent(context: any, showing_or_not: any): any {
  let chart: any, prs: any, slide_idx: any;
  slide_idx = $get($d({"not showing": 4, showing: 3}), showing_or_not);
  prs = $call(Presentation, test_pptx("cht-datalabels"));
  chart = $get($get(prs.slides, slide_idx).shapes, 0).chart;
  context.dataLabels = $get($get(chart.plots, 0).series, 0).dataLabels;
}
step("given", "a DataLabels object {showing_or_not} percentage as data_labels", ["showing_or_not"], given_a_DataLabels_object_showing_or_not_percent);

function given_a_DataLabels_object_showing_or_not_ser_name(context: any, showing_or_not: any): any {
  let chart: any, prs: any, series_idx: any;
  series_idx = $get($d({"not showing": 0, showing: 1}), showing_or_not);
  prs = $call(Presentation, test_pptx("cht-datalabels"));
  chart = $get($get(prs.slides, 2).shapes, 0).chart;
  context.dataLabels = $get($get(chart.plots, 0).series, series_idx).dataLabels;
}
step("given", "a DataLabels object {showing_or_not} series-name as data_labels", ["showing_or_not"], given_a_DataLabels_object_showing_or_not_ser_name);

function given_a_DataLabels_object_showing_or_not_value(context: any, showing_or_not: any): any {
  let chart: any, prs: any, series_idx: any;
  series_idx = $get($d({"not showing": 0, showing: 1}), showing_or_not);
  prs = $call(Presentation, test_pptx("cht-datalabels"));
  chart = $get($get(prs.slides, 2).shapes, 0).chart;
  context.dataLabels = $get($get(chart.plots, 0).series, series_idx).dataLabels;
}
step("given", "a DataLabels object {showing_or_not} value as data_labels", ["showing_or_not"], given_a_DataLabels_object_showing_or_not_value);

function given_a_DataLabels_object_with_pos_position(context: any, pos: any): any {
  let chart: any, prs: any, slide_idx: any;
  slide_idx = $get($d({inherited: 0, "inside-base": 1}), pos);
  prs = $call(Presentation, test_pptx("cht-datalabels"));
  chart = $get($get(prs.slides, slide_idx).shapes, 0).chart;
  context.dataLabels = $get(chart.plots, 0).dataLabels;
}
step("given", "a DataLabels object with {pos} position as data_labels", ["pos"], given_a_DataLabels_object_with_pos_position);

function given_a_data_label_having_or_not_custom_font(context: any, having_or_not: any): any {
  let point_idx: any, points: any, prs: any;
  point_idx = $get($d({"having a": 0, "having no": 1}), having_or_not);
  prs = $call(Presentation, test_pptx("cht-point-props"));
  points = $get($get($get($get(prs.slides, 2).shapes, 0).chart.plots, 0).series, 0).points;
  context.dataLabel = $get(points, point_idx).dataLabel;
}
step("given", "a data label {having_or_not} custom font as data_label", ["having_or_not"], given_a_data_label_having_or_not_custom_font);

function given_a_data_label_having_or_not_custom_text(context: any, having_or_not: any): any {
  let plot: any, point_idx: any, prs: any;
  point_idx = $get($d({having: 0, "having no": 1}), having_or_not);
  prs = $call(Presentation, test_pptx("cht-point-props"));
  plot = $get($get($get(prs.slides, 0).shapes, 0).chart.plots, 0);
  context.dataLabel = $get($get(plot.series, 0).points, point_idx).dataLabel;
}
step("given", "a data label {having_or_not} custom text as data_label", ["having_or_not"], given_a_data_label_having_or_not_custom_text);

function given_a_data_label_with_pos_position_as_data_label(context: any, pos: any): any {
  let plot: any, point_idx: any, prs: any;
  point_idx = $get($d({inherited: 0, centered: 1, below: 2}), pos);
  prs = $call(Presentation, test_pptx("cht-point-props"));
  plot = $get($get($get(prs.slides, 1).shapes, 0).chart.plots, 0);
  context.dataLabel = $get($get(plot.series, 0).points, point_idx).dataLabel;
}
step("given", "a data label with {pos} position as data_label", ["pos"], given_a_data_label_with_pos_position_as_data_label);

function when_I_assign_value_to_data_label_has_text_frame(context: any, value: any): any {
  let new_value: any;
  new_value = $get($d({True: true, False: false}), value);
  context.dataLabel.hasTextFrame = new_value;
}
step("when", "I assign {value} to data_label.has_text_frame", ["value"], when_I_assign_value_to_data_label_has_text_frame);

function when_I_assign_value_to_data_label_position(context: any, value: any): any {
  let new_value: any;
  new_value = ($truthy($eq(value, "None")) ? null : $getattr(XL_DATA_LABEL_POSITION, value));
  context.dataLabel.position = new_value;
}
step("when", "I assign {value} to data_label.position", ["value"], when_I_assign_value_to_data_label_position);

function when_I_assign_value_to_data_labels_position(context: any, value: any): any {
  let new_value: any;
  new_value = ($truthy($eq(value, "None")) ? null : $getattr(XL_DATA_LABEL_POSITION, value));
  context.dataLabels.position = new_value;
}
step("when", "I assign {value} to data_labels.position", ["value"], when_I_assign_value_to_data_labels_position);

function when_I_assign_value_to_data_labels_show_category_name(context: any, value: any): any {
  context.dataLabels.showCategoryName = $pyEval(value, $scope);
}
step("when", "I assign {value} to data_labels.show_category_name", ["value"], when_I_assign_value_to_data_labels_show_category_name);

function when_I_assign_value_to_data_labels_show_legend_key(context: any, value: any): any {
  context.dataLabels.showLegendKey = $pyEval(value, $scope);
}
step("when", "I assign {value} to data_labels.show_legend_key", ["value"], when_I_assign_value_to_data_labels_show_legend_key);

function when_I_assign_value_to_data_labels_show_percentage(context: any, value: any): any {
  context.dataLabels.showPercentage = $pyEval(value, $scope);
}
step("when", "I assign {value} to data_labels.show_percentage", ["value"], when_I_assign_value_to_data_labels_show_percentage);

function when_I_assign_value_to_data_labels_show_series_name(context: any, value: any): any {
  context.dataLabels.showSeriesName = $pyEval(value, $scope);
}
step("when", "I assign {value} to data_labels.show_series_name", ["value"], when_I_assign_value_to_data_labels_show_series_name);

function when_I_assign_value_to_data_labels_show_value(context: any, value: any): any {
  context.dataLabels.showValue = $pyEval(value, $scope);
}
step("when", "I assign {value} to data_labels.show_value", ["value"], when_I_assign_value_to_data_labels_show_value);

function then_data_label_font_is_a_Font_object(context: any): any {
  let font: any;
  font = context.dataLabel.font;
  $assertEqual($typename(font), "Font");
}
step("then", "data_label.font is a Font object", [], then_data_label_font_is_a_Font_object);

function then_data_label_has_text_frame_is_value(context: any, value: any): any {
  let data_label: any, expected_value: any;
  expected_value = $get($d({True: true, False: false}), value);
  data_label = context.dataLabel;
  $assert($is(data_label.hasTextFrame, expected_value));
}
step("then", "data_label.has_text_frame is {value}", ["value"], then_data_label_has_text_frame_is_value);

function then_data_label_position_is_value(context: any, value: any): any {
  let data_label: any, expected_value: any;
  expected_value = ($truthy($eq(value, "None")) ? null : $getattr(XL_DATA_LABEL_POSITION, value));
  data_label = context.dataLabel;
  $assert($is(data_label.position, expected_value), $mod("got %s", data_label.position));
}
step("then", "data_label.position is {value}", ["value"], then_data_label_position_is_value);

function then_data_label_text_frame_is_a_TextFrame_object(context: any): any {
  let text_frame: any;
  text_frame = context.dataLabel.textFrame;
  $assertEqual($typename(text_frame), "TextFrame");
}
step("then", "data_label.text_frame is a TextFrame object", [], then_data_label_text_frame_is_a_TextFrame_object);

function then_data_labels_position_is_value(context: any, value: any): any {
  let data_labels: any, expected_value: any;
  expected_value = ($truthy($eq(value, "None")) ? null : $getattr(XL_DATA_LABEL_POSITION, value));
  data_labels = context.dataLabels;
  $assert($is(data_labels.position, expected_value), $mod("got %s", data_labels.position));
}
step("then", "data_labels.position is {value}", ["value"], then_data_labels_position_is_value);

function then_data_labels_show_category_name_is_value(context: any, value: any): any {
  let actual: any, expected: any;
  [actual, expected] = $list([context.dataLabels.showCategoryName, $pyEval(value, $scope)]);
  $assert($is(actual, expected), $mod("data_labels.show_category_name is %s", actual));
}
step("then", "data_labels.show_category_name is {value}", ["value"], then_data_labels_show_category_name_is_value);

function then_data_labels_show_legend_key_is_value(context: any, value: any): any {
  let actual: any, expected: any;
  [actual, expected] = $list([context.dataLabels.showLegendKey, $pyEval(value, $scope)]);
  $assert($is(actual, expected), $mod("data_labels.show_legend_key is %s", actual));
}
step("then", "data_labels.show_legend_key is {value}", ["value"], then_data_labels_show_legend_key_is_value);

function then_data_labels_show_percentage_is_value(context: any, value: any): any {
  let actual: any, expected: any;
  [actual, expected] = $list([context.dataLabels.showPercentage, $pyEval(value, $scope)]);
  $assert($is(actual, expected), $mod("data_labels.show_percentage is %s", actual));
}
step("then", "data_labels.show_percentage is {value}", ["value"], then_data_labels_show_percentage_is_value);

function then_data_labels_show_series_name_is_value(context: any, value: any): any {
  let actual: any, expected: any;
  [actual, expected] = $list([context.dataLabels.showSeriesName, $pyEval(value, $scope)]);
  $assert($is(actual, expected), $mod("data_labels.show_series_name is %s", actual));
}
step("then", "data_labels.show_series_name is {value}", ["value"], then_data_labels_show_series_name_is_value);

function then_data_labels_show_value_is_value(context: any, value: any): any {
  let actual: any, expected: any;
  [actual, expected] = $list([context.dataLabels.showValue, $pyEval(value, $scope)]);
  $assert($is(actual, expected), $mod("data_labels.show_value is %s", actual));
}
step("then", "data_labels.show_value is {value}", ["value"], then_data_labels_show_value_is_value);

// the names python's eval() sees in this module
const $scope = { Presentation, XL_DATA_LABEL_POSITION };
