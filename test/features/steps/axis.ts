// Machine-translated from python-pptx features/steps/axis.py by
// scripts/py2ts_steps.py. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $assert, $assertEqual, $call, $d, $eq, $float, $get, $getattr, $int, $is, $mod, $setattr, $truthy, $typename } from "../../pyshim.js";
import { AttributeError, IndexError, KeyError, NotImplementedError, ValueError } from "../../../src/errors.js";
import { step } from "../runner.js";
import { test_pptx } from "./helpers.js";
import { XL_AXIS_CROSSES, XL_CATEGORY_TYPE } from "../../../src/enum/chart.js";
import { Presentation } from "../../../src/index.js";
void $py;


function given_a_axis_type_axis(context: any, axis_type: any): any {
  let chart: any, prs: any;
  prs = $call(Presentation, test_pptx("cht-axis-props"));
  chart = $get($get(prs.slides, 0).shapes, 0).chart;
  context.axis = $get($d({category: chart.categoryAxis, value: chart.valueAxis}), axis_type);
}
step("given", "a {axis_type} axis", ["axis_type"], given_a_axis_type_axis);

function given_a_major_gridlines(context: any): any {
  let axis: any, prs: any;
  prs = $call(Presentation, test_pptx("cht-gridlines-props"));
  axis = $get($get(prs.slides, 0).shapes, 0).chart.valueAxis;
  context.gridlines = axis.majorGridlines;
}
step("given", "a major gridlines", [], given_a_major_gridlines);

function given_a_value_axis_having_cat_ax_crossing_of(context: any, crossing: any): any {
  let prs: any, slide_idx: any;
  slide_idx = $get($d({automatic: 0, maximum: 2, minimum: 3, "2.75": 4, "-1.5": 5}), crossing);
  prs = $call(Presentation, test_pptx("cht-axis-props"));
  context.valueAxis = $get($get(prs.slides, slide_idx).shapes, 0).chart.valueAxis;
}
step("given", "a value axis having category axis crossing of {crossing}", ["crossing"], given_a_value_axis_having_cat_ax_crossing_of);

function given_an_axis(context: any): any {
  let chart: any, prs: any;
  prs = $call(Presentation, test_pptx("cht-axis-props"));
  chart = $get($get(prs.slides, 0).shapes, 0).chart;
  context.axis = chart.valueAxis;
}
step("given", "an axis", [], given_an_axis);

function given_an_axis_having_a_or_no_title(context: any, a_or_no: any): any {
  let chart: any, prs: any;
  prs = $call(Presentation, test_pptx("cht-axis-props"));
  chart = $get($get(prs.slides, 7).shapes, 0).chart;
  context.axis = $get($d({a: chart.valueAxis, no: chart.categoryAxis}), a_or_no);
}
step("given", "an axis having {a_or_no} title", ["a_or_no"], given_an_axis_having_a_or_no_title);

function given_an_axis_having_major_or_minor_gridlines(context: any, major_or_minor: any): any {
  let chart: any, prs: any;
  prs = $call(Presentation, test_pptx("cht-axis-props"));
  chart = $get($get(prs.slides, 0).shapes, 0).chart;
  context.axis = chart.valueAxis;
}
step("given", "an axis having {major_or_minor} gridlines", ["major_or_minor"], given_an_axis_having_major_or_minor_gridlines);

function given_an_axis_having_major_or_minor_unit_of_value(context: any, major_or_minor: any, value: any): any {
  let chart: any, prs: any, slide_idx: any;
  slide_idx = ($truthy($eq(value, "Auto")) ? 0 : 1);
  prs = $call(Presentation, test_pptx("cht-axis-props"));
  chart = $get($get(prs.slides, slide_idx).shapes, 0).chart;
  context.axis = chart.valueAxis;
}
step("given", "an axis having {major_or_minor} unit of {value}", ["major_or_minor", "value"], given_an_axis_having_major_or_minor_unit_of_value);

function given_an_axis_having_reverse_order_turned_on_or_off(context: any, status: any): any {
  let chart: any, prs: any;
  prs = $call(Presentation, test_pptx("cht-axis-props"));
  chart = $get($get(prs.slides, 0).shapes, 0).chart;
  context.axis = $get($d({on: chart.valueAxis, off: chart.categoryAxis}), status);
}
step("given", "an axis having reverse-order turned {status}", ["status"], given_an_axis_having_reverse_order_turned_on_or_off);

function given_an_axis_of_type_cls_name(context: any, cls_name: any): any {
  let chart: any, prs: any, slide_idx: any;
  slide_idx = $get($d({CategoryAxis: 0, DateAxis: 6}), cls_name);
  prs = $call(Presentation, test_pptx("cht-axis-props"));
  chart = $get($get(prs.slides, slide_idx).shapes, 0).chart;
  context.axis = chart.categoryAxis;
}
step("given", "an axis of type {cls_name}", ["cls_name"], given_an_axis_of_type_cls_name);

function given_an_axis_not_having_major_or_minor_gridlines(context: any, major_or_minor: any): any {
  let chart: any, prs: any;
  prs = $call(Presentation, test_pptx("cht-axis-props"));
  chart = $get($get(prs.slides, 0).shapes, 0).chart;
  context.axis = chart.categoryAxis;
}
step("given", "an axis not having {major_or_minor} gridlines", ["major_or_minor"], given_an_axis_not_having_major_or_minor_gridlines);

function given_an_axis_title(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("cht-axis-props"));
  context.axisTitle = $get($get(prs.slides, 7).shapes, 0).chart.valueAxis.axisTitle;
}
step("given", "an axis title", [], given_an_axis_title);

function given_an_axis_title_having_a_or_no_text_frame(context: any, a_or_no: any): any {
  let axis: any, chart: any, prs: any;
  prs = $call(Presentation, test_pptx("cht-axis-props"));
  chart = $get($get(prs.slides, 7).shapes, 0).chart;
  axis = $get($d({a: chart.valueAxis, no: chart.categoryAxis}), a_or_no);
  context.axisTitle = axis.axisTitle;
}
step("given", "an axis title having {a_or_no} text frame", ["a_or_no"], given_an_axis_title_having_a_or_no_text_frame);

function given_tick_labels_having_an_offset_of_setting(context: any, setting: any): any {
  let chart: any, prs: any, slide_idx: any;
  slide_idx = $get($d({"no explicit setting": 0, "420": 1}), setting);
  prs = $call(Presentation, test_pptx("cht-ticklabels-props"));
  chart = $get($get(prs.slides, slide_idx).shapes, 0).chart;
  context.tickLabels = chart.categoryAxis.tickLabels;
}
step("given", "tick labels having an offset of {setting}", ["setting"], given_tick_labels_having_an_offset_of_setting);

function when_I_assign_value_to_axis_has_title(context: any, value: any): any {
  context.axis.hasTitle = $get($d({True: true, False: false}), value);
}
step("when", "I assign {value} to axis.has_title", ["value"], when_I_assign_value_to_axis_has_title);

function when_I_assign_value_to_axis_has_major_or_minor_gridlines(context: any, value: any, major_or_minor: any): any {
  let axis: any, new_value: any, propname: any;
  axis = context.axis;
  propname = $mod("has_%s_gridlines", major_or_minor);
  new_value = $get($d({True: true, False: false}), value);
  $setattr(axis, propname, new_value);
}
step("when", "I assign {value} to axis.has_{major_or_minor}_gridlines", ["value", "major_or_minor"], when_I_assign_value_to_axis_has_major_or_minor_gridlines);

function when_I_assign_value_to_axis_major_or_minor_unit(context: any, value: any, major_or_minor: any): any {
  let axis: any, new_value: any, propname: any;
  axis = context.axis;
  propname = $mod("%s_unit", major_or_minor);
  new_value = $get($d({"8.4": 8.4, "5": 5, None: null}), value);
  $setattr(axis, propname, new_value);
}
step("when", "I assign {value} to axis.{major_or_minor}_unit", ["value", "major_or_minor"], when_I_assign_value_to_axis_major_or_minor_unit);

function when_I_assign_value_to_axis_reverse_order(context: any, value: any): any {
  context.axis.reverseOrder = $get($d({True: true, False: false}), value);
}
step("when", "I assign {value} to axis.reverse_order", ["value"], when_I_assign_value_to_axis_reverse_order);

function when_I_assign_value_to_axis_title_has_text_frame(context: any, value: any): any {
  context.axisTitle.hasTextFrame = $get($d({True: true, False: false}), value);
}
step("when", "I assign {value} to axis_title.has_text_frame", ["value"], when_I_assign_value_to_axis_title_has_text_frame);

function when_I_assign_value_to_tick_labels_offset(context: any, value: any): any {
  let new_value: any;
  new_value = $int(value);
  context.tickLabels.offset = new_value;
}
step("when", "I assign {value} to tick_labels.offset", ["value"], when_I_assign_value_to_tick_labels_offset);

function when_I_assign_member_to_value_axis_crosses(context: any, member: any): any {
  let value_axis: any;
  value_axis = context.valueAxis;
  value_axis.crosses = $getattr(XL_AXIS_CROSSES, member);
}
step("when", "I assign {member} to value_axis.crosses", ["member"], when_I_assign_member_to_value_axis_crosses);

function when_I_assign_value_to_value_axis_crosses_at(context: any, value: any): any {
  let new_value: any;
  new_value = ($truthy($eq(value, "None")) ? null : $float(value));
  context.valueAxis.crossesAt = new_value;
}
step("when", "I assign {value} to value_axis.crosses_at", ["value"], when_I_assign_value_to_value_axis_crosses_at);

function then_axis_axis_title_is_an_AxisTitle_object(context: any): any {
  let class_name: any;
  class_name = $typename(context.axis.axisTitle);
  $assertEqual(class_name, "AxisTitle", $mod("got %s", class_name));
}
step("then", "axis.axis_title is an AxisTitle object", [], then_axis_axis_title_is_an_AxisTitle_object);

function then_axis_category_type_is_XL_CATEGORY_TYPE_member(context: any, member: any): any {
  let category_type: any, expected_value: any;
  expected_value = $getattr(XL_CATEGORY_TYPE, member);
  category_type = context.axis.categoryType;
  $assert($is(category_type, expected_value), $mod("got %s", category_type));
}
step("then", "axis.category_type is XL_CATEGORY_TYPE.{member}", ["member"], then_axis_category_type_is_XL_CATEGORY_TYPE_member);

function then_axis_format_is_a_ChartFormat_object(context: any): any {
  let axis: any;
  axis = context.axis;
  $assertEqual($typename(axis.format), "ChartFormat");
}
step("then", "axis.format is a ChartFormat object", [], then_axis_format_is_a_ChartFormat_object);

function then_axis_format_fill_is_a_FillFormat_object(context: any): any {
  let axis: any;
  axis = context.axis;
  $assertEqual($typename(axis.format.fill), "FillFormat");
}
step("then", "axis.format.fill is a FillFormat object", [], then_axis_format_fill_is_a_FillFormat_object);

function then_axis_format_line_is_a_LineFormat_object(context: any): any {
  let axis: any;
  axis = context.axis;
  $assertEqual($typename(axis.format.line), "LineFormat");
}
step("then", "axis.format.line is a LineFormat object", [], then_axis_format_line_is_a_LineFormat_object);

function then_axis_has_title_is_value(context: any, value: any): any {
  let actual_value: any, axis: any, expected_value: any;
  axis = context.axis;
  actual_value = axis.hasTitle;
  expected_value = $get($d({True: true, False: false}), value);
  $assert($is(actual_value, expected_value), $mod("got %s", actual_value));
}
step("then", "axis.has_title is {value}", ["value"], then_axis_has_title_is_value);

function then_axis_has_major_or_minor_gridlines_is_expected_value(context: any, major_or_minor: any, value: any): any {
  let actual_value: any, axis: any, expected_value: any;
  axis = context.axis;
  actual_value = $get($d({major: axis.hasMajorGridlines, minor: axis.hasMinorGridlines}), major_or_minor);
  expected_value = $get($d({True: true, False: false}), value);
  $assert($is(actual_value, expected_value), $mod("got %s", actual_value));
}
step("then", "axis.has_{major_or_minor}_gridlines is {value}", ["major_or_minor", "value"], then_axis_has_major_or_minor_gridlines_is_expected_value);

function then_axis_major_gridlines_is_a_MajorGridlines_object(context: any): any {
  let axis: any;
  axis = context.axis;
  $assertEqual($typename(axis.majorGridlines), "MajorGridlines");
}
step("then", "axis.major_gridlines is a MajorGridlines object", [], then_axis_major_gridlines_is_a_MajorGridlines_object);

function then_axis_major_or_minor_unit_is_value(context: any, major_or_minor: any, value: any): any {
  let actual_value: any, axis: any, expected_value: any, propname: any;
  axis = context.axis;
  propname = $mod("%s_unit", major_or_minor);
  actual_value = $getattr(axis, propname);
  expected_value = $get($d({"20.0": 20.0, "8.4": 8.4, "5.0": 5.0, "4.2": 4.2, None: null}), value);
  $assertEqual(actual_value, expected_value, $mod("got %s", actual_value));
}
step("then", "axis.{major_or_minor}_unit is {value}", ["major_or_minor", "value"], then_axis_major_or_minor_unit_is_value);

function then_axis_reverse_order_is_value(context: any, value: any): any {
  let actual_value: any, axis: any, expected_value: any;
  axis = context.axis;
  actual_value = axis.reverseOrder;
  expected_value = $get($d({True: true, False: false}), value);
  $assert($is(actual_value, expected_value), $mod("got %s", actual_value));
}
step("then", "axis.reverse_order is {value}", ["value"], then_axis_reverse_order_is_value);

function then_axis_title_format_is_a_ChartFormat_object(context: any): any {
  let class_name: any;
  class_name = $typename(context.axisTitle.format);
  $assertEqual(class_name, "ChartFormat", $mod("got %s", class_name));
}
step("then", "axis_title.format is a ChartFormat object", [], then_axis_title_format_is_a_ChartFormat_object);

function then_axis_title_format_fill_is_a_FillFormat_object(context: any): any {
  let class_name: any;
  class_name = $typename(context.axisTitle.format.fill);
  $assertEqual(class_name, "FillFormat", $mod("got %s", class_name));
}
step("then", "axis_title.format.fill is a FillFormat object", [], then_axis_title_format_fill_is_a_FillFormat_object);

function then_axis_title_format_line_is_a_LineFormat_object(context: any): any {
  let class_name: any;
  class_name = $typename(context.axisTitle.format.line);
  $assertEqual(class_name, "LineFormat", $mod("got %s", class_name));
}
step("then", "axis_title.format.line is a LineFormat object", [], then_axis_title_format_line_is_a_LineFormat_object);

function then_axis_title_has_text_frame_is_value(context: any, value: any): any {
  let actual_value: any, expected_value: any;
  actual_value = context.axisTitle.hasTextFrame;
  expected_value = $get($d({True: true, False: false}), value);
  $assert($is(actual_value, expected_value), $mod("got %s", actual_value));
}
step("then", "axis_title.has_text_frame is {value}", ["value"], then_axis_title_has_text_frame_is_value);

function then_axis_title_text_frame_is_a_TextFrame_object(context: any): any {
  let class_name: any;
  class_name = $typename(context.axisTitle.textFrame);
  $assertEqual(class_name, "TextFrame", $mod("got %s", class_name));
}
step("then", "axis_title.text_frame is a TextFrame object", [], then_axis_title_text_frame_is_a_TextFrame_object);

function then_gridlines_format_is_a_ChartFormat_object(context: any): any {
  let gridlines: any;
  gridlines = context.gridlines;
  $assertEqual($typename(gridlines.format), "ChartFormat");
}
step("then", "gridlines.format is a ChartFormat object", [], then_gridlines_format_is_a_ChartFormat_object);

function then_gridlines_format_fill_is_a_FillFormat_object(context: any): any {
  let gridlines: any;
  gridlines = context.gridlines;
  $assertEqual($typename(gridlines.format.fill), "FillFormat");
}
step("then", "gridlines.format.fill is a FillFormat object", [], then_gridlines_format_fill_is_a_FillFormat_object);

function then_gridlines_format_line_is_a_LineFormat_object(context: any): any {
  let gridlines: any;
  gridlines = context.gridlines;
  $assertEqual($typename(gridlines.format.line), "LineFormat");
}
step("then", "gridlines.format.line is a LineFormat object", [], then_gridlines_format_line_is_a_LineFormat_object);

function then_tick_labels_offset_is_expected_value(context: any, value: any): any {
  let expected_value: any, tick_labels: any;
  expected_value = $int(value);
  tick_labels = context.tickLabels;
  $assertEqual(tick_labels.offset, expected_value, $mod("got %s", tick_labels.offset));
}
step("then", "tick_labels.offset is {value}", ["value"], then_tick_labels_offset_is_expected_value);

function then_value_axis_crosses_is_value(context: any, member: any): any {
  let expected_value: any, value_axis: any;
  value_axis = context.valueAxis;
  expected_value = $getattr(XL_AXIS_CROSSES, member);
  $assertEqual(value_axis.crosses, expected_value, $mod("got %s", value_axis.crosses));
}
step("then", "value_axis.crosses is {member}", ["member"], then_value_axis_crosses_is_value);

function then_value_axis_crosses_at_is_value(context: any, value: any): any {
  let expected_value: any, value_axis: any;
  value_axis = context.valueAxis;
  expected_value = ($truthy($eq(value, "None")) ? null : $float(value));
  $assertEqual(value_axis.crossesAt, expected_value, $mod("got %s", value_axis.crossesAt));
}
step("then", "value_axis.crosses_at is {value}", ["value"], then_value_axis_crosses_at_is_value);

// the names python's eval() sees in this module
const $scope = { Presentation, XL_AXIS_CROSSES, XL_CATEGORY_TYPE };
