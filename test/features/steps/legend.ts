// Machine-translated from python-pptx features/steps/legend.py by
// scripts/py2ts_steps.py. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $assert, $assertEqual, $call, $d, $float, $get, $getattr, $is, $isinstance, $mod } from "../../pyshim.js";
import { AttributeError, IndexError, KeyError, NotImplementedError, ValueError } from "../../../src/errors.js";
import { step } from "../runner.js";
import { test_pptx } from "./helpers.js";
import { XL_LEGEND_POSITION } from "../../../src/enum/chart.js";
import { Presentation } from "../../../src/index.js";
import { Font } from "../../../src/text/text.js";
void $py;


function given_a_legend(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("cht-legend-props"));
  context.legend = $get($get(prs.slides, 0).shapes, 0).chart.legend;
}
step("given", "a legend", [], given_a_legend);

function given_a_legend_having_horizontal_offset_of_value(context: any, value: any): any {
  let prs: any, slide_idx: any;
  slide_idx = $get($d({none: 0, "-0.5": 1, "0.42": 2}), value);
  prs = $call(Presentation, test_pptx("cht-legend-props"));
  context.legend = $get($get(prs.slides, slide_idx).shapes, 0).chart.legend;
}
step("given", "a legend having horizontal offset of {value}", ["value"], given_a_legend_having_horizontal_offset_of_value);

function given_a_legend_positioned_location_the_chart(context: any, location: any): any {
  let prs: any, slide_idx: any;
  slide_idx = $get($d({"at an unspecified location of": 0, below: 1, "to the right of": 2}), location);
  prs = $call(Presentation, test_pptx("cht-legend-props"));
  context.legend = $get($get(prs.slides, slide_idx).shapes, 0).chart.legend;
}
step("given", "a legend positioned {location} the chart", ["location"], given_a_legend_positioned_location_the_chart);

function given_a_legend_with_overlay_setting_of_setting(context: any, setting: any): any {
  let prs: any, slide_idx: any;
  slide_idx = $get($d({"no explicit setting": 0, True: 1, False: 2}), setting);
  prs = $call(Presentation, test_pptx("cht-legend-props"));
  context.legend = $get($get(prs.slides, slide_idx).shapes, 0).chart.legend;
}
step("given", "a legend with overlay setting of {setting}", ["setting"], given_a_legend_with_overlay_setting_of_setting);

function when_I_assign_value_to_legend_horz_offset(context: any, value: any): any {
  let new_value: any;
  new_value = $float(value);
  context.legend.horzOffset = new_value;
}
step("when", "I assign {value} to legend.horz_offset", ["value"], when_I_assign_value_to_legend_horz_offset);

function when_I_assign_value_to_legend_include_in_layout(context: any, value: any): any {
  let new_value: any;
  new_value = $get($d({True: true, False: false}), value);
  context.legend.includeInLayout = new_value;
}
step("when", "I assign {value} to legend.include_in_layout", ["value"], when_I_assign_value_to_legend_include_in_layout);

function when_I_assign_value_to_legend_position(context: any, value: any): any {
  let enum_value: any;
  enum_value = $getattr(XL_LEGEND_POSITION, value);
  context.legend.position = enum_value;
}
step("when", "I assign {value} to legend.position", ["value"], when_I_assign_value_to_legend_position);

function then_legend_font_is_a_Font_object(context: any): any {
  let legend: any;
  legend = context.legend;
  $assert($isinstance(legend.font, Font));
}
step("then", "legend.font is a Font object", [], then_legend_font_is_a_Font_object);

function then_legend_horz_offset_is_value(context: any, value: any): any {
  let expected_value: any, legend: any;
  expected_value = $float(value);
  legend = context.legend;
  $assertEqual(legend.horzOffset, expected_value);
}
step("then", "legend.horz_offset is {value}", ["value"], then_legend_horz_offset_is_value);

function then_legend_include_in_layout_is_value(context: any, value: any): any {
  let expected_value: any, legend: any;
  expected_value = $get($d({True: true, False: false}), value);
  legend = context.legend;
  $assert($is(legend.includeInLayout, expected_value));
}
step("then", "legend.include_in_layout is {value}", ["value"], then_legend_include_in_layout_is_value);

function then_legend_position_is_value(context: any, value: any): any {
  let expected_position: any, legend: any;
  expected_position = $getattr(XL_LEGEND_POSITION, value);
  legend = context.legend;
  $assert($is(legend.position, expected_position), $mod("got %s", legend.position));
}
step("then", "legend.position is {value}", ["value"], then_legend_position_is_value);

// the names python's eval() sees in this module
const $scope = { Font, Presentation, XL_LEGEND_POSITION };
