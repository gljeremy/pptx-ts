// Machine-translated from python-pptx features/steps/fill.py by
// scripts/py2ts_steps.py. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $assertEqual, $call, $d, $get, $mod, $pyEval, $round, $typename } from "../../pyshim.js";
import { AttributeError, IndexError, KeyError, NotImplementedError, ValueError } from "../../../src/errors.js";
import { step } from "../runner.js";
import { test_pptx } from "./helpers.js";
import { MSO_FILL, MSO_PATTERN } from "../../../src/enum/dml.js";
import { Presentation } from "../../../src/index.js";
void $py;


function given_a_FillFormat_object_as_fill(context: any): any {
  let fill: any;
  fill = $get($get($call(Presentation, test_pptx("dml-fill")).slides, 0).shapes, 0).fill;
  context.fill = fill;
}
step("given", "a FillFormat object as fill", [], given_a_FillFormat_object_as_fill);

function given_a_FillFormat_object_as_fill_having_pattern(context: any, pattern: any): any {
  let fill: any, shape_idx: any, slide: any;
  shape_idx = $get($d({"no pattern": 0, "MSO_PATTERN.DIVOT": 1, "MSO_PATTERN.WAVE": 2}), pattern);
  slide = $get($call(Presentation, test_pptx("dml-fill")).slides, 1);
  fill = $get(slide.shapes, shape_idx).fill;
  context.fill = fill;
}
step("given", "a FillFormat object as fill having {pattern} fill", ["pattern"], given_a_FillFormat_object_as_fill_having_pattern);

function given_type_FillFormat_object_as_fill(context: any, type: any): any {
  let shape: any, shape_idx: any;
  shape_idx = $get($d({"an inheriting": 0, "a no-fill": 1, "a solid": 2, "a picture": 3, "a gradient": 4, "a patterned": 5}), type);
  shape = $get($get($call(Presentation, test_pptx("dml-fill")).slides, 0).shapes, shape_idx);
  context.fill = shape.fill;
}
step("given", "{type} FillFormat object as fill", ["type"], given_type_FillFormat_object_as_fill);

function given_a_GradientStop_object_as_stop(context: any): any {
  let shape: any;
  shape = $get($get($call(Presentation, test_pptx("dml-fill")).slides, 0).shapes, 4);
  context.stop = $get(shape.fill.gradientStops, 0);
}
step("given", "a _GradientStop object as stop", [], given_a_GradientStop_object_as_stop);

function when_I_assign_value_to_fill_gradient_angle(context: any, value: any): any {
  context.fill.gradientAngle = $pyEval(value, $scope);
}
step("when", "I assign {value} to fill.gradient_angle", ["value"], when_I_assign_value_to_fill_gradient_angle);

function when_I_assign_value_to_fill_pattern(context: any, value: any): any {
  let pattern: any;
  pattern = $get($d({None: null, "MSO_PATTERN.CROSS": MSO_PATTERN.CROSS, "MSO_PATTERN.DIVOT": MSO_PATTERN.DIVOT, "MSO_PATTERN.WAVE": MSO_PATTERN.WAVE}), value);
  context.fill.pattern = pattern;
}
step("when", "I assign {value} to fill.pattern", ["value"], when_I_assign_value_to_fill_pattern);

function when_I_assign_value_to_stop_position(context: any, value: any): any {
  context.stop.position = $pyEval(value, $scope);
}
step("when", "I assign {value} to stop.position", ["value"], when_I_assign_value_to_stop_position);

function when_I_call_fill_background(context: any): any {
  context.fill.background();
}
step("when", "I call fill.background()", [], when_I_call_fill_background);

function when_I_call_fill_gradient(context: any): any {
  context.fill.gradient();
}
step("when", "I call fill.gradient()", [], when_I_call_fill_gradient);

function when_I_call_fill_patterned(context: any): any {
  context.fill.patterned();
}
step("when", "I call fill.patterned()", [], when_I_call_fill_patterned);

function when_I_call_fill_solid(context: any): any {
  context.fill.solid();
}
step("when", "I call fill.solid()", [], when_I_call_fill_solid);

function then_fill_back_color_is_a_ColorFormat_object(context: any): any {
  let actual_value: any, expected_value: any;
  actual_value = $typename(context.fill.backColor);
  expected_value = "ColorFormat";
  $assertEqual(actual_value, expected_value, $mod("fill.back_color is a %s object", actual_value));
}
step("then", "fill.back_color is a ColorFormat object", [], then_fill_back_color_is_a_ColorFormat_object);

function then_fill_fore_color_is_a_ColorFormat_object(context: any): any {
  let actual_value: any, expected_value: any;
  actual_value = $typename(context.fill.foreColor);
  expected_value = "ColorFormat";
  $assertEqual(actual_value, expected_value, $mod("fill.fore_color is a %s object", actual_value));
}
step("then", "fill.fore_color is a ColorFormat object", [], then_fill_fore_color_is_a_ColorFormat_object);

function then_fill_gradient_angle_eq_value(context: any, value: any): any {
  let actual_value: any, expected_value: any;
  expected_value = $round($pyEval(value, $scope), 2);
  actual_value = $round(context.fill.gradientAngle, 2);
  $assertEqual(actual_value, expected_value, $mod("fill.gradient_angle == %s", actual_value));
}
step("then", "fill.gradient_angle == {value}", ["value"], then_fill_gradient_angle_eq_value);

function then_fill_gradient_stops_is_a_GradientStops_object(context: any): any {
  let actual_value: any, expected_value: any;
  expected_value = "_GradientStops";
  actual_value = $typename(context.fill.gradientStops);
  $assertEqual(actual_value, expected_value, $mod("fill.gradient_stops is a %s object", actual_value));
}
step("then", "fill.gradient_stops is a _GradientStops object", [], then_fill_gradient_stops_is_a_GradientStops_object);

function then_fill_pattern_is_value(context: any, value: any): any {
  let expected_value: any, fill_pattern: any;
  fill_pattern = context.fill.pattern;
  expected_value = $get($d({None: null, "MSO_PATTERN.CROSS": MSO_PATTERN.CROSS, "MSO_PATTERN.DIVOT": MSO_PATTERN.DIVOT, "MSO_PATTERN.WAVE": MSO_PATTERN.WAVE}), value);
  $assertEqual(fill_pattern, expected_value, $mod("expected fill pattern %s, got %s", [expected_value, fill_pattern]));
}
step("then", "fill.pattern is {value}", ["value"], then_fill_pattern_is_value);

function then_fill_type_eq_value(context: any, value: any): any {
  let actual_value: any, expected_value: any;
  expected_value = $pyEval(value, $scope);
  actual_value = context.fill.type;
  $assertEqual(actual_value, expected_value, $mod("fill.type is %s", actual_value));
}
step("then", "fill.type == {value}", ["value"], then_fill_type_eq_value);

function then_stop_color_is_a_ColorFormat_object(context: any): any {
  let actual_value: any, expected_value: any;
  expected_value = "ColorFormat";
  actual_value = $typename(context.stop.color);
  $assertEqual(actual_value, expected_value, $mod("stop.color is a %s object", actual_value));
}
step("then", "stop.color is a ColorFormat object", [], then_stop_color_is_a_ColorFormat_object);

function then_stop_position_eq_value(context: any, value: any): any {
  let actual_value: any, expected_value: any;
  expected_value = $pyEval(value, $scope);
  actual_value = context.stop.position;
  $assertEqual(actual_value, expected_value, $mod("stop.position == %s", actual_value));
}
step("then", "stop.position == {value}", ["value"], then_stop_position_eq_value);

// the names python's eval() sees in this module
const $scope = { MSO_FILL, MSO_PATTERN, Presentation };
