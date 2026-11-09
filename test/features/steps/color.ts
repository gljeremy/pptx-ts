// Machine-translated from python-pptx features/steps/color.py by
// scripts/py2ts_steps.py. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $assertEqual, $call, $get, $mod, $repr } from "../../pyshim.js";
import { AttributeError, IndexError, KeyError, NotImplementedError, ValueError } from "../../../src/errors.js";
import { step } from "../runner.js";
import { test_pptx } from "./helpers.js";
import { RGBColor } from "../../../src/dml/color.js";
import { MSO_THEME_COLOR } from "../../../src/enum/dml.js";
import { Presentation } from "../../../src/index.js";
void $py;


function given_a_ColorFormat_object_as_color(context: any): any {
  let color: any, shape: any;
  shape = $get($get($call(Presentation, test_pptx("dml-fill")).slides, 0).shapes, 2);
  color = shape.fill.foreColor;
  context.color = color;
}
step("given", "a ColorFormat object as color", [], given_a_ColorFormat_object_as_color);

function when_assign_MSO_THEME_COLOR_ACCENT_6_to_color_theme_color(context: any): any {
  context.color.themeColor = MSO_THEME_COLOR.ACCENT_6;
}
step("when", "I assign MSO_THEME_COLOR.ACCENT_6 to color.theme_color", [], when_assign_MSO_THEME_COLOR_ACCENT_6_to_color_theme_color);

function when_I_assign_RGBColor_to_color_rgb(context: any): any {
  context.color.rgb = $call(RGBColor, 12, 34, 56);
}
step("when", "I assign RGBColor(12, 34, 56) to color.rgb", [], when_I_assign_RGBColor_to_color_rgb);

function when_I_assign_0_42_to_color_brightness(context: any): any {
  context.color.brightness = 0.42;
}
step("when", "I assign 0.42 to color.brightness", [], when_I_assign_0_42_to_color_brightness);

function then_color_brightness_is_value(context: any): any {
  let brightness: any, expected_value: any;
  brightness = context.color.brightness;
  expected_value = 0.42;
  $assertEqual(brightness, expected_value, $mod("expected %s, got %s", [expected_value, brightness]));
}
step("then", "color.brightness is 0.42", [], then_color_brightness_is_value);

function then_color_rgb_is_RGBColor_12_34_56(context: any): any {
  let expected_value: any, rgb: any;
  rgb = context.color.rgb;
  expected_value = $call(RGBColor, 12, 34, 56);
  $assertEqual(rgb, expected_value, $mod("expected %s, got %s", [$repr(expected_value), $repr(rgb)]));
}
step("then", "color.rgb is RGBColor(12, 34, 56)", [], then_color_rgb_is_RGBColor_12_34_56);

function then_color_theme_color_is_MSO_THEME_COLOR_ACCENT_6(context: any): any {
  let expected_value: any, theme_color: any;
  theme_color = context.color.themeColor;
  expected_value = MSO_THEME_COLOR.ACCENT_6;
  $assertEqual(theme_color, expected_value, $mod("expected %s, got %s", [expected_value, theme_color]));
}
step("then", "color.theme_color is MSO_THEME_COLOR.ACCENT_6", [], then_color_theme_color_is_MSO_THEME_COLOR_ACCENT_6);

// the names python's eval() sees in this module
const $scope = { MSO_THEME_COLOR, Presentation, RGBColor };
