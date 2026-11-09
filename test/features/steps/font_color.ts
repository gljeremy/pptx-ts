// Machine-translated from python-pptx features/steps/font_color.py by
// scripts/py2ts_steps.py. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $assertEqual, $call, $d, $eq, $float, $get, $truthy } from "../../pyshim.js";
import { AttributeError, IndexError, KeyError, NotImplementedError, ValueError } from "../../../src/errors.js";
import { step } from "../runner.js";
import { test_pptx } from "./helpers.js";
import { RGBColor } from "../../../src/dml/color.js";
import { MSO_COLOR_TYPE, MSO_THEME_COLOR } from "../../../src/enum/dml.js";
import { Presentation } from "../../../src/index.js";
void $py;


let font_color_pptx_path = test_pptx("font-color");

function step_given_font_with_color_type(context: any, color_type: any): any {
  let textbox: any;
  context.textboxIdx = $get($d({no: 0, "an RGB": 1, "a theme": 2}), color_type);
  context.prs = $call(Presentation, font_color_pptx_path);
  textbox = $get($get(context.prs.slides, 0).shapes, context.textboxIdx);
  context.font = $get($get(textbox.textFrame.paragraphs, 0).runs, 0).font;
}
step("given", "a font with {color_type} color", ["color_type"], step_given_font_with_color_type);

function step_font_with_color_brightness(context: any, setting: any): any {
  let textbox: any, textbox_idx: any;
  textbox_idx = $get($d({"no brightness adjustment": 2, "25% darker": 3, "40% lighter": 4}), setting);
  context.prs = $call(Presentation, font_color_pptx_path);
  textbox = $get($get(context.prs.slides, 0).shapes, textbox_idx);
  context.font = $get($get(textbox.textFrame.paragraphs, 0).runs, 0).font;
}
step("given", "a font with a color brightness setting of {setting}", ["setting"], step_font_with_color_brightness);

function step_set_font_color_brightness(context: any, value: any): any {
  context.font.color.brightness = $float(value);
}
step("when", "I set the font color brightness to {value}", ["value"], step_set_font_color_brightness);

function step_set_font_color_value(context: any, color_type: any): any {
  if ($truthy($eq(color_type, "RGB"))) {
    context.font.color.rgb = $call(RGBColor, 18, 52, 86);
  } else {
    if ($truthy($eq(color_type, "theme color"))) {
      context.font.color.themeColor = MSO_THEME_COLOR.DARK_1;
    }
  }
}
step("when", "I set the font {color_type} value", ["color_type"], step_set_font_color_value);

function step_color_value_matches_RGB_color(context: any): any {
  $assertEqual(context.font.color.rgb, $call(RGBColor, 255, 102, 0));
}
step("then", "its color value matches its RGB color", [], step_color_value_matches_RGB_color);

function step_color_value_matches_theme_color(context: any): any {
  $assertEqual(context.font.color.themeColor, MSO_THEME_COLOR.ACCENT_1);
}
step("then", "its color value matches its theme color", [], step_color_value_matches_theme_color);

function step_then_font_color_type_is_value(context: any, color_type: any): any {
  let expected_value: any, font: any, textbox: any;
  expected_value = $get($d({None: null, RGB: MSO_COLOR_TYPE.RGB, "theme color": MSO_COLOR_TYPE.SCHEME}), color_type);
  textbox = $get($get(context.prs.slides, 0).shapes, context.textboxIdx);
  font = $get($get(textbox.textFrame.paragraphs, 0).runs, 0).font;
  $assertEqual(font.color.type, expected_value);
}
step("then", "the font's color type is {color_type}", ["color_type"], step_then_font_color_type_is_value);

function step_color_brightness_value_matches(context: any, value: any): any {
  $assertEqual(context.font.color.brightness, $float(value));
}
step("then", "its color brightness value is {value}", ["value"], step_color_brightness_value_matches);

function step_color_type_value_matches(context: any, color_type: any): any {
  let font: any, textbox: any;
  textbox = $get($get(context.prs.slides, 0).shapes, context.textboxIdx);
  font = $get($get(textbox.textFrame.paragraphs, 0).runs, 0).font;
  if ($truthy($eq(color_type, "RGB"))) {
    $assertEqual(font.color.rgb, $call(RGBColor, 18, 52, 86));
  } else {
    $assertEqual(font.color.themeColor, MSO_THEME_COLOR.DARK_1);
  }
}
step("then", "the font's {color_type} value matches the value I set", ["color_type"], step_color_type_value_matches);

function step_color_brightness_matches(context: any, value: any): any {
  let font: any, textbox: any;
  textbox = $get($get(context.prs.slides, 0).shapes, context.textboxIdx);
  font = $get($get(textbox.textFrame.paragraphs, 0).runs, 0).font;
  $assertEqual(font.color.brightness, $float(value));
}
step("then", "the font's color brightness is {value}", ["value"], step_color_brightness_matches);

// the names python's eval() sees in this module
const $scope = { MSO_COLOR_TYPE, MSO_THEME_COLOR, Presentation, RGBColor };
