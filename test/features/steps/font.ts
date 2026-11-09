// Machine-translated from python-pptx features/steps/font.py by
// scripts/py2ts_steps.py. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $assert, $assertEqual, $call, $d, $eq, $get, $getattr, $index, $is, $mod, $slice, $truthy } from "../../pyshim.js";
import { AttributeError, IndexError, KeyError, NotImplementedError, ValueError } from "../../../src/errors.js";
import { step } from "../runner.js";
import { test_pptx } from "./helpers.js";
import { MSO_LANGUAGE_ID } from "../../../src/enum/lang.js";
import { MSO_UNDERLINE } from "../../../src/enum/text.js";
import { Presentation } from "../../../src/index.js";
void $py;


function given_a_font(context: any): any {
  let prs: any, run: any, slide: any, textbox: any;
  prs = $call(Presentation, test_pptx("txt-font-props"));
  slide = $get(prs.slides, 1);
  textbox = $get(slide.shapes, 0);
  run = $get($get(textbox.textFrame.paragraphs, 0).runs, 0);
  context.font = run.font;
}
step("given", "a font", [], given_a_font);

function given_a_font_having_language_id_value(context: any, value: any): any {
  let prs: any, run: any, shape_idx: any, textbox: any;
  shape_idx = $get($d({"of no explicit setting": 0, "MSO_LANGUAGE_ID.FRENCH": 1, "MSO_LANGUAGE_ID.POLISH": 2}), value);
  prs = $call(Presentation, test_pptx("txt-font-props"));
  textbox = $get($get(prs.slides, 4).shapes, shape_idx);
  run = $get($get(textbox.textFrame.paragraphs, 0).runs, 0);
  context.font = run.font;
}
step("given", "a font having language id {value}", ["value"], given_a_font_having_language_id_value);

function given_a_font_having_size_of_value(context: any, value: any): any {
  let prs: any, run: any, shape_idx: any, slide: any, textbox: any;
  shape_idx = $get($d({"no explicit value": 0, "42pt": 1}), value);
  prs = $call(Presentation, test_pptx("txt-font-props"));
  slide = $get(prs.slides, 1);
  textbox = $get(slide.shapes, shape_idx);
  run = $get($get(textbox.textFrame.paragraphs, 0).runs, 0);
  context.font = run.font;
}
step("given", "a font having size of {value}", ["value"], given_a_font_having_size_of_value);

function given_a_font_with_bold_set_state(context: any, state: any): any {
  let paragraph: any, prs: any, shape_idx: any;
  shape_idx = $index(["on", "off", "to inherit"], state);
  prs = $call(Presentation, test_pptx("txt-font-props"));
  paragraph = $get($get($get(prs.slides, 2).shapes, shape_idx).textFrame.paragraphs, 0);
  context.font = $get(paragraph.runs, 0).font;
}
step("given", "a font with bold set {state}", ["state"], given_a_font_with_bold_set_state);

function given_run_with_italic_set_to_state(context: any, state: any): any {
  let prs: any, run_idx: any, runs: any;
  run_idx = $index(["on", "off", "to inherit"], state);
  prs = $call(Presentation, test_pptx("txt-font-props"));
  runs = $get($get($get(prs.slides, 0).shapes, 0).textFrame.paragraphs, 0).runs;
  context.font = $get(runs, run_idx).font;
}
step("given", "a font with italic set {state}", ["state"], given_run_with_italic_set_to_state);

function given_run_with_underline_set_to_state(context: any, state: any): any {
  let prs: any, run_idx: any, runs: any;
  run_idx = $index(["on", "off", "to inherit", "to DOUBLE_LINE", "to WAVY_LINE"], state);
  prs = $call(Presentation, test_pptx("txt-font-props"));
  runs = $get($get($get(prs.slides, 3).shapes, 0).textFrame.paragraphs, 0).runs;
  console.log($get(runs, run_idx)._r.xml);
  context.font = $get(runs, run_idx).font;
}
step("given", "a font with underline set {state}", ["state"], given_run_with_underline_set_to_state);

function when_I_assign_value_to_font_bold(context: any, value: any): any {
  let new_value: any;
  new_value = $get($d({True: true, False: false, None: null}), value);
  context.font.bold = new_value;
}
step("when", "I assign {value} to font.bold", ["value"], when_I_assign_value_to_font_bold);

function when_I_assign_value_to_font_italic(context: any, value: any): any {
  let new_value: any;
  new_value = $get($d({True: true, False: false, None: null}), value);
  context.font.italic = new_value;
}
step("when", "I assign {value} to font.italic", ["value"], when_I_assign_value_to_font_italic);

function when_I_assign_value_to_font_language_id(context: any, value: any): any {
  let new_value: any;
  new_value = ($truthy($eq(value, "None")) ? null : $getattr(MSO_LANGUAGE_ID, $slice(value, 16, null, null)));
  context.font.languageId = new_value;
}
step("when", "I assign {value} to font.language_id", ["value"], when_I_assign_value_to_font_language_id);

function when_I_assign_value_to_font_underline(context: any, value: any): any {
  let new_value: any;
  new_value = $get($d({True: true, False: false, None: null, DOUBLE_LINE: MSO_UNDERLINE.DOUBLE_LINE, NONE: MSO_UNDERLINE.NONE, SINGLE_LINE: MSO_UNDERLINE.SINGLE_LINE}), value);
  context.font.underline = new_value;
}
step("when", "I assign {value} to font.underline", ["value"], when_I_assign_value_to_font_underline);

function then_font_bold_is_value(context: any, value: any): any {
  let expected_value: any, font: any;
  expected_value = $get($d({True: true, False: false, None: null}), value);
  font = context.font;
  $assert($is(font.bold, expected_value));
}
step("then", "font.bold is {value}", ["value"], then_font_bold_is_value);

function then_font_italic_is_value(context: any, value: any): any {
  let expected_value: any, font: any;
  expected_value = $get($d({True: true, False: false, None: null}), value);
  font = context.font;
  $assert($is(font.italic, expected_value));
}
step("then", "font.italic is {value}", ["value"], then_font_italic_is_value);

function then_font_language_id_is_MSO_LANGUAGE_ID_(context: any, member: any): any {
  let expected_value: any, font: any;
  expected_value = $getattr(MSO_LANGUAGE_ID, member);
  font = context.font;
  $assert($is(font.languageId, expected_value));
}
step("then", "font.language_id is MSO_LANGUAGE_ID.{member}", ["member"], then_font_language_id_is_MSO_LANGUAGE_ID_);

function then_font_size_is_value(context: any, value_str: any): any {
  let expected_value: any, font: any, value: any;
  expected_value = $get($d({"42.0 points": 42.0, None: null}), value_str);
  font = context.font;
  value = ($truthy($is(font.size, null)) ? font.size : $getattr(font.size, "pt"));
  $assertEqual(value, expected_value, $mod("expected %s, got %s", [expected_value, value]));
}
step("then", "font.size is {value_str}", ["value_str"], then_font_size_is_value);

function then_font_underline_is_value(context: any, value: any): any {
  let expected_value: any, font: any;
  expected_value = $get($d({True: true, False: false, None: null, DOUBLE_LINE: MSO_UNDERLINE.DOUBLE_LINE, SINGLE_LINE: MSO_UNDERLINE.SINGLE_LINE, WAVY_LINE: MSO_UNDERLINE.WAVY_LINE}), value);
  font = context.font;
  console.log(font._rPr.xml);
  $assert($is(font.underline, expected_value), $mod("got %s", font.underline));
}
step("then", "font.underline is {value}", ["value"], then_font_underline_is_value);

// the names python's eval() sees in this module
const $scope = { MSO_LANGUAGE_ID, MSO_UNDERLINE, Presentation };
