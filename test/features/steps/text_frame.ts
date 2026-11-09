// Machine-translated from python-pptx features/steps/text_frame.py by
// scripts/py2ts_steps.py. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $assert, $assertEqual, $call, $d, $float, $get, $getattr, $in, $is, $kw, $list, $m, $mod, $pyEval, $setattr } from "../../pyshim.js";
import { AttributeError, IndexError, KeyError, NotImplementedError, ValueError } from "../../../src/errors.js";
import { step } from "../runner.js";
import { test_file, test_pptx } from "./helpers.js";
import { MSO_AUTO_SIZE } from "../../../src/enum/text.js";
import { Presentation } from "../../../src/index.js";
import { Inches, Pt } from "../../../src/util.js";
void $py;


function given_a_text_frame(context: any): any {
  context.textFrame = $get($get($call(Presentation, test_pptx("txt-text")).slides, 0).shapes, 0).textFrame;
}
step("given", "a TextFrame object as text_frame", [], given_a_text_frame);

function given_a_TextFrame_object_containing_value_as_text_frame(context: any, value: any): any {
  let prs: any, shape_idx: any;
  shape_idx = $get($d({abc: 0, "a\nb\nc": 1}), $pyEval(value, $scope));
  prs = $call(Presentation, test_pptx("txt-text-frame"));
  context.textFrame = $get($get(prs.slides, 1).shapes, shape_idx).textFrame;
}
step("given", "a TextFrame object containing {value} as text_frame", ["value"], given_a_TextFrame_object_containing_value_as_text_frame);

function given_a_TextFrame_object_having_auto_size_of_setting(context: any, setting: any): any {
  let prs: any, shape: any, shape_idx: any;
  shape_idx = $get($d({None: 0, "no auto-size": 1, "fit shape to text": 2, "fit text to shape": 3}), setting);
  prs = $call(Presentation, test_pptx("txt-text-frame"));
  shape = $get($get(prs.slides, 0).shapes, shape_idx);
  context.textFrame = shape.textFrame;
}
step("given", "a TextFrame object having auto-size of {setting} as text_frame", ["setting"], given_a_TextFrame_object_having_auto_size_of_setting);

function given_a_text_frame_with_more_text_than_will_fit(context: any): any {
  let prs: any, shape: any;
  prs = $call(Presentation, test_pptx("txt-fit-text"));
  shape = $get($get(prs.slides, 0).shapes, 0);
  context.textFrame = shape.textFrame;
}
step("given", "a text frame with more text than will fit", [], given_a_text_frame_with_more_text_than_will_fit);

function when_I_assign_value_to_text_frame_auto_size(context: any, value: any): any {
  let text_frame: any;
  text_frame = context.textFrame;
  text_frame.autoSize = $get($d({None: null, "MSO_AUTO_SIZE.NONE": MSO_AUTO_SIZE.NONE, "MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT": MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT, "MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE": MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE}), value);
}
step("when", "I assign {value} to text_frame.auto_size", ["value"], when_I_assign_value_to_text_frame_auto_size);

function when_I_assign_text_frame_margin_side_eq_inches(context: any, side: any, inches: any): any {
  let attr_name: any;
  attr_name = $mod("margin_%s", side);
  $setattr(context.textFrame, attr_name, $call(Inches, $float(inches)));
}
step("when", "I assign text_frame.margin_{side} = Inches({inches})", ["side", "inches"], when_I_assign_text_frame_margin_side_eq_inches);

function when_I_assign_text_frame_text_eq_value(context: any, value: any): any {
  context.textFrame.text = $pyEval(value, $scope);
}
step("when", "I assign text_frame.text = {value}", ["value"], when_I_assign_text_frame_text_eq_value);

function when_I_assign_value_to_text_frame_word_wrap(context: any, value: any): any {
  let new_value: any;
  new_value = $get($d({True: true, False: false, None: null}), value);
  context.textFrame.wordWrap = new_value;
}
step("when", "I assign {value} to text_frame.word_wrap", ["value"], when_I_assign_value_to_text_frame_word_wrap);

function when_I_call_TextFrame_fit_text(context: any): any {
  let font_file: any;
  font_file = test_file("calibriz.ttf");
  $m(context.textFrame, "fitText", $kw({bold: true, italic: true, font_file: font_file}));
}
step("when", "I call TextFrame.fit_text()", [], when_I_call_TextFrame_fit_text);

function then_text_frame_autosize_is_value(context: any, value: any): any {
  let expected_value: any, text_frame: any;
  expected_value = $get($d({None: null, "MSO_AUTO_SIZE.NONE": MSO_AUTO_SIZE.NONE, "MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT": MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT, "MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE": MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE}), value);
  text_frame = context.textFrame;
  $assertEqual(text_frame.autoSize, expected_value, $mod("got %s", text_frame.autoSize));
}
step("then", "text_frame.auto_size is {value}", ["value"], then_text_frame_autosize_is_value);

function then_text_frame_margin_side_inches_eq_inches(context: any, side: any, inches: any): any {
  let actual: any, attr_name: any, expected: any;
  attr_name = $mod("margin_%s", side);
  actual = $getattr($getattr(context.textFrame, attr_name), "inches");
  expected = $float(inches);
  $assertEqual(actual, expected, $mod("text_frame.margin_%s.inches == %s", [side, actual]));
}
step("then", "text_frame.margin_{side}.inches == {inches}", ["side", "inches"], then_text_frame_margin_side_inches_eq_inches);

function then_text_frame_text_eq_value(context: any, value: any): any {
  let actual: any, expected: any;
  [actual, expected] = $list([context.textFrame.text, $pyEval(value, $scope)]);
  $assertEqual(actual, expected, $mod("text_frame.text == \"%s\"", actual));
}
step("then", "text_frame.text == {value}", ["value"], then_text_frame_text_eq_value);

function then_text_frame_word_wrap_is_value(context: any, value: any): any {
  let expected_value: any, text_frame: any;
  expected_value = $get($d({True: true, False: false, None: null}), value);
  text_frame = context.textFrame;
  $assert($is(text_frame.wordWrap, expected_value));
}
step("then", "text_frame.word_wrap is {value}", ["value"], then_text_frame_word_wrap_is_value);

function then_the_size_of_the_text_is_10pt(context: any): any {
  let text_frame: any;
  text_frame = context.textFrame;
  for (const paragraph of $list(text_frame.paragraphs)) {
    for (const run of $list(paragraph.runs)) {
      $assert($in(run.font.size, [$call(Pt, 10.0), $call(Pt, 11.0)]), $mod("got %s", $getattr(run.font.size, "pt")));
    }
  }
}
step("then", "the size of the text is 10pt or 11pt", [], then_the_size_of_the_text_is_10pt);

// the names python's eval() sees in this module
const $scope = { Inches, MSO_AUTO_SIZE, Presentation, Pt };
