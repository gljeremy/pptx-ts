// Machine-translated from python-pptx features/steps/text.py by
// scripts/py2ts_steps.py. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $assert, $assertEqual, $call, $d, $eq, $get, $getattr, $int, $is, $isinstance, $list, $mod, $pyEval, $setattr, $truthy } from "../../pyshim.js";
import { AttributeError, IndexError, KeyError, NotImplementedError, ValueError } from "../../../src/errors.js";
import { step } from "../runner.js";
import { test_pptx } from "./helpers.js";
import { PP_ALIGN } from "../../../src/enum/text.js";
import { Presentation } from "../../../src/index.js";
import { Emu } from "../../../src/util.js";
void $py;


function given_a_Paragraph_object_as_paragraph(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("txt-text"));
  context.paragraph = $get($get($get(prs.slides, 0).shapes, 0).textFrame.paragraphs, 0);
}
step("given", "a _Paragraph object as paragraph", [], given_a_Paragraph_object_as_paragraph);

function given_a_Paragraph_object_containing_value_as_paragraph(context: any, value: any): any {
  let paragraph_idx: any, prs: any;
  prs = $call(Presentation, test_pptx("txt-text"));
  paragraph_idx = $get($d({abc: 0, "a\u000bb\u000bc": 1}), $pyEval(value, $scope));
  context.paragraph = $get($get($get(prs.slides, 0).shapes, 1).textFrame.paragraphs, paragraph_idx);
}
step("given", "a _Paragraph object containing {value} as paragraph", ["value"], given_a_Paragraph_object_containing_value_as_paragraph);

function given_a_paragraph_having_line_spacing_of_setting(context: any, setting: any): any {
  let paragraph_idx: any, prs: any, text_frame: any;
  paragraph_idx = $get($d({"no explicit setting": 0, "1.5 lines": 1, "20 pt": 2}), setting);
  prs = $call(Presentation, test_pptx("txt-paragraph-spacing"));
  text_frame = $get($get(prs.slides, 2).shapes, 0).textFrame;
  context.paragraph = $get(text_frame.paragraphs, paragraph_idx);
}
step("given", "a paragraph having line spacing of {setting}", ["setting"], given_a_paragraph_having_line_spacing_of_setting);

function given_a_paragraph_having_space_before_after_of_setting(context: any, before_after: any, setting: any): any {
  let paragraph_idx: any, prs: any, slide_idx: any, text_frame: any;
  slide_idx = $get($d({before: 0, after: 1}), before_after);
  paragraph_idx = $get($d({"no explicit setting": 0, "6 pt": 1}), setting);
  prs = $call(Presentation, test_pptx("txt-paragraph-spacing"));
  text_frame = $get($get(prs.slides, slide_idx).shapes, 0).textFrame;
  context.paragraph = $get(text_frame.paragraphs, paragraph_idx);
}
step("given", "a paragraph having space {before_after} of {setting}", ["before_after", "setting"], given_a_paragraph_having_space_before_after_of_setting);

function given_a_Run_object_as_run(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("txt-text"));
  context.run = $get($get($get($get(prs.slides, 0).shapes, 0).textFrame.paragraphs, 0).runs, 0);
}
step("given", "a _Run object as run", [], given_a_Run_object_as_run);

function given_a_Run_object_containing_text_as_run(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("txt-text"));
  context.run = $get($get($get($get(prs.slides, 0).shapes, 0).textFrame.paragraphs, 0).runs, 0);
}
step("given", "a _Run object containing text as run", [], given_a_Run_object_containing_text_as_run);

function given_a_text_run(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("txt-text"));
  context.run = $get($get($get($get(prs.slides, 0).shapes, 0).textFrame.paragraphs, 0).runs, 0);
}
step("given", "a text run", [], given_a_text_run);

function given_a_text_run_in_a_table_cell(context: any): any {
  let cell: any, prs: any;
  prs = $call(Presentation, test_pptx("txt-text"));
  cell = $get($get(prs.slides, 1).shapes, 0).table.cell(0, 0);
  context.run = $get($get(cell.textFrame.paragraphs, 0).runs, 0);
}
step("given", "a text run in a table cell", [], given_a_text_run_in_a_table_cell);

function given_a_text_run_having_a_hyperlink(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("txt-text"));
  context.run = $get($get($get($get(prs.slides, 0).shapes, 0).textFrame.paragraphs, 1).runs, 0);
}
step("given", "a text run having a hyperlink", [], given_a_text_run_having_a_hyperlink);

function when_assign_typeface_name_to_font(context: any): any {
  context.font.name = "Verdana";
}
step("when", "I assign a typeface name to the font", [], when_assign_typeface_name_to_font);

function when_assign_None_to_hyperlink_address(context: any): any {
  context.run.hyperlink.address = null;
}
step("when", "I assign None to hyperlink.address", [], when_assign_None_to_hyperlink_address);

function when_I_assign_paragraph_alignment_eq_center(context: any): any {
  context.paragraph.alignment = PP_ALIGN.CENTER;
}
step("when", "I assign paragraph.alignment = PP_ALIGN.CENTER", [], when_I_assign_paragraph_alignment_eq_center);

function when_I_assign_paragraph_leve_eq_1(context: any): any {
  context.paragraph.level = 1;
}
step("when", "I assign paragraph.level = 1", [], when_I_assign_paragraph_leve_eq_1);

function when_I_assign_paragraph_text_eq_value(context: any, value: any): any {
  context.paragraph.text = $pyEval(value, $scope);
}
step("when", "I assign paragraph.text = {value}", ["value"], when_I_assign_paragraph_text_eq_value);

function when_I_assign_run_text_eq_value(context: any, value: any): any {
  context.run.text = $pyEval(value, $scope);
}
step("when", "I assign run.text = {value}", ["value"], when_I_assign_run_text_eq_value);

function when_I_assign_value_to_paragraph_line_spacing(context: any, value_str: any): any {
  let paragraph: any, value: any;
  value = $get($d({"1.5": 1.5, "2.0": 2.0, "254000": $call(Emu, 254000), "304800": $call(Emu, 304800), None: null}), value_str);
  paragraph = context.paragraph;
  paragraph.lineSpacing = value;
}
step("when", "I assign {value_str} to paragraph.line_spacing", ["value_str"], when_I_assign_value_to_paragraph_line_spacing);

function when_I_assign_value_to_paragraph_space_before_after(context: any, value_str: any, before_after: any): any {
  let attr_name: any, paragraph: any, value: any;
  value = $get($d({"76200": 76200, "38100": 38100, None: null}), value_str);
  attr_name = $get($d({before: "space_before", after: "space_after"}), before_after);
  paragraph = context.paragraph;
  $setattr(paragraph, attr_name, value);
}
step("when", "I assign {value_str} to paragraph.space_{before_after}", ["value_str", "before_after"], when_I_assign_value_to_paragraph_space_before_after);

function when_set_hyperlink_address(context: any): any {
  let hlink: any, run: any;
  context.runText = "python-pptx @ GitHub";
  context.address = "https://github.com/scanny/python-pptx";
  run = context.run;
  run.text = context.runText;
  hlink = run.hyperlink;
  hlink.address = context.address;
}
step("when", "I set the hyperlink address", [], when_set_hyperlink_address);

function then_paragraph_alignment_eq_center(context: any): any {
  let actual: any, expected: any;
  [actual, expected] = $list([context.paragraph.alignment, PP_ALIGN.CENTER]);
  $assertEqual(actual, expected, $mod("paragraph.alignment == %s", actual));
}
step("then", "paragraph.alignment == PP_ALIGN.CENTER", [], then_paragraph_alignment_eq_center);

function then_paragraph_level_eq_1(context: any): any {
  let actual: any, expected: any;
  [actual, expected] = $list([context.paragraph.level, 1]);
  $assertEqual(actual, expected, $mod("paragraph.level == %s", actual));
}
step("then", "paragraph.level == 1", [], then_paragraph_level_eq_1);

function then_paragraph_line_spacing_is_value(context: any, value_str: any): any {
  let paragraph: any, value: any;
  value = $get($d({None: null, "1.0": 1.0, "1.5": 1.5, "2.0": 2.0, "254000": 254000, "304800": 304800}), value_str);
  paragraph = context.paragraph;
  $assertEqual(paragraph.lineSpacing, value);
}
step("then", "paragraph.line_spacing is {value_str}", ["value_str"], then_paragraph_line_spacing_is_value);

function then_paragraph_line_spacing_pt_result(context: any, result: any): any {
  let exception: any, line_spacing: any, value: any;
  [value, exception] = $list($get($d({"raises AttributeError": [null, AttributeError], "is 20.0": [20.0, null], "is 24.0": [24.0, null]}), result));
  line_spacing = context.paragraph.lineSpacing;
  if (exception !== null && line_spacing !== null) {
    $assert(line_spacing === 1.5 || line_spacing === 2.0);
    return;
  }
  if ($truthy(!$is(value, null))) {
    $assertEqual($getattr(line_spacing, "pt"), value);
  }
  if ($truthy(!$is(exception, null))) {
    try {
      $getattr(line_spacing, "pt");
      throw new Error("did not raise");
    } catch (_e: any) {
      if ($isinstance(_e, exception)) {
      } else {
        throw _e;
      }
    }
  }
}
step("then", "paragraph.line_spacing.pt {result}", ["result"], then_paragraph_line_spacing_pt_result);

function then_paragraph_space_before_is_value(context: any, before_after: any, value_str: any): any {
  let attr_name: any, paragraph: any, value: any;
  attr_name = $get($d({before: "space_before", after: "space_after"}), before_after);
  value = ($truthy($eq(value_str, "None")) ? null : $int(value_str));
  paragraph = context.paragraph;
  $assertEqual($getattr(paragraph, attr_name), value);
}
step("then", "paragraph.space_{before_after} is {value_str}", ["before_after", "value_str"], then_paragraph_space_before_is_value);

function then_paragraph_space_before_pt_is_value(context: any, before_after: any, result: any): any {
  let attr_name: any, exception: any, space_before_after: any, value: any;
  attr_name = $get($d({before: "space_before", after: "space_after"}), before_after);
  [value, exception] = $list($get($d({"raises AttributeError": [null, AttributeError], "== 6.0": [6.0, null]}), result));
  space_before_after = $getattr(context.paragraph, attr_name);
  if ($truthy(!$is(value, null))) {
    $assertEqual($getattr(space_before_after, "pt"), value);
  }
  if ($truthy(!$is(exception, null))) {
    try {
      $getattr(space_before_after, "pt");
      throw new Error("did not raise");
    } catch (_e: any) {
      if ($isinstance(_e, exception)) {
      } else {
        throw _e;
      }
    }
  }
}
step("then", "paragraph.space_{before_after}.pt {result}", ["before_after", "result"], then_paragraph_space_before_pt_is_value);

function then_paragraph_text_eq_value(context: any, value: any): any {
  let actual: any, expected: any;
  [actual, expected] = $list([context.paragraph.text, $pyEval(value, $scope)]);
  $assertEqual(actual, expected, $mod("paragraph.text == \"%s\"", actual));
}
step("then", "paragraph.text == {value}", ["value"], then_paragraph_text_eq_value);

function then_paragraph_text_matches_the_assigned_string(context: any): any {
  let paragraph: any;
  paragraph = context.paragraph;
  $assertEqual(paragraph.text, " Boo Far \n Faz Foo ");
}
step("then", "paragraph.text matches the assigned string", [], then_paragraph_text_matches_the_assigned_string);

function then_run_text_is_value(context: any, value: any): any {
  let actual: any, expected: any;
  [actual, expected] = $list([context.run.text, $pyEval(value, $scope)]);
  $assertEqual(actual, expected, $mod("run.text == %s", [actual]));
}
step("then", "run.text == {value}", ["value"], then_run_text_is_value);

function then_run_text_is_a_hyperlink(context: any): any {
  let hlink: any, run: any;
  run = context.run;
  hlink = run.hyperlink;
  $assertEqual(run.text, context.runText);
  $assertEqual(hlink.address, context.address);
}
step("then", "run.text is a hyperlink", [], then_run_text_is_a_hyperlink);

function then_run_text_is_not_a_hyperlink(context: any): any {
  let hlink: any;
  hlink = context.run.hyperlink;
  $assert($is(hlink.address, null));
}
step("then", "run.text is not a hyperlink", [], then_run_text_is_not_a_hyperlink);

function then_font_name_matches_typeface_I_set(context: any): any {
  $assertEqual(context.font.name, "Verdana");
}
step("then", "the font name matches the typeface I set", [], then_font_name_matches_typeface_I_set);

// the names python's eval() sees in this module
const $scope = { Emu, PP_ALIGN, Presentation };
