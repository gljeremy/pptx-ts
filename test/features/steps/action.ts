// Machine-translated from python-pptx features/steps/action.py by
// scripts/py2ts_steps.py. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $assert, $assertEqual, $call, $d, $eq, $get, $getattr, $index, $int, $isinstance, $mod, $startswith, $truthy } from "../../pyshim.js";
import { AttributeError, IndexError, KeyError, NotImplementedError, ValueError } from "../../../src/errors.js";
import { step } from "../runner.js";
import { test_file } from "./helpers.js";
import { Hyperlink } from "../../../src/action.js";
import { PP_ACTION } from "../../../src/enum/action.js";
import { Presentation } from "../../../src/index.js";
void $py;


function given_an_ActionSetting_object_as_click_action(context: any, action: any): any {
  let shape_idx: any, slides: any;
  shape_idx = $get($d({NONE: 0, NAMED_SLIDE: 6}), action);
  slides = $call(Presentation, test_file("act-props.pptm")).slides;
  context.slides = slides;
  context.clickAction = $get($get(slides, 2).shapes, shape_idx).clickAction;
}
step("given", "an ActionSetting object having action {action} as click_action", ["action"], given_an_ActionSetting_object_as_click_action);

function given_another_slide_in_the_deck_as_slide(context: any): any {
  context.slide = $get(context.slides, 1);
}
step("given", "another slide in the deck as slide", [], given_another_slide_in_the_deck_as_slide);

function given_a_shape_having_click_action_action(context: any, action: any): any {
  let shape_idx: any, slides: any;
  shape_idx = $index(["none", "first slide", "last slide", "previous slide", "next slide", "last slide viewed", "named slide", "end show", "hyperlink", "other presentation", "open file", "custom slide show", "OLE action", "run macro", "run program", "play media"], action);
  slides = $call(Presentation, test_file("act-props.pptm")).slides;
  context.slides = slides;
  context.clickAction = $get($get(slides, 2).shapes, shape_idx).clickAction;
}
step("given", "a shape having click action {action}", ["action"], given_a_shape_having_click_action_action);

function when_I_assign_value_to_click_action_hyperlink_address(context: any, value: any): any {
  value = ($truthy($eq(value, "None")) ? null : value);
  context.clickAction.hyperlink.address = value;
}
step("when", "I assign {value} to click_action.hyperlink.address", ["value"], when_I_assign_value_to_click_action_hyperlink_address);

function when_I_assign_value_to_click_action_target_slide(context: any, value: any): any {
  let rhs: any;
  rhs = $get($d({None: null, slide: context.slide}), value);
  context.clickAction.targetSlide = rhs;
}
step("when", "I assign {value} to click_action.target_slide", ["value"], when_I_assign_value_to_click_action_target_slide);

function then_click_action_action_is_value(context: any, member_name: any): any {
  let click_action: any, expected_value: any;
  click_action = context.clickAction;
  expected_value = $getattr(PP_ACTION, member_name);
  $assertEqual(click_action.action, expected_value);
}
step("then", "click_action.action is {member_name}", ["member_name"], then_click_action_action_is_value);

function then_click_action_hyperlink_is_a_Hyperlink_object(context: any): any {
  let hyperlink: any;
  hyperlink = context.clickAction.hyperlink;
  $assert($isinstance(hyperlink, Hyperlink));
}
step("then", "click_action.hyperlink is a Hyperlink object", [], then_click_action_hyperlink_is_a_Hyperlink_object);

function then_click_action_hyperlink_address_is_value(context: any, value: any): any {
  let expected_value: any, hyperlink: any;
  expected_value = ($truthy($eq(value, "None")) ? null : value);
  hyperlink = context.clickAction.hyperlink;
  $assertEqual(hyperlink.address, expected_value, $mod("expected %s, got %s", [expected_value, hyperlink.address]));
}
step("then", "click_action.hyperlink.address is {value}", ["value"], then_click_action_hyperlink_address_is_value);

function then_click_action_target_slide_is_value(context: any, value: any): any {
  let click_action: any, expected_value: any, idx: any;
  if ($truthy($startswith(value, "slides["))) {
    idx = $get(value, 7);
    expected_value = $get(context.slides, $int(idx));
  } else {
    if ($truthy($eq(value, "None"))) {
      expected_value = null;
    } else {
      expected_value = context.slide;
    }
  }
  click_action = context.clickAction;
  $assertEqual(click_action.targetSlide, expected_value);
}
step("then", "click_action.target_slide is {value}", ["value"], then_click_action_target_slide_is_value);

// the names python's eval() sees in this module
const $scope = { Hyperlink, PP_ACTION, Presentation };
