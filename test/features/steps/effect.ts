// Machine-translated from python-pptx features/steps/effect.py by
// scripts/py2ts_steps.py. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $assert, $call, $d, $get, $is, $mod, $pyEval } from "../../pyshim.js";
import { AttributeError, IndexError, KeyError, NotImplementedError, ValueError } from "../../../src/errors.js";
import { step } from "../runner.js";
import { test_pptx } from "./helpers.js";
import { Presentation } from "../../../src/index.js";
void $py;


function given_a_ShadowFormat_object_that_inherits_or_not(context: any, inherits: any): any {
  let shape: any, shape_idx: any;
  shape_idx = $get($d({inherits: 0, "does not inherit": 1}), inherits);
  shape = $get($get($call(Presentation, test_pptx("dml-effect")).slides, 0).shapes, shape_idx);
  context.shadow = shape.shadow;
}
step("given", "a ShadowFormat object that {inherits} as shadow", ["inherits"], given_a_ShadowFormat_object_that_inherits_or_not);

function when_I_assign_value_to_shadow_inherit(context: any, value: any): any {
  context.shadow.inherit = $pyEval(value, $scope);
}
step("when", "I assign {value} to shadow.inherit", ["value"], when_I_assign_value_to_shadow_inherit);

function then_shadow_inherit_is_bool_val(context: any, bool_str: any): any {
  let actual_value: any, expected_value: any;
  expected_value = $pyEval(bool_str, $scope);
  actual_value = context.shadow.inherit;
  $assert($is(actual_value, expected_value), $mod("shadow.inherit is %s", actual_value));
}
step("then", "shadow.inherit is {bool_str}", ["bool_str"], then_shadow_inherit_is_bool_val);

// the names python's eval() sees in this module
const $scope = { Presentation };
