// Machine-translated from python-pptx features/steps/background.py by
// scripts/py2ts_steps.py. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $assertEqual, $call, $d, $get, $is, $mod, $truthy, $typename } from "../../pyshim.js";
import { AttributeError, IndexError, KeyError, NotImplementedError, ValueError } from "../../../src/errors.js";
import { step } from "../runner.js";
import { test_pptx } from "./helpers.js";
import { Presentation } from "../../../src/index.js";
void $py;


function given_a_Background_object_having_type_background(context: any, type: any): any {
  let prs: any, sld_idx: any, slide: any;
  sld_idx = $get($d({no: 0, "a fill": 1, "a style reference": null}), type);
  prs = $call(Presentation, test_pptx("sld-background"));
  slide = ($truthy($is(sld_idx, null)) ? $get(prs.slideMasters, 0) : $get(prs.slides, sld_idx));
  context.background = slide.background;
}
step("given", "a _Background object having {type} background as background", ["type"], given_a_Background_object_having_type_background);

function then_background_fill_is_a_Fill_object(context: any): any {
  let cls_name: any;
  cls_name = $typename(context.background.fill);
  $assertEqual(cls_name, "FillFormat", $mod("background.fill is a %s object", cls_name));
}
step("then", "background.fill is a FillFormat object", [], then_background_fill_is_a_Fill_object);

// the names python's eval() sees in this module
const $scope = { Presentation };
