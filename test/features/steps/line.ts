// Machine-translated from python-pptx features/steps/line.py by
// scripts/py2ts_steps.py. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $assert, $assertEqual, $call, $d, $get, $isinstance, $mod, $typename } from "../../pyshim.js";
import { AttributeError, IndexError, KeyError, NotImplementedError, ValueError } from "../../../src/errors.js";
import { step } from "../runner.js";
import { test_pptx } from "./helpers.js";
import { MSO_LINE } from "../../../src/enum/dml.js";
import { Presentation } from "../../../src/index.js";
import { Length, Pt } from "../../../src/util.js";
void $py;


function given_a_LineFormat_object_as_line(context: any): any {
  let line: any;
  line = $get($get($call(Presentation, test_pptx("dml-line")).slides, 0).shapes, 0).line;
  context.line = line;
}
step("given", "a LineFormat object as line", [], given_a_LineFormat_object_as_line);

function given_a_LineFormat_object_as_line_having_dash_style(context: any, current: any): any {
  let shape: any, shape_idx: any;
  shape_idx = $get($d({"no explicit": 0, solid: 1, dashed: 2, "dash-dot": 3}), current);
  shape = $get($get($call(Presentation, test_pptx("dml-line")).slides, 3).shapes, shape_idx);
  context.line = shape.line;
}
step("given", "a LineFormat object as line having {current} dash style", ["current"], given_a_LineFormat_object_as_line_having_dash_style);

function given_a_LineFormat_object_as_line_having_width(context: any, line_width: any): any {
  let prs: any, shape: any, shape_idx: any;
  shape_idx = $get($d({"no explicit": 0, "1 pt": 1}), line_width);
  prs = $call(Presentation, test_pptx("dml-line"));
  shape = $get($get(prs.slides, 2).shapes, shape_idx);
  context.line = shape.line;
}
step("given", "a LineFormat object as line having {line_width} width", ["line_width"], given_a_LineFormat_object_as_line_having_width);

function when_I_assign_value_to_line_dash_style(context: any, value_key: any): any {
  let value: any;
  value = $get($d({None: null, "MSO_LINE.DASH": MSO_LINE.DASH, "MSO_LINE.DASH_DOT": MSO_LINE.DASH_DOT, "MSO_LINE.SOLID": MSO_LINE.SOLID}), value_key);
  context.line.dashStyle = value;
}
step("when", "I assign {value_key} to line.dash_style", ["value_key"], when_I_assign_value_to_line_dash_style);

function when_I_assign_value_to_line_width(context: any, line_width: any): any {
  let value: any;
  value = $get($d({None: null, "1 pt": $call(Pt, 1), "2.34 pt": $call(Pt, 2.34)}), line_width);
  context.line.width = value;
}
step("when", "I assign {line_width} to line.width", ["line_width"], when_I_assign_value_to_line_width);

function then_line_color_is_a_ColorFormat_object(context: any): any {
  let class_name: any, expected_value: any;
  class_name = $typename(context.line.color);
  expected_value = "ColorFormat";
  $assertEqual(class_name, expected_value, $mod("expected '%s', got '%s'", [expected_value, class_name]));
}
step("then", "line.color is a ColorFormat object", [], then_line_color_is_a_ColorFormat_object);

function then_line_dash_style_is_value(context: any, dash_style: any): any {
  let actual_value: any, expected_value: any;
  expected_value = $get($d({None: null, "MSO_LINE.DASH": MSO_LINE.DASH, "MSO_LINE.DASH_DOT": MSO_LINE.DASH_DOT, "MSO_LINE.SOLID": MSO_LINE.SOLID}), dash_style);
  actual_value = context.line.dashStyle;
  $assertEqual(actual_value, expected_value, $mod("expected %s, got %s", [expected_value, actual_value]));
}
step("then", "line.dash_style is {dash_style}", ["dash_style"], then_line_dash_style_is_value);

function then_line_fill_is_a_FillFormat_object(context: any): any {
  let class_name: any, expected_value: any;
  class_name = $typename(context.line.fill);
  expected_value = "FillFormat";
  $assertEqual(class_name, expected_value, $mod("expected '%s', got '%s'", [expected_value, class_name]));
}
step("then", "line.fill is a FillFormat object", [], then_line_fill_is_a_FillFormat_object);

function then_line_width_is_value(context: any, line_width: any): any {
  let expected_value: any;
  expected_value = $get($d({"0": 0, "1 pt": $call(Pt, 1), "2.34 pt": $call(Pt, 2.34)}), line_width);
  line_width = context.line.width;
  $assertEqual(line_width, expected_value);
  $assert($isinstance(line_width, Length));
}
step("then", "line.width is {line_width}", ["line_width"], then_line_width_is_value);

// the names python's eval() sees in this module
const $scope = { Length, MSO_LINE, Presentation, Pt };
