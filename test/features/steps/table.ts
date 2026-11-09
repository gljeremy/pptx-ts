// Machine-translated from python-pptx features/steps/table.py by
// scripts/py2ts_steps.py. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $assert, $assertEqual, $call, $d, $float, $get, $getattr, $int, $is, $len, $list, $m, $mod, $pyEval, $setattr, $split, $typename } from "../../pyshim.js";
import { AttributeError, IndexError, KeyError, NotImplementedError, ValueError } from "../../../src/errors.js";
import { step } from "../runner.js";
import { test_pptx } from "./helpers.js";
import { MSO_ANCHOR } from "../../../src/enum/text.js";
import { Presentation } from "../../../src/index.js";
import { Inches } from "../../../src/util.js";
void $py;


function given_a_2x2_Table_object_as_table(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("shp-shapes"));
  context.table_ = $get($get(prs.slides, 0).shapes, 3).table;
}
step("given", "a 2x2 Table object as table", [], given_a_2x2_Table_object_as_table);
step("given", "a Table object as table", [], given_a_2x2_Table_object_as_table);

function given_a_2x3_MergeOriginCell_object_as_cell(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("tbl-cell"));
  context.cell = $get($get(prs.slides, 1).shapes, 1).table.cell(0, 0);
}
step("given", "a 2x3 _MergeOriginCell object as cell", [], given_a_2x3_MergeOriginCell_object_as_cell);

function given_a_3x3_table_with_cells_a_to_i_as_table(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("tbl-cell"));
  context.table_ = $get($get(prs.slides, 2).shapes, 0).table;
}
step("given", "a 3x3 Table object with cells a to i as table", [], given_a_3x3_table_with_cells_a_to_i_as_table);
step("given", "a 3x3 Table object as table", [], given_a_3x3_table_with_cells_a_to_i_as_table);

function given_a_Cell_object_as_cell(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("shp-shapes"));
  context.cell = $get($get(prs.slides, 0).shapes, 3).table.cell(0, 0);
}
step("given", "a _Cell object as cell", [], given_a_Cell_object_as_cell);

function given_a_Cell_object_containing_unladen_swallows_as_cell(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("tbl-cell"));
  context.cell = $get($get(prs.slides, 0).shapes, 0).table.cell(1, 0);
}
step("given", "a _Cell object containing \"unladen swallows\" as cell", [], given_a_Cell_object_containing_unladen_swallows_as_cell);

function given_a_Cell_object_with_known_margins_as_cell(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("tbl-cell"));
  context.cell = $get($get(prs.slides, 0).shapes, 0).table.cell(0, 0);
}
step("given", "a _Cell object with known margins as cell", [], given_a_Cell_object_with_known_margins_as_cell);

function given_a_Cell_object_with_setting_vertical_alignment(context: any, setting: any): any {
  let cell_coordinates: any, prs: any;
  cell_coordinates = $get($d({inherited: [0, 1], middle: [0, 2], bottom: [0, 3]}), setting);
  prs = $call(Presentation, test_pptx("tbl-cell"));
  context.cell = $m($get($get(prs.slides, 0).shapes, 0).table, "cell", ...$list(cell_coordinates));
}
step("given", "a _Cell object with {setting} vertical alignment as cell", ["setting"], given_a_Cell_object_with_setting_vertical_alignment);

function given_a_role_Cell_object_as_cell(context: any, role: any): any {
  let coordinates: any, table: any;
  coordinates = $get($d({"merge-origin": [0, 0], spanned: [0, 1], unmerged: [2, 2]}), role);
  table = $get($get($call(Presentation, test_pptx("tbl-cell")).slides, 1).shapes, 0).table;
  context.cell = $m(table, "cell", ...$list(coordinates));
  context.otherCell = $m(table, "cell", ...$list(coordinates));
}
step("given", "a {role} _Cell object as cell", ["role"], given_a_role_Cell_object_as_cell);

function given_a_second_proxy_instance_for_that_cell_as_other_cell(context: any): any {
  $assert(context.otherCell);
}
step("given", "a second proxy instance for that cell as other_cell", [], given_a_second_proxy_instance_for_that_cell_as_other_cell);

function given_a_Column_object_as_column(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("shp-shapes"));
  context.column = $get($get($get(prs.slides, 0).shapes, 3).table.columns, 0);
}
step("given", "a _Column object as column", [], given_a_Column_object_as_column);

function when_I_assign_cell_margin_side_eq_value(context: any, value: any, side: any): any {
  $setattr(context.cell, $mod("margin_%s", side), $pyEval(value, $scope));
}
step("when", "I assign cell.margin_{side} = {value}", ["value", "side"], when_I_assign_cell_margin_side_eq_value);

function when_I_assign_cell_text(context: any): any {
  context.cell.text = "test text";
}
step("when", "I assign cell.text = \"test text\"", [], when_I_assign_cell_text);

function when_I_assign_cell_vertical_anchor_eq_value(context: any, value: any): any {
  context.cell.verticalAnchor = $pyEval(value, $scope);
}
step("when", "I assign cell.vertical_anchor = {value}", ["value"], when_I_assign_cell_vertical_anchor_eq_value);

function when_I_assign_column_width_eq_value(context: any, value: any): any {
  context.column.width = $pyEval(value, $scope);
}
step("when", "I assign column.width = {value}", ["value"], when_I_assign_column_width_eq_value);

function when_I_assign_origin_cell_eq_table_cell_0_0(context: any): any {
  context.originCell = context.table_.cell(0, 0);
}
step("when", "I assign origin_cell = table.cell(0, 0)", [], when_I_assign_origin_cell_eq_table_cell_0_0);

function when_I_assign_other_cell_eq_table_cell_1_1(context: any): any {
  context.otherCell = context.table_.cell(1, 1);
}
step("when", "I assign other_cell = table.cell(1, 1)", [], when_I_assign_other_cell_eq_table_cell_1_1);

function when_I_assign_table_first_col_eq_True(context: any): any {
  context.table_.firstCol = true;
}
step("when", "I assign table.first_col = True", [], when_I_assign_table_first_col_eq_True);

function when_I_assign_table_first_row_eq_True(context: any): any {
  context.table_.firstRow = true;
}
step("when", "I assign table.first_row = True", [], when_I_assign_table_first_row_eq_True);

function when_I_assign_table_horz_banding_eq_True(context: any): any {
  context.table_.horzBanding = true;
}
step("when", "I assign table.horz_banding = True", [], when_I_assign_table_horz_banding_eq_True);

function when_I_assign_table_last_col_eq_True(context: any): any {
  context.table_.lastCol = true;
}
step("when", "I assign table.last_col = True", [], when_I_assign_table_last_col_eq_True);

function when_I_assign_table_last_row_eq_True(context: any): any {
  context.table_.lastRow = true;
}
step("when", "I assign table.last_row = True", [], when_I_assign_table_last_row_eq_True);

function when_I_assign_table_vert_banding_eq_True(context: any): any {
  context.table_.vertBanding = true;
}
step("when", "I assign table.vert_banding = True", [], when_I_assign_table_vert_banding_eq_True);

function when_I_call_cell_split_other_cell(context: any): any {
  $split(context.cell);
}
step("when", "I call cell.split()", [], when_I_call_cell_split_other_cell);

function when_I_call_origin_cell_merge_other_cell(context: any): any {
  context.originCell.merge(context.otherCell);
}
step("when", "I call origin_cell.merge(other_cell)", [], when_I_call_origin_cell_merge_other_cell);

function then_cell_eq_other_cell(context: any): any {
  let cell: any, other_cell: any;
  [cell, other_cell] = $list([context.cell, context.otherCell]);
  $assertEqual(cell, other_cell, "cell != other_cell");
}
step("then", "cell == other_cell", [], then_cell_eq_other_cell);

function then_cell_fill_is_a_FillFormat_object(context: any): any {
  let actual: any, expected: any;
  actual = $typename(context.cell.fill);
  expected = "FillFormat";
  $assertEqual(actual, expected, $mod("cell.fill is a %s object", actual));
}
step("then", "cell.fill is a FillFormat object", [], then_cell_fill_is_a_FillFormat_object);

function then_cell_margin_side_eq_Inches_num(context: any, side: any, num_lit: any): any {
  let actual: any, expected: any;
  actual = $getattr(context.cell, $mod("margin_%s", side));
  expected = $call(Inches, $float(num_lit));
  $assertEqual(actual, expected, $mod("cell.margin_%s == %s", [side, $getattr(actual, "inches")]));
}
step("then", "cell.margin_{side} == Inches({num_lit})", ["side", "num_lit"], then_cell_margin_side_eq_Inches_num);

function then_cell_ref_is_merge_origin_is(context: any, cell_ref: any, bool_lit: any): any {
  let actual: any, expected: any;
  expected = $pyEval(bool_lit, $scope);
  actual = $getattr(context, cell_ref).isMergeOrigin;
  $assert($is(actual, expected), $mod("%s.is_merge_origin is %s", [cell_ref, actual]));
}
step("then", "{cell_ref}.is_merge_origin is {bool_lit}", ["cell_ref", "bool_lit"], then_cell_ref_is_merge_origin_is);

function then_cell_is_spanned_is(context: any, cell_ref: any, bool_lit: any): any {
  let actual: any, expected: any;
  expected = $pyEval(bool_lit, $scope);
  actual = $getattr(context, cell_ref).isSpanned;
  $assert($is(actual, expected), $mod("%s.is_spanned is %s", [cell_ref, actual]));
}
step("then", "{cell_ref}.is_spanned is {bool_lit}", ["cell_ref", "bool_lit"], then_cell_is_spanned_is);

function then_cell_text_eq_value(context: any, value: any): any {
  let actual: any, expected: any;
  [actual, expected] = $list([context.cell.text, $pyEval(value, $scope)]);
  $assertEqual(actual, expected, $mod("cell.text == \"%s\"", actual));
}
step("then", "cell.text == {value}", ["value"], then_cell_text_eq_value);

function then_cell_span_height_eq(context: any, int_lit: any): any {
  let actual: any, expected: any;
  expected = $int(int_lit);
  actual = context.cell.spanHeight;
  $assert($is(actual, expected), $mod("cell.span_height == %s", actual));
}
step("then", "cell.span_height == {int_lit}", ["int_lit"], then_cell_span_height_eq);

function then_cell_span_width_eq(context: any, int_lit: any): any {
  let actual: any, expected: any;
  expected = $int(int_lit);
  actual = context.cell.spanWidth;
  $assert($is(actual, expected), $mod("cell.span_width == %s", actual));
}
step("then", "cell.span_width == {int_lit}", ["int_lit"], then_cell_span_width_eq);

function then_cell_vertical_anchor_eq_value(context: any, value: any): any {
  let actual: any, expected: any;
  actual = context.cell.verticalAnchor;
  expected = $pyEval(value, $scope);
  $assertEqual(actual, expected, $mod("cell.vertical_anchor == %s", actual));
}
step("then", "cell.vertical_anchor == {value}", ["value"], then_cell_vertical_anchor_eq_value);

function then_column_width_inches_eq(context: any, float_lit: any): any {
  let actual: any, expected: any;
  actual = $getattr(context.column.width, "inches");
  expected = $float(float_lit);
  $assertEqual(actual, expected, $mod("column.width.inches == %s", actual));
}
step("then", "column.width.inches == {float_lit}", ["float_lit"], then_column_width_inches_eq);

function then_len_list_table_iter_cells_eq(context: any, int_lit: any): any {
  let actual: any, expected: any;
  actual = $len($list(context.table_.iterCells()));
  expected = $int(int_lit);
  $assert($is(actual, expected), $mod("len(list(table.iter_cells())) == %s", actual));
}
step("then", "len(list(table.iter_cells())) == {int_lit}", ["int_lit"], then_len_list_table_iter_cells_eq);

function then_origin_cell_text_eq_value(context: any, value: any): any {
  let actual: any, expected: any;
  [actual, expected] = $list([context.originCell.text, $pyEval(value, $scope)]);
  $assertEqual(actual, expected, $mod("origin_cell.text == \"%s\"", actual));
}
step("then", "origin_cell.text == {value}", ["value"], then_origin_cell_text_eq_value);

function then_other_cell_text_eq_value(context: any, value: any): any {
  let actual: any, expected: any;
  [actual, expected] = $list([context.otherCell.text, $pyEval(value, $scope)]);
  $assertEqual(actual, expected, $mod("other_cell.text == \"%s\"", actual));
}
step("then", "other_cell.text == {value}", ["value"], then_other_cell_text_eq_value);

function then_table_cell_0_0_is_a_type_object(context: any, type_name: any): any {
  let actual: any, expected: any;
  actual = $typename(context.table_.cell(0, 0));
  expected = type_name;
  $assertEqual(actual, expected, $mod("table.cell(0, 0) is a %s object", actual));
}
step("then", "table.cell(0, 0) is a {type_name} object", ["type_name"], then_table_cell_0_0_is_a_type_object);

function then_table_columns_is_a_type_object(context: any, type_name: any): any {
  let actual: any, expected: any;
  actual = $typename(context.table_.columns);
  expected = type_name;
  $assertEqual(actual, expected, $mod("table.columns is a %s object", actual));
}
step("then", "table.columns is a {type_name} object", ["type_name"], then_table_columns_is_a_type_object);

function then_table_first_col_is_value(context: any, bool_lit: any): any {
  let actual: any, expected: any;
  actual = context.table_.firstCol;
  expected = $pyEval(bool_lit, $scope);
  $assert($is(actual, expected), $mod("table.first_col is %s", actual));
}
step("then", "table.first_col is {bool_lit}", ["bool_lit"], then_table_first_col_is_value);

function then_table_first_row_is_value(context: any, bool_lit: any): any {
  let actual: any, expected: any;
  actual = context.table_.firstRow;
  expected = $pyEval(bool_lit, $scope);
  $assert($is(actual, expected), $mod("table.first_row is %s", actual));
}
step("then", "table.first_row is {bool_lit}", ["bool_lit"], then_table_first_row_is_value);

function then_table_horz_banding_is_value(context: any, bool_lit: any): any {
  let actual: any, expected: any;
  actual = context.table_.horzBanding;
  expected = $pyEval(bool_lit, $scope);
  $assert($is(actual, expected), $mod("table.horz_banding is %s", actual));
}
step("then", "table.horz_banding is {bool_lit}", ["bool_lit"], then_table_horz_banding_is_value);

function then_table_last_col_is_value(context: any, bool_lit: any): any {
  let actual: any, expected: any;
  actual = context.table_.lastCol;
  expected = $pyEval(bool_lit, $scope);
  $assert($is(actual, expected), $mod("table.last_col is %s", actual));
}
step("then", "table.last_col is {bool_lit}", ["bool_lit"], then_table_last_col_is_value);

function then_table_last_row_is_value(context: any, bool_lit: any): any {
  let actual: any, expected: any;
  actual = context.table_.lastRow;
  expected = $pyEval(bool_lit, $scope);
  $assert($is(actual, expected), $mod("table.last_row is %s", actual));
}
step("then", "table.last_row is {bool_lit}", ["bool_lit"], then_table_last_row_is_value);

function then_table_rows_is_a_type_object(context: any, type_name: any): any {
  let actual: any, expected: any;
  actual = $typename(context.table_.rows);
  expected = type_name;
  $assertEqual(actual, expected, $mod("table.rows is a %s object", actual));
}
step("then", "table.rows is a {type_name} object", ["type_name"], then_table_rows_is_a_type_object);

function then_table_vert_banding_is_value(context: any, bool_lit: any): any {
  let actual: any, expected: any;
  actual = context.table_.vertBanding;
  expected = $pyEval(bool_lit, $scope);
  $assert($is(actual, expected), $mod("table.vert_banding is %s", actual));
}
step("then", "table.vert_banding is {bool_lit}", ["bool_lit"], then_table_vert_banding_is_value);

// the names python's eval() sees in this module
const $scope = { Inches, MSO_ANCHOR, Presentation };
