// Machine-translated from python-pptx features/steps/placeholder.py by
// scripts/py2ts_steps.py. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $abs, $assert, $assertEqual, $call, $cm1, $cmp, $d, $enter, $exit, $float, $get, $getattr, $hashlib, $index, $int, $is, $isinstance, $len, $list, $mod, $open, $read, $split, $typename } from "../../pyshim.js";
import { AttributeError, IndexError, KeyError, NotImplementedError, ValueError } from "../../../src/errors.js";
import { step } from "../runner.js";
import { saved_pptx_path, test_file, test_pptx, test_text } from "./helpers.js";
import { CategoryChartData } from "../../../src/chart/data.js";
import { XL_CHART_TYPE } from "../../../src/enum/chart.js";
import { MSO_SHAPE_TYPE, PP_PLACEHOLDER } from "../../../src/enum/shapes.js";
import { Presentation } from "../../../src/index.js";
import { _PlaceholderFormat } from "../../../src/shapes/base.js";
void $py;


function given_a_bullet_body_placeholder(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("ph-unpopulated-placeholders"));
  context.prs = prs;
  context.sld = $get(prs.slides, 2);
  context.body = $get($get(prs.slides, 2).shapes.placeholders, 10);
}
step("given", "a bullet body placeholder", [], given_a_bullet_body_placeholder);

function given_a_known_placeholder_shape(context: any, placeholder_type: any): any {
  context.executeSteps($mod("given an unpopulated %s placeholder shape", placeholder_type));
}
step("given", "a known {placeholder_type} placeholder shape", ["placeholder_type"], given_a_known_placeholder_shape);

function given_layout_placeholder_with_directly_set_pos_and_size(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("ph-inherit-props"));
  context.placeholder = $get($get(prs.slideLayouts, 0).placeholders, 1);
}
step("given", "a layout placeholder having directly set position and size", [], given_layout_placeholder_with_directly_set_pos_and_size);

function given_layout_placeholder_with_no_direct_pos_or_size_settings(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("ph-inherit-props"));
  context.placeholder = $get($get(prs.slideLayouts, 0).placeholders, 0);
}
step("given", "a layout placeholder having no direct position or size settings", [], given_layout_placeholder_with_no_direct_pos_or_size_settings);

function given_a_master_placeholder(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("ph-inherit-props"));
  context.masterPlaceholder = $get(prs.slideMaster.placeholders, 1);
}
step("given", "a master placeholder", [], given_a_master_placeholder);

function given_notes_slide_placeholder_having_directly_set_pos_and_size(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("ph-inherit-props"));
  context.placeholder = $get($get(prs.slides, 1).notesSlide.placeholders, 1);
}
step("given", "a notes slide placeholder having directly set position and size", [], given_notes_slide_placeholder_having_directly_set_pos_and_size);

function given_notes_slide_placeholder_having_no_direct_pos_or_size(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("ph-inherit-props"));
  context.placeholder = $get($get(prs.slides, 0).notesSlide.placeholders, 1);
}
step("given", "a notes slide placeholder having no direct position or size settings", [], given_notes_slide_placeholder_having_no_direct_pos_or_size);

function given_slide_placeholder_with_directly_set_pos_and_size(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("ph-inherit-props"));
  context.placeholder = $get($get(prs.slides, 0).placeholders, 10);
}
step("given", "a slide placeholder having directly set position and size", [], given_slide_placeholder_with_directly_set_pos_and_size);

function given_slide_placeholder_with_no_direct_pos_or_size_settings(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("ph-inherit-props"));
  context.placeholder = $get($get(prs.slides, 0).placeholders, 0);
}
step("given", "a slide placeholder having no direct position or size settings", [], given_slide_placeholder_with_no_direct_pos_or_size_settings);

function given_a_slide_with_an_unpopulated_type_placeholder(context: any, type_: any): any {
  let prs: any, slide_idx: any;
  slide_idx = $index(["title", "content", "text", "chart", "table", "smart art", "media", "clip art", "picture"], type_);
  prs = $call(Presentation, test_pptx("ph-unpopulated-placeholders"));
  context.shape = $get($get(prs.slides, slide_idx).shapes, 0);
}
step("given", "a slide with an unpopulated {type_} placeholder", ["type_"], given_a_slide_with_an_unpopulated_type_placeholder);

function given_a_slide_with_a_type_ph_with_content(context: any, type_: any, content: any): any {
  let prs: any, slide_idx: any;
  slide_idx = $index(["picture", "clip art", "table", "chart", "title", "content", "text", "smart art", "media"], type_);
  prs = $call(Presentation, test_pptx("ph-populated-placeholders"));
  context.shape = $get($get(prs.slides, slide_idx).shapes, 0);
}
step("given", "a slide with a {type_} placeholder populated with {content}", ["type_", "content"], given_a_slide_with_a_type_ph_with_content);

function given_an_unpopulated_placeholder_shape(context: any, placeholder_type: any): any {
  let prs: any, slide_idx: any;
  slide_idx = $index(["title", "content", "text", "chart", "table", "smart art", "media", "clip art", "picture"], placeholder_type);
  prs = $call(Presentation, test_pptx("ph-unpopulated-placeholders"));
  context.shape = $get($get(prs.slides, slide_idx).shapes, 0);
}
step("given", "an unpopulated {placeholder_type} placeholder shape", ["placeholder_type"], given_an_unpopulated_placeholder_shape);

function when_I_call_placeholder_insert_chart(context: any): any {
  let chart_data: any, placeholder: any;
  chart_data = $call(CategoryChartData);
  chart_data.categories = ["Yes", "No"];
  chart_data.addSeries("Series 1", [42, 24]);
  placeholder = context.shape;
  context.placeholder = placeholder.insertChart(XL_CHART_TYPE.PIE, chart_data);
}
step("when", "I call placeholder.insert_chart(XL_CHART_TYPE.PIE, chart_data)", [], when_I_call_placeholder_insert_chart);

function when_I_call_placeholder_insert_picture(context: any, filename: any): any {
  let f: any, path: any, placeholder: any;
  placeholder = context.shape;
  path = test_file(filename);
  const $cm1 = $open(path, "rb");
  f = $enter($cm1);
  try {
    context.imageSha1 = $hashlib.sha1($read(f)).hexdigest();
  } finally {
    $exit($cm1);
  }
  context.placeholder = placeholder.insertPicture(path);
}
step("when", "I call placeholder.insert_picture('{filename}')", ["filename"], when_I_call_placeholder_insert_picture);

function when_I_call_placeholder_insert_table(context: any): any {
  let placeholder: any;
  placeholder = context.shape;
  context.placeholder = placeholder.insertTable(2, 3);
}
step("when", "I call placeholder.insert_table(rows=2, cols=3)", [], when_I_call_placeholder_insert_table);

function when_I_indent_the_first_paragraph(context: any): any {
  $get(context.body.textFrame.paragraphs, 0).level = 1;
}
step("when", "I indent the first paragraph", [], when_I_indent_the_first_paragraph);

function step_when_set_slide_title_text(context: any): any {
  context.slide.shapes.title.text = test_text;
}
step("when", "I set the title text of the slide", [], step_when_set_slide_title_text);

function then_I_can_get_the_placeholder_dimensions(context: any): any {
  let placeholder: any;
  placeholder = context.masterPlaceholder;
  $assertEqual(placeholder.width, 6923112, $mod("got %d", placeholder.width));
  $assertEqual(placeholder.height, 3484984, $mod("got %d", placeholder.height));
}
step("then", "I can get the placeholder dimensions", [], then_I_can_get_the_placeholder_dimensions);

function then_I_can_get_the_placeholder_position(context: any): any {
  let placeholder: any;
  placeholder = context.masterPlaceholder;
  $assertEqual(placeholder.left, 1110444, $mod("got %d", placeholder.left));
  $assertEqual(placeholder.top, 1686508, $mod("got %d", placeholder.top));
}
step("then", "I can get the placeholder position", [], then_I_can_get_the_placeholder_position);

function then_I_get_direct_settings_when_query_pos_and_size(context: any): any {
  let placeholder: any;
  placeholder = context.placeholder;
  $assertEqual(placeholder.left, 468312, $mod("got %s", placeholder.left));
  $assertEqual(placeholder.top, 1700212, $mod("got %s", placeholder.top));
  $assertEqual(placeholder.width, 8208143, $mod("got %s", placeholder.width));
  $assertEqual(placeholder.height, 4537099, $mod("got %s", placeholder.height));
}
step("then", "I get the direct settings when I query position and size", [], then_I_get_direct_settings_when_query_pos_and_size);

function then_I_get_inherited_settings_when_I_query_position_and_size(context: any): any {
  let placeholder: any;
  placeholder = context.placeholder;
  $assertEqual(placeholder.left, 457200, $mod("got %s", placeholder.left));
  $assertEqual(placeholder.top, 274638, $mod("got %s", placeholder.top));
  $assertEqual(placeholder.width, 8229600, $mod("got %s", placeholder.width));
  $assertEqual(placeholder.height, 1143000, $mod("got %s", placeholder.height));
}
step("then", "I get inherited settings when I query position and size", [], then_I_get_inherited_settings_when_I_query_position_and_size);

function then_placeholder_format_idx_is_value(context: any, value: any): any {
  let expected_value: any, placeholder_format: any;
  expected_value = $int(value);
  placeholder_format = context.shape.placeholderFormat;
  $assertEqual(placeholder_format.idx, expected_value);
}
step("then", "placeholder_format.idx is {value}", ["value"], then_placeholder_format_idx_is_value);

function then_placeholder_format_type_is_value(context: any, value: any): any {
  let expected_value: any, placeholder_format: any;
  expected_value = $getattr(PP_PLACEHOLDER, $get($split(value, "."), 1));
  placeholder_format = context.shape.placeholderFormat;
  $assertEqual(placeholder_format.type, expected_value);
}
step("then", "placeholder_format.type is {value}", ["value"], then_placeholder_format_type_is_value);

function then_shape_placeholder_format_is_its_PlaceholderFormat_object(context: any): any {
  let placeholder_format: any, shape: any;
  shape = context.shape;
  placeholder_format = shape.placeholderFormat;
  $assert($isinstance(placeholder_format, _PlaceholderFormat));
  $assert($is(placeholder_format.element, shape.element.ph));
}
step("then", "shape.placeholder_format is its _PlaceholderFormat object", [], then_shape_placeholder_format_is_its_PlaceholderFormat_object);

function then_shape_shape_type_is_PLACEHOLDER(context: any): any {
  let shape: any;
  shape = context.shape;
  $assertEqual(shape.shapeType, MSO_SHAPE_TYPE.PLACEHOLDER);
}
step("then", "shape.shape_type is MSO_SHAPE_TYPE.PLACEHOLDER", [], then_shape_shape_type_is_PLACEHOLDER);

function then_slide_shapes_0_is_a_cls_proxy_for_that_placeholder(context: any, cls: any): any {
  let clsname: any, placeholder: any;
  placeholder = context.shape;
  clsname = $typename(placeholder);
  $assertEqual(clsname, cls, $mod("got %s", clsname));
}
step("then", "slide.shapes[0] is a {cls} proxy object for that placeholder", ["cls"], then_slide_shapes_0_is_a_cls_proxy_for_that_placeholder);

function then_the_chart_is_a_pie_chart(context: any): any {
  let chart: any;
  chart = context.chart;
  $assertEqual(chart.chartType, XL_CHART_TYPE.PIE);
}
step("then", "the chart is a pie chart", [], then_the_chart_is_a_pie_chart);

function then_the_return_value_is_a_PlaceholderType_object(context: any, type: any): any {
  let expected_type_name: any, placeholder_type_name: any;
  expected_type_name = $mod("Placeholder%s", type);
  placeholder_type_name = $typename(context.placeholder);
  $assertEqual(placeholder_type_name, expected_type_name);
}
step("then", "the return value is a Placeholder{type} object", ["type"], then_the_return_value_is_a_PlaceholderType_object);

function then_the_paragraph_is_indented(context: any): any {
  let p: any, prs: any;
  prs = $call(Presentation, saved_pptx_path);
  p = $get($get($get(prs.slides, 2).shapes.placeholders, 10).textFrame.paragraphs, 0);
  $assertEqual(p.level, 1);
}
step("then", "the paragraph is indented", [], then_the_paragraph_is_indented);

function then_the_placeholder_contains_the_chart(context: any): any {
  let placeholder_graphic_frame: any;
  placeholder_graphic_frame = context.placeholder;
  $assert(placeholder_graphic_frame.hasChart);
  context.chart = placeholder_graphic_frame.chart;
}
step("then", "the placeholder contains the chart", [], then_the_placeholder_contains_the_chart);

function then_the_placeholder_contains_the_image(context: any): any {
  let placeholder_picture: any;
  placeholder_picture = context.placeholder;
  $assertEqual(placeholder_picture.image.sha1, context.imageSha1);
}
step("then", "the placeholder contains the image", [], then_the_placeholder_contains_the_image);

function then_the_placeholder_contains_the_table(context: any): any {
  let placeholder_graphic_frame: any;
  placeholder_graphic_frame = context.placeholder;
  $assert(placeholder_graphic_frame.hasTable);
  context.table_ = placeholder_graphic_frame.table;
}
step("then", "the placeholder contains the table", [], then_the_placeholder_contains_the_table);

function then_the_placeholders_position_and_size_are_inherited(context: any): any {
  let expected_values: any, placeholder: any, value: any;
  placeholder = context.shape;
  expected_values = [["left", 2743200], ["top", 2057400], ["width", 3657600], ["height", 2743200]];
  for (const [prop_name, expected_value] of $list(expected_values)) {
    value = $getattr(placeholder, prop_name);
    $assertEqual(value, expected_value, $mod("got %s", value));
  }
}
step("then", "the placeholder's position and size are inherited from its layout", [], then_the_placeholders_position_and_size_are_inherited);

function then_the_sides_crop_is_value(context: any, sides: any, value: any): any {
  let difference: any, expected_value: any, placeholder_picture: any, side_prop_names: any;
  side_prop_names = $get($d({"top and bottom": ["crop_top", "crop_bottom"], "left and right": ["crop_left", "crop_right"]}), sides);
  expected_value = $float(value);
  placeholder_picture = context.placeholder;
  for (const prop_name of $list(side_prop_names)) {
    value = $getattr(placeholder_picture, prop_name);
    difference = $abs((expected_value - value));
    $assert($cmp(difference, "<", 2e-06), $mod("got %s for %s", [value, prop_name]));
  }
}
step("then", "the {sides} crop is {value}", ["sides", "value"], then_the_sides_crop_is_value);

function then_the_table_has_2_rows_and_3_columns(context: any): any {
  let table: any;
  table = context.table_;
  $assertEqual($len(table.rows), 2);
  $assertEqual($len(table.columns), 3);
}
step("then", "the table has 2 rows and 3 columns", [], then_the_table_has_2_rows_and_3_columns);

function step_then_text_appears_in_title_placeholder(context: any): any {
  let prs: any, title_shape: any, title_text: any;
  prs = $call(Presentation, saved_pptx_path);
  title_shape = $get(prs.slides, 0).shapes.title;
  title_text = $get($get(title_shape.textFrame.paragraphs, 0).runs, 0).text;
  $assertEqual(title_text, test_text);
}
step("then", "the text appears in the title placeholder", [], step_then_text_appears_in_title_placeholder);

// the names python's eval() sees in this module
const $scope = { CategoryChartData, MSO_SHAPE_TYPE, PP_PLACEHOLDER, Presentation, XL_CHART_TYPE, _PlaceholderFormat };
