// Machine-translated from python-pptx features/steps/shapes.py by
// scripts/py2ts_steps.py. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $add, $assert, $assertEqual, $call, $cm1, $d, $dictget, $enter, $enumerate, $exit, $get, $getattr, $index, $int, $is, $isinstance, $kw, $len, $list, $m, $mod, $mul, $next, $open, $range, $read, $t2, $t3, $t4, $truthy, $typename, BytesIO } from "../../pyshim.js";
import { AttributeError, IndexError, KeyError, NotImplementedError, ValueError } from "../../../src/errors.js";
import { step } from "../runner.js";
import { saved_pptx_path, test_file, test_image, test_pptx } from "./helpers.js";
import { CategoryChartData } from "../../../src/chart/data.js";
import { XL_CHART_TYPE } from "../../../src/enum/chart.js";
import { MSO_CONNECTOR, MSO_SHAPE, PP_PLACEHOLDER, PROG_ID } from "../../../src/enum/shapes.js";
import { Presentation } from "../../../src/index.js";
import { BaseShape } from "../../../src/shapes/base.js";
import { Emu, Inches } from "../../../src/util.js";
void $py;


function given_a_BaseShapes_object_as_shapes(context: any): any {
  let prs: any;
  prs = $call(Presentation);
  context.shapes = prs.slides.addSlide($get(prs.slideLayouts, 6)).shapes;
}
step("given", "a _BaseShapes object as shapes", [], given_a_BaseShapes_object_as_shapes);

function given_a_GroupShapes_object_of_length_3_as_shapes(context: any): any {
  let group_shape: any, prs: any;
  prs = $call(Presentation, test_pptx("shp-groupshape"));
  group_shape = $get($get(prs.slides, 0).shapes, 0);
  context.shapes = group_shape.shapes;
}
step("given", "a GroupShapes object of length 3 as shapes", [], given_a_GroupShapes_object_of_length_3_as_shapes);
step("given", "a GroupShapes object as shapes", [], given_a_GroupShapes_object_of_length_3_as_shapes);

function given_a_LayoutPlaceholders_object_of_length_2_as_shapes(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("lyt-shapes"));
  context.shapes = $get(prs.slideLayouts, 0).placeholders;
}
step("given", "a LayoutPlaceholders object of length 2 as shapes", [], given_a_LayoutPlaceholders_object_of_length_2_as_shapes);

function given_a_LayoutShapes_object_of_length_3_as_shapes(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("lyt-shapes"));
  context.shapes = $get(prs.slideLayouts, 0).shapes;
}
step("given", "a LayoutShapes object of length 3 as shapes", [], given_a_LayoutShapes_object_of_length_3_as_shapes);

function given_a_MasterPlaceholders_object_of_length_2_as_shapes(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("mst-placeholders"));
  context.shapes = $get(prs.slideMasters, 0).placeholders;
}
step("given", "a MasterPlaceholders object of length 2 as shapes", [], given_a_MasterPlaceholders_object_of_length_2_as_shapes);

function given_a_MasterShapes_object_of_length_2_as_shapes(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("mst-shapes"));
  context.shapes = $get(prs.slideMasters, 0).shapes;
}
step("given", "a MasterShapes object of length 2 as shapes", [], given_a_MasterShapes_object_of_length_2_as_shapes);

function given_a_PROG_ID_member_file_as_ole_object_file(context: any, PROG_ID_member: any): any {
  let f: any, filename: any;
  filename = $get($d({DOCX: "shp-embedded-docx.docx", PPTX: "shp-embedded-pptx.pptx", XLSX: "shp-embedded-xlsx.xlsx"}), PROG_ID_member);
  const $cm1 = $open(test_file(filename), "rb");
  f = $enter($cm1);
  try {
    context.oleObjectFile = new BytesIO($read(f));
  } finally {
    $exit($cm1);
  }
  context.PROG_ID_member = PROG_ID_member;
}
step("given", "a {PROG_ID_member} file as ole_object_file", ["PROG_ID_member"], given_a_PROG_ID_member_file_as_ole_object_file);

function given_a_SlidePlaceholders_object_of_length_2_as_shapes(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("shp-shapes"));
  context.shapes = $get(prs.slides, 0).placeholders;
}
step("given", "a SlidePlaceholders object of length 2 as shapes", [], given_a_SlidePlaceholders_object_of_length_2_as_shapes);

function given_a_SlideShapes_object_as_shapes(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("shp-shapes"));
  context.shapes = $get(prs.slides, 0).shapes;
}
step("given", "a SlideShapes object as shapes", [], given_a_SlideShapes_object_as_shapes);

function given_a_SlideShapes_object_containing_a_or_no_movies(context: any, a_or_no: any): any {
  let pptx: any, prs: any;
  pptx = $get($d({"one or more": "shp-movie-props", no: "shp-shapes"}), a_or_no);
  prs = $call(Presentation, test_pptx(pptx));
  context.prs = prs;
  context.shapes = $get(prs.slides, 0).shapes;
}
step("given", "a SlideShapes object containing {a_or_no} movies", ["a_or_no"], given_a_SlideShapes_object_containing_a_or_no_movies);

function given_a_SlideShapes_object_of_length_6_as_shapes(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("shp-shapes"));
  context.shapes = $get(prs.slides, 0).shapes;
}
step("given", "a SlideShapes object of length 6 shapes as shapes", [], given_a_SlideShapes_object_of_length_6_as_shapes);

function given_a_SlideShapes_obj_having_type_shape_at_off_idx(context: any, type: any, idx: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("shp-shapes"));
  context.shapes = $get(prs.slides, 1).shapes;
}
step("given", "a SlideShapes object having a {type} shape at offset {idx}", ["type", "idx"], given_a_SlideShapes_obj_having_type_shape_at_off_idx);

function when_I_add_100_shapes(context: any): any {
  let CX: any, CY: any, X_INCR: any, X_ORIG: any, Y_INCR: any, Y_ORIG: any, corners: any, shapes: any, x: any, y: any;
  const $t2 = $call(Inches, 0.0625);
  X_ORIG = $t2;
  Y_ORIG = $t2;
  const $t3 = $call(Inches, 0.5);
  X_INCR = $t3;
  Y_INCR = $t3;
  const $t4 = $call(Inches, 0.375);
  CX = $t4;
  CY = $t4;
  const iter_corner = function* (): any {
    let x: any, y: any;
    y = Y_ORIG;
    while ($truthy(true)) {
      for (const i of $list($range(20))) {
        x = $add(X_ORIG, $mul(X_INCR, i));
        yield [x, y];
      }
      y = $add(y, Y_INCR);
    }
  };
  shapes = context.shapes;
  corners = iter_corner();
  for (const i of $list($range(100))) {
    [x, y] = $list($next(corners));
    shapes.addShape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, CX, CY);
  }
}
step("when", "I add 100 shapes", [], when_I_add_100_shapes);

function when_I_call_shapes_add_table(context: any): any {
  let cx: any, cy: any, shapes: any, x: any, y: any;
  shapes = context.slide.shapes;
  [x, y] = $list([$call(Inches, 1.0), $call(Inches, 2.0)]);
  [cx, cy] = $list([$call(Inches, 3.0), $call(Inches, 1.0)]);
  shapes.addTable(2, 2, x, y, cx, cy);
}
step("when", "I add a table to the slide's shape collection", [], when_I_call_shapes_add_table);

function when_I_assign_shape_ole_format_to_ole_format(context: any): any {
  context.oleFormat = context.shape.oleFormat;
}
step("when", "I assign shape.ole_format to ole_format", [], when_I_assign_shape_ole_format_to_ole_format);

function when_I_assign_shapes_add_chart_to_shape(context: any): any {
  let chart_data: any;
  chart_data = $call(CategoryChartData);
  chart_data.categories = ["Foo", "Bar"];
  chart_data.addSeries("East", [1.0, 2.0]);
  chart_data.addSeries("West", [3.0, 4.0]);
  context.shape = context.shapes.addChart(XL_CHART_TYPE.COLUMN_CLUSTERED, $call(Inches, 1), $call(Inches, 1), $call(Inches, 8), $call(Inches, 5), chart_data);
}
step("when", "I assign shapes.add_chart() to shape", [], when_I_assign_shapes_add_chart_to_shape);

function when_I_assign_shapes_add_connector_to_shape(context: any): any {
  context.shape = context.shapes.addConnector(MSO_CONNECTOR.CURVE, 4, 3, 2, 1);
}
step("when", "I assign shapes.add_connector() to shape", [], when_I_assign_shapes_add_connector_to_shape);

function when_I_assign_shapes_add_group_shape_to_shape(context: any): any {
  context.shape = context.shapes.addGroupShape();
}
step("when", "I assign shapes.add_group_shape() to shape", [], when_I_assign_shapes_add_group_shape_to_shape);

function when_I_assign_shapes_add_ole_object_to_shape(context: any): any {
  context.shape = context.shapes.addOleObject(context.oleObjectFile, $getattr(PROG_ID, context.PROG_ID_member), 4, 3, 2, 1);
}
step("when", "I assign shapes.add_ole_object(ole_object_file) to shape", [], when_I_assign_shapes_add_ole_object_to_shape);

function when_I_assign_shapes_add_picture_to_shape(context: any): any {
  context.shape = context.shapes.addPicture(test_image("sonic.gif"), $call(Inches, 1), $call(Inches, 2));
}
step("when", "I assign shapes.add_picture() to shape", [], when_I_assign_shapes_add_picture_to_shape);

function when_I_assign_shapes_add_shape_to_shape(context: any): any {
  context.shape = context.shapes.addShape(MSO_SHAPE.ROUNDED_RECTANGLE, $call(Inches, 2), $call(Inches, 3), $call(Inches, 1), $call(Inches, 0.5));
}
step("when", "I assign shapes.add_shape() to shape", [], when_I_assign_shapes_add_shape_to_shape);

function when_I_assign_shapes_add_textbox_to_shape(context: any): any {
  context.shape = context.shapes.addTextbox($call(Inches, 1), $call(Inches, 2), $call(Inches, 3), $call(Inches, 0.5));
}
step("when", "I assign shapes.add_textbox() to shape", [], when_I_assign_shapes_add_textbox_to_shape);

function when_I_assign_shapes_build_freeform_to_builder(context: any): any {
  let builder: any, shapes: any;
  shapes = context.shapes;
  builder = shapes.buildFreeform();
  context.builder = builder;
}
step("when", "I assign shapes.build_freeform() to builder", [], when_I_assign_shapes_build_freeform_to_builder);

function when_I_assign_shapes_build_freeform_scale_to_builder(context: any): any {
  let builder: any, shapes: any;
  shapes = context.shapes;
  builder = $m(shapes, "buildFreeform", $kw({scale: 100.0}));
  context.builder = builder;
}
step("when", "I assign shapes.build_freeform(scale=100.0) to builder", [], when_I_assign_shapes_build_freeform_scale_to_builder);

function when_I_assign_shapes_build_freeform_scale_rectnglr_to_builder(context: any): any {
  let builder: any, shapes: any;
  shapes = context.shapes;
  builder = $m(shapes, "buildFreeform", $kw({scale: [200.0, 100.0]}));
  context.builder = builder;
}
step("when", "I assign shapes.build_freeform(scale=(200.0, 100.0)) to builder", [], when_I_assign_shapes_build_freeform_scale_rectnglr_to_builder);

function when_I_assign_shapes_build_freeform_start_x_start_y_to_builder(context: any): any {
  let builder: any, shapes: any;
  shapes = context.shapes;
  builder = shapes.buildFreeform(25, 125);
  context.builder = builder;
}
step("when", "I assign shapes.build_freeform(start_x=25, start_y=125) to builder", [], when_I_assign_shapes_build_freeform_start_x_start_y_to_builder);

function when_I_assign_True_to_shapes_turbo_add_enabled(context: any): any {
  context.shapes.turboAddEnabled = true;
}
step("when", "I assign True to shapes.turbo_add_enabled", [], when_I_assign_True_to_shapes_turbo_add_enabled);

function when_I_call_shapes_add_chart(context: any, type_: any): any {
  let chart_type: any;
  chart_type = $getattr(XL_CHART_TYPE, type_);
  context.chart = context.shapes.addChart(chart_type, 0, 0, 0, 0, context.chartData).chart;
}
step("when", "I call shapes.add_chart({type_}, chart_data)", ["type_"], when_I_call_shapes_add_chart);

function when_I_call_shapes_add_connector(context: any): any {
  context.connector = context.shapes.addConnector(MSO_CONNECTOR.STRAIGHT, 1, 2, 3, 4);
}
step("when", "I call shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 1, 2, 3, 4)", [], when_I_call_shapes_add_connector);

function when_I_call_shapes_add_movie(context: any): any {
  let cx: any, cy: any, shapes: any, x: any, y: any;
  shapes = context.shapes;
  [x, y, cx, cy] = $list([$call(Emu, 2590800), $call(Emu, 571500), $call(Emu, 3962400), $call(Emu, 5715000)]);
  context.movie = shapes.addMovie(test_file("just-two-mice.mp4"), x, y, cx, cy, test_file("just-two-mice.png"));
}
step("when", "I call shapes.add_movie(file, x, y, cx, cy, poster_frame)", [], when_I_call_shapes_add_movie);

function then_iterating_shapes_produces_count_objects_of_type_class_name(context: any, count: any, class_name: any): any {
  let actual_class_name: any, actual_count: any, expected_class_name: any, expected_count: any, idx: any, shapes: any;
  shapes = context.shapes;
  [expected_count, expected_class_name] = $list([$int(count), class_name]);
  idx = -1;
  let shape: any;
  for ([idx, shape] of $list($enumerate(shapes))) {
    actual_class_name = $typename(shape);
    $assertEqual(actual_class_name, expected_class_name, $mod("shape.__class__.__name__ == %s", actual_class_name));
  }
  actual_count = (idx + 1);
  $assertEqual(actual_count, expected_count, $mod("got %d items", actual_count));
}
step("then", "iterating shapes produces {count} objects of type {class_name}", ["count", "class_name"], then_iterating_shapes_produces_count_objects_of_type_class_name);

function then_iterating_shapes_produces_count_objects_that_subclass_BaseShape(context: any, count: any): any {
  let actual_count: any, class_name: any, expected_count: any, idx: any, shapes: any;
  shapes = context.shapes;
  expected_count = $int(count);
  idx = -1;
  let shape: any;
  for ([idx, shape] of $list($enumerate(shapes))) {
    class_name = $typename(shape);
    $assert($isinstance(shape, BaseShape), $mod("%s does not subclass BaseShape", class_name));
  }
  actual_count = (idx + 1);
  $assertEqual(actual_count, expected_count, $mod("got %d items", actual_count));
}
step("then", "iterating shapes produces {count} objects that subclass BaseShape", ["count"], then_iterating_shapes_produces_count_objects_that_subclass_BaseShape);

function then_len_shapes_eq_value(context: any, value: any): any {
  let actual_len: any, expected_len: any;
  expected_len = $int(value);
  actual_len = $len(context.shapes);
  $assertEqual(actual_len, expected_len, $mod("len(shapes) == %s", actual_len));
}
step("then", "len(shapes) == {value}", ["value"], then_len_shapes_eq_value);

function then_shape_is_a_type_object(context: any, clsname: any): any {
  let actual_class_name: any, expected_class_name: any;
  actual_class_name = $typename(context.shape);
  expected_class_name = clsname;
  $assertEqual(actual_class_name, expected_class_name, $mod("shape is a %s object", actual_class_name));
}
step("then", "shape is a {clsname} object", ["clsname"], then_shape_is_a_type_object);

function then_shapes_minus_1_eq_shape(context: any): any {
  let shape: any, shapes: any;
  [shapes, shape] = $list([context.shapes, context.shape]);
  $assertEqual($get(shapes, -1), shape);
}
step("then", "shapes[-1] == shape", [], then_shapes_minus_1_eq_shape);

function then_shapes_idx_is_a_type_object(context: any, idx: any, type_: any): any {
  let shapes: any, type_name: any;
  shapes = context.shapes;
  type_name = $typename($get(shapes, $int(idx)));
  $assertEqual(type_name, type_, $mod("got %s", type_name));
}
step("then", "shapes[{idx}] is a {type_} object", ["idx", "type_"], then_shapes_idx_is_a_type_object);

function then_shapes_get_10_is_the_body_placeholder(context: any): any {
  let body_placeholder: any, shapes: any, title_placeholder: any;
  shapes = context.shapes;
  title_placeholder = $dictget(shapes, $kw({idx: 0}));
  body_placeholder = $dictget(shapes, $kw({idx: 10}));
  $assert($is(title_placeholder._element, $get(shapes, 0)._element));
  $assert($is(body_placeholder._element, $get(shapes, 1)._element));
}
step("then", "shapes.get(idx=10) is the body placeholder", [], then_shapes_get_10_is_the_body_placeholder);

function then_shapes_get_PP_PLACEHOLDER_BODY_is_the_body_ph(context: any): any {
  let body_placeholder: any, shapes: any, title_placeholder: any;
  shapes = context.shapes;
  title_placeholder = $dictget(shapes, PP_PLACEHOLDER.TITLE);
  body_placeholder = $dictget(shapes, PP_PLACEHOLDER.BODY);
  $assert($is(title_placeholder._element, $get(shapes, 0)._element));
  $assert($is(body_placeholder._element, $get(shapes, 1)._element));
}
step("then", "shapes.get(PP_PLACEHOLDER.BODY) is the body placeholder", [], then_shapes_get_PP_PLACEHOLDER_BODY_is_the_body_ph);

function then_shapes_index_for_each_shape_matches_sequence_position(context: any): any {
  let shapes: any;
  shapes = context.shapes;
  for (const [idx, shape] of $list($enumerate(shapes))) {
    $assertEqual(idx, $index(shapes, shape), $mod("index doesn't match for idx == %s", idx));
  }
}
step("then", "shapes.index(shape) for each shape matches its sequence position", [], then_shapes_index_for_each_shape_matches_sequence_position);

function then_shapes_title_is_the_title_placeholder(context: any): any {
  let shapes: any, title_placeholder: any;
  shapes = context.shapes;
  title_placeholder = shapes.title;
  $assert($is(title_placeholder.element, $get(shapes, 0).element));
  $assertEqual(title_placeholder.shapeId, 4);
}
step("then", "shapes.title is the title placeholder", [], then_shapes_title_is_the_title_placeholder);

function then_shapes_turbo_add_enabled_is_False(context: any): any {
  let shapes: any;
  shapes = context.shapes;
  $assert($is(shapes.turboAddEnabled, false));
}
step("then", "shapes.turbo_add_enabled is False", [], then_shapes_turbo_add_enabled_is_False);

function then_the_table_appears_in_the_slide(context: any): any {
  let expected_table_graphic_frame: any, prs: any;
  prs = $call(Presentation, saved_pptx_path);
  expected_table_graphic_frame = $get($get(prs.slides, 0).shapes, 0);
  $assert(expected_table_graphic_frame.hasTable);
}
step("then", "the table appears in the slide", [], then_the_table_appears_in_the_slide);

// the names python's eval() sees in this module
const $scope = { BaseShape, CategoryChartData, Emu, Inches, MSO_CONNECTOR, MSO_SHAPE, PP_PLACEHOLDER, PROG_ID, Presentation, XL_CHART_TYPE };
