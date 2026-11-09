// Machine-translated from python-pptx features/steps/shape.py by
// scripts/py2ts_steps.py. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $assert, $assertEqual, $call, $cm1, $d, $enter, $eq, $exit, $float, $get, $getattr, $hashlib, $in, $int, $is, $isinstance, $len, $list, $mod, $open, $pyEval, $read, $round, $set, $setattr, $truthy, $typename } from "../../pyshim.js";
import { AttributeError, IndexError, KeyError, NotImplementedError, ValueError } from "../../../src/errors.js";
import { step } from "../runner.js";
import { cls_qname, test_file, test_pptx } from "./helpers.js";
import { ActionSetting } from "../../../src/action.js";
import { MSO_SHAPE, MSO_SHAPE_TYPE, PP_MEDIA_TYPE } from "../../../src/enum/shapes.js";
import { Presentation } from "../../../src/index.js";
import { Emu } from "../../../src/util.js";
void $py;


function given_an_autoshape(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("shp-autoshape-adjustments"));
  context.shape = $get($get(prs.slides, 0).shapes, 0);
}
step("given", "an autoshape", [], given_an_autoshape);

function given_builder_start_x_builder_start_y_is_x_y(context: any, x_str: any, y_str: any): any {
  let builder: any;
  builder = context.builder;
  [builder._startX, builder._startY] = $list([$int(x_str), $int(y_str)]);
}
step("given", "(builder._start_x, builder._start_y) is ({x_str}, {y_str})", ["x_str", "y_str"], given_builder_start_x_builder_start_y_is_x_y);

function given_builder_x_scale_builder_y_scale_is_p_q(context: any, p_str: any, q_str: any): any {
  let builder: any;
  builder = context.builder;
  [builder._xScale, builder._yScale] = $list([$float(p_str), $float(q_str)]);
}
step("given", "(builder._x_scale, builder._y_scale) is ({p_str}, {q_str})", ["p_str", "q_str"], given_builder_x_scale_builder_y_scale_is_p_q);

function given_a_chevron_shape(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("shp-autoshape-adjustments"));
  context.shape = $get($get(prs.slides, 0).shapes, 0);
}
step("given", "a chevron shape", [], given_a_chevron_shape);

function given_a_Connector_object_as_shape(context: any): any {
  let prs: any, sld: any;
  prs = $call(Presentation, test_pptx("shp-common-props"));
  sld = $get(prs.slides, 0);
  context.shape = $get(sld.shapes, 4);
}
step("given", "a Connector object as shape", [], given_a_Connector_object_as_shape);

function given_a_connector_and_a_1_inch_square_picture_at_0_0(context: any): any {
  let prs: any, shapes: any;
  prs = $call(Presentation, test_pptx("shp-connector-props"));
  shapes = $get(prs.slides, 1).shapes;
  context.picture = $get(shapes, 0);
  context.connector = $get(shapes, 1);
}
step("given", "a connector and a 1 inch square picture at 0, 0", [], given_a_connector_and_a_1_inch_square_picture_at_0_0);

function given_a_connector_having_its_begin_point_at_x_y(context: any, x: any, y: any): any {
  let prs: any, sld: any;
  prs = $call(Presentation, test_pptx("shp-connector-props"));
  sld = $get(prs.slides, 0);
  context.connector = $get(sld.shapes, 0);
}
step("given", "a connector having its begin point at ({x}, {y})", ["x", "y"], given_a_connector_having_its_begin_point_at_x_y);

function given_a_connector_having_its_end_point_at_x_y(context: any, x: any, y: any): any {
  let prs: any, sld: any;
  prs = $call(Presentation, test_pptx("shp-connector-props"));
  sld = $get(prs.slides, 0);
  context.connector = $get(sld.shapes, 0);
}
step("given", "a connector having its end point at ({x}, {y})", ["x", "y"], given_a_connector_having_its_end_point_at_x_y);

function given_an_empty_GroupShape_object_as_shape(context: any): any {
  let prs: any, sld: any;
  prs = $call(Presentation, test_pptx("shp-common-props"));
  sld = $get(prs.slides, 0);
  context.shape = sld.shapes.addGroupShape();
}
step("given", "an empty GroupShape object as shape", [], given_an_empty_GroupShape_object_as_shape);

function given_a_FreeformBuilder_object_as_builder(context: any): any {
  let builder: any, shapes: any;
  shapes = $get($call(Presentation, test_pptx("shp-freeform")).slides, 0).shapes;
  builder = shapes.buildFreeform();
  context.builder = builder;
}
step("given", "a FreeformBuilder object as builder", [], given_a_FreeformBuilder_object_as_builder);

function given_a_GraphicFrame_object_as_shape(context: any): any {
  let prs: any, sld: any;
  prs = $call(Presentation, test_pptx("shp-common-props"));
  sld = $get(prs.slides, 0);
  context.shape = $get(sld.shapes, 2);
}
step("given", "a GraphicFrame object as shape", [], given_a_GraphicFrame_object_as_shape);

function given_a_GraphicFrame_object_containing_a_chart_as_shape(context: any): any {
  let prs: any, sld: any;
  prs = $call(Presentation, test_pptx("shp-access-chart"));
  sld = $get(prs.slides, 0);
  context.shape = $get(sld.shapes, 0);
}
step("given", "a GraphicFrame object containing a chart as shape", [], given_a_GraphicFrame_object_containing_a_chart_as_shape);

function given_a_GraphicFrame_object_containing_a_table_as_shape(context: any): any {
  let prs: any, sld: any;
  prs = $call(Presentation, test_pptx("shp-access-chart"));
  sld = $get(prs.slides, 1);
  context.shape = $get(sld.shapes, 0);
}
step("given", "a GraphicFrame object containing a table as shape", [], given_a_GraphicFrame_object_containing_a_table_as_shape);

function given_a_GraphicFrame_object_containing_an_embedded_xlsx_object_as_shape(context: any): any {
  let prs: any, sld: any;
  prs = $call(Presentation, test_pptx("shp-access-ole-object"));
  sld = $get(prs.slides, 0);
  context.shape = $get(sld.shapes, 0);
}
step("given", "a GraphicFrame object containing an embedded XLSX object as shape", [], given_a_GraphicFrame_object_containing_an_embedded_xlsx_object_as_shape);
step("given", "a GraphicFrame object containing an OLE object as shape", [], given_a_GraphicFrame_object_containing_an_embedded_xlsx_object_as_shape);

function given_a_GroupShape_object_as_group_shape(context: any): any {
  let prs: any, sld: any;
  prs = $call(Presentation, test_pptx("shp-groupshape"));
  sld = $get(prs.slides, 0);
  context.groupShape = $get(sld.shapes, 0);
}
step("given", "a GroupShape object as group_shape", [], given_a_GroupShape_object_as_group_shape);

function given_a_GroupShape_object_as_shape(context: any): any {
  let prs: any, sld: any;
  prs = $call(Presentation, test_pptx("shp-common-props"));
  sld = $get(prs.slides, 0);
  context.shape = $get(sld.shapes, 3);
}
step("given", "a GroupShape object as shape", [], given_a_GroupShape_object_as_shape);

function given_a_movie_shape(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("shp-movie-props"));
  context.movie = $get($get(prs.slides, 0).shapes, 0);
}
step("given", "a movie shape", [], given_a_movie_shape);

function given_an_OleFormat_object_for_an_embedded_XLSX_as_ole_format(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("shp-access-ole-object"));
  context.oleFormat = $get($get(prs.slides, 0).shapes, 0).oleFormat;
}
step("given", "an _OleFormat object for an embedded XLSX as ole_format", [], given_an_OleFormat_object_for_an_embedded_XLSX_as_ole_format);
step("given", "an _OleFormat object for an OLE object as ole_format", [], given_an_OleFormat_object_for_an_embedded_XLSX_as_ole_format);

function given_a_Picture_object_as_picture(context: any): any {
  let slide: any;
  slide = $get($call(Presentation, test_pptx("shp-picture")).slides, 0);
  context.picture = $get(slide.shapes, 0);
}
step("given", "a Picture object as picture", [], given_a_Picture_object_as_picture);

function given_a_Picture_object_as_shape(context: any): any {
  let slide: any;
  slide = $get($call(Presentation, test_pptx("shp-common-props")).slides, 0);
  context.shape = $get(slide.shapes, 1);
}
step("given", "a Picture object as shape", [], given_a_Picture_object_as_shape);

function given_a_Picture_object_with_crop_or_no_as_picture(context: any, crop_or_no: any): any {
  let shape_idx: any, slide: any;
  shape_idx = $get($d({"no cropping": 0, cropping: 1}), crop_or_no);
  slide = $get($call(Presentation, test_pptx("shp-picture")).slides, 0);
  context.picture = $get(slide.shapes, shape_idx);
}
step("given", "a Picture object with {crop_or_no} as picture", ["crop_or_no"], given_a_Picture_object_with_crop_or_no_as_picture);

function given_a_rotated_shape_type_object_as_shape(context: any, shape_type: any): any {
  let prs: any, shape_idx: any, sld: any;
  shape_idx = $get($d({Shape: 0, Picture: 1, GraphicFrame: 2, GroupShape: 3, Connector: 4}), shape_type);
  prs = $call(Presentation, test_pptx("shp-common-props"));
  sld = $get(prs.slides, 1);
  context.shape = $get(sld.shapes, shape_idx);
}
step("given", "a rotated {shape_type} object as shape", ["shape_type"], given_a_rotated_shape_type_object_as_shape);

function given_a_Shape_object_as_shape(context: any): any {
  let prs: any, sld: any;
  prs = $call(Presentation, test_pptx("shp-common-props"));
  sld = $get(prs.slides, 0);
  context.shape = $get(sld.shapes, 0);
}
step("given", "a Shape object as shape", [], given_a_Shape_object_as_shape);

function given_a_Shape_object_having_text_as_shape(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("shp-autoshape-props"));
  context.shape = $get($get(prs.slides, 0).shapes, 0);
}
step("given", "a Shape object having text as shape", [], given_a_Shape_object_having_text_as_shape);

function given_a_shape_on_a_slide(context: any, shape_type: any): any {
  let prs: any, shape_idx: any, sld: any;
  shape_idx = $get($d({Shape: 0, Picture: 1, GraphicFrame: 2, GroupShape: 3, Connector: 4}), shape_type);
  prs = $call(Presentation, test_pptx("shp-common-props"));
  sld = $get(prs.slides, 0);
  context.shape = $get(sld.shapes, shape_idx);
  context.slide = sld;
}
step("given", "a {shape_type} object on a slide as shape", ["shape_type"], given_a_shape_on_a_slide);

function given_a_textbox(context: any): any {
  let prs: any, sld: any;
  prs = $call(Presentation, test_pptx("shp-common-props"));
  sld = $get(prs.slides, 0);
  context.shape = $get(sld.shapes, 5);
}
step("given", "a textbox", [], given_a_textbox);

function given_a_shape_of_known_position_and_size(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("shp-pos-and-size"));
  context.shape = $get($get(prs.slides, 0).shapes, 0);
}
step("given", "a shape of known position and size", [], given_a_shape_of_known_position_and_size);

function when_I_add_a_cx_cy_shape_at_x_y(context: any, cx: any, cy: any, x: any, y: any): any {
  context.shape.shapes.addShape(MSO_SHAPE.ROUNDED_RECTANGLE, $int(x), $int(y), $int(cx), $int(cy));
}
step("when", "I add a {cx} x {cy} shape at ({x}, {y})", ["cx", "cy", "x", "y"], when_I_add_a_cx_cy_shape_at_x_y);

function when_I_assign_to_shape_adjustments(context: any): any {
  $set(context.shape.adjustments, 0, 0.15);
}
step("when", "I assign 0.15 to shape.adjustments[0]", [], when_I_assign_to_shape_adjustments);

function when_I_assign_builder_convert_to_shape_to_shape(context: any): any {
  let builder: any;
  builder = context.builder;
  context.shape = builder.convertToShape();
}
step("when", "I assign builder.convert_to_shape() to shape", [], when_I_assign_builder_convert_to_shape_to_shape);

function when_I_assign_builder_convert_to_shape_origin_x_y(context: any, x_str: any, y_str: any): any {
  let builder: any, origin_x: any, origin_y: any;
  builder = context.builder;
  [origin_x, origin_y] = $list([$int(x_str), $int(y_str)]);
  context.shape = builder.convertToShape(origin_x, origin_y);
}
step("when", "I assign builder.convert_to_shape({x_str}, {y_str}) to shape", ["x_str", "y_str"], when_I_assign_builder_convert_to_shape_origin_x_y);

function when_I_assign_shape_text_eq_value(context: any, value: any): any {
  context.shape.text = $pyEval(value, $scope);
}
step("when", "I assign shape.text = {value}", ["value"], when_I_assign_shape_text_eq_value);

function when_I_assign_value_to_connector_begin_x(context: any, value: any): any {
  context.connector.beginX = $int(value);
}
step("when", "I assign {value} to connector.begin_x", ["value"], when_I_assign_value_to_connector_begin_x);

function when_I_assign_value_to_connector_begin_y(context: any, value: any): any {
  context.connector.beginY = $int(value);
}
step("when", "I assign {value} to connector.begin_y", ["value"], when_I_assign_value_to_connector_begin_y);

function when_I_assign_value_to_connector_end_x(context: any, value: any): any {
  context.connector.endX = $int(value);
}
step("when", "I assign {value} to connector.end_x", ["value"], when_I_assign_value_to_connector_end_x);

function when_I_assign_value_to_connector_end_y(context: any, value: any): any {
  context.connector.endY = $int(value);
}
step("when", "I assign {value} to connector.end_y", ["value"], when_I_assign_value_to_connector_end_y);

function when_I_assign_value_to_picture_crop_side(context: any, value: any, side: any): any {
  let new_value: any;
  new_value = ($truthy($eq(value, "None")) ? null : ($truthy($in(".", value)) ? $float(value) : $int(value)));
  $setattr(context.picture, $mod("crop_%s", side), new_value);
}
step("when", "I assign {value} to picture.crop_{side}", ["value", "side"], when_I_assign_value_to_picture_crop_side);

function when_I_assign_value_to_shape_height(context: any, value: any): any {
  context.shape.height = $int(value);
}
step("when", "I assign {value} to shape.height", ["value"], when_I_assign_value_to_shape_height);

function when_I_assign_value_to_shape_left(context: any, value: any): any {
  context.shape.left = $int(value);
}
step("when", "I assign {value} to shape.left", ["value"], when_I_assign_value_to_shape_left);

function when_I_assign_value_to_shape_name(context: any, value: any): any {
  context.shape.name = value;
}
step("when", "I assign '{value}' to shape.name", ["value"], when_I_assign_value_to_shape_name);

function when_I_assign_value_to_shape_rotation(context: any, value: any): any {
  context.shape.rotation = $float(value);
}
step("when", "I assign {value} to shape.rotation", ["value"], when_I_assign_value_to_shape_rotation);

function when_I_assign_value_to_shape_top(context: any, value: any): any {
  context.shape.top = $int(value);
}
step("when", "I assign {value} to shape.top", ["value"], when_I_assign_value_to_shape_top);

function when_I_assign_value_to_shape_width(context: any, value: any): any {
  context.shape.width = $int(value);
}
step("when", "I assign {value} to shape.width", ["value"], when_I_assign_value_to_shape_width);

function when_I_call_builder_add_line_segments_100_25_25_100(context: any): any {
  let builder: any;
  builder = context.builder;
  builder.addLineSegments([[100, 25], [25, 100]]);
}
step("when", "I call builder.add_line_segments([(100, 25), (25, 100)])", [], when_I_call_builder_add_line_segments_100_25_25_100);

function when_I_call_connector_begin_connect_picture_3(context: any): any {
  let connector: any, picture: any;
  [connector, picture] = $list([context.connector, context.picture]);
  connector.beginConnect(picture, 3);
}
step("when", "I call connector.begin_connect(picture, 3)", [], when_I_call_connector_begin_connect_picture_3);

function when_I_call_connector_end_connect_picture_3(context: any): any {
  let connector: any, picture: any;
  [connector, picture] = $list([context.connector, context.picture]);
  connector.endConnect(picture, 3);
}
step("when", "I call connector.end_connect(picture, 3)", [], when_I_call_connector_end_connect_picture_3);

function then_accessing_shape_click_action_raises_TypeError(context: any): any {
  try {
    context.shape.clickAction;
  } catch (_e: any) {
    return;
  }
  throw new Error("Accessing GroupShape.click_action did not raise");
}
step("then", "accessing shape.click_action raises TypeError", [], then_accessing_shape_click_action_raises_TypeError);

function then_builder_is_a_FreeformBuilder_object(context: any): any {
  let builder: any, class_name: any, expected_value: any;
  builder = context.builder;
  class_name = $typename(builder);
  expected_value = "FreeformBuilder";
  $assertEqual(class_name, expected_value, $mod("Expected class name '%s', got '%s'", [expected_value, class_name]));
}
step("then", "builder is a FreeformBuilder object", [], then_builder_is_a_FreeformBuilder_object);

function then_builder_start_x_builder_start_y_is_x_y(context: any, x_str: any, y_str: any): any {
  let actual_value: any, builder: any, expected_value: any;
  builder = context.builder;
  actual_value = [builder._startX, builder._startY];
  expected_value = [$int(x_str), $int(y_str)];
  $assertEqual(actual_value, expected_value, $mod("Expected %s, got %s", [expected_value, actual_value]));
}
step("then", "(builder._start_x, builder._start_y) is ({x_str}, {y_str})", ["x_str", "y_str"], then_builder_start_x_builder_start_y_is_x_y);

function then_builder_x_scale_builder_y_scale_is_x_y(context: any, p_str: any, q_str: any): any {
  let actual_value: any, builder: any, expected_value: any;
  builder = context.builder;
  actual_value = [builder._xScale, builder._yScale];
  expected_value = [$float(p_str), $float(q_str)];
  $assertEqual(actual_value, expected_value, $mod("Expected %s, got %s", [expected_value, actual_value]));
}
step("then", "(builder._x_scale, builder._y_scale) is ({p_str}, {q_str})", ["p_str", "q_str"], then_builder_x_scale_builder_y_scale_is_x_y);

function then_connector_is_a_Connector_object(context: any): any {
  $assertEqual($typename(context.connector), "Connector");
}
step("then", "connector is a Connector object", [], then_connector_is_a_Connector_object);

function then_connector_begin_x_equals_value(context: any, value: any): any {
  $assertEqual(context.connector.beginX, $int(value));
}
step("then", "connector.begin_x == {value}", ["value"], then_connector_begin_x_equals_value);

function then_connector_begin_x_is_an_Emu_object_with_value_x(context: any, x: any): any {
  let begin_x: any;
  begin_x = context.connector.beginX;
  $assert($isinstance(begin_x, Emu));
  $assertEqual(begin_x, $int(x));
}
step("then", "connector.begin_x is an Emu object with value {x}", ["x"], then_connector_begin_x_is_an_Emu_object_with_value_x);

function then_connector_begin_y_equals_value(context: any, value: any): any {
  $assertEqual(context.connector.beginY, $int(value));
}
step("then", "connector.begin_y == {value}", ["value"], then_connector_begin_y_equals_value);

function then_connector_begin_y_is_an_Emu_object_with_value_y(context: any, y: any): any {
  let begin_y: any;
  begin_y = context.connector.beginY;
  $assert($isinstance(begin_y, Emu));
  $assertEqual(begin_y, $int(y));
}
step("then", "connector.begin_y is an Emu object with value {y}", ["y"], then_connector_begin_y_is_an_Emu_object_with_value_y);

function then_connector_end_x_equals_value(context: any, value: any): any {
  $assertEqual(context.connector.endX, $int(value));
}
step("then", "connector.end_x == {value}", ["value"], then_connector_end_x_equals_value);

function then_connector_end_x_is_an_Emu_object_with_value_x(context: any, x: any): any {
  let end_x: any;
  end_x = context.connector.endX;
  $assert($isinstance(end_x, Emu));
  $assertEqual(end_x, $int(x));
}
step("then", "connector.end_x is an Emu object with value {x}", ["x"], then_connector_end_x_is_an_Emu_object_with_value_x);

function then_connector_end_y_equals_value(context: any, value: any): any {
  $assertEqual(context.connector.endY, $int(value));
}
step("then", "connector.end_y == {value}", ["value"], then_connector_end_y_equals_value);

function then_connector_end_y_is_an_Emu_object_with_value_y(context: any, y: any): any {
  let end_y: any;
  end_y = context.connector.endY;
  $assert($isinstance(end_y, Emu));
  $assertEqual(end_y, $int(y));
}
step("then", "connector.end_y is an Emu object with value {y}", ["y"], then_connector_end_y_is_an_Emu_object_with_value_y);

function then_group_shape_shapes_is_a_GroupShapes_object(context: any): any {
  let class_name: any;
  class_name = $typename(context.groupShape.shapes);
  $assertEqual(class_name, "GroupShapes", $mod("got %s", class_name));
}
step("then", "group_shape.shapes is a GroupShapes object", [], then_group_shape_shapes_is_a_GroupShapes_object);

function then_len_ole_format_blob_eq_value(context: any, value: any): any {
  let actual: any;
  actual = $len(context.oleFormat.blob);
  $assertEqual(actual, $int(value));
}
step("then", "len(ole_format.blob) == {value}", ["value"], then_len_ole_format_blob_eq_value);

function then_movie_is_a_Movie_object(context: any): any {
  let class_name: any;
  class_name = $typename(context.movie);
  $assertEqual(class_name, "Movie", $mod("got %s", class_name));
}
step("then", "movie is a Movie object", [], then_movie_is_a_Movie_object);

function then_movie_left_movie_top_eq_x_y(context: any): any {
  let movie: any, position: any;
  movie = context.movie;
  position = [movie.left, movie.top];
  $assertEqual(position, [$call(Emu, 2590800), $call(Emu, 571500)], $mod("got %s", position));
}
step("then", "movie.left, movie.top == x, y", [], then_movie_left_movie_top_eq_x_y);

function then_movie_media_format_is_a_MediaFormat_object(context: any): any {
  let class_name: any;
  class_name = $typename(context.movie.mediaFormat);
  $assertEqual(class_name, "_MediaFormat", $mod("got %s", class_name));
}
step("then", "movie.media_format is a _MediaFormat object", [], then_movie_media_format_is_a_MediaFormat_object);

function then_movie_media_type_is_PP_MEDIA_TYPE_MOVIE(context: any): any {
  let media_type: any;
  media_type = context.movie.mediaType;
  $assertEqual(media_type, PP_MEDIA_TYPE.MOVIE, $mod("got %s", media_type));
}
step("then", "movie.media_type is PP_MEDIA_TYPE.MOVIE", [], then_movie_media_type_is_PP_MEDIA_TYPE_MOVIE);

function then_movie_poster_frame_is_the_same_image_as_poster_frame(context: any): any {
  let actual_sha1: any, expected_sha1: any, f: any;
  actual_sha1 = context.movie.posterFrame.sha1;
  const $cm1 = $open(test_file("just-two-mice.png"), "rb");
  f = $enter($cm1);
  try {
    expected_sha1 = $hashlib.sha1($read(f)).hexdigest();
  } finally {
    $exit($cm1);
  }
  $assertEqual(actual_sha1, expected_sha1, "not the same image");
}
step("then", "movie.poster_frame is the same image as poster_frame", [], then_movie_poster_frame_is_the_same_image_as_poster_frame);

function then_movie_shape_type_is_MSO_SHAPE_TYPE_MEDIA(context: any): any {
  let shape_type: any;
  shape_type = context.movie.shapeType;
  $assertEqual(shape_type, MSO_SHAPE_TYPE.MEDIA, $mod("got %s", shape_type));
}
step("then", "movie.shape_type is MSO_SHAPE_TYPE.MEDIA", [], then_movie_shape_type_is_MSO_SHAPE_TYPE_MEDIA);

function then_movie_width_movie_height_eq_cx_cy(context: any): any {
  let movie: any, size: any;
  movie = context.movie;
  size = [movie.width, movie.height];
  $assertEqual(size, [$call(Emu, 3962400), $call(Emu, 5715000)], $mod("got %s", size));
}
step("then", "movie.width, movie.height == cx, cy", [], then_movie_width_movie_height_eq_cx_cy);

function then_ole_format_bytes_matches_ole_object_file_byte_for_byte(context: any): any {
  $assertEqual(context.oleFormat.blob, context.oleObjectFile.getvalue());
}
step("then", "ole_format.blob matches ole_object_file byte-for-byte", [], then_ole_format_bytes_matches_ole_object_file_byte_for_byte);

function then_ole_format_prog_id_eq_value(context: any, expected_value: any): any {
  let actual_value: any;
  actual_value = context.oleFormat.progId;
  $assertEqual(actual_value, expected_value, $mod("expected %r, got %r", [expected_value, actual_value]));
}
step("then", "ole_format.prog_id == \"{expected_value}\"", ["expected_value"], then_ole_format_prog_id_eq_value);

function then_ole_format_show_as_icon_is_value(context: any, value: any): any {
  let actual_value: any, expected_value: any;
  expected_value = $get($d({True: true, False: false}), value);
  actual_value = context.oleFormat.showAsIcon;
  $assertEqual(actual_value, expected_value, $mod("expected %r, got %r", [expected_value, actual_value]));
}
step("then", "ole_format.show_as_icon is {value}", ["value"], then_ole_format_show_as_icon_is_value);

function then_picture_crop_side_eq_value(context: any, side: any, value: any): any {
  let actual_value: any, expected_value: any;
  expected_value = $round($float(value), 5);
  actual_value = $round($getattr(context.picture, $mod("crop_%s", side)), 5);
  $assertEqual(actual_value, expected_value, $mod("picture.crop_%s == %s", [side, actual_value]));
}
step("then", "picture.crop_{side} == {value}", ["side", "value"], then_picture_crop_side_eq_value);

function then_picture_image_is_an_Image_object(context: any): any {
  let class_name: any;
  class_name = $typename(context.picture.image);
  $assertEqual(class_name, "Image", $mod("picture.image is a %s object", class_name));
}
step("then", "picture.image is an Image object", [], then_picture_image_is_an_Image_object);

function then_shape_adjustments_is_value(context: any): any {
  let shape: any;
  shape = context.shape;
  $assertEqual($get(shape.adjustments, 0), 0.15);
}
step("then", "shape.adjustments[0] is 0.15", [], then_shape_adjustments_is_value);

function then_shape_chart_is_a_Chart_object(context: any): any {
  let chart: any, class_name: any;
  chart = context.shape.chart;
  class_name = $typename(chart);
  $assertEqual(class_name, "Chart", $mod("got %s", class_name));
}
step("then", "shape.chart is a Chart object", [], then_shape_chart_is_a_Chart_object);

function then_shape_click_action_is_an_ActionSetting_object(context: any): any {
  $assert($isinstance(context.shape.clickAction, ActionSetting));
}
step("then", "shape.click_action is an ActionSetting object", [], then_shape_click_action_is_an_ActionSetting_object);

function then_shape_has_chart_is_value(context: any, value: any): any {
  let actual_value: any, expected_value: any;
  expected_value = $get($d({True: true, False: false}), value);
  actual_value = context.shape.hasChart;
  $assert($is(actual_value, expected_value), $mod("shape.has_chart is %s", actual_value));
}
step("then", "shape.has_chart is {value}", ["value"], then_shape_has_chart_is_value);

function then_shape_has_table_is_value(context: any, value: any): any {
  let actual_value: any, expected_value: any;
  expected_value = $get($d({True: true, False: false}), value);
  actual_value = context.shape.hasTable;
  $assert($is(actual_value, expected_value), $mod("shape.has_table is %s", actual_value));
}
step("then", "shape.has_table is {value}", ["value"], then_shape_has_table_is_value);

function then_shape_has_text_frame_is(context: any, value_str: any): any {
  let expected_value: any, has_text_frame: any;
  expected_value = $get($d({True: true, False: false}), value_str);
  has_text_frame = context.shape.hasTextFrame;
  $assert($is(has_text_frame, expected_value), $mod("got %s", has_text_frame));
}
step("then", "shape.has_text_frame is {value_str}", ["value_str"], then_shape_has_text_frame_is);

function then_shape_height_eq_value(context: any, value: any): any {
  let actual_height: any, expected_height: any;
  expected_height = $int(value);
  actual_height = context.shape.height;
  $assertEqual(actual_height, expected_height, $mod("shape.height == %s", actual_height));
}
step("then", "shape.height == {value}", ["value"], then_shape_height_eq_value);

function then_shape_left_eq_value(context: any, value: any): any {
  let actual_left: any, expected_left: any;
  expected_left = $int(value);
  actual_left = context.shape.left;
  $assertEqual(actual_left, expected_left, $mod("shape.left == %s", actual_left));
}
step("then", "shape.left == {value}", ["value"], then_shape_left_eq_value);

function then_shape_line_is_a_LineFormat_object(context: any): any {
  let expected_cls_name: any, line_format: any, line_format_cls_name: any, shape: any;
  shape = context.shape;
  line_format = shape.line;
  line_format_cls_name = cls_qname(line_format);
  expected_cls_name = "pptx.dml.line.LineFormat";
  $assertEqual(line_format_cls_name, expected_cls_name, $mod("expected '%s', got '%s'", [expected_cls_name, line_format_cls_name]));
}
step("then", "shape.line is a LineFormat object", [], then_shape_line_is_a_LineFormat_object);

function then_shape_name_eq_value(context: any, expected_value: any): any {
  let msg: any, shape: any;
  shape = context.shape;
  msg = $mod("expected shape name '%s', got '%s'", [shape.name, expected_value]);
  $assertEqual(shape.name, expected_value, msg);
}
step("then", "shape.name == '{expected_value}'", ["expected_value"], then_shape_name_eq_value);

function then_shape_ole_format_is_an_OleFormat_object(context: any): any {
  let cls_name: any, expected_cls_name: any;
  cls_name = $typename(context.shape.oleFormat);
  expected_cls_name = "_OleFormat";
  $assertEqual(cls_name, expected_cls_name, $mod("expected %r, got %r", [expected_cls_name, cls_name]));
}
step("then", "shape.ole_format is an _OleFormat object", [], then_shape_ole_format_is_an_OleFormat_object);

function then_shape_part_is_a_SlidePart_object(context: any): any {
  let cls_name: any, expected_cls_name: any;
  cls_name = $typename(context.shape.part);
  expected_cls_name = "SlidePart";
  $assertEqual(cls_name, expected_cls_name, $mod("expected '%s', got '%s'", [expected_cls_name, cls_name]));
}
step("then", "shape.part is a SlidePart object", [], then_shape_part_is_a_SlidePart_object);

function then_shape_part_is_slide_part(context: any): any {
  $assert($is(context.shape.part, context.slide.part));
}
step("then", "shape.part is slide.part", [], then_shape_part_is_slide_part);

function then_shape_rotation_eq_value(context: any, value: any): any {
  let expected_value: any, shape: any;
  shape = context.shape;
  expected_value = $float(value);
  $assertEqual(shape.rotation, expected_value, $mod("got %s", expected_value));
}
step("then", "shape.rotation == {value}", ["value"], then_shape_rotation_eq_value);

function then_shape_shadow_is_a_ShadowFormat_object(context: any): any {
  let cls_name: any;
  cls_name = $typename(context.shape.shadow);
  $assertEqual(cls_name, "ShadowFormat", $mod("shape.shadow is a '%s' object", cls_name));
}
step("then", "shape.shadow is a ShadowFormat object", [], then_shape_shadow_is_a_ShadowFormat_object);

function then_shape_shadow_raises_NotImplementedError(context: any): any {
  try {
    context.shape.shadow;
  } catch (_e: any) {
    return;
  }
  throw new Error("shape.shadow did not raise");
}
step("then", "shape.shadow raises NotImplementedError", [], then_shape_shadow_raises_NotImplementedError);

function then_shape_shape_id_equals(context: any, value_str: any): any {
  let expected_value: any, shape_id: any;
  expected_value = $int(value_str);
  shape_id = context.shape.shapeId;
  $assertEqual(shape_id, expected_value, $mod("got %s", shape_id));
}
step("then", "shape.shape_id == {value_str}", ["value_str"], then_shape_shape_id_equals);

function then_shape_shape_type_is_MSO_SHAPE_TYPE_member(context: any, member_name: any): any {
  let actual_shape_type: any, expected_shape_type: any;
  expected_shape_type = $getattr(MSO_SHAPE_TYPE, member_name);
  actual_shape_type = context.shape.shapeType;
  $assertEqual(actual_shape_type, expected_shape_type, $mod("shape.shape_type == %s", actual_shape_type));
}
step("then", "shape.shape_type == MSO_SHAPE_TYPE.{member_name}", ["member_name"], then_shape_shape_type_is_MSO_SHAPE_TYPE_member);

function then_shape_text_eq_value(context: any, value: any): any {
  let actual: any, expected: any;
  [actual, expected] = $list([context.shape.text, $pyEval(value, $scope)]);
  $assertEqual(actual, expected, $mod("shape.text == \"%s\"", actual));
}
step("then", "shape.text == {value}", ["value"], then_shape_text_eq_value);

function then_shape_top_eq_value(context: any, value: any): any {
  let actual_top: any, expected_top: any;
  expected_top = $int(value);
  actual_top = context.shape.top;
  $assertEqual(actual_top, expected_top, $mod("shape.top == %s", actual_top));
}
step("then", "shape.top == {value}", ["value"], then_shape_top_eq_value);

function then_shape_width_eq_value(context: any, value: any): any {
  let actual_width: any, expected_width: any;
  expected_width = $int(value);
  actual_width = context.shape.width;
  $assertEqual(actual_width, expected_width, $mod("shape.width == %s", actual_width));
}
step("then", "shape.width == {value}", ["value"], then_shape_width_eq_value);

// the names python's eval() sees in this module
const $scope = { ActionSetting, Emu, MSO_SHAPE, MSO_SHAPE_TYPE, PP_MEDIA_TYPE, Presentation };
