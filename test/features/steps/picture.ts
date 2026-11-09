// Machine-translated from python-pptx features/steps/picture.py by
// scripts/py2ts_steps.py. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $assert, $assertEqual, $call, $cm1, $d, $enter, $exit, $get, $getattr, $in, $list, $mod, $open, $read, $set_, $truthy, $typename, BytesIO } from "../../pyshim.js";
import { AttributeError, IndexError, KeyError, NotImplementedError, ValueError } from "../../../src/errors.js";
import { step } from "../runner.js";
import { saved_pptx_path, test_image, test_pptx } from "./helpers.js";
import { MSO_AUTO_SHAPE_TYPE } from "../../../src/enum/shapes.js";
import { Presentation } from "../../../src/index.js";
import { Package } from "../../../src/package.js";
import { Inches } from "../../../src/util.js";
void $py;


function given_a_picture_object_masked_by_shape_as_picture(context: any, shape: any): any {
  let prs: any, shape_idx: any;
  shape_idx = $get($d({rectangle: 0, circle: 1}), shape);
  prs = $call(Presentation, test_pptx("shp-picture"));
  context.picture = $get($get(prs.slides, 1).shapes, shape_idx);
}
step("given", "a Picture object masked by a {shape} as picture", ["shape"], given_a_picture_object_masked_by_shape_as_picture);

function given_a_picture_of_known_position_and_size(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("shp-pos-and-size"));
  context.picture = $get($get(prs.slides, 1).shapes, 0);
}
step("given", "a picture of known position and size", [], given_a_picture_of_known_position_and_size);

function when_I_add_the_image_filename_using_shapes_add_picture(context: any, filename: any): any {
  let shapes: any;
  shapes = context.slide.shapes;
  shapes.addPicture(test_image(filename), $call(Inches, 1.25), $call(Inches, 1.25));
}
step("when", "I add the image {filename} using shapes.add_picture()", ["filename"], when_I_add_the_image_filename_using_shapes_add_picture);

function when_I_add_the_stream_image_filename_using_add_picture(context: any, filename: any): any {
  let f: any, shapes: any, stream: any;
  shapes = context.slide.shapes;
  const $cm1 = $open(test_image(filename), "rb");
  f = $enter($cm1);
  try {
    stream = new BytesIO($read(f));
  } finally {
    $exit($cm1);
  }
  shapes.addPicture(stream, $call(Inches, 1.25), $call(Inches, 1.25));
}
step("when", "I add the stream image {filename} using shapes.add_picture()", ["filename"], when_I_add_the_stream_image_filename_using_add_picture);

function when_I_assign_member_to_picture_auto_shape_type(context: any, member: any): any {
  context.picture.autoShapeType = $getattr(MSO_AUTO_SHAPE_TYPE, member);
}
step("when", "I assign MSO_AUTO_SHAPE_TYPE.{member} to picture.auto_shape_type", ["member"], when_I_assign_member_to_picture_auto_shape_type);

function step_then_a_ext_image_part_appears_in_the_pptx_file(context: any, ext: any): any {
  let image_partname: any, partnames: any, pkg: any;
  pkg = Package.open(saved_pptx_path);
  partnames = $set_($list(pkg.iterParts()).map((p: any) => p.partname));
  image_partname = $mod("/ppt/media/image1.%s", ext);
  $assert($in(image_partname, partnames), $mod("got %s", $list(partnames).filter((p: any) => $truthy($in("image", p))).map((p: any) => p)));
}
step("then", "a {ext} image part appears in the pptx file", ["ext"], step_then_a_ext_image_part_appears_in_the_pptx_file);

function then_picture_auto_shape_type_eq_shape_type_member(context: any, member: any): any {
  let actual: any, expected: any;
  expected = $getattr(MSO_AUTO_SHAPE_TYPE, member);
  actual = context.picture.autoShapeType;
  $assertEqual(actual, expected, $mod("shape.auto_shape_type == %s", actual));
}
step("then", "picture.auto_shape_type == MSO_AUTO_SHAPE_TYPE.{member}", ["member"], then_picture_auto_shape_type_eq_shape_type_member);

function then_the_picture_appears_in_the_slide(context: any): any {
  let cls_names: any, prs: any, shapes: any, slide: any;
  prs = $call(Presentation, saved_pptx_path);
  slide = $get(prs.slides, 0);
  shapes = slide.shapes;
  cls_names = $list(shapes).map((sp: any) => $typename(sp));
  $assert($in("Picture", cls_names));
}
step("then", "the picture appears in the slide", [], then_the_picture_appears_in_the_slide);

// the names python's eval() sees in this module
const $scope = { Inches, MSO_AUTO_SHAPE_TYPE, Package, Presentation };
