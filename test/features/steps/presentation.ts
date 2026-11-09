// Machine-translated from python-pptx features/steps/presentation.py by
// scripts/py2ts_steps.py. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $assert, $assertEqual, $call, $cast, $cm1, $cm2, $cm3, $cmp, $enter, $exit, $get, $int, $is, $isinstance, $len, $mod, $open, $os, $path, $read, $truthy, $typename, $zipfile, BytesIO } from "../../pyshim.js";
import { AttributeError, IndexError, KeyError, NotImplementedError, ValueError } from "../../../src/errors.js";
import { step } from "../runner.js";
import { saved_pptx_path, test_file, test_pptx } from "./helpers.js";
import { Presentation } from "../../../src/index.js";
import { RELATIONSHIP_TYPE as RT } from "../../../src/opc/constants.js";
import { Inches } from "../../../src/util.js";
void $py;



function given_clean_working_dir(context: any): any {
  if ($truthy($path.isfile(saved_pptx_path))) {
    $os.remove(saved_pptx_path);
  }
}
step("given", "a clean working directory", [], given_clean_working_dir);

function given_a_presentation(context: any): any {
  context.presentation = $call(Presentation, test_pptx("prs-properties"));
}
step("given", "a presentation", [], given_a_presentation);

function given_a_presentation_having_a_notes_master(context: any): any {
  context.prs = $call(Presentation, test_pptx("prs-notes"));
}
step("given", "a presentation having a notes master", [], given_a_presentation_having_a_notes_master);

function given_a_presentation_having_no_notes_master(context: any): any {
  context.prs = $call(Presentation, test_pptx("prs-properties"));
}
step("given", "a presentation having no notes master", [], given_a_presentation_having_no_notes_master);

function given_prs_with_image_jpg_MIME_type(context: any): any {
  context.prs = $call(Presentation, test_pptx("test-image-jpg-mime"));
}
step("given", "a presentation with an image/jpg MIME-type", [], given_prs_with_image_jpg_MIME_type);

function given_prs_with_ext_rels(context: any): any {
  context.prs = $call(Presentation, test_pptx("ext-rels"));
}
step("given", "a presentation with external relationships", [], given_prs_with_ext_rels);

function given_initialized_pptx_env(context: any): any {
}
step("given", "an initialized pptx environment", [], given_initialized_pptx_env);

function when_change_slide_width_and_height(context: any): any {
  let presentation: any;
  presentation = context.presentation;
  presentation.slideWidth = $call(Inches, 4);
  presentation.slideHeight = $call(Inches, 3);
}
step("when", "I change the slide width and height", [], when_change_slide_width_and_height);

function when_construct_default_prs(context: any): any {
  context.prs = $call(Presentation);
}
step("when", "I construct a Presentation instance with no path argument", [], when_construct_default_prs);

function when_open_basic_pptx(context: any): any {
  context.prs = $call(Presentation, test_pptx("test"));
}
step("when", "I open a basic PowerPoint presentation", [], when_open_basic_pptx);

function when_I_open_a_presentation_extracted_into_a_directory(context: any): any {
  context.prs = $call(Presentation, test_file("extracted-pptx"));
}
step("when", "I open a presentation extracted into a directory", [], when_I_open_a_presentation_extracted_into_a_directory);

function when_open_presentation_stream(context: any): any {
  let f: any, stream: any;
  const $cm1 = $open(test_pptx("test"), "rb");
  f = $enter($cm1);
  try {
    stream = new BytesIO($read(f));
  } finally {
    $exit($cm1);
  }
  context.prs = $call(Presentation, stream);
  stream.close();
}
step("when", "I open a presentation contained in a stream", [], when_open_presentation_stream);

function when_save_and_reload_prs(context: any): any {
  if ($truthy($path.isfile(saved_pptx_path))) {
    $os.remove(saved_pptx_path);
  }
  context.prs.save(saved_pptx_path);
  context.prs = $call(Presentation, saved_pptx_path);
}
step("when", "I save and reload the presentation", [], when_save_and_reload_prs);

function when_save_stream_to_a_file(context: any): any {
  let f: any;
  if ($truthy($path.isfile(saved_pptx_path))) {
    $os.remove(saved_pptx_path);
  }
  context.stream.seek(0);
  const $cm2 = $open(saved_pptx_path, "wb");
  f = $enter($cm2);
  try {
    f.write($read(context.stream));
  } finally {
    $exit($cm2);
  }
}
step("when", "I save that stream to a file", [], when_save_stream_to_a_file);

function when_save_presentation(context: any): any {
  if ($truthy($path.isfile(saved_pptx_path))) {
    $os.remove(saved_pptx_path);
  }
  context.prs.save(saved_pptx_path);
}
step("when", "I save the presentation", [], when_save_presentation);

function when_save_presentation_to_stream(context: any): any {
  context.stream = new BytesIO();
  context.prs.save(context.stream);
}
step("when", "I save the presentation to a stream", [], when_save_presentation_to_stream);

function then_receive_prs_based_on_def_tmpl(context: any): any {
  let prs: any, slide_layouts: any, slide_masters: any;
  prs = context.prs;
  $assert(!$is(prs, null));
  slide_masters = prs.slideMasters;
  $assert(!$is(slide_masters, null));
  $assertEqual($len(slide_masters), 1);
  slide_layouts = $get(slide_masters, 0).slideLayouts;
  $assert(!$is(slide_layouts, null));
  $assertEqual($len(slide_layouts), 11);
}
step("then", "I receive a presentation based on the default template", [], then_receive_prs_based_on_def_tmpl);

function then_slide_height_matches_known_value(context: any): any {
  let presentation: any;
  presentation = context.presentation;
  $assertEqual(presentation.slideHeight, 6858000);
}
step("then", "its slide height matches its known value", [], then_slide_height_matches_known_value);

function then_slide_width_matches_known_value(context: any): any {
  let presentation: any;
  presentation = context.presentation;
  $assertEqual(presentation.slideWidth, 9144000);
}
step("then", "its slide width matches its known value", [], then_slide_width_matches_known_value);

function then_see_pptx_file_in_working_dir(context: any): any {
  let actual: any, minimum: any;
  $assert($path.isfile(saved_pptx_path));
  minimum = 30000;
  actual = $path.getsize(saved_pptx_path);
  $assert($cmp(actual, ">", minimum));
}
step("then", "I see the pptx file in the working directory", [], then_see_pptx_file_in_working_dir);

function then_len_notes_master_shapes_is_shape_count(context: any, shape_count: any): any {
  let actual: any, expected: any, notes_master: any;
  notes_master = context.prs.notesMaster;
  expected = $int(shape_count);
  actual = $len(notes_master.shapes);
  $assertEqual(actual, expected, $mod("got %s", actual));
}
step("then", "len(notes_master.shapes) is {shape_count}", ["shape_count"], then_len_notes_master_shapes_is_shape_count);

function then_prs_notes_master_is_a_NotesMaster_object(context: any): any {
  let prs: any;
  prs = context.prs;
  $assertEqual($typename(prs.notesMaster), "NotesMaster");
}
step("then", "prs.notes_master is a NotesMaster object", [], then_prs_notes_master_is_a_NotesMaster_object);

function then_prs_slides_is_a_Slides_object(context: any): any {
  let prs: any;
  prs = context.presentation;
  $assertEqual($typename(prs.slides), "Slides");
}
step("then", "prs.slides is a Slides object", [], then_prs_slides_is_a_Slides_object);

function then_prs_slide_masters_is_a_SlideMasters_object(context: any): any {
  let prs: any;
  prs = context.presentation;
  $assertEqual($typename(prs.slideMasters), "SlideMasters");
}
step("then", "prs.slide_masters is a SlideMasters object", [], then_prs_slide_masters_is_a_SlideMasters_object);

function then_ext_rels_are_preserved(context: any): any {
  let prs: any, rel: any, sld: any;
  prs = context.prs;
  sld = $get(prs.slides, 0);
  rel = $get(sld.part._rels, "rId2");
  $assert(rel.isExternal);
  $assertEqual(rel.reltype, RT.HYPERLINK);
  $assertEqual(rel.targetRef, "https://github.com/scanny/python-pptx");
}
step("then", "the external relationships are still there", [], then_ext_rels_are_preserved);

function then_the_package_has_the_expected_number_of_rels_parts(context: any): any {
  let member_count: any, z: any;
  const $cm3 = $zipfile.ZipFile(saved_pptx_path, "r");
  z = $enter($cm3);
  try {
    member_count = $len(z.namelist());
  } finally {
    $exit($cm3);
  }
  $assertEqual(member_count, 18, $mod("expected 18, got %d", member_count));
}
step("then", "the package has the expected number of .rels parts", [], then_the_package_has_the_expected_number_of_rels_parts);

function then_I_can_access_the_JPEG_image(context: any): any {
  let picture: any, prs: any, slide: any;
  prs = $cast("presentation.Presentation", context.prs);
  slide = $get(prs.slides, 0);
  picture = $cast("Picture", $get(slide.shapes, 0));
  try {
    picture.image;
  } catch (_e: any) {
    if ($isinstance(_e, AttributeError)) {
      throw new Error("JPEG image not recognized");
    } else {
      throw _e;
    }
  }
}
step("then", "I can access the JPEG image", [], then_I_can_access_the_JPEG_image);

function then_slide_height_matches_new_value(context: any): any {
  let presentation: any;
  presentation = context.presentation;
  $assertEqual(presentation.slideHeight, $call(Inches, 3));
}
step("then", "the slide height matches the new value", [], then_slide_height_matches_new_value);

function then_slide_width_matches_new_value(context: any): any {
  let presentation: any;
  presentation = context.presentation;
  $assertEqual(presentation.slideWidth, $call(Inches, 4));
}
step("then", "the slide width matches the new value", [], then_slide_width_matches_new_value);

// the names python's eval() sees in this module
const $scope = { Inches, Presentation, RT };
