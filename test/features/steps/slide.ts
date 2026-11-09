// Machine-translated from python-pptx features/steps/slide.py by
// scripts/py2ts_steps.py. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $assert, $assertEqual, $call, $d, $eq, $get, $in, $int, $is, $len, $list, $mod, $t1, $t2, $t3, $truthy, $typename } from "../../pyshim.js";
import { AttributeError, IndexError, KeyError, NotImplementedError, ValueError } from "../../../src/errors.js";
import { step } from "../runner.js";
import { test_pptx } from "./helpers.js";
import { Presentation } from "../../../src/index.js";
void $py;


function given_a_blank_slide(context: any): any {
  context.prs = $call(Presentation, test_pptx("sld-blank"));
  context.slide = $get(context.prs.slides, 0);
}
step("given", "a blank slide", [], given_a_blank_slide);

function given_a_notes_slide(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("sld-notes"));
  context.notesSlide = $get(prs.slides, 0).notesSlide;
}
step("given", "a notes slide", [], given_a_notes_slide);

function given_a_slide(context: any): any {
  let presentation: any;
  presentation = $call(Presentation, test_pptx("shp-shapes"));
  context.slide = $get(presentation.slides, 0);
}
step("given", "a slide", [], given_a_slide);

function given_a_slide_having_a_notes_slide(context: any): any {
  context.slide = $get($call(Presentation, test_pptx("sld-notes")).slides, 0);
}
step("given", "a slide having a notes slide", [], given_a_slide_having_a_notes_slide);

function given_a_slide_having_no_notes_slide(context: any): any {
  context.slide = $get($call(Presentation, test_pptx("sld-notes")).slides, 1);
}
step("given", "a slide having no notes slide", [], given_a_slide_having_no_notes_slide);

function given_a_slide_having_a_title(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("shp-shapes"));
  [context.prs, context.slide] = $list([prs, $get(prs.slides, 0)]);
}
step("given", "a slide having a title", [], given_a_slide_having_a_title);

function given_a_slide_having_name_name(context: any, name: any): any {
  let presentation: any, slide_idx: any;
  slide_idx = ($truthy($eq(name, "Overview")) ? 0 : 1);
  presentation = $call(Presentation, test_pptx("sld-slide"));
  context.slide = $get(presentation.slides, slide_idx);
}
step("given", "a slide having name {name}", ["name"], given_a_slide_having_name_name);

function given_a_slide_having_slide_id_256(context: any): any {
  let presentation: any;
  presentation = $call(Presentation, test_pptx("shp-shapes"));
  context.slide = $get(presentation.slides, 0);
}
step("given", "a slide having slide id 256", [], given_a_slide_having_slide_id_256);

function given_a_Slide_object_based_on_slide_layout_as_slide(context: any): any {
  context.slide = $get(context.prs.slides, 0);
}
step("given", "a Slide object based on slide_layout as slide", [], given_a_Slide_object_based_on_slide_layout_as_slide);

function given_a_Slide_object_having_background_as_slide(context: any, def_or_ovr: any): any {
  let slide_idx: any;
  slide_idx = $get($d({"the default": 0, "an overridden": 1}), def_or_ovr);
  context.slide = $get($call(Presentation, test_pptx("sld-slide")).slides, slide_idx);
}
step("given", "a Slide object having {def_or_ovr} background as slide", ["def_or_ovr"], given_a_Slide_object_having_background_as_slide);

function given_a_SlideLayout_object_as_slide(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("sld-slide"));
  const $t1 = $get(prs.slideLayouts, 0);
  context.slide = $t1;
  context.slideLayout = $t1;
}
step("given", "a SlideLayout object as slide_layout", [], given_a_SlideLayout_object_as_slide);
step("given", "a SlideLayout object as slide", [], given_a_SlideLayout_object_as_slide);

function given_a_SlideLayout_object_having_name_as_slide(context: any, name: any): any {
  let prs: any, slide_layout_idx: any;
  slide_layout_idx = ($truthy($eq(name, "of no explicit value")) ? 0 : 1);
  prs = $call(Presentation, test_pptx("sld-slide"));
  context.slide = $get(prs.slideLayouts, slide_layout_idx);
}
step("given", "a SlideLayout object having name {name} as slide", ["name"], given_a_SlideLayout_object_having_name_as_slide);

function given_a_SlideLayout_object_used_by_slides_as_slide_layout(context: any, which_slides: any): any {
  let slide_layout_idx: any;
  slide_layout_idx = $get($d({"a slide": 0, "no slides": 1}), which_slides);
  context.prs = $call(Presentation, test_pptx("sld-slide"));
  context.slideLayout = $get(context.prs.slideLayouts, slide_layout_idx);
}
step("given", "a SlideLayout object used by {which_slides} as slide_layout", ["which_slides"], given_a_SlideLayout_object_used_by_slides_as_slide_layout);

function given_a_SlideMaster_object_as_slide(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("sld-slide"));
  const $t2 = $get(prs.slideMasters, 0);
  context.slide = $t2;
  context.slideMaster = $t2;
}
step("given", "a SlideMaster object as slide_master", [], given_a_SlideMaster_object_as_slide);
step("given", "a SlideMaster object as slide", [], given_a_SlideMaster_object_as_slide);

function then_len_notes_slide_shapes_is_count(context: any, count: any): any {
  let shapes: any;
  shapes = context.notesSlide.shapes;
  $assertEqual($len(shapes), $int(count));
}
step("then", "len(notes_slide.shapes) is {count}", ["count"], then_len_notes_slide_shapes_is_count);

function then_notes_slide_notes_placeholder_is_a_NotesSlidePlacehldr_obj(context: any): any {
  let cls_name: any, notes_slide: any;
  notes_slide = context.notesSlide;
  cls_name = $typename(notes_slide.notesPlaceholder);
  $assertEqual(cls_name, "NotesSlidePlaceholder", $mod("got %s", cls_name));
}
step("then", "notes_slide.notes_placeholder is a NotesSlidePlaceholder object", [], then_notes_slide_notes_placeholder_is_a_NotesSlidePlacehldr_obj);

function then_notes_slide_notes_text_frame_is_a_TextFrame_object(context: any): any {
  let cls_name: any, notes_slide: any;
  notes_slide = context.notesSlide;
  cls_name = $typename(notes_slide.notesTextFrame);
  $assertEqual(cls_name, "TextFrame", $mod("got %s", cls_name));
}
step("then", "notes_slide.notes_text_frame is a TextFrame object", [], then_notes_slide_notes_text_frame_is_a_TextFrame_object);

function then_notes_slide_placeholders_is_a_NotesSlidePlaceholders_object(context: any): any {
  let notes_slide: any;
  notes_slide = context.notesSlide;
  $assertEqual($typename(notes_slide.placeholders), "NotesSlidePlaceholders");
}
step("then", "notes_slide.placeholders is a NotesSlidePlaceholders object", [], then_notes_slide_placeholders_is_a_NotesSlidePlaceholders_object);

function then_slide_in_slide_layout_used_by_slides_is_True(context: any): any {
  $assert($in(context.slide, context.slideLayout.usedBySlides));
}
step("then", "slide in slide_layout.used_by_slides is True", [], then_slide_in_slide_layout_used_by_slides_is_True);

function then_slide_background_is_a_Background_object(context: any): any {
  let cls_name: any;
  cls_name = $typename(context.slide.background);
  $assertEqual(cls_name, "_Background", $mod("slide.background is a %s object", cls_name));
}
step("then", "slide.background is a _Background object", [], then_slide_background_is_a_Background_object);

function then_slide_follow_master_background_is_value(context: any, value: any): any {
  let actual_value: any, expected_value: any;
  expected_value = $get($d({True: true, False: false}), value);
  actual_value = context.slide.followMasterBackground;
  $assert($is(actual_value, expected_value), $mod("slide.follow_master_background is %s", actual_value));
}
step("then", "slide.follow_master_background is {value}", ["value"], then_slide_follow_master_background_is_value);

function then_slide_has_notes_slide_is_value(context: any, value: any): any {
  let expected_value: any, slide: any;
  expected_value = $get($d({True: true, False: false}), value);
  slide = context.slide;
  $assert($is(slide.hasNotesSlide, expected_value));
}
step("then", "slide.has_notes_slide is {value}", ["value"], then_slide_has_notes_slide_is_value);

function then_slide_name_is_value(context: any, value: any): any {
  let actual_name: any, expected_name: any;
  expected_name = ($truthy($eq(value, "the empty string")) ? "" : value);
  actual_name = context.slide.name;
  $assertEqual(actual_name, expected_name, $mod("slide.name == %s", actual_name));
}
step("then", "slide.name is {value}", ["value"], then_slide_name_is_value);

function then_slide_notes_slide_is_a_NotesSlide_object(context: any): any {
  let notes_slide: any;
  const $t3 = context.slide.notesSlide;
  notes_slide = $t3;
  context.notesSlide = $t3;
  $assertEqual($typename(notes_slide), "NotesSlide");
}
step("then", "slide.notes_slide is a NotesSlide object", [], then_slide_notes_slide_is_a_NotesSlide_object);

function then_slide_placeholders_is_a_clsname_object(context: any, clsname: any): any {
  let actual_clsname: any, expected_clsname: any;
  actual_clsname = $typename(context.slide.placeholders);
  expected_clsname = clsname;
  $assertEqual(actual_clsname, expected_clsname, $mod("slide.placeholders is a %s object", actual_clsname));
}
step("then", "slide.placeholders is a {clsname} object", ["clsname"], then_slide_placeholders_is_a_clsname_object);

function then_slide_shapes_is_a_clsname_object(context: any, clsname: any): any {
  let actual_clsname: any, expected_clsname: any;
  actual_clsname = $typename(context.slide.shapes);
  expected_clsname = clsname;
  $assertEqual(actual_clsname, expected_clsname, $mod("slide.shapes is a %s object", actual_clsname));
}
step("then", "slide.shapes is a {clsname} object", ["clsname"], then_slide_shapes_is_a_clsname_object);

function then_slide_slide_id_is_256(context: any): any {
  let slide: any;
  slide = context.slide;
  $assertEqual(slide.slideId, 256);
}
step("then", "slide.slide_id is 256", [], then_slide_slide_id_is_256);

function then_slide_slide_layout_is_the_one_passed_in_the_call(context: any): any {
  let slide: any;
  slide = $get(context.prs.slides, 3);
  $assertEqual(slide.slideLayout, context.slideLayout);
}
step("then", "slide.slide_layout is the one passed in the call", [], then_slide_slide_layout_is_the_one_passed_in_the_call);

function then_slide_layout_slide_master_is_a_SlideMaster_object(context: any): any {
  let slide_layout: any;
  slide_layout = context.slideLayout;
  $assertEqual($typename(slide_layout.slideMaster), "SlideMaster");
}
step("then", "slide_layout.slide_master is a SlideMaster object", [], then_slide_layout_slide_master_is_a_SlideMaster_object);

function then_slide_master_slide_layouts_is_a_SlideLayouts_object(context: any): any {
  let slide_master: any;
  slide_master = context.slideMaster;
  $assertEqual($typename(slide_master.slideLayouts), "SlideLayouts");
}
step("then", "slide_master.slide_layouts is a SlideLayouts object", [], then_slide_master_slide_layouts_is_a_SlideLayouts_object);

function then_slide_layout_used_by_slides_eq_empty_tuple(context: any): any {
  $assertEqual(context.slideLayout.usedBySlides, []);
}
step("then", "slide_layout.used_by_slides == ()", [], then_slide_layout_used_by_slides_eq_empty_tuple);

// the names python's eval() sees in this module
const $scope = { Presentation };
