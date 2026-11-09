// Machine-translated from python-pptx features/steps/slides.py by
// scripts/py2ts_steps.py. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $assert, $assertEqual, $call, $dictget, $enumerate, $get, $index, $int, $is, $kw, $len, $list, $mod, $remove, $typename } from "../../pyshim.js";
import { AttributeError, IndexError, KeyError, NotImplementedError, ValueError } from "../../../src/errors.js";
import { step } from "../runner.js";
import { test_pptx } from "./helpers.js";
import { Presentation } from "../../../src/index.js";
void $py;


function given_a_SlideLayouts_object_containing_2_layouts(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("mst-slide-layouts"));
  context.slideLayouts = prs.slideMaster.slideLayouts;
}
step("given", "a SlideLayouts object containing 2 layouts as slide_layouts", [], given_a_SlideLayouts_object_containing_2_layouts);

function given_a_SlideMasters_object_containing_2_masters(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("prs-slide-masters"));
  context.slideMasters = prs.slideMasters;
}
step("given", "a SlideMasters object containing 2 masters", [], given_a_SlideMasters_object_containing_2_masters);

function given_a_Slides_object_containing_3_slides(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("sld-slides"));
  context.prs = prs;
  context.slides = prs.slides;
}
step("given", "a Slides object containing 3 slides", [], given_a_Slides_object_containing_3_slides);

function when_I_call_slides_add_slide(context: any): any {
  context.slideLayout = $get($get(context.prs.slideMasters, 0).slideLayouts, 0);
  context.slides.addSlide(context.slideLayout);
}
step("when", "I call slides.add_slide()", [], when_I_call_slides_add_slide);

function when_I_call_slide_layouts_remove(context: any): any {
  let slide_layouts: any;
  slide_layouts = context.slideLayouts;
  $remove(slide_layouts, $get(slide_layouts, 1));
}
step("when", "I call slide_layouts.remove(slide_layouts[1])", [], when_I_call_slide_layouts_remove);

function then_iterating_produces_3_NotesSlidePlaceholder_objects(context: any): any {
  let idx: any, typename: any;
  idx = -1;
  let placeholder: any;
  for ([idx, placeholder] of $list($enumerate(context.notesSlide.placeholders))) {
    typename = $typename(placeholder);
    $assertEqual(typename, "NotesSlidePlaceholder", $mod("got %s", typename));
  }
  $assertEqual(idx, 2);
}
step("then", "iterating produces 3 NotesSlidePlaceholder objects", [], then_iterating_produces_3_NotesSlidePlaceholder_objects);

function then_iterating_slide_layouts_produces_2_SlideLayout_objects(context: any): any {
  let idx: any, slide_layouts: any;
  slide_layouts = context.slideLayouts;
  idx = -1;
  let slide_layout: any;
  for ([idx, slide_layout] of $list($enumerate(slide_layouts))) {
    $assertEqual($typename(slide_layout), "SlideLayout");
  }
  $assertEqual(idx, 1);
}
step("then", "iterating slide_layouts produces 2 SlideLayout objects", [], then_iterating_slide_layouts_produces_2_SlideLayout_objects);

function then_iterating_slide_masters_produces_2_SlideMaster_objects(context: any): any {
  let idx: any, slide_masters: any;
  slide_masters = context.slideMasters;
  idx = -1;
  let slide_master: any;
  for ([idx, slide_master] of $list($enumerate(slide_masters))) {
    $assertEqual($typename(slide_master), "SlideMaster");
  }
  $assertEqual(idx, 1);
}
step("then", "iterating slide_masters produces 2 SlideMaster objects", [], then_iterating_slide_masters_produces_2_SlideMaster_objects);

function then_iterating_slides_produces_3_Slide_objects(context: any): any {
  let idx: any, slides: any;
  slides = context.slides;
  idx = -1;
  let slide: any;
  for ([idx, slide] of $list($enumerate(slides))) {
    $assertEqual($typename(slide), "Slide");
  }
  $assertEqual(idx, 2);
}
step("then", "iterating slides produces 3 Slide objects", [], then_iterating_slides_produces_3_Slide_objects);

function then_len_slides_is_count(context: any, count: any): any {
  let slides: any;
  slides = context.slides;
  $assertEqual($len(slides), $int(count));
}
step("then", "len(slides) is {count}", ["count"], then_len_slides_is_count);

function then_len_slide_layouts_is_2(context: any, n: any): any {
  $assertEqual($len(context.slideLayouts), $int(n));
}
step("then", "len(slide_layouts) is {n}", ["n"], then_len_slide_layouts_is_2);

function then_len_slide_masters_is_2(context: any): any {
  let slide_masters: any;
  slide_masters = context.slideMasters;
  $assertEqual($len(slide_masters), 2);
}
step("then", "len(slide_masters) is 2", [], then_len_slide_masters_is_2);

function then_slide_layouts_1_is_a_SlideLayout_object(context: any): any {
  let slide_layouts: any;
  slide_layouts = context.slideLayouts;
  $assertEqual($typename($get(slide_layouts, 1)), "SlideLayout");
}
step("then", "slide_layouts[1] is a SlideLayout object", [], then_slide_layouts_1_is_a_SlideLayout_object);

function then_slide_layouts_get_by_name_is_slide_layout(context: any): any {
  let slide_layouts: any;
  slide_layouts = context.slideLayouts;
  $assert($is(slide_layouts.getByName($get(slide_layouts, 1).name), $get(slide_layouts, 1)));
}
step("then", "slide_layouts.get_by_name(slide_layouts[1].name) is slide_layouts[1]", [], then_slide_layouts_get_by_name_is_slide_layout);

function then_slide_layouts_index_is_1(context: any): any {
  let slide_layouts: any;
  slide_layouts = context.slideLayouts;
  $assertEqual($index(slide_layouts, $get(slide_layouts, 1)), 1);
}
step("then", "slide_layouts.index(slide_layouts[1]) == 1", [], then_slide_layouts_index_is_1);

function then_slide_masters_1_is_a_SlideMaster_object(context: any): any {
  let slide_masters: any;
  slide_masters = context.slideMasters;
  $assertEqual($typename($get(slide_masters, 1)), "SlideMaster");
}
step("then", "slide_masters[1] is a SlideMaster object", [], then_slide_masters_1_is_a_SlideMaster_object);

function then_slides_get_256_is_slides_0(context: any): any {
  let slides: any;
  slides = context.slides;
  $assert($is($dictget(slides, 256), $get(slides, 0)));
}
step("then", "slides.get(256) is slides[0]", [], then_slides_get_256_is_slides_0);

function then_slides_get_666_default_slides_2_is_slides_2(context: any): any {
  let slides: any;
  slides = context.slides;
  $assert($is($dictget(slides, 666, $kw({default: $get(slides, 2)})), $get(slides, 2)));
}
step("then", "slides.get(666, default=slides[2]) is slides[2]", [], then_slides_get_666_default_slides_2_is_slides_2);

function then_slides_2_is_a_Slide_object(context: any): any {
  let slides: any;
  slides = context.slides;
  $assertEqual($typename($get(slides, 2)), "Slide");
}
step("then", "slides[2] is a Slide object", [], then_slides_2_is_a_Slide_object);

// the names python's eval() sees in this module
const $scope = { Presentation };
