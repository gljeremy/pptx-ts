// Machine-translated from python-pptx tests/test_slide.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "./pyshim.js";
import { $append, $assert, $assertEqual, $call, $dictget, $enumerate, $get, $getattr, $in, $index, $is, $isinstance, $iter, $kw, $len, $list, $mod, $raises, $remove, $slice, $str, $t1, $t2, $truthy, call, classMock, instanceMock, methodMock, propertyMock, pytestClass } from "./pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../src/exc.js";
import { defineChildren } from "../src/oxml/xmlchemy.js";
import { FillFormat as FillFormat$imported } from "../src/dml/fill.js";
import { PP_PLACEHOLDER as PP_PLACEHOLDER$imported } from "../src/enum/shapes.js";
import { Package as Package$imported } from "../src/package.js";
import { PresentationPart as PresentationPart$imported } from "../src/parts/presentation.js";
import { SlideLayoutPart as SlideLayoutPart$imported, SlideMasterPart as SlideMasterPart$imported, SlidePart as SlidePart$imported } from "../src/parts/slide.js";
import { Presentation as Presentation$imported } from "../src/presentation.js";
import { BaseShape as BaseShape$imported } from "../src/shapes/base.js";
import { LayoutPlaceholder as LayoutPlaceholder$imported, NotesSlidePlaceholder as NotesSlidePlaceholder$imported } from "../src/shapes/placeholder.js";
import { LayoutPlaceholders as LayoutPlaceholders$imported, LayoutShapes as LayoutShapes$imported, MasterPlaceholders as MasterPlaceholders$imported, MasterShapes as MasterShapes$imported, NotesSlidePlaceholders as NotesSlidePlaceholders$imported, NotesSlideShapes as NotesSlideShapes$imported, SlidePlaceholders as SlidePlaceholders$imported, SlideShapes as SlideShapes$imported } from "../src/shapes/shapetree.js";
import { NotesMaster as NotesMaster$imported, NotesSlide as NotesSlide$imported, Slide as Slide$imported, SlideLayout as SlideLayout$imported, SlideLayouts as SlideLayouts$imported, SlideMaster as SlideMaster$imported, SlideMasters as SlideMasters$imported, Slides as Slides$imported, _Background as _Background$imported, _BaseMaster as _BaseMaster$imported, _BaseSlide as _BaseSlide$imported } from "../src/slide.js";
import { TextFrame as TextFrame$imported } from "../src/text/text.js";
import { element, xml } from "./unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const FillFormat = FillFormat$imported;
const PP_PLACEHOLDER = PP_PLACEHOLDER$imported;
const Package = Package$imported;
const PresentationPart = PresentationPart$imported;
const SlideLayoutPart = SlideLayoutPart$imported;
const SlideMasterPart = SlideMasterPart$imported;
const SlidePart = SlidePart$imported;
const Presentation = Presentation$imported;
const BaseShape = BaseShape$imported;
const LayoutPlaceholder = LayoutPlaceholder$imported;
const NotesSlidePlaceholder = NotesSlidePlaceholder$imported;
const LayoutPlaceholders = LayoutPlaceholders$imported;
const LayoutShapes = LayoutShapes$imported;
const MasterPlaceholders = MasterPlaceholders$imported;
const MasterShapes = MasterShapes$imported;
const NotesSlidePlaceholders = NotesSlidePlaceholders$imported;
const NotesSlideShapes = NotesSlideShapes$imported;
const SlidePlaceholders = SlidePlaceholders$imported;
const SlideShapes = SlideShapes$imported;
const NotesMaster = NotesMaster$imported;
const NotesSlide = NotesSlide$imported;
const Slide = Slide$imported;
const SlideLayout = SlideLayout$imported;
const SlideLayouts = SlideLayouts$imported;
const SlideMaster = SlideMaster$imported;
const SlideMasters = SlideMasters$imported;
const Slides = Slides$imported;
const _Background = _Background$imported;
const _BaseMaster = _BaseMaster$imported;
const _BaseSlide = _BaseSlide$imported;
const TextFrame = TextFrame$imported;
void $py;


const Describe_BaseSlide: any = {
  fixtures: {
    background_fixture: { args: ["_Background_", "background_"], fn: (self: any, _Background_: any, background_: any): any => {
        let cSld: any, sld: any, slide: any;
        sld = element("p:sld/p:cSld");
        slide = $call(_BaseSlide, sld, null);
        cSld = $get(sld.xpath("//p:cSld"), 0);
        _Background_.returnValue = background_;
        return [slide, _Background_, cSld, background_];
      } },
    name_get_fixture: { args: ["request"], params: [["p:sld/p:cSld", ""], ["p:sld/p:cSld{name=Foobar}", "Foobar"]], fn: (self: any, request: any): any => {
        let base_slide: any, expected_name: any, sld_cxml: any;
        [sld_cxml, expected_name] = $list(request.param);
        base_slide = $call(_BaseSlide, element(sld_cxml), null);
        return [base_slide, expected_name];
      } },
    name_set_fixture: { args: ["request"], params: [["p:sld/p:cSld", "foo", "p:sld/p:cSld{name=foo}"], ["p:sld/p:cSld{name=foo}", "bar", "p:sld/p:cSld{name=bar}"], ["p:sld/p:cSld{name=bar}", "", "p:sld/p:cSld"], ["p:sld/p:cSld{name=bar}", null, "p:sld/p:cSld"], ["p:sld/p:cSld", "", "p:sld/p:cSld"], ["p:sld/p:cSld", null, "p:sld/p:cSld"]], fn: (self: any, request: any): any => {
        let base_slide: any, expected_cxml: any, expected_xml: any, new_value: any, xSld_cxml: any;
        [xSld_cxml, new_value, expected_cxml] = $list(request.param);
        base_slide = $call(_BaseSlide, element(xSld_cxml), null);
        expected_xml = xml(expected_cxml);
        return [base_slide, new_value, expected_xml];
      } },
    _Background_: { args: ["request"], fn: (self: any, request: any): any => {
        return classMock(request, "pptx.slide._Background");
      } },
    background_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _Background);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_name", args: ["name_get_fixture"], fn: (self: any, name_get_fixture: any): any => {
        let base_slide: any, expected_value: any;
        [base_slide, expected_value] = $list(name_get_fixture);
        $assertEqual(base_slide.name, expected_value);
      } },
    { name: "it_can_change_its_name", args: ["name_set_fixture"], fn: (self: any, name_set_fixture: any): any => {
        let base_slide: any, expected_xml: any, new_value: any;
        [base_slide, new_value, expected_xml] = $list(name_set_fixture);
        base_slide.name = new_value;
        $assertEqual(base_slide._element.xml, expected_xml);
      } },
    { name: "it_provides_access_to_its_background", args: ["background_fixture"], fn: (self: any, background_fixture: any): any => {
        let _Background_: any, background: any, background_: any, cSld: any, slide: any;
        [slide, _Background_, cSld, background_] = $list(background_fixture);
        background = slide.background;
        _Background_.assertCalledOnceWith(cSld);
        $assert($is(background, background_));
      } },
  ],
};
pytestClass("Describe_BaseSlide", Describe_BaseSlide);

const Describe_BaseMaster: any = {
  fixtures: {
    placeholders_fixture: { args: ["MasterPlaceholders_", "placeholders_"], fn: (self: any, MasterPlaceholders_: any, placeholders_: any): any => {
        let base_master: any, master_element: any, spTree: any;
        master_element = element("p:sldMaster/p:cSld/p:spTree");
        base_master = $call(_BaseMaster, master_element, null);
        spTree = $get(master_element.xpath("//p:spTree"), 0);
        return [base_master, MasterPlaceholders_, spTree, placeholders_];
      } },
    shapes_fixture: { args: ["MasterShapes_", "shapes_"], fn: (self: any, MasterShapes_: any, shapes_: any): any => {
        let base_master: any, master_element: any, spTree: any;
        master_element = element("p:sldMaster/p:cSld/p:spTree");
        base_master = $call(_BaseMaster, master_element, null);
        spTree = $get(master_element.xpath("//p:spTree"), 0);
        return [base_master, MasterShapes_, spTree, shapes_];
      } },
    subclass_fixture: { args: [], fn: (self: any): any => {
        return $call(_BaseMaster, null, null);
      } },
    MasterPlaceholders_: { args: ["request", "placeholders_"], fn: (self: any, request: any, placeholders_: any): any => {
        return $call(classMock, request, "pptx.slide.MasterPlaceholders", $kw({return_value: placeholders_}));
      } },
    MasterShapes_: { args: ["request", "shapes_"], fn: (self: any, request: any, shapes_: any): any => {
        return $call(classMock, request, "pptx.slide.MasterShapes", $kw({return_value: shapes_}));
      } },
    placeholders_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, MasterPlaceholders);
      } },
    shapes_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, MasterShapes);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_is_a_BaseSlide_subclass", args: ["subclass_fixture"], fn: (self: any, subclass_fixture: any): any => {
        let base_master: any;
        base_master = subclass_fixture;
        $assert($isinstance(base_master, _BaseSlide));
      } },
    { name: "it_provides_access_to_its_placeholders", args: ["placeholders_fixture"], fn: (self: any, placeholders_fixture: any): any => {
        let MasterPlaceholders_: any, base_master: any, placeholders: any, placeholders_: any, spTree: any;
        [base_master, MasterPlaceholders_, spTree, placeholders_] = $list(placeholders_fixture);
        placeholders = base_master.placeholders;
        MasterPlaceholders_.assertCalledOnceWith(spTree, base_master);
        $assert($is(placeholders, placeholders_));
      } },
    { name: "it_provides_access_to_its_shapes", args: ["shapes_fixture"], fn: (self: any, shapes_fixture: any): any => {
        let MasterShapes_: any, base_master: any, shapes: any, shapes_: any, spTree: any;
        [base_master, MasterShapes_, spTree, shapes_] = $list(shapes_fixture);
        shapes = base_master.shapes;
        MasterShapes_.assertCalledOnceWith(spTree, base_master);
        $assert($is(shapes, shapes_));
      } },
  ],
};
pytestClass("Describe_BaseMaster", Describe_BaseMaster);

const DescribeNotesSlide: any = {
  fixtures: {
    notes_ph_fixture: { args: ["request", "placeholders_prop_"], params: [[["SLIDE_IMAGE", "BODY", "SLIDE_NUMBER"], 1], [["DATE", "SLIDE_IMAGE", "FOOTER"], null]], fn: (self: any, request: any, placeholders_prop_: any): any => {
        let expected_value: any, match: any, notes_slide: any, placeholder_: any, placeholders_: any, type_names: any;
        [type_names, match] = $list(request.param);
        notes_slide = $call(NotesSlide, null, null);
        placeholders_ = [];
        for (const type_name of $list(type_names)) {
          placeholder_ = $call(instanceMock, request, NotesSlidePlaceholder, $kw({name: $mod("%s-placeholder", type_name)}));
          placeholder_.placeholderFormat.type = $getattr(PP_PLACEHOLDER, type_name);
          $append(placeholders_, placeholder_);
        }
        placeholders_prop_.returnValue = placeholders_;
        expected_value = ($truthy($is(match, null)) ? null : $get(placeholders_, match));
        return [notes_slide, expected_value];
      } },
    notes_tf_fixture: { args: ["request", "notes_placeholder_prop_", "placeholder_", "text_frame_"], params: [true, false], fn: (self: any, request: any, notes_placeholder_prop_: any, placeholder_: any, text_frame_: any): any => {
        let expected_value: any, has_text_frame: any, notes_slide: any;
        has_text_frame = request.param;
        notes_slide = $call(NotesSlide, null, null);
        if ($truthy(has_text_frame)) {
          notes_placeholder_prop_.returnValue = placeholder_;
          placeholder_.textFrame = text_frame_;
          expected_value = text_frame_;
        } else {
          notes_placeholder_prop_.returnValue = null;
          expected_value = null;
        }
        return [notes_slide, expected_value];
      } },
    placeholders_fixture: { args: ["NotesSlidePlaceholders_", "placeholders_"], fn: (self: any, NotesSlidePlaceholders_: any, placeholders_: any): any => {
        let notes: any, notes_slide: any, spTree: any;
        notes = element("p:notes/p:cSld/p:spTree");
        notes_slide = $call(NotesSlide, notes, null);
        spTree = $get(notes.xpath("//p:spTree"), 0);
        return [notes_slide, NotesSlidePlaceholders_, spTree, placeholders_];
      } },
    shapes_fixture: { args: ["NotesSlideShapes_", "shapes_"], fn: (self: any, NotesSlideShapes_: any, shapes_: any): any => {
        let notes: any, notes_slide: any, spTree: any;
        notes = element("p:notes/p:cSld/p:spTree");
        notes_slide = $call(NotesSlide, notes, null);
        spTree = $get(notes.xpath("//p:spTree"), 0);
        return [notes_slide, NotesSlideShapes_, spTree, shapes_];
      } },
    notes_master_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, NotesMaster);
      } },
    notes_placeholder_prop_: { args: ["request", "placeholder_"], fn: (self: any, request: any, placeholder_: any): any => {
        return $call(propertyMock, request, NotesSlide, "notes_placeholder", $kw({return_value: placeholder_}));
      } },
    NotesSlidePlaceholders_: { args: ["request", "placeholders_"], fn: (self: any, request: any, placeholders_: any): any => {
        return $call(classMock, request, "pptx.slide.NotesSlidePlaceholders", $kw({return_value: placeholders_}));
      } },
    NotesSlideShapes_: { args: ["request", "shapes_"], fn: (self: any, request: any, shapes_: any): any => {
        return $call(classMock, request, "pptx.slide.NotesSlideShapes", $kw({return_value: shapes_}));
      } },
    placeholder_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, NotesSlidePlaceholder);
      } },
    placeholders_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, NotesSlidePlaceholders);
      } },
    placeholders_prop_: { args: ["request", "placeholders_"], fn: (self: any, request: any, placeholders_: any): any => {
        return $call(propertyMock, request, NotesSlide, "placeholders", $kw({return_value: placeholders_}));
      } },
    shapes_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, NotesSlideShapes);
      } },
    text_frame_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, TextFrame);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_clone_the_notes_master_placeholders", args: ["request", "notes_master_", "shapes_"], fn: (self: any, request: any, notes_master_: any, shapes_: any): any => {
        let notes_slide: any, placeholders: any;
        const $t1 = [$call(BaseShape, element("p:sp/p:nvSpPr/p:nvPr/p:ph{type=body}"), null), $call(BaseShape, element("p:sp/p:nvSpPr/p:nvPr/p:ph{type=dt}"), null)];
        placeholders = $t1;
        notes_master_.placeholders = $t1;
        $call(propertyMock, request, NotesSlide, "shapes", $kw({return_value: shapes_}));
        notes_slide = $call(NotesSlide, null, null);
        notes_slide.cloneMasterPlaceholders(notes_master_);
        $assertEqual(shapes_.clonePlaceholder.callArgsList, [call($get(placeholders, 0))]);
      } },
    { name: "it_provides_access_to_its_shapes", args: ["shapes_fixture"], fn: (self: any, shapes_fixture: any): any => {
        let NotesSlideShapes_: any, notes_slide: any, shapes: any, shapes_: any, spTree: any;
        [notes_slide, NotesSlideShapes_, spTree, shapes_] = $list(shapes_fixture);
        shapes = notes_slide.shapes;
        NotesSlideShapes_.assertCalledOnceWith(spTree, notes_slide);
        $assert($is(shapes, shapes_));
      } },
    { name: "it_provides_access_to_its_placeholders", args: ["placeholders_fixture"], fn: (self: any, placeholders_fixture: any): any => {
        let NotesSlidePlaceholders_: any, notes_slide: any, placeholders: any, placeholders_: any, spTree: any;
        [notes_slide, NotesSlidePlaceholders_, spTree, placeholders_] = $list(placeholders_fixture);
        placeholders = notes_slide.placeholders;
        NotesSlidePlaceholders_.assertCalledOnceWith(spTree, notes_slide);
        $assert($is(placeholders, placeholders_));
      } },
    { name: "it_provides_access_to_its_notes_placeholder", args: ["notes_ph_fixture"], fn: (self: any, notes_ph_fixture: any): any => {
        let expected_value: any, notes_slide: any, placeholder: any;
        [notes_slide, expected_value] = $list(notes_ph_fixture);
        placeholder = notes_slide.notesPlaceholder;
        $assert($is(placeholder, expected_value));
      } },
    { name: "it_provides_access_to_its_notes_text_frame", args: ["notes_tf_fixture"], fn: (self: any, notes_tf_fixture: any): any => {
        let expected_value: any, notes_slide: any, text_frame: any;
        [notes_slide, expected_value] = $list(notes_tf_fixture);
        text_frame = notes_slide.notesTextFrame;
        $assert($is(text_frame, expected_value));
      } },
  ],
};
pytestClass("DescribeNotesSlide", DescribeNotesSlide);

const DescribeSlide: any = {
  fixtures: {
    background_fixture: { args: ["_BaseSlide_background_", "background_"], fn: (self: any, _BaseSlide_background_: any, background_: any): any => {
        let slide: any;
        slide = $call(Slide, null, null);
        _BaseSlide_background_.returnValue = background_;
        return [slide, _BaseSlide_background_, background_];
      } },
    follow_get_fixture: { args: ["request"], params: [["p:sld/p:cSld", true], ["p:sld/p:cSld/p:bg", false]], fn: (self: any, request: any): any => {
        let expected_value: any, pSld_cxml: any, slide: any;
        [pSld_cxml, expected_value] = $list(request.param);
        slide = $call(Slide, element(pSld_cxml), null);
        return [slide, expected_value];
      } },
    has_notes_slide_fixture: { args: ["part_prop_", "slide_part_"], fn: (self: any, part_prop_: any, slide_part_: any): any => {
        let expected_value: any, slide: any;
        slide = $call(Slide, null, null);
        const $t2 = 42;
        expected_value = $t2;
        slide_part_.hasNotesSlide = $t2;
        return [slide, expected_value];
      } },
    layout_fixture: { args: ["slide_layout_", "part_prop_", "slide_part_"], fn: (self: any, slide_layout_: any, part_prop_: any, slide_part_: any): any => {
        let slide: any;
        slide = $call(Slide, null, null);
        part_prop_.returnValue = slide_part_;
        slide_part_.slideLayout = slide_layout_;
        return [slide, slide_layout_];
      } },
    notes_slide_fixture: { args: ["notes_slide_", "part_prop_", "slide_part_"], fn: (self: any, notes_slide_: any, part_prop_: any, slide_part_: any): any => {
        let slide: any;
        slide = $call(Slide, null, null);
        slide_part_.notesSlide = notes_slide_;
        return [slide, notes_slide_];
      } },
    placeholders_fixture: { args: ["SlidePlaceholders_", "placeholders_"], fn: (self: any, SlidePlaceholders_: any, placeholders_: any): any => {
        let sld: any, slide: any, spTree: any;
        sld = element("p:sld/p:cSld/p:spTree");
        slide = $call(Slide, sld, null);
        spTree = $get(sld.xpath("//p:spTree"), 0);
        return [slide, SlidePlaceholders_, spTree, placeholders_];
      } },
    shapes_fixture: { args: ["SlideShapes_", "shapes_"], fn: (self: any, SlideShapes_: any, shapes_: any): any => {
        let sld: any, slide: any, spTree: any;
        sld = element("p:sld/p:cSld/p:spTree");
        spTree = $get(sld.xpath("//p:spTree"), 0);
        slide = $call(Slide, sld, null);
        return [slide, SlideShapes_, spTree, shapes_];
      } },
    slide_id_fixture: { args: ["part_prop_", "slide_part_"], fn: (self: any, part_prop_: any, slide_part_: any): any => {
        let slide: any, slide_id: any;
        slide = $call(Slide, null, null);
        slide_id = 256;
        slide_part_.slideId = slide_id;
        return [slide, slide_id];
      } },
    subclass_fixture: { args: [], fn: (self: any): any => {
        return $call(Slide, null, null);
      } },
    _BaseSlide_background_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, _BaseSlide, "background");
      } },
    background_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _Background);
      } },
    notes_slide_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, NotesSlide);
      } },
    part_prop_: { args: ["request", "slide_part_"], fn: (self: any, request: any, slide_part_: any): any => {
        return $call(propertyMock, request, Slide, "part", $kw({return_value: slide_part_}));
      } },
    placeholders_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlidePlaceholders);
      } },
    SlidePlaceholders_: { args: ["request", "placeholders_"], fn: (self: any, request: any, placeholders_: any): any => {
        return $call(classMock, request, "pptx.slide.SlidePlaceholders", $kw({return_value: placeholders_}));
      } },
    SlideShapes_: { args: ["request", "shapes_"], fn: (self: any, request: any, shapes_: any): any => {
        return $call(classMock, request, "pptx.slide.SlideShapes", $kw({return_value: shapes_}));
      } },
    shapes_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlideShapes);
      } },
    slide_layout_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlideLayout);
      } },
    slide_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlidePart);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_is_a_BaseSlide_subclass", args: ["subclass_fixture"], fn: (self: any, subclass_fixture: any): any => {
        let slide: any;
        slide = subclass_fixture;
        $assert($isinstance(slide, _BaseSlide));
      } },
    { name: "it_provides_access_to_its_background", args: ["background_fixture"], fn: (self: any, background_fixture: any): any => {
        let _BaseSlide_background_: any, background: any, background_: any, slide: any;
        [slide, _BaseSlide_background_, background_] = $list(background_fixture);
        background = slide.background;
        _BaseSlide_background_.assertCalledOnceWith();
        $assert($is(background, background_));
      } },
    { name: "it_knows_whether_it_follows_the_mstr_bkgd", args: ["follow_get_fixture"], fn: (self: any, follow_get_fixture: any): any => {
        let expected_value: any, follows: any, slide: any;
        [slide, expected_value] = $list(follow_get_fixture);
        follows = slide.followMasterBackground;
        $assert($is(follows, expected_value));
      } },
    { name: "it_knows_whether_it_has_a_notes_slide", args: ["has_notes_slide_fixture"], fn: (self: any, has_notes_slide_fixture: any): any => {
        let expected_value: any, slide: any;
        [slide, expected_value] = $list(has_notes_slide_fixture);
        $assert($is(slide.hasNotesSlide, expected_value));
      } },
    { name: "it_knows_its_slide_id", args: ["slide_id_fixture"], fn: (self: any, slide_id_fixture: any): any => {
        let expected_value: any, slide: any;
        [slide, expected_value] = $list(slide_id_fixture);
        $assertEqual(slide.slideId, expected_value);
      } },
    { name: "it_provides_access_to_its_shapes", args: ["shapes_fixture"], fn: (self: any, shapes_fixture: any): any => {
        let SlideShapes_: any, shapes: any, shapes_: any, slide: any, spTree: any;
        [slide, SlideShapes_, spTree, shapes_] = $list(shapes_fixture);
        shapes = slide.shapes;
        SlideShapes_.assertCalledOnceWith(spTree, slide);
        $assert($is(shapes, shapes_));
      } },
    { name: "it_provides_access_to_its_placeholders", args: ["placeholders_fixture"], fn: (self: any, placeholders_fixture: any): any => {
        let SlidePlaceholders_: any, placeholders: any, placeholders_: any, slide: any, spTree: any;
        [slide, SlidePlaceholders_, spTree, placeholders_] = $list(placeholders_fixture);
        placeholders = slide.placeholders;
        SlidePlaceholders_.assertCalledOnceWith(spTree, slide);
        $assert($is(placeholders, placeholders_));
      } },
    { name: "it_provides_access_to_its_slide_layout", args: ["layout_fixture"], fn: (self: any, layout_fixture: any): any => {
        let slide: any, slide_layout_: any;
        [slide, slide_layout_] = $list(layout_fixture);
        $assert($is(slide.slideLayout, slide_layout_));
      } },
    { name: "it_provides_access_to_its_notes_slide", args: ["notes_slide_fixture"], fn: (self: any, notes_slide_fixture: any): any => {
        let notes_slide_: any, slide: any;
        [slide, notes_slide_] = $list(notes_slide_fixture);
        $assert($is(slide.notesSlide, notes_slide_));
      } },
  ],
};
pytestClass("DescribeSlide", DescribeSlide);

const DescribeSlides: any = {
  fixtures: {
    add_fixture: { args: ["slide_layout_", "part_prop_", "slide_"], fn: (self: any, slide_layout_: any, part_prop_: any, slide_: any): any => {
        let clone_layout_placeholders_: any, expected_xml: any, part_: any, slides: any;
        slides = $call(Slides, element("p:sldIdLst/p:sldId{r:id=rId1}"), null);
        part_ = part_prop_.returnValue;
        clone_layout_placeholders_ = slide_.shapes.cloneLayoutPlaceholders;
        expected_xml = xml("p:sldIdLst/(p:sldId{r:id=rId1},p:sldId{r:id=rId2,id=256})");
        part_.addSlide.returnValue = ["rId2", slide_];
        return [slides, slide_layout_, part_, clone_layout_placeholders_, expected_xml, slide_];
      } },
    get_fixture: { args: ["request", "part_prop_", "prs_part_", "slide_"], params: [true, false], fn: (self: any, request: any, part_prop_: any, prs_part_: any, slide_: any): any => {
        let default$: any, expected_value: any, found: any, slide_id: any, slides: any;
        found = request.param;
        slides = $call(Slides, null, null);
        [slide_id, default$] = $list([256, "foobar"]);
        expected_value = ($truthy(found) ? slide_ : default$);
        prs_part_.getSlide.returnValue = ($truthy(found) ? slide_ : null);
        return [slides, slide_id, default$, prs_part_, expected_value];
      } },
    getitem_fixture: { args: ["prs_part_", "slide_", "part_prop_"], fn: (self: any, prs_part_: any, slide_: any, part_prop_: any): any => {
        let sldIdLst: any, slides: any;
        sldIdLst = element("p:sldIdLst/p:sldId{r:id=rId1}");
        slides = $call(Slides, sldIdLst, null);
        prs_part_.relatedSlide.returnValue = slide_;
        return [slides, prs_part_, "rId1", slide_];
      } },
    getitem_raises_fixture: { args: [], fn: (self: any): any => {
        let sldIdLst: any, slides: any;
        sldIdLst = element("p:sldIdLst/p:sldId{r:id=rId1}");
        slides = $call(Slides, sldIdLst, null);
        return slides;
      } },
    index_fixture: { args: ["request", "part_prop_"], params: [0, 1], fn: (self: any, request: any, part_prop_: any): any => {
        let _slides: any, idx: any, sldIdLst: any, slides: any;
        idx = request.param;
        sldIdLst = element("p:sldIdLst/(p:sldId{r:id=a},p:sldId{r:id=b})");
        slides = $call(Slides, sldIdLst, null);
        _slides = [$call(Slide, element("p:sld"), null), $call(Slide, element("p:sld"), null)];
        part_prop_.returnValue.relatedSlide.sideEffect = _slides;
        return [slides, $get(_slides, idx), idx];
      } },
    iter_fixture: { args: ["part_prop_", "slide_"], fn: (self: any, part_prop_: any, slide_: any): any => {
        let _slides: any, calls: any, related_slide_: any, sldIdLst: any, slides: any;
        sldIdLst = element("p:sldIdLst/(p:sldId{r:id=a},p:sldId{r:id=b})");
        slides = $call(Slides, sldIdLst, null);
        related_slide_ = part_prop_.returnValue.relatedSlide;
        related_slide_.returnValue = slide_;
        calls = [call("a"), call("b")];
        _slides = [slide_, slide_];
        return [slides, related_slide_, calls, _slides];
      } },
    len_fixture: { args: ["request"], params: [["p:sldIdLst", 0], ["p:sldIdLst/p:sldId{r:id=a}", 1], ["p:sldIdLst/(p:sldId{r:id=a},p:sldId{r:id=b})", 2]], fn: (self: any, request: any): any => {
        let expected_value: any, sldIdLst_cxml: any, slides: any;
        [sldIdLst_cxml, expected_value] = $list(request.param);
        slides = $call(Slides, element(sldIdLst_cxml), null);
        return [slides, expected_value];
      } },
    raises_fixture: { args: [], fn: (self: any): any => {
        let slide: any, slides: any;
        slides = $call(Slides, element("p:sldIdLst"), null);
        slide = $call(Slide, element("p:sld"), null);
        return [slides, slide];
      } },
    part_prop_: { args: ["request", "prs_part_"], fn: (self: any, request: any, prs_part_: any): any => {
        return $call(propertyMock, request, Slides, "part", $kw({return_value: prs_part_}));
      } },
    prs_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, PresentationPart);
      } },
    slide_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Slide);
      } },
    slide_layout_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlideLayout);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_supports_indexed_access", args: ["getitem_fixture"], fn: (self: any, getitem_fixture: any): any => {
        let prs_part_: any, rId: any, slide: any, slide_: any, slides: any;
        [slides, prs_part_, rId, slide_] = $list(getitem_fixture);
        slide = $get(slides, 0);
        prs_part_.relatedSlide.assertCalledOnceWith(rId);
        $assert($is(slide, slide_));
      } },
    { name: "it_raises_on_slide_index_out_of_range", args: ["getitem_raises_fixture"], fn: (self: any, getitem_raises_fixture: any): any => {
        let slides: any;
        slides = getitem_raises_fixture;
        $raises(IndexError, () => {
          $get(slides, 2);
        });
      } },
    { name: "it_knows_the_index_of_a_slide_it_contains", args: ["index_fixture"], fn: (self: any, index_fixture: any): any => {
        let expected_value: any, index: any, slide: any, slides: any;
        [slides, slide, expected_value] = $list(index_fixture);
        index = $index(slides, slide);
        $assertEqual(index, expected_value);
      } },
    { name: "it_raises_on_slide_not_in_collection", args: ["raises_fixture"], fn: (self: any, raises_fixture: any): any => {
        let slide: any, slides: any;
        [slides, slide] = $list(raises_fixture);
        $raises(ValueError, () => {
          $index(slides, slide);
        });
      } },
    { name: "it_can_iterate_its_slides", args: ["iter_fixture"], fn: (self: any, iter_fixture: any): any => {
        let calls: any, expected_value: any, related_slide_: any, slide_lst: any, slides: any;
        [slides, related_slide_, calls, expected_value] = $list(iter_fixture);
        slide_lst = $list(slides).map((s: any) => s);
        $assertEqual(related_slide_.callArgsList, calls);
        $assertEqual(slide_lst, expected_value);
      } },
    { name: "it_supports_len", args: ["len_fixture"], fn: (self: any, len_fixture: any): any => {
        let expected_value: any, slides: any;
        [slides, expected_value] = $list(len_fixture);
        $assertEqual($len(slides), expected_value);
      } },
    { name: "it_can_add_a_new_slide", args: ["add_fixture"], fn: (self: any, add_fixture: any): any => {
        let clone_layout_placeholders_: any, expected_xml: any, part_: any, slide: any, slide_: any, slide_layout_: any, slides: any;
        [slides, slide_layout_, part_] = $list($slice(add_fixture, null, 3, null));
        [clone_layout_placeholders_, expected_xml, slide_] = $list($slice(add_fixture, 3, null, null));
        slide = slides.addSlide(slide_layout_);
        part_.addSlide.assertCalledOnceWith(slide_layout_);
        clone_layout_placeholders_.assertCalledOnceWith(slide_layout_);
        $assertEqual(slides._sldIdLst.xml, expected_xml);
        $assert($is(slide, slide_));
      } },
    { name: "it_finds_a_slide_by_slide_id", args: ["get_fixture"], fn: (self: any, get_fixture: any): any => {
        let default$: any, expected_value: any, prs_part_: any, slide: any, slide_id: any, slides: any;
        [slides, slide_id, default$, prs_part_, expected_value] = $list(get_fixture);
        slide = $dictget(slides, slide_id, default$);
        prs_part_.getSlide.assertCalledOnceWith(slide_id);
        $assert($is(slide, expected_value));
      } },
  ],
};
pytestClass("DescribeSlides", DescribeSlides);

const DescribeSlideLayout: any = {
  fixtures: {
    cloneable_fixture: { args: ["request", "placeholders_prop_", "placeholder_", "placeholder_2_"], params: [[[PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.BODY], [0, 1]], [[PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.DATE], [0]], [[PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.OBJECT], [1]], [[PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.FOOTER], []]], fn: (self: any, request: any, placeholders_prop_: any, placeholder_: any, placeholder_2_: any): any => {
        let _placeholders: any, expected_indices: any, expected_placeholders: any, ph_types: any, slide_layout: any;
        [ph_types, expected_indices] = $list(request.param);
        slide_layout = $call(SlideLayout, null, null);
        placeholder_.element.phType = $get(ph_types, 0);
        placeholder_2_.element.phType = $get(ph_types, 1);
        _placeholders = [placeholder_, placeholder_2_];
        expected_placeholders = $list(expected_indices).map((idx: any) => $get(_placeholders, idx));
        placeholders_prop_.returnValue = _placeholders;
        return [slide_layout, expected_placeholders];
      } },
    used_by_fixture: { args: ["request", "presentation_", "slide_", "slide_2_"], params: [[], [0], [1], [0, 1]], fn: (self: any, request: any, presentation_: any, slide_: any, slide_2_: any): any => {
        let expected_value: any, slide_layout: any, slides: any, used_by_idxs: any;
        used_by_idxs = request.param;
        slides = [slide_, slide_2_];
        slide_layout = $call(SlideLayout, null, null);
        for (const [idx, s] of $list($enumerate(slides))) {
          s.slideLayout = ($truthy($in(idx, used_by_idxs)) ? slide_layout : null);
        }
        presentation_.slides = slides;
        expected_value = $list($list($enumerate(slides)).filter(([i, s]: any) => $truthy($in(i, used_by_idxs))).map(([i, s]: any) => s));
        return [presentation_, slide_layout, expected_value];
      } },
    LayoutPlaceholders_: { args: ["request", "placeholders_"], fn: (self: any, request: any, placeholders_: any): any => {
        return $call(classMock, request, "pptx.slide.LayoutPlaceholders", $kw({return_value: placeholders_}));
      } },
    LayoutShapes_: { args: ["request", "shapes_"], fn: (self: any, request: any, shapes_: any): any => {
        return $call(classMock, request, "pptx.slide.LayoutShapes", $kw({return_value: shapes_}));
      } },
    package_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Package);
      } },
    part_prop_: { args: ["request", "slide_layout_part_"], fn: (self: any, request: any, slide_layout_part_: any): any => {
        return $call(propertyMock, request, SlideLayout, "part", $kw({return_value: slide_layout_part_}));
      } },
    placeholder_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, LayoutPlaceholder);
      } },
    placeholder_2_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, LayoutPlaceholder);
      } },
    placeholders_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, LayoutPlaceholders);
      } },
    placeholders_prop_: { args: ["request", "placeholders_"], fn: (self: any, request: any, placeholders_: any): any => {
        return $call(propertyMock, request, SlideLayout, "placeholders", $kw({return_value: placeholders_}));
      } },
    presentation_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Presentation);
      } },
    presentation_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, PresentationPart);
      } },
    shapes_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, LayoutShapes);
      } },
    slide_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Slide);
      } },
    slide_2_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Slide);
      } },
    slide_layout_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlideLayoutPart);
      } },
    slide_master_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlideMaster);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_is_a_BaseSlide_subclass", args: [], fn: (self: any): any => {
        let slide_layout: any;
        slide_layout = $call(SlideLayout, null, null);
        $assert($isinstance(slide_layout, _BaseSlide));
      } },
    { name: "it_can_iterate_its_clonable_placeholders", args: ["cloneable_fixture"], fn: (self: any, cloneable_fixture: any): any => {
        let cloneable: any, expected_placeholders: any, slide_layout: any;
        [slide_layout, expected_placeholders] = $list(cloneable_fixture);
        cloneable = $list(slide_layout.iterCloneablePlaceholders());
        $assertEqual(cloneable, expected_placeholders);
      } },
    { name: "it_provides_access_to_its_placeholders", args: ["LayoutPlaceholders_", "placeholders_"], fn: (self: any, LayoutPlaceholders_: any, placeholders_: any): any => {
        let placeholders: any, sldLayout: any, slide_layout: any, spTree: any;
        sldLayout = element("p:sldLayout/p:cSld/p:spTree");
        spTree = $get(sldLayout.xpath("//p:spTree"), 0);
        slide_layout = $call(SlideLayout, sldLayout, null);
        placeholders = slide_layout.placeholders;
        LayoutPlaceholders_.assertCalledOnceWith(spTree, slide_layout);
        $assert($is(placeholders, placeholders_));
      } },
    { name: "it_provides_access_to_its_shapes", args: ["LayoutShapes_", "shapes_"], fn: (self: any, LayoutShapes_: any, shapes_: any): any => {
        let shapes: any, sldLayout: any, slide_layout: any, spTree: any;
        sldLayout = element("p:sldLayout/p:cSld/p:spTree");
        spTree = $get(sldLayout.xpath("//p:spTree"), 0);
        slide_layout = $call(SlideLayout, sldLayout, null);
        shapes = slide_layout.shapes;
        LayoutShapes_.assertCalledOnceWith(spTree, slide_layout);
        $assert($is(shapes, shapes_));
      } },
    { name: "it_provides_access_to_its_slide_master", args: ["slide_master_", "part_prop_"], fn: (self: any, slide_master_: any, part_prop_: any): any => {
        let slide_layout: any, slide_master: any;
        part_prop_.returnValue.slideMaster = slide_master_;
        slide_layout = $call(SlideLayout, null, null);
        slide_master = slide_layout.slideMaster;
        $assert($is(slide_master, slide_master_));
      } },
    { name: "it_knows_which_slides_are_based_on_it", args: ["used_by_fixture", "part_prop_", "slide_layout_part_", "package_", "presentation_part_", "presentation_"], fn: (self: any, used_by_fixture: any, part_prop_: any, slide_layout_part_: any, package_: any, presentation_part_: any, presentation_: any): any => {
        let expected_value: any, slide_layout: any, used_by_slides: any;
        [presentation_, slide_layout, expected_value] = $list(used_by_fixture);
        part_prop_.returnValue = slide_layout_part_;
        slide_layout_part_.package = package_;
        package_.presentationPart = presentation_part_;
        presentation_part_.presentation = presentation_;
        used_by_slides = slide_layout.usedBySlides;
        $assertEqual(used_by_slides, expected_value);
      } },
  ],
};
pytestClass("DescribeSlideLayout", DescribeSlideLayout);

const DescribeSlideLayouts: any = {
  fixtures: {
    len_fixture: { args: ["request"], params: [["p:sldLayoutIdLst", 0], ["p:sldLayoutIdLst/p:sldLayoutId", 1], ["p:sldLayoutIdLst/(p:sldLayoutId,p:sldLayoutId)", 2]], fn: (self: any, request: any): any => {
        let expected_value: any, sldLayoutIdLst_cxml: any, slide_layouts: any;
        [sldLayoutIdLst_cxml, expected_value] = $list(request.param);
        slide_layouts = $call(SlideLayouts, element(sldLayoutIdLst_cxml), null);
        return [slide_layouts, expected_value];
      } },
    index_: { args: ["request"], fn: (self: any, request: any): any => {
        return methodMock(request, SlideLayouts, "index");
      } },
    _iter_: { args: ["request"], fn: (self: any, request: any): any => {
        return methodMock(request, SlideLayouts, "__iter__");
      } },
    part_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, SlideLayouts, "part");
      } },
    slide_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Slide);
      } },
    slide_layout_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlideLayout);
      } },
    slide_layout_2_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlideLayout);
      } },
    slide_master_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlideMaster);
      } },
    slide_master_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlideMasterPart);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_supports_len", args: ["len_fixture"], fn: (self: any, len_fixture: any): any => {
        let expected_value: any, slide_layouts: any;
        [slide_layouts, expected_value] = $list(len_fixture);
        $assertEqual($len(slide_layouts), expected_value);
      } },
    { name: "it_can_iterate_its_slide_layouts", args: ["part_prop_", "slide_master_part_"], fn: (self: any, part_prop_: any, slide_master_part_: any): any => {
        let _slide_layouts: any, related_slide_layout_: any, sldLayoutIdLst: any, slide_layout_lst: any, slide_layouts: any;
        sldLayoutIdLst = element("p:sldLayoutIdLst/(p:sldLayoutId{r:id=a},p:sldLayoutId{r:id=b})");
        _slide_layouts = [$call(SlideLayout, element("p:sldLayout"), null), $call(SlideLayout, element("p:sldLayout"), null)];
        part_prop_.returnValue = slide_master_part_;
        related_slide_layout_ = slide_master_part_.relatedSlideLayout;
        related_slide_layout_.sideEffect = _slide_layouts;
        slide_layouts = $call(SlideLayouts, sldLayoutIdLst, null);
        slide_layout_lst = $list(slide_layouts).map((sl: any) => sl);
        $assertEqual(related_slide_layout_.callArgsList, [call("a"), call("b")]);
        $assertEqual(slide_layout_lst, _slide_layouts);
      } },
    { name: "it_supports_indexed_access", args: ["slide_layout_", "part_prop_", "slide_master_part_"], fn: (self: any, slide_layout_: any, part_prop_: any, slide_master_part_: any): any => {
        let slide_layout: any, slide_layouts: any;
        part_prop_.returnValue = slide_master_part_;
        slide_master_part_.relatedSlideLayout.returnValue = slide_layout_;
        slide_layouts = $call(SlideLayouts, element("p:sldLayoutIdLst/p:sldLayoutId{r:id=rId1}"), null);
        slide_layout = $get(slide_layouts, 0);
        slide_master_part_.relatedSlideLayout.assertCalledOnceWith("rId1");
        $assert($is(slide_layout, slide_layout_));
      } },
    { name: "but_it_raises_on_index_out_of_range", args: ["part_prop_"], fn: (self: any, part_prop_: any): any => {
        let slide_layouts: any;
        slide_layouts = $call(SlideLayouts, element("p:sldLayoutIdLst/p:sldLayoutId{r:id=rId1}"), null);
        $raises(IndexError, () => {
          $get(slide_layouts, 1);
        });
      } },
    { name: "it_can_find_a_slide_layout_by_name", args: ["_iter_", "slide_layout_", "slide_layout_2_"], fn: (self: any, _iter_: any, slide_layout_: any, slide_layout_2_: any): any => {
        let slide_layout: any, slide_layouts: any;
        _iter_.returnValue = $iter([slide_layout_, slide_layout_2_]);
        slide_layout_2_.name = "pick me!";
        slide_layouts = $call(SlideLayouts, null, null);
        slide_layout = slide_layouts.getByName("pick me!");
        $assert($is(slide_layout, slide_layout_2_));
      } },
    { name: "but_it_returns_the_default_value_when_no_layout_has_that_name", args: ["_iter_", "slide_layout_", "slide_layout_2_"], fn: (self: any, _iter_: any, slide_layout_: any, slide_layout_2_: any): any => {
        let slide_layout: any, slide_layouts: any;
        _iter_.returnValue = $iter([slide_layout_, slide_layout_2_]);
        slide_layout_2_.name = "not the droid you're looking for";
        slide_layouts = $call(SlideLayouts, null, null);
        slide_layout = slide_layouts.getByName("pick me!");
        $assert($is(slide_layout, null));
        slide_layout = slide_layouts.getByName("pick me!", "default-value");
        $assertEqual(slide_layout, "default-value");
      } },
    { name: "it_knows_the_index_of_each_of_its_slide_layouts", args: ["_iter_", "slide_layout_", "slide_layout_2_"], fn: (self: any, _iter_: any, slide_layout_: any, slide_layout_2_: any): any => {
        let index: any, slide_layouts: any;
        _iter_.returnValue = $iter([slide_layout_, slide_layout_2_]);
        slide_layouts = $call(SlideLayouts, null, null);
        index = $index(slide_layouts, slide_layout_2_);
        $assertEqual(index, 1);
      } },
    { name: "but_it_raises_on_slide_layout_not_in_collection", args: ["_iter_", "slide_layout_", "slide_layout_2_"], fn: (self: any, _iter_: any, slide_layout_: any, slide_layout_2_: any): any => {
        let e: any, slide_layouts: any;
        _iter_.returnValue = $iter([slide_layout_]);
        slide_layouts = $call(SlideLayouts, null, null);
        e = $raises(ValueError, () => {
          $index(slide_layouts, slide_layout_2_);
        });
        $assertEqual($str(e.value), "layout not in this SlideLayouts collection");
      } },
    { name: "it_can_remove_an_unused_slide_layout", args: ["slide_layout_", "index_", "slide_master_", "slide_master_part_"], fn: (self: any, slide_layout_: any, index_: any, slide_master_: any, slide_master_part_: any): any => {
        let sldLayoutIdLst: any, slide_layouts: any;
        slide_layout_.usedBySlides = [];
        index_.returnValue = 0;
        sldLayoutIdLst = element("p:sldLayoutIdLst/(p:sldLayoutId{r:id=rId1},p:sldLayoutId{r:id=rId2})");
        slide_layout_.slideMaster = slide_master_;
        slide_master_.part = slide_master_part_;
        slide_layouts = $call(SlideLayouts, sldLayoutIdLst, null);
        $remove(slide_layouts, slide_layout_);
        $assertEqual(slide_layouts._sldLayoutIdLst.xml, xml("p:sldLayoutIdLst/p:sldLayoutId{r:id=rId2}"));
        slide_master_part_.dropRel.assertCalledOnceWith("rId1");
      } },
    { name: "but_it_raises_on_attempt_to_remove_slide_layout_in_use", args: ["slide_layout_", "slide_"], fn: (self: any, slide_layout_: any, slide_: any): any => {
        let slide_layouts: any;
        slide_layout_.usedBySlides = [slide_];
        slide_layouts = $call(SlideLayouts, null, null);
        $raises(ValueError, () => {
          $remove(slide_layouts, slide_layout_);
        });
      } },
  ],
};
pytestClass("DescribeSlideLayouts", DescribeSlideLayouts);

const DescribeSlideMaster: any = {
  fixtures: {
    layouts_fixture: { args: ["SlideLayouts_", "slide_layouts_"], fn: (self: any, SlideLayouts_: any, slide_layouts_: any): any => {
        let sldMaster: any, sldMasterIdLst: any, slide_master: any;
        sldMaster = element("p:sldMaster/p:sldLayoutIdLst");
        slide_master = $call(SlideMaster, sldMaster, null);
        sldMasterIdLst = sldMaster.sldLayoutIdLst;
        return [slide_master, SlideLayouts_, sldMasterIdLst, slide_layouts_];
      } },
    subclass_fixture: { args: [], fn: (self: any): any => {
        return $call(SlideMaster, null, null);
      } },
    SlideLayouts_: { args: ["request", "slide_layouts_"], fn: (self: any, request: any, slide_layouts_: any): any => {
        return $call(classMock, request, "pptx.slide.SlideLayouts", $kw({return_value: slide_layouts_}));
      } },
    slide_layouts_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlideLayouts);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_is_a_BaseMaster_subclass", args: ["subclass_fixture"], fn: (self: any, subclass_fixture: any): any => {
        let slide_master: any;
        slide_master = subclass_fixture;
        $assert($isinstance(slide_master, _BaseMaster));
      } },
    { name: "it_provides_access_to_its_slide_layouts", args: ["layouts_fixture"], fn: (self: any, layouts_fixture: any): any => {
        let SlideLayouts_: any, sldLayoutIdLst: any, slide_layouts: any, slide_layouts_: any, slide_master: any;
        [slide_master, SlideLayouts_, sldLayoutIdLst, slide_layouts_] = $list(layouts_fixture);
        slide_layouts = slide_master.slideLayouts;
        SlideLayouts_.assertCalledOnceWith(sldLayoutIdLst, slide_master);
        $assert($is(slide_layouts, slide_layouts_));
      } },
  ],
};
pytestClass("DescribeSlideMaster", DescribeSlideMaster);

const DescribeSlideMasters: any = {
  fixtures: {
    getitem_fixture: { args: ["part_", "slide_master_", "part_prop_"], fn: (self: any, part_: any, slide_master_: any, part_prop_: any): any => {
        let slide_masters: any;
        slide_masters = $call(SlideMasters, element("p:sldMasterIdLst/p:sldMasterId{r:id=rId1}"), null);
        part_.relatedSlideMaster.returnValue = slide_master_;
        return [slide_masters, part_, slide_master_, "rId1"];
      } },
    getitem_raises_fixture: { args: ["part_prop_"], fn: (self: any, part_prop_: any): any => {
        return $call(SlideMasters, element("p:sldMasterIdLst/p:sldMasterId{r:id=rId1}"), null);
      } },
    iter_fixture: { args: ["part_prop_"], fn: (self: any, part_prop_: any): any => {
        let _slide_masters: any, calls: any, related_slide_master_: any, sldMasterIdLst: any, slide_masters: any;
        sldMasterIdLst = element("p:sldMasterIdLst/(p:sldMasterId{r:id=a},p:sldMasterId{r:id=b})");
        slide_masters = $call(SlideMasters, sldMasterIdLst, null);
        related_slide_master_ = part_prop_.returnValue.relatedSlideMaster;
        calls = [call("a"), call("b")];
        _slide_masters = [$call(SlideMaster, element("p:sldMaster"), null), $call(SlideMaster, element("p:sldMaster"), null)];
        related_slide_master_.sideEffect = _slide_masters;
        return [slide_masters, related_slide_master_, calls, _slide_masters];
      } },
    len_fixture: { args: ["request"], params: [["p:sldMasterIdLst", 0], ["p:sldMasterIdLst/p:sldMasterId", 1], ["p:sldMasterIdLst/(p:sldMasterId,p:sldMasterId)", 2]], fn: (self: any, request: any): any => {
        let expected_value: any, sldMasterIdLst_cxml: any, slide_masters: any;
        [sldMasterIdLst_cxml, expected_value] = $list(request.param);
        slide_masters = $call(SlideMasters, element(sldMasterIdLst_cxml), null);
        return [slide_masters, expected_value];
      } },
    part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, PresentationPart);
      } },
    part_prop_: { args: ["request", "part_"], fn: (self: any, request: any, part_: any): any => {
        return $call(propertyMock, request, SlideMasters, "part", $kw({return_value: part_}));
      } },
    slide_master_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlideMaster);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_how_many_masters_it_contains", args: ["len_fixture"], fn: (self: any, len_fixture: any): any => {
        let expected_value: any, slide_masters: any;
        [slide_masters, expected_value] = $list(len_fixture);
        $assertEqual($len(slide_masters), expected_value);
      } },
    { name: "it_can_iterate_the_slide_masters", args: ["iter_fixture"], fn: (self: any, iter_fixture: any): any => {
        let _slide_masters: any, calls: any, expected_values: any, related_slide_master_: any, slide_masters: any;
        [slide_masters, related_slide_master_, calls, expected_values] = $list(iter_fixture);
        _slide_masters = $list(slide_masters).map((sm: any) => sm);
        $assertEqual(related_slide_master_.callArgsList, calls);
        $assertEqual(_slide_masters, expected_values);
      } },
    { name: "it_supports_indexed_access", args: ["getitem_fixture"], fn: (self: any, getitem_fixture: any): any => {
        let part_: any, rId: any, slide_master: any, slide_master_: any, slide_masters: any;
        [slide_masters, part_, slide_master_, rId] = $list(getitem_fixture);
        slide_master = $get(slide_masters, 0);
        part_.relatedSlideMaster.assertCalledOnceWith(rId);
        $assert($is(slide_master, slide_master_));
      } },
    { name: "it_raises_on_index_out_of_range", args: ["getitem_raises_fixture"], fn: (self: any, getitem_raises_fixture: any): any => {
        let slides: any;
        slides = getitem_raises_fixture;
        $raises(IndexError, () => {
          $get(slides, 1);
        });
      } },
  ],
};
pytestClass("DescribeSlideMasters", DescribeSlideMasters);

const Describe_Background: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_access_to_its_fill", args: ["request", "cSld_xml", "expected_cxml"], parametrize: [{ names: ["cSld_xml", "expected_cxml"], values: [["p:cSld{a:b=c}", "p:cSld{a:b=c}/p:bg/p:bgPr/(a:noFill,a:effectLst)"], ["p:cSld{a:b=c}/p:bg/p:bgRef", "p:cSld{a:b=c}/p:bg/p:bgPr/(a:noFill,a:effectLst)"], ["p:cSld/p:bg/p:bgPr/a:solidFill", "p:cSld/p:bg/p:bgPr/a:solidFill"]] }], fn: (self: any, request: any, cSld_xml: any, expected_cxml: any): any => {
        let background: any, cSld: any, fill: any, fill_: any, from_fill_parent_: any;
        fill_ = instanceMock(request, FillFormat);
        from_fill_parent_ = $call(methodMock, request, FillFormat, "from_fill_parent", $kw({autospec: false, return_value: fill_}));
        cSld = element(cSld_xml);
        background = $call(_Background, cSld);
        fill = background.fill;
        $assertEqual(cSld.xml, xml(expected_cxml));
        from_fill_parent_.assertCalledOnceWith($get(cSld.xpath("p:bg/p:bgPr"), 0));
        $assert($is(fill, fill_));
      } },
  ],
};
pytestClass("Describe_Background", Describe_Background);

