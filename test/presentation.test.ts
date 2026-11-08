// Machine-translated from python-pptx tests/test_presentation.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "./pyshim.js";
import { $assert, $assertEqual, $call, $get, $is, $kw, $list, $slice, classMock, instanceMock, propertyMock, pytestClass } from "./pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../src/exc.js";
import { defineChildren } from "../src/oxml/xmlchemy.js";
import { CorePropertiesPart as CorePropertiesPart$imported } from "../src/parts/coreprops.js";
import { PresentationPart as PresentationPart$imported } from "../src/parts/presentation.js";
import { NotesMasterPart as NotesMasterPart$imported } from "../src/parts/slide.js";
import { Presentation as Presentation$imported } from "../src/presentation.js";
import { SlideLayouts as SlideLayouts$imported, SlideMaster as SlideMaster$imported, SlideMasters as SlideMasters$imported, Slides as Slides$imported } from "../src/slide.js";
import { element, xml } from "./unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const CorePropertiesPart = CorePropertiesPart$imported;
const PresentationPart = PresentationPart$imported;
const NotesMasterPart = NotesMasterPart$imported;
const Presentation = Presentation$imported;
const SlideLayouts = SlideLayouts$imported;
const SlideMaster = SlideMaster$imported;
const SlideMasters = SlideMasters$imported;
const Slides = Slides$imported;
void $py;


const DescribePresentation: any = {
  fixtures: {
    core_props_fixture: { args: ["prs_part_", "core_properties_"], fn: (self: any, prs_part_: any, core_properties_: any): any => {
        let prs: any;
        prs = $call(Presentation, null, prs_part_);
        prs_part_.coreProperties = core_properties_;
        return [prs, core_properties_];
      } },
    layouts_fixture: { args: ["masters_prop_", "slide_layouts_"], fn: (self: any, masters_prop_: any, slide_layouts_: any): any => {
        let prs: any;
        prs = $call(Presentation, null, null);
        masters_prop_.returnValue.__getitem__.returnValue.slideLayouts = slide_layouts_;
        return [prs, slide_layouts_];
      } },
    master_fixture: { args: ["masters_prop_", "slide_master_"], fn: (self: any, masters_prop_: any, slide_master_: any): any => {
        let getitem_: any, prs: any;
        prs = $call(Presentation, null, null);
        getitem_ = masters_prop_.returnValue.__getitem__;
        getitem_.returnValue = slide_master_;
        return [prs, getitem_, slide_master_];
      } },
    masters_fixture: { args: ["request", "SlideMasters_", "slide_masters_"], params: [["p:presentation", "p:presentation/p:sldMasterIdLst"], ["p:presentation/p:sldMasterIdLst", "p:presentation/p:sldMasterIdLst"]], fn: (self: any, request: any, SlideMasters_: any, slide_masters_: any): any => {
        let expected_cxml: any, expected_xml: any, prs: any, prs_cxml: any;
        [prs_cxml, expected_cxml] = $list(request.param);
        prs = $call(Presentation, element(prs_cxml), null);
        expected_xml = xml(expected_cxml);
        return [prs, SlideMasters_, slide_masters_, expected_xml];
      } },
    notes_master_fixture: { args: ["prs_part_", "notes_master_"], fn: (self: any, prs_part_: any, notes_master_: any): any => {
        let prs: any;
        prs = $call(Presentation, null, prs_part_);
        prs_part_.notesMaster = notes_master_;
        return [prs, notes_master_];
      } },
    part_fixture: { args: ["prs_part_"], fn: (self: any, prs_part_: any): any => {
        let prs: any;
        prs = $call(Presentation, null, prs_part_);
        return [prs, prs_part_];
      } },
    save_fixture: { args: ["prs_part_"], fn: (self: any, prs_part_: any): any => {
        let file_: any, prs: any;
        prs = $call(Presentation, null, prs_part_);
        file_ = "foobar.docx";
        return [prs, file_, prs_part_];
      } },
    sld_height_get_fixture: { args: ["request"], params: [["p:presentation", null], ["p:presentation/p:sldSz{cy=42}", 42]], fn: (self: any, request: any): any => {
        let expected_value: any, prs: any, prs_cxml: any;
        [prs_cxml, expected_value] = $list(request.param);
        prs = $call(Presentation, element(prs_cxml), null);
        return [prs, expected_value];
      } },
    sld_height_set_fixture: { args: ["request"], params: [["p:presentation", "p:presentation/p:sldSz{cy=914400}"], ["p:presentation/p:sldSz{cy=424242}", "p:presentation/p:sldSz{cy=914400}"]], fn: (self: any, request: any): any => {
        let expected_cxml: any, expected_xml: any, prs: any, prs_cxml: any;
        [prs_cxml, expected_cxml] = $list(request.param);
        prs = $call(Presentation, element(prs_cxml), null);
        expected_xml = xml(expected_cxml);
        return [prs, 914400, expected_xml];
      } },
    sld_width_get_fixture: { args: ["request"], params: [["p:presentation", null], ["p:presentation/p:sldSz{cx=42}", 42]], fn: (self: any, request: any): any => {
        let expected_value: any, prs: any, prs_cxml: any;
        [prs_cxml, expected_value] = $list(request.param);
        prs = $call(Presentation, element(prs_cxml), null);
        return [prs, expected_value];
      } },
    sld_width_set_fixture: { args: ["request"], params: [["p:presentation", "p:presentation/p:sldSz{cx=914400}"], ["p:presentation/p:sldSz{cx=424242}", "p:presentation/p:sldSz{cx=914400}"]], fn: (self: any, request: any): any => {
        let expected_cxml: any, expected_xml: any, prs: any, prs_cxml: any;
        [prs_cxml, expected_cxml] = $list(request.param);
        prs = $call(Presentation, element(prs_cxml), null);
        expected_xml = xml(expected_cxml);
        return [prs, 914400, expected_xml];
      } },
    slides_fixture: { args: ["request", "part_prop_", "Slides_", "slides_"], params: [["p:presentation", [], "p:presentation/p:sldIdLst"], ["p:presentation/p:sldIdLst/p:sldId{r:id=a}", ["a"], "p:presentation/p:sldIdLst/p:sldId{r:id=a}"], ["p:presentation/p:sldIdLst/(p:sldId{r:id=a},p:sldId{r:id=b})", ["a", "b"], "p:presentation/p:sldIdLst/(p:sldId{r:id=a},p:sldId{r:id=b})"]], fn: (self: any, request: any, part_prop_: any, Slides_: any, slides_: any): any => {
        let expected_cxml: any, expected_xml: any, prs: any, prs_cxml: any, rIds: any, rename_slide_parts_: any;
        [prs_cxml, rIds, expected_cxml] = $list(request.param);
        prs = $call(Presentation, element(prs_cxml), null);
        rename_slide_parts_ = part_prop_.returnValue.renameSlideParts;
        expected_xml = xml(expected_cxml);
        return [prs, rename_slide_parts_, rIds, Slides_, slides_, expected_xml];
      } },
    core_properties_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, CorePropertiesPart);
      } },
    masters_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, Presentation, "slide_masters");
      } },
    notes_master_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, NotesMasterPart);
      } },
    part_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, Presentation, "part");
      } },
    prs_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, PresentationPart);
      } },
    slide_layouts_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlideLayouts);
      } },
    SlideMasters_: { args: ["request", "slide_masters_"], fn: (self: any, request: any, slide_masters_: any): any => {
        return $call(classMock, request, "pptx.presentation.SlideMasters", $kw({return_value: slide_masters_}));
      } },
    slide_master_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlideMaster);
      } },
    slide_masters_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlideMasters);
      } },
    Slides_: { args: ["request", "slides_"], fn: (self: any, request: any, slides_: any): any => {
        return $call(classMock, request, "pptx.presentation.Slides", $kw({return_value: slides_}));
      } },
    slides_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Slides);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_the_height_of_its_slides", args: ["sld_height_get_fixture"], fn: (self: any, sld_height_get_fixture: any): any => {
        let expected_value: any, prs: any;
        [prs, expected_value] = $list(sld_height_get_fixture);
        $assertEqual(prs.slideHeight, expected_value);
      } },
    { name: "it_can_change_the_height_of_its_slides", args: ["sld_height_set_fixture"], fn: (self: any, sld_height_set_fixture: any): any => {
        let expected_xml: any, prs: any, slide_height: any;
        [prs, slide_height, expected_xml] = $list(sld_height_set_fixture);
        prs.slideHeight = slide_height;
        $assertEqual(prs._element.xml, expected_xml);
      } },
    { name: "it_knows_the_width_of_its_slides", args: ["sld_width_get_fixture"], fn: (self: any, sld_width_get_fixture: any): any => {
        let expected_value: any, prs: any;
        [prs, expected_value] = $list(sld_width_get_fixture);
        $assertEqual(prs.slideWidth, expected_value);
      } },
    { name: "it_can_change_the_width_of_its_slides", args: ["sld_width_set_fixture"], fn: (self: any, sld_width_set_fixture: any): any => {
        let expected_xml: any, prs: any, slide_width: any;
        [prs, slide_width, expected_xml] = $list(sld_width_set_fixture);
        prs.slideWidth = slide_width;
        $assertEqual(prs._element.xml, expected_xml);
      } },
    { name: "it_knows_its_part", args: ["part_fixture"], fn: (self: any, part_fixture: any): any => {
        let prs: any, prs_part_: any;
        [prs, prs_part_] = $list(part_fixture);
        $assert($is(prs.part, prs_part_));
      } },
    { name: "it_provides_access_to_its_core_properties", args: ["core_props_fixture"], fn: (self: any, core_props_fixture: any): any => {
        let core_properties_: any, prs: any;
        [prs, core_properties_] = $list(core_props_fixture);
        $assert($is(prs.coreProperties, core_properties_));
      } },
    { name: "it_provides_access_to_its_notes_master", args: ["notes_master_fixture"], fn: (self: any, notes_master_fixture: any): any => {
        let notes_master_: any, prs: any;
        [prs, notes_master_] = $list(notes_master_fixture);
        $assert($is(prs.notesMaster, notes_master_));
      } },
    { name: "it_provides_access_to_its_slides", args: ["slides_fixture"], fn: (self: any, slides_fixture: any): any => {
        let Slides_: any, expected_xml: any, prs: any, rIds: any, rename_slide_parts_: any, slides: any, slides_: any;
        [prs, rename_slide_parts_, rIds] = $list($slice(slides_fixture, null, 3, null));
        [Slides_, slides_, expected_xml] = $list($slice(slides_fixture, 3, null, null));
        slides = prs.slides;
        rename_slide_parts_.assertCalledOnceWith(rIds);
        Slides_.assertCalledOnceWith($get(prs._element.xpath("p:sldIdLst"), 0), prs);
        $assertEqual(prs._element.xml, expected_xml);
        $assert($is(slides, slides_));
      } },
    { name: "it_provides_access_to_its_slide_layouts", args: ["layouts_fixture"], fn: (self: any, layouts_fixture: any): any => {
        let prs: any, slide_layouts_: any;
        [prs, slide_layouts_] = $list(layouts_fixture);
        $assert($is(prs.slideLayouts, slide_layouts_));
      } },
    { name: "it_provides_access_to_its_slide_master", args: ["master_fixture"], fn: (self: any, master_fixture: any): any => {
        let getitem_: any, prs: any, slide_master: any, slide_master_: any;
        [prs, getitem_, slide_master_] = $list(master_fixture);
        slide_master = prs.slideMaster;
        getitem_.assertCalledOnceWith(0);
        $assert($is(slide_master, slide_master_));
      } },
    { name: "it_provides_access_to_its_slide_masters", args: ["masters_fixture"], fn: (self: any, masters_fixture: any): any => {
        let SlideMasters_: any, expected_xml: any, prs: any, slide_masters: any, slide_masters_: any;
        [prs, SlideMasters_, slide_masters_, expected_xml] = $list(masters_fixture);
        slide_masters = prs.slideMasters;
        SlideMasters_.assertCalledOnceWith($get(prs._element.xpath("p:sldMasterIdLst"), 0), prs);
        $assert($is(slide_masters, slide_masters_));
        $assertEqual(prs._element.xml, expected_xml);
      } },
    { name: "it_can_save_the_presentation_to_a_file", args: ["save_fixture"], fn: (self: any, save_fixture: any): any => {
        let file_: any, prs: any, prs_part_: any;
        [prs, file_, prs_part_] = $list(save_fixture);
        prs.save(file_);
        prs_part_.save.assertCalledOnceWith(file_);
      } },
  ],
};
pytestClass("DescribePresentation", DescribePresentation);

