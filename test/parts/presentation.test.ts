// Machine-translated from python-pptx tests/parts/test_presentation.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $call, $is, $iter, $kw, $len, $list, $mod, $raises, $range, $truthy, call, classMock, instanceMock, methodMock, propertyMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { RELATIONSHIP_TYPE as RT$imported } from "../../src/opc/constants.js";
import { PackURI as PackURI$imported } from "../../src/opc/packuri.js";
import { Package as Package$imported } from "../../src/package.js";
import { CorePropertiesPart as CorePropertiesPart$imported } from "../../src/parts/coreprops.js";
import { PresentationPart as PresentationPart$imported } from "../../src/parts/presentation.js";
import { NotesMasterPart as NotesMasterPart$imported, SlideMasterPart as SlideMasterPart$imported, SlidePart as SlidePart$imported } from "../../src/parts/slide.js";
import { Presentation as Presentation$imported } from "../../src/presentation.js";
import { NotesMaster as NotesMaster$imported, Slide as Slide$imported, SlideLayout as SlideLayout$imported, SlideMaster as SlideMaster$imported } from "../../src/slide.js";
import { element } from "../unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const RT = RT$imported;
const PackURI = PackURI$imported;
const Package = Package$imported;
const CorePropertiesPart = CorePropertiesPart$imported;
const PresentationPart = PresentationPart$imported;
const NotesMasterPart = NotesMasterPart$imported;
const SlideMasterPart = SlideMasterPart$imported;
const SlidePart = SlidePart$imported;
const Presentation = Presentation$imported;
const NotesMaster = NotesMaster$imported;
const Slide = Slide$imported;
const SlideLayout = SlideLayout$imported;
const SlideMaster = SlideMaster$imported;
void $py;


const DescribePresentationPart: any = {
  fixtures: {
    notes_master_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, NotesMasterPart);
      } },
    package_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Package);
      } },
    part_related_by_: { args: ["request"], fn: (self: any, request: any): any => {
        return methodMock(request, PresentationPart, "part_related_by");
      } },
    relate_to_: { args: ["request"], fn: (self: any, request: any): any => {
        return methodMock(request, PresentationPart, "relate_to");
      } },
    related_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return methodMock(request, PresentationPart, "related_part");
      } },
    slide_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Slide);
      } },
    slide_master_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlideMaster);
      } },
    slide_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlidePart);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_access_to_its_presentation", args: ["request"], fn: (self: any, request: any): any => {
        let Presentation_: any, prs: any, prs_: any, prs_elm: any, prs_part: any;
        prs_ = instanceMock(request, Presentation);
        Presentation_ = $call(classMock, request, "pptx.parts.presentation.Presentation", $kw({return_value: prs_}));
        prs_elm = element("p:presentation");
        prs_part = $call(PresentationPart, null, null, null, prs_elm);
        prs = prs_part.presentation;
        Presentation_.assertCalledOnceWith(prs_elm, prs_part);
        $assert($is(prs, prs_));
      } },
    { name: "it_provides_access_to_its_core_properties", args: ["request", "package_"], fn: (self: any, request: any, package_: any): any => {
        let core_properties_: any, prs_part: any;
        core_properties_ = instanceMock(request, CorePropertiesPart);
        package_.coreProperties = core_properties_;
        prs_part = $call(PresentationPart, null, null, package_, null);
        $assert($is(prs_part.coreProperties, core_properties_));
      } },
    { name: "it_provides_access_to_an_existing_notes_master_part", args: ["notes_master_part_", "part_related_by_"], fn: (self: any, notes_master_part_: any, part_related_by_: any): any => {
        let notes_master_part: any, prs_part: any;
        prs_part = $call(PresentationPart, null, null, null, null);
        part_related_by_.returnValue = notes_master_part_;
        notes_master_part = prs_part.notesMasterPart;
        prs_part.partRelatedBy.assertCalledOnceWith(prs_part, RT.NOTES_MASTER);
        $assert($is(notes_master_part, notes_master_part_));
      } },
    { name: "but_it_adds_a_notes_master_part_when_needed", args: ["request", "package_", "notes_master_part_", "part_related_by_", "relate_to_"], fn: (self: any, request: any, package_: any, notes_master_part_: any, part_related_by_: any, relate_to_: any): any => {
        let NotesMasterPart_: any, notes_master_part: any, prs_part: any;
        NotesMasterPart_ = classMock(request, "pptx.parts.presentation.NotesMasterPart");
        NotesMasterPart_.createDefault.returnValue = notes_master_part_;
        part_related_by_.sideEffect = KeyError;
        prs_part = $call(PresentationPart, null, null, package_, null);
        notes_master_part = prs_part.notesMasterPart;
        NotesMasterPart_.createDefault.assertCalledOnceWith(package_);
        relate_to_.assertCalledOnceWith(prs_part, notes_master_part_, RT.NOTES_MASTER);
        $assert($is(notes_master_part, notes_master_part_));
      } },
    { name: "it_provides_access_to_its_notes_master", args: ["request", "notes_master_part_"], fn: (self: any, request: any, notes_master_part_: any): any => {
        let notes_master_: any, prs_part: any;
        notes_master_ = instanceMock(request, NotesMaster);
        $call(propertyMock, request, PresentationPart, "notes_master_part", $kw({return_value: notes_master_part_}));
        notes_master_part_.notesMaster = notes_master_;
        prs_part = $call(PresentationPart, null, null, null, null);
        $assert($is(prs_part.notesMaster, notes_master_));
      } },
    { name: "it_provides_access_to_a_related_slide", args: ["request", "slide_", "related_part_"], fn: (self: any, request: any, slide_: any, related_part_: any): any => {
        let prs_part: any, slide: any, slide_part_: any;
        slide_part_ = $call(instanceMock, request, SlidePart, $kw({slide: slide_}));
        related_part_.returnValue = slide_part_;
        prs_part = $call(PresentationPart, null, null, null, null);
        slide = prs_part.relatedSlide("rId42");
        related_part_.assertCalledOnceWith(prs_part, "rId42");
        $assert($is(slide, slide_));
      } },
    { name: "it_provides_access_to_a_related_master", args: ["request", "slide_master_", "related_part_"], fn: (self: any, request: any, slide_master_: any, related_part_: any): any => {
        let prs_part: any, slide_master: any, slide_master_part_: any;
        slide_master_part_ = $call(instanceMock, request, SlideMasterPart, $kw({slide_master: slide_master_}));
        related_part_.returnValue = slide_master_part_;
        prs_part = $call(PresentationPart, null, null, null, null);
        slide_master = prs_part.relatedSlideMaster("rId42");
        related_part_.assertCalledOnceWith(prs_part, "rId42");
        $assert($is(slide_master, slide_master_));
      } },
    { name: "it_can_rename_related_slide_parts", args: ["request", "related_part_"], fn: (self: any, request: any, related_part_: any): any => {
        let prs_part: any, rIds: any, slide_parts: any;
        rIds = $list($list($range(5, 0, -1)).map((n: any) => $mod("rId%d", n)));
        slide_parts = $list($list($range(5)).map((_: any) => instanceMock(request, SlidePart)));
        related_part_.sideEffect = $iter(slide_parts);
        prs_part = $call(PresentationPart, null, null, null, null);
        prs_part.renameSlideParts(rIds);
        $assertEqual(related_part_.callArgsList, $list(rIds).map((rId: any) => call(prs_part, rId)));
        $assertEqual($list(slide_parts).map((s: any) => s.partname), $list($range($len(rIds))).map((i: any) => $call(PackURI, $mod("/ppt/slides/slide%d.xml", (i + 1)))));
      } },
    { name: "it_can_save_the_package_to_a_file", args: ["package_"], fn: (self: any, package_: any): any => {
        $call(PresentationPart, null, null, package_, null).save("prs.pptx");
        package_.save.assertCalledOnceWith("prs.pptx");
      } },
    { name: "it_can_add_a_new_slide", args: ["request", "package_", "slide_part_", "slide_", "relate_to_"], fn: (self: any, request: any, package_: any, slide_part_: any, slide_: any, relate_to_: any): any => {
        let SlidePart_: any, partname: any, prs_part: any, rId: any, slide: any, slide_layout_: any, slide_layout_part_: any;
        slide_layout_ = instanceMock(request, SlideLayout);
        partname = $call(PackURI, "/ppt/slides/slide9.xml");
        $call(propertyMock, request, PresentationPart, "_next_slide_partname", $kw({return_value: partname}));
        SlidePart_ = classMock(request, "pptx.parts.presentation.SlidePart");
        SlidePart_.new.returnValue = slide_part_;
        relate_to_.returnValue = "rId42";
        slide_layout_part_ = slide_layout_.part;
        slide_part_.slide = slide_;
        prs_part = $call(PresentationPart, null, null, package_, null);
        [rId, slide] = $list(prs_part.addSlide(slide_layout_));
        SlidePart_.new.assertCalledOnceWith(partname, package_, slide_layout_part_);
        prs_part.relateTo.assertCalledOnceWith(prs_part, slide_part_, RT.SLIDE);
        $assertEqual(rId, "rId42");
        $assert($is(slide, slide_));
      } },
    { name: "it_finds_the_slide_id_of_a_slide_part", args: ["slide_part_", "related_part_"], fn: (self: any, slide_part_: any, related_part_: any): any => {
        let _slide_id: any, prs_elm: any, prs_part: any;
        prs_elm = element("p:presentation/p:sldIdLst/(p:sldId{r:id=a,id=256},p:sldId{r:id=b,id=257},p:sldId{r:id=c,id=258})");
        related_part_.sideEffect = $iter([null, slide_part_, null]);
        prs_part = $call(PresentationPart, null, null, null, prs_elm);
        _slide_id = prs_part.slideId(slide_part_);
        $assertEqual(related_part_.callArgsList, [call(prs_part, "a"), call(prs_part, "b")]);
        $assertEqual(_slide_id, 257);
      } },
    { name: "it_raises_on_slide_id_not_found", args: ["slide_part_", "related_part_"], fn: (self: any, slide_part_: any, related_part_: any): any => {
        let prs_elm: any, prs_part: any;
        prs_elm = element("p:presentation/p:sldIdLst/(p:sldId{r:id=a,id=256},p:sldId{r:id=b,id=257},p:sldId{r:id=c,id=258})");
        related_part_.returnValue = "not the slide you're looking for";
        prs_part = $call(PresentationPart, null, null, null, prs_elm);
        $raises(ValueError, () => {
          prs_part.slideId(slide_part_);
        });
      } },
    { name: "it_finds_a_slide_by_slide_id", args: ["is_present", "slide_", "slide_part_", "related_part_"], parametrize: [{ names: ["is_present"], values: [true, false] }], fn: (self: any, is_present: any, slide_: any, slide_part_: any, related_part_: any): any => {
        let expected_value: any, prs_elm: any, prs_part: any, slide: any, slide_id: any;
        prs_elm = element("p:presentation/p:sldIdLst/(p:sldId{r:id=a,id=256},p:sldId{r:id=b,id=257},p:sldId{r:id=c,id=258})");
        slide_id = ($truthy(is_present) ? 257 : 666);
        expected_value = ($truthy(is_present) ? slide_ : null);
        related_part_.returnValue = slide_part_;
        slide_part_.slide = slide_;
        prs_part = $call(PresentationPart, null, null, null, prs_elm);
        slide = prs_part.getSlide(slide_id);
        $assertEqual(slide, expected_value);
      } },
    { name: "it_knows_the_next_slide_partname_to_help", args: [], fn: (self: any): any => {
        let prs_elm: any, prs_part: any;
        prs_elm = element("p:presentation/p:sldIdLst/(p:sldId,p:sldId)");
        prs_part = $call(PresentationPart, null, null, null, prs_elm);
        $assertEqual(prs_part._nextSlidePartname, $call(PackURI, "/ppt/slides/slide3.xml"));
      } },
  ],
};
pytestClass("DescribePresentationPart", DescribePresentationPart);

