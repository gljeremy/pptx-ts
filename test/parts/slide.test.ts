// Machine-translated from python-pptx tests/parts/test_slide.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $append, $assert, $assertEqual, $b, $call, $is, $isinstance, $kw, $list, $open, $read, $t1, $truthy, call, classMock, initializerMock, instanceMock, methodMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
// pptx.chart.data is ported in stage 5; the test only needs a class to mock
class ChartData { xmlBytes(_chartType: any): Uint8Array { return new Uint8Array(); } get xlsxBlob(): Uint8Array { return new Uint8Array(); } }
import { XL_CHART_TYPE as XCT$imported } from "../../src/enum/chart.js";
import { PROG_ID as PROG_ID$imported } from "../../src/enum/shapes.js";
import { Video as Video$imported } from "../../src/media.js";
import { CONTENT_TYPE as CT$imported, RELATIONSHIP_TYPE as RT$imported } from "../../src/opc/constants.js";
import { Part as Part$imported } from "../../src/opc/package.js";
import { PackURI as PackURI$imported } from "../../src/opc/packuri.js";
import { CT_NotesMaster as CT_NotesMaster$imported, CT_NotesSlide as CT_NotesSlide$imported, CT_Slide as CT_Slide$imported } from "../../src/oxml/slide.js";
import { CT_OfficeStyleSheet as CT_OfficeStyleSheet$imported } from "../../src/oxml/theme.js";
import { Package as Package$imported } from "../../src/package.js";
import { ChartPart as ChartPart$imported } from "../../src/parts/chart.js";
import { EmbeddedPackagePart as EmbeddedPackagePart$imported } from "../../src/parts/embeddedpackage.js";
import { Image as Image$imported, ImagePart as ImagePart$imported } from "../../src/parts/image.js";
import { MediaPart as MediaPart$imported } from "../../src/parts/media.js";
import { PresentationPart as PresentationPart$imported } from "../../src/parts/presentation.js";
import { BaseSlidePart as BaseSlidePart$imported, NotesMasterPart as NotesMasterPart$imported, NotesSlidePart as NotesSlidePart$imported, SlideLayoutPart as SlideLayoutPart$imported, SlideMasterPart as SlideMasterPart$imported, SlidePart } from "../../src/parts/slide.js";
import { NotesMaster as NotesMaster$imported, NotesSlide as NotesSlide$imported, Slide as Slide$imported, SlideLayout as SlideLayout$imported, SlideMaster as SlideMaster$imported } from "../../src/slide.js";
import { element } from "../unitutil/cxml.js";
import { absjoin, testFileDir } from "../unitutil/file.js";
// python binds imported names at import time; patches made later don't change them
const XCT = XCT$imported;
const PROG_ID = PROG_ID$imported;
const Video = Video$imported;
const CT = CT$imported;
const RT = RT$imported;
const Part = Part$imported;
const PackURI = PackURI$imported;
const CT_NotesMaster = CT_NotesMaster$imported;
const CT_NotesSlide = CT_NotesSlide$imported;
const CT_Slide = CT_Slide$imported;
const CT_OfficeStyleSheet = CT_OfficeStyleSheet$imported;
const Package = Package$imported;
const ChartPart = ChartPart$imported;
const EmbeddedPackagePart = EmbeddedPackagePart$imported;
const Image = Image$imported;
const ImagePart = ImagePart$imported;
const MediaPart = MediaPart$imported;
const PresentationPart = PresentationPart$imported;
const BaseSlidePart = BaseSlidePart$imported;
const NotesMasterPart = NotesMasterPart$imported;
const NotesSlidePart = NotesSlidePart$imported;
const SlideLayoutPart = SlideLayoutPart$imported;
const SlideMasterPart = SlideMasterPart$imported;
const NotesMaster = NotesMaster$imported;
const NotesSlide = NotesSlide$imported;
const Slide = Slide$imported;
const SlideLayout = SlideLayout$imported;
const SlideMaster = SlideMaster$imported;
void $py;


const DescribeBaseSlidePart: any = {
  fixtures: {
    image_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, ImagePart);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_name", args: [], fn: (self: any): any => {
        let slide_part: any;
        slide_part = $call(BaseSlidePart, null, null, null, element("p:sld/p:cSld{name=Foobar}"));
        $assertEqual(slide_part.name, "Foobar");
      } },
    { name: "it_can_get_a_related_image_by_rId", args: ["request", "image_part_"], fn: (self: any, request: any, image_part_: any): any => {
        let image: any, image_: any, related_part_: any, slide_part: any;
        image_ = instanceMock(request, Image);
        image_part_.image = image_;
        related_part_ = $call(methodMock, request, BaseSlidePart, "related_part", $kw({return_value: image_part_}));
        slide_part = $call(BaseSlidePart, null, null, null, null);
        image = slide_part.getImage("rId42");
        related_part_.assertCalledOnceWith(slide_part, "rId42");
        $assert($is(image, image_));
      } },
    { name: "it_can_add_an_image_part", args: ["request", "image_part_"], fn: (self: any, request: any, image_part_: any): any => {
        let image_part: any, package_: any, rId: any, relate_to_: any, slide_part: any;
        package_ = instanceMock(request, Package);
        package_.getOrAddImagePart.returnValue = image_part_;
        relate_to_ = $call(methodMock, request, BaseSlidePart, "relate_to", $kw({return_value: "rId6"}));
        slide_part = $call(BaseSlidePart, null, null, package_, null);
        [image_part, rId] = $list(slide_part.getOrAddImagePart("foobar.png"));
        package_.getOrAddImagePart.assertCalledOnceWith("foobar.png");
        relate_to_.assertCalledOnceWith(slide_part, image_part_, RT.IMAGE);
        $assert($is(image_part, image_part_));
        $assertEqual(rId, "rId6");
      } },
  ],
};
pytestClass("DescribeBaseSlidePart", DescribeBaseSlidePart);

const DescribeNotesMasterPart: any = {
  fixtures: {
    notes_master_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, NotesMasterPart);
      } },
    package_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Package);
      } },
    theme_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Part);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_create_a_notes_master_part", args: ["request", "package_", "notes_master_part_", "theme_part_"], fn: (self: any, request: any, package_: any, notes_master_part_: any, theme_part_: any): any => {
        let notes_master_part: any;
        $call(methodMock, request, NotesMasterPart, "_new", $kw({autospec: false, return_value: notes_master_part_}));
        $call(methodMock, request, NotesMasterPart, "_new_theme_part", $kw({autospec: false, return_value: theme_part_}));
        notes_master_part = NotesMasterPart.createDefault(package_);
        NotesMasterPart._new.assertCalledOnceWith(package_);
        NotesMasterPart._newThemePart.assertCalledOnceWith(package_);
        notes_master_part.relateTo.assertCalledOnceWith(theme_part_, RT.THEME);
        $assert($is(notes_master_part, notes_master_part_));
      } },
    { name: "it_provides_access_to_its_notes_master", args: ["request"], fn: (self: any, request: any): any => {
        let NotesMaster_: any, notesMaster: any, notes_master: any, notes_master_: any, notes_master_part: any;
        notes_master_ = instanceMock(request, NotesMaster);
        NotesMaster_ = $call(classMock, request, "pptx.parts.slide.NotesMaster", $kw({return_value: notes_master_}));
        notesMaster = element("p:notesMaster");
        notes_master_part = $call(NotesMasterPart, null, null, null, notesMaster);
        notes_master = notes_master_part.notesMaster;
        NotesMaster_.assertCalledOnceWith(notesMaster, notes_master_part);
        $assert($is(notes_master, notes_master_));
      } },
    { name: "it_creates_a_new_notes_master_part_to_help", args: ["request", "package_", "notes_master_part_"], fn: (self: any, request: any, package_: any, notes_master_part_: any): any => {
        let NotesMasterPart_: any, notesMaster: any, notes_master_part: any;
        NotesMasterPart_ = $call(classMock, request, "pptx.parts.slide.NotesMasterPart", $kw({return_value: notes_master_part_}));
        notesMaster = element("p:notesMaster");
        $call(methodMock, request, CT_NotesMaster, "new_default", $kw({autospec: false, return_value: notesMaster}));
        notes_master_part = NotesMasterPart._new(package_);
        CT_NotesMaster.newDefault.assertCalledOnceWith();
        NotesMasterPart_.assertCalledOnceWith($call(PackURI, "/ppt/notesMasters/notesMaster1.xml"), CT.PML_NOTES_MASTER, package_, notesMaster);
        $assert($is(notes_master_part, notes_master_part_));
      } },
    { name: "it_creates_a_new_theme_part_to_help", args: ["request", "package_", "theme_part_"], fn: (self: any, request: any, package_: any, theme_part_: any): any => {
        let XmlPart_: any, partname: any, pn_tmpl: any, theme_elm: any, theme_part: any;
        XmlPart_ = $call(classMock, request, "pptx.parts.slide.XmlPart", $kw({return_value: theme_part_}));
        theme_elm = element("p:theme");
        $call(methodMock, request, CT_OfficeStyleSheet, "new_default", $kw({autospec: false, return_value: theme_elm}));
        pn_tmpl = "/ppt/theme/theme%d.xml";
        partname = $call(PackURI, "/ppt/theme/theme2.xml");
        package_.nextPartname.returnValue = partname;
        theme_part = NotesMasterPart._newThemePart(package_);
        package_.nextPartname.assertCalledOnceWith(pn_tmpl);
        CT_OfficeStyleSheet.newDefault.assertCalledOnceWith();
        XmlPart_.assertCalledOnceWith(partname, CT.OFC_THEME, package_, theme_elm);
        $assert($is(theme_part, theme_part_));
      } },
  ],
};
pytestClass("DescribeNotesMasterPart", DescribeNotesMasterPart);

const DescribeNotesSlidePart: any = {
  fixtures: {
    notes_master_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, NotesMaster);
      } },
    notes_master_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, NotesMasterPart);
      } },
    notes_slide_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, NotesSlide);
      } },
    notes_slide_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, NotesSlidePart);
      } },
    package_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Package);
      } },
    slide_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlidePart);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_create_a_notes_slide_part", args: ["request", "package_", "slide_part_", "notes_master_part_", "notes_slide_", "notes_master_", "notes_slide_part_"], fn: (self: any, request: any, package_: any, slide_part_: any, notes_master_part_: any, notes_slide_: any, notes_master_: any, notes_slide_part_: any): any => {
        let _add_notes_slide_part_: any, notes_slide_part: any, presentation_part_: any;
        presentation_part_ = instanceMock(request, PresentationPart);
        package_.presentationPart = presentation_part_;
        presentation_part_.notesMasterPart = notes_master_part_;
        _add_notes_slide_part_ = $call(methodMock, request, NotesSlidePart, "_add_notes_slide_part", $kw({autospec: false, return_value: notes_slide_part_}));
        notes_slide_part_.notesSlide = notes_slide_;
        notes_master_part_.notesMaster = notes_master_;
        notes_slide_part = NotesSlidePart.new(package_, slide_part_);
        _add_notes_slide_part_.assertCalledOnceWith(package_, slide_part_, notes_master_part_);
        notes_slide_.cloneMasterPlaceholders.assertCalledOnceWith(notes_master_);
        $assert($is(notes_slide_part, notes_slide_part_));
      } },
    { name: "it_provides_access_to_the_notes_master", args: ["request", "notes_master_", "notes_master_part_"], fn: (self: any, request: any, notes_master_: any, notes_master_part_: any): any => {
        let notes_master: any, notes_slide_part: any, part_related_by_: any;
        part_related_by_ = $call(methodMock, request, NotesSlidePart, "part_related_by", $kw({return_value: notes_master_part_}));
        notes_slide_part = $call(NotesSlidePart, null, null, null, null);
        notes_master_part_.notesMaster = notes_master_;
        notes_master = notes_slide_part.notesMaster;
        part_related_by_.assertCalledOnceWith(notes_slide_part, RT.NOTES_MASTER);
        $assert($is(notes_master, notes_master_));
      } },
    { name: "it_provides_access_to_its_notes_slide", args: ["request", "notes_slide_"], fn: (self: any, request: any, notes_slide_: any): any => {
        let NotesSlide_: any, notes: any, notes_slide: any, notes_slide_part: any;
        NotesSlide_ = $call(classMock, request, "pptx.parts.slide.NotesSlide", $kw({return_value: notes_slide_}));
        notes = element("p:notes");
        notes_slide_part = $call(NotesSlidePart, null, null, null, notes);
        notes_slide = notes_slide_part.notesSlide;
        NotesSlide_.assertCalledOnceWith(notes, notes_slide_part);
        $assert($is(notes_slide, notes_slide_));
      } },
    { name: "it_adds_a_notes_slide_part_to_help", args: ["request", "package_", "slide_part_", "notes_master_part_", "notes_slide_part_"], fn: (self: any, request: any, package_: any, slide_part_: any, notes_master_part_: any, notes_slide_part_: any): any => {
        let NotesSlidePart_: any, new_: any, notes: any, notes_slide_part: any;
        NotesSlidePart_ = $call(classMock, request, "pptx.parts.slide.NotesSlidePart", $kw({return_value: notes_slide_part_}));
        notes = element("p:notes");
        new_ = $call(methodMock, request, CT_NotesSlide, "new", $kw({autospec: false, return_value: notes}));
        package_.nextPartname.returnValue = $call(PackURI, "/ppt/notesSlides/notesSlide42.xml");
        notes_slide_part = NotesSlidePart._addNotesSlidePart(package_, slide_part_, notes_master_part_);
        package_.nextPartname.assertCalledOnceWith("/ppt/notesSlides/notesSlide%d.xml");
        new_.assertCalledOnceWith();
        NotesSlidePart_.assertCalledOnceWith($call(PackURI, "/ppt/notesSlides/notesSlide42.xml"), CT.PML_NOTES_SLIDE, package_, notes);
        $assertEqual(notes_slide_part_.relateTo.callArgsList, [call(notes_master_part_, RT.NOTES_MASTER), call(slide_part_, RT.SLIDE)]);
        $assert($is(notes_slide_part, notes_slide_part_));
      } },
  ],
};
pytestClass("DescribeNotesSlidePart", DescribeNotesSlidePart);

const DescribeSlidePart: any = {
  fixtures: {
    has_notes_slide_fixture: { args: ["request", "part_related_by_"], params: [true, false], fn: (self: any, request: any, part_related_by_: any): any => {
        let expected_value: any, has_notes_slide: any, slide_part: any;
        has_notes_slide = request.param;
        slide_part = $call(SlidePart, null, null, null, null);
        part_related_by_.sideEffect = ($truthy(has_notes_slide) ? null : KeyError);
        expected_value = has_notes_slide;
        return [slide_part, expected_value];
      } },
    layout_fixture: { args: ["slide_layout_", "part_related_by_"], fn: (self: any, slide_layout_: any, part_related_by_: any): any => {
        let slide_part: any;
        slide_part = $call(SlidePart, null, null, null, null);
        part_related_by_.returnValue.slideLayout = slide_layout_;
        return [slide_part, slide_layout_];
      } },
    notes_slide_fixture: { args: ["request", "notes_slide_", "part_related_by_", "_add_notes_slide_part_", "notes_slide_part_"], params: [true, false], fn: (self: any, request: any, notes_slide_: any, part_related_by_: any, _add_notes_slide_part_: any, notes_slide_part_: any): any => {
        let add_calls: any, has_notes_slide: any, slide_part: any;
        has_notes_slide = request.param;
        slide_part = $call(SlidePart, null, null, null, null);
        part_related_by_.returnValue = notes_slide_part_;
        add_calls = [];
        if ($truthy(!$truthy(has_notes_slide))) {
          part_related_by_.sideEffect = KeyError;
          $append(add_calls, call(slide_part));
        }
        notes_slide_part_.notesSlide = notes_slide_;
        return [slide_part, add_calls, notes_slide_];
      } },
    slide_fixture: { args: ["Slide_", "slide_"], fn: (self: any, Slide_: any, slide_: any): any => {
        let sld: any, slide_part: any;
        sld = element("p:sld");
        slide_part = $call(SlidePart, null, null, null, sld);
        return [slide_part, Slide_, sld, slide_];
      } },
    slide_id_fixture: { args: ["package_", "presentation_part_"], fn: (self: any, package_: any, presentation_part_: any): any => {
        let slide_id: any, slide_part: any;
        slide_part = $call(SlidePart, null, null, package_, null);
        slide_id = 256;
        package_.presentationPart = presentation_part_;
        presentation_part_.slideId.returnValue = slide_id;
        return [slide_part, presentation_part_, slide_id];
      } },
    _add_notes_slide_part_: { args: ["request", "notes_slide_part_"], fn: (self: any, request: any, notes_slide_part_: any): any => {
        return $call(methodMock, request, SlidePart, "_add_notes_slide_part", $kw({return_value: notes_slide_part_, autospec: true}));
      } },
    media_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, MediaPart);
      } },
    notes_slide_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, NotesSlide);
      } },
    NotesSlidePart_: { args: ["request", "notes_slide_part_"], fn: (self: any, request: any, notes_slide_part_: any): any => {
        return $call(classMock, request, "pptx.parts.slide.NotesSlidePart", $kw({return_value: notes_slide_part_}));
      } },
    notes_slide_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, NotesSlidePart);
      } },
    package_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Package);
      } },
    part_related_by_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, SlidePart, "part_related_by", $kw({autospec: true}));
      } },
    presentation_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, PresentationPart);
      } },
    relate_to_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, SlidePart, "relate_to", $kw({autospec: true}));
      } },
    Slide_: { args: ["request", "slide_"], fn: (self: any, request: any, slide_: any): any => {
        return $call(classMock, request, "pptx.parts.slide.Slide", $kw({return_value: slide_}));
      } },
    slide_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Slide);
      } },
    slide_layout_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlideLayout);
      } },
    video_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Video);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_slide_id", args: ["slide_id_fixture"], fn: (self: any, slide_id_fixture: any): any => {
        let _slide_id: any, presentation_part_: any, slide_id: any, slide_part: any;
        [slide_part, presentation_part_, slide_id] = $list(slide_id_fixture);
        _slide_id = slide_part.slideId;
        presentation_part_.slideId.assertCalledOnceWith(slide_part);
        $assert($is(_slide_id, slide_id));
      } },
    { name: "it_knows_whether_it_has_a_notes_slide", args: ["has_notes_slide_fixture"], fn: (self: any, has_notes_slide_fixture: any): any => {
        let expected_value: any, slide_part: any, value: any;
        [slide_part, expected_value] = $list(has_notes_slide_fixture);
        value = slide_part.hasNotesSlide;
        slide_part.partRelatedBy.assertCalledOnceWith(slide_part, RT.NOTES_SLIDE);
        $assert($is(value, expected_value));
      } },
    { name: "it_can_add_a_chart_part", args: ["request", "package_", "relate_to_"], fn: (self: any, request: any, package_: any, relate_to_: any): any => {
        let ChartPart_: any, chart_data_: any, chart_part_: any, rId: any, slide_part: any;
        chart_data_ = instanceMock(request, ChartData);
        chart_part_ = instanceMock(request, ChartPart);
        ChartPart_ = classMock(request, "pptx.parts.slide.ChartPart");
        ChartPart_.new.returnValue = chart_part_;
        relate_to_.returnValue = "rId42";
        slide_part = $call(SlidePart, null, null, package_, null);
        rId = slide_part.addChartPart(XCT.RADAR, chart_data_);
        ChartPart_.new.assertCalledOnceWith(XCT.RADAR, chart_data_, package_);
        relate_to_.assertCalledOnceWith(slide_part, chart_part_, RT.CHART);
        $assertEqual(rId, "rId42");
      } },
    { name: "it_can_add_an_embedded_ole_object_part", args: ["request", "package_", "relate_to_", "prog_id", "rel_type"], parametrize: [{ names: ["prog_id", "rel_type"], values: [[PROG_ID.DOCX, RT.PACKAGE], [PROG_ID.PPTX, RT.PACKAGE], [PROG_ID.XLSX, RT.PACKAGE], ["Foo.Bar.18", RT.OLE_OBJECT]] }], fn: (self: any, request: any, package_: any, relate_to_: any, prog_id: any, rel_type: any): any => {
        let EmbeddedPackagePart_: any, _blob_from_file_: any, _rId: any, embedded_package_part_: any, slide_part: any;
        _blob_from_file_ = $call(methodMock, request, SlidePart, "_blob_from_file", $kw({return_value: $b("012345")}));
        embedded_package_part_ = instanceMock(request, EmbeddedPackagePart);
        EmbeddedPackagePart_ = classMock(request, "pptx.parts.slide.EmbeddedPackagePart");
        EmbeddedPackagePart_.factory.returnValue = embedded_package_part_;
        relate_to_.returnValue = "rId9";
        slide_part = $call(SlidePart, null, null, package_, null);
        _rId = slide_part.addEmbeddedOleObjectPart(prog_id, "workbook.xlsx");
        _blob_from_file_.assertCalledOnceWith(slide_part, "workbook.xlsx");
        EmbeddedPackagePart_.factory.assertCalledOnceWith(prog_id, $b("012345"), package_);
        relate_to_.assertCalledOnceWith(slide_part, embedded_package_part_, rel_type);
        $assertEqual(_rId, "rId9");
      } },
    { name: "it_can_get_or_add_a_video_part", args: ["package_", "video_", "relate_to_", "media_part_"], fn: (self: any, package_: any, video_: any, relate_to_: any, media_part_: any): any => {
        let media_rId: any, result: any, slide_part: any, video_rId: any;
        [media_rId, video_rId] = $list(["rId1", "rId2"]);
        package_.getOrAddMediaPart.returnValue = media_part_;
        relate_to_.sideEffect = [media_rId, video_rId];
        slide_part = $call(SlidePart, null, null, package_, null);
        result = slide_part.getOrAddVideoMediaPart(video_);
        package_.getOrAddMediaPart.assertCalledOnceWith(video_);
        $assertEqual(relate_to_.callArgsList, [call(slide_part, media_part_, RT.MEDIA), call(slide_part, media_part_, RT.VIDEO)]);
        $assertEqual(result, [media_rId, video_rId]);
      } },
    { name: "it_can_create_a_new_slide_part", args: ["request", "package_", "relate_to_"], fn: (self: any, request: any, package_: any, relate_to_: any): any => {
        let CT_Slide_: any, _init_: any, partname: any, sld: any, slide_layout_part_: any, slide_part: any;
        partname = $call(PackURI, "/foobar.xml");
        _init_ = initializerMock(request, SlidePart);
        slide_layout_part_ = instanceMock(request, SlideLayoutPart);
        CT_Slide_ = classMock(request, "pptx.parts.slide.CT_Slide");
        const $t1 = element("c:sld");
        CT_Slide_.new.returnValue = $t1;
        sld = $t1;
        slide_part = SlidePart.new(partname, package_, slide_layout_part_);
        _init_.assertCalledOnceWith(slide_part, partname, CT.PML_SLIDE, package_, sld);
        slide_part.relateTo.assertCalledOnceWith(slide_part, slide_layout_part_, RT.SLIDE_LAYOUT);
        $assert($isinstance(slide_part, SlidePart));
      } },
    { name: "it_provides_access_to_its_slide", args: ["slide_fixture"], fn: (self: any, slide_fixture: any): any => {
        let Slide_: any, sld: any, slide: any, slide_: any, slide_part: any;
        [slide_part, Slide_, sld, slide_] = $list(slide_fixture);
        slide = slide_part.slide;
        Slide_.assertCalledOnceWith(sld, slide_part);
        $assert($is(slide, slide_));
      } },
    { name: "it_provides_access_to_the_slide_layout", args: ["layout_fixture"], fn: (self: any, layout_fixture: any): any => {
        let slide_layout: any, slide_layout_: any, slide_part: any;
        [slide_part, slide_layout_] = $list(layout_fixture);
        slide_layout = slide_part.slideLayout;
        slide_part.partRelatedBy.assertCalledOnceWith(slide_part, RT.SLIDE_LAYOUT);
        $assert($is(slide_layout, slide_layout_));
      } },
    { name: "it_knows_the_minimal_element_xml_for_a_slide", args: [], fn: (self: any): any => {
        let expected_xml: any, f: any, path: any, sld: any;
        path = absjoin(testFileDir, "minimal_slide.xml");
        sld = CT_Slide.new();
        f = $open(path, "r");
        expected_xml = $read(f);
        $assertEqual(sld.xml, expected_xml);
      } },
    { name: "it_gets_its_notes_slide_to_help", args: ["notes_slide_fixture"], fn: (self: any, notes_slide_fixture: any): any => {
        let calls: any, notes_slide: any, notes_slide_: any, slide_part: any;
        [slide_part, calls, notes_slide_] = $list(notes_slide_fixture);
        notes_slide = slide_part.notesSlide;
        slide_part.partRelatedBy.assertCalledOnceWith(slide_part, RT.NOTES_SLIDE);
        $assertEqual(slide_part._addNotesSlidePart.callArgsList, calls);
        $assert($is(notes_slide, notes_slide_));
      } },
    { name: "it_adds_a_notes_slide_part_to_help", args: ["package_", "NotesSlidePart_", "notes_slide_part_", "relate_to_"], fn: (self: any, package_: any, NotesSlidePart_: any, notes_slide_part_: any, relate_to_: any): any => {
        let notes_slide_part: any, slide_part: any;
        NotesSlidePart_.new.returnValue = notes_slide_part_;
        slide_part = $call(SlidePart, null, null, package_, null);
        notes_slide_part = slide_part._addNotesSlidePart();
        $assert($is(notes_slide_part, notes_slide_part_));
        NotesSlidePart_.new.assertCalledOnceWith(package_, slide_part);
        relate_to_.assertCalledOnceWith(slide_part, notes_slide_part, RT.NOTES_SLIDE);
      } },
  ],
};
pytestClass("DescribeSlidePart", DescribeSlidePart);

const DescribeSlideLayoutPart: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_access_to_its_slide_master", args: ["request"], fn: (self: any, request: any): any => {
        let part_related_by_: any, slide_layout_part: any, slide_master: any, slide_master_: any, slide_master_part_: any;
        slide_master_ = instanceMock(request, SlideMaster);
        slide_master_part_ = $call(instanceMock, request, SlideMasterPart, $kw({slide_master: slide_master_}));
        part_related_by_ = $call(methodMock, request, SlideLayoutPart, "part_related_by", $kw({return_value: slide_master_part_}));
        slide_layout_part = $call(SlideLayoutPart, null, null, null, null);
        slide_master = slide_layout_part.slideMaster;
        part_related_by_.assertCalledOnceWith(slide_layout_part, RT.SLIDE_MASTER);
        $assert($is(slide_master, slide_master_));
      } },
    { name: "it_provides_access_to_its_slide_layout", args: ["request"], fn: (self: any, request: any): any => {
        let SlideLayout_: any, sldLayout: any, slide_layout: any, slide_layout_: any, slide_layout_part: any;
        slide_layout_ = instanceMock(request, SlideLayout);
        SlideLayout_ = $call(classMock, request, "pptx.parts.slide.SlideLayout", $kw({return_value: slide_layout_}));
        sldLayout = element("p:sldLayout");
        slide_layout_part = $call(SlideLayoutPart, null, null, null, sldLayout);
        slide_layout = slide_layout_part.slideLayout;
        SlideLayout_.assertCalledOnceWith(sldLayout, slide_layout_part);
        $assert($is(slide_layout, slide_layout_));
      } },
  ],
};
pytestClass("DescribeSlideLayoutPart", DescribeSlideLayoutPart);

const DescribeSlideMasterPart: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_access_to_its_slide_master", args: ["request"], fn: (self: any, request: any): any => {
        let SlideMaster_: any, sldMaster: any, slide_master: any, slide_master_: any, slide_master_part: any;
        slide_master_ = instanceMock(request, SlideMaster);
        SlideMaster_ = $call(classMock, request, "pptx.parts.slide.SlideMaster", $kw({return_value: slide_master_}));
        sldMaster = element("p:sldMaster");
        slide_master_part = $call(SlideMasterPart, null, null, null, sldMaster);
        slide_master = slide_master_part.slideMaster;
        SlideMaster_.assertCalledOnceWith(sldMaster, slide_master_part);
        $assert($is(slide_master, slide_master_));
      } },
    { name: "it_provides_access_to_a_related_slide_layout", args: ["request"], fn: (self: any, request: any): any => {
        let related_part_: any, slide_layout: any, slide_layout_: any, slide_layout_part_: any, slide_master_part: any;
        slide_layout_ = instanceMock(request, SlideLayout);
        slide_layout_part_ = $call(instanceMock, request, SlideLayoutPart, $kw({slide_layout: slide_layout_}));
        related_part_ = $call(methodMock, request, SlideMasterPart, "related_part", $kw({return_value: slide_layout_part_}));
        slide_master_part = $call(SlideMasterPart, null, null, null, null);
        slide_layout = slide_master_part.relatedSlideLayout("rId42");
        related_part_.assertCalledOnceWith(slide_master_part, "rId42");
        $assert($is(slide_layout, slide_layout_));
      } },
  ],
};
pytestClass("DescribeSlideMasterPart", DescribeSlideMasterPart);

