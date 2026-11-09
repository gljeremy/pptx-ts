// Machine-translated from python-pptx tests/shapes/test_shapetree.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $call, $d, $dictget, $endswith, $eq, $get, $icon, $index, $is, $iter, $kw, $len, $list, $mod, $mul, $raises, $slice, $t1, $t2, $t3, $t4, $t5, $truthy, ANY, BytesIO, call, classMock, functionMock, initializerMock, instanceMock, methodMock, propertyMock, pytestClass } from "../pyshim.js";
import { templateBytes } from "../../src/templates/index.js";
import { $assert, $assertEqual, $call, $d, $dictget, $endswith, $get, $index, $is, $iter, $kw, $len, $list, $mod, $mul, $raises, $slice, $t1, $t2, $t3, $t4, $t5, $truthy, ANY, BytesIO, call, classMock, functionMock, initializerMock, instanceMock, methodMock, propertyMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { ChartData as ChartData$imported } from "../../src/chart/data.js";
import { XL_CHART_TYPE as XL_CHART_TYPE$imported } from "../../src/enum/chart.js";
import { MSO_AUTO_SHAPE_TYPE as MSO_AUTO_SHAPE_TYPE$imported, MSO_CONNECTOR as MSO_CONNECTOR$imported, PP_PLACEHOLDER as PP_PLACEHOLDER$imported, PROG_ID as PROG_ID$imported } from "../../src/enum/shapes.js";
import { SPEAKER_IMAGE_BYTES as SPEAKER_IMAGE_BYTES$imported, Video as Video$imported } from "../../src/media.js";
import { parseXml as parseXml$imported } from "../../src/oxml/index.js";
import { CT_GroupShape as CT_GroupShape$imported } from "../../src/oxml/shapes/groupshape.js";
import { CT_Picture as CT_Picture$imported } from "../../src/oxml/shapes/picture.js";
import { BaseShapeElement as BaseShapeElement$imported, ST_Direction as ST_Direction$imported } from "../../src/oxml/shapes/shared.js";
import { ImagePart as ImagePart$imported } from "../../src/parts/image.js";
import { SlidePart as SlidePart$imported } from "../../src/parts/slide.js";
import { AutoShapeType as AutoShapeType$imported, Shape as Shape$imported } from "../../src/shapes/autoshape.js";
import { BaseShape as BaseShape$imported } from "../../src/shapes/base.js";
import { Connector as Connector$imported } from "../../src/shapes/connector.js";
import { FreeformBuilder as FreeformBuilder$imported } from "../../src/shapes/freeform.js";
import { GraphicFrame as GraphicFrame$imported } from "../../src/shapes/graphfrm.js";
import { GroupShape as GroupShape$imported } from "../../src/shapes/group.js";
import { Movie as Movie$imported, Picture as Picture$imported } from "../../src/shapes/picture.js";
import { LayoutPlaceholder as LayoutPlaceholder$imported, MasterPlaceholder as MasterPlaceholder$imported, NotesSlidePlaceholder as NotesSlidePlaceholder$imported, _BaseSlidePlaceholder as _BaseSlidePlaceholder$imported } from "../../src/shapes/placeholder.js";
import { BasePlaceholders as BasePlaceholders$imported, BaseShapeFactory as BaseShapeFactory$imported, GroupShapes as GroupShapes$imported, LayoutPlaceholders as LayoutPlaceholders$imported, LayoutShapes as LayoutShapes$imported, MasterPlaceholders as MasterPlaceholders$imported, MasterShapes as MasterShapes$imported, NotesSlidePlaceholders as NotesSlidePlaceholders$imported, NotesSlideShapes as NotesSlideShapes$imported, SlidePlaceholders as SlidePlaceholders$imported, SlideShapeFactory as SlideShapeFactory$imported, SlideShapes as SlideShapes$imported, _BaseGroupShapes as _BaseGroupShapes$imported, _BaseShapes as _BaseShapes$imported, _LayoutShapeFactory as _LayoutShapeFactory$imported, _MasterShapeFactory as _MasterShapeFactory$imported, _MoviePicElementCreator, _NotesSlideShapeFactory as _NotesSlideShapeFactory$imported, _OleObjectElementCreator, _SlidePlaceholderFactory as _SlidePlaceholderFactory$imported } from "../../src/shapes/shapetree.js";
import { SlideLayout as SlideLayout$imported, SlideMaster as SlideMaster$imported } from "../../src/slide.js";
import { Table as Table$imported } from "../../src/table.js";
import { Emu as Emu$imported, Inches as Inches$imported } from "../../src/util.js";
import { aPh, aPic, anNvPr, anNvSpPr, anSp } from "../oxml/unitdata/shape.js";
import { element, xml } from "../unitutil/cxml.js";
import { snippetSeq } from "../unitutil/file.js";
// python binds imported names at import time; patches made later don't change them
const ChartData = ChartData$imported;
const XL_CHART_TYPE = XL_CHART_TYPE$imported;
const MSO_AUTO_SHAPE_TYPE = MSO_AUTO_SHAPE_TYPE$imported;
const MSO_CONNECTOR = MSO_CONNECTOR$imported;
const PP_PLACEHOLDER = PP_PLACEHOLDER$imported;
const PROG_ID = PROG_ID$imported;
const SPEAKER_IMAGE_BYTES = SPEAKER_IMAGE_BYTES$imported;
const Video = Video$imported;
const parseXml = parseXml$imported;
const CT_GroupShape = CT_GroupShape$imported;
const CT_Picture = CT_Picture$imported;
const BaseShapeElement = BaseShapeElement$imported;
const ST_Direction = ST_Direction$imported;
const ImagePart = ImagePart$imported;
const SlidePart = SlidePart$imported;
const AutoShapeType = AutoShapeType$imported;
const Shape = Shape$imported;
const BaseShape = BaseShape$imported;
const Connector = Connector$imported;
const FreeformBuilder = FreeformBuilder$imported;
const GraphicFrame = GraphicFrame$imported;
const GroupShape = GroupShape$imported;
const Movie = Movie$imported;
const Picture = Picture$imported;
const LayoutPlaceholder = LayoutPlaceholder$imported;
const MasterPlaceholder = MasterPlaceholder$imported;
const NotesSlidePlaceholder = NotesSlidePlaceholder$imported;
const _BaseSlidePlaceholder = _BaseSlidePlaceholder$imported;
const BasePlaceholders = BasePlaceholders$imported;
const BaseShapeFactory = BaseShapeFactory$imported;
const GroupShapes = GroupShapes$imported;
const LayoutPlaceholders = LayoutPlaceholders$imported;
const LayoutShapes = LayoutShapes$imported;
const MasterPlaceholders = MasterPlaceholders$imported;
const MasterShapes = MasterShapes$imported;
const NotesSlidePlaceholders = NotesSlidePlaceholders$imported;
const NotesSlideShapes = NotesSlideShapes$imported;
const SlidePlaceholders = SlidePlaceholders$imported;
const SlideShapeFactory = SlideShapeFactory$imported;
const SlideShapes = SlideShapes$imported;
const _BaseGroupShapes = _BaseGroupShapes$imported;
const _BaseShapes = _BaseShapes$imported;
const _LayoutShapeFactory = _LayoutShapeFactory$imported;
const _MasterShapeFactory = _MasterShapeFactory$imported;
const _NotesSlideShapeFactory = _NotesSlideShapeFactory$imported;
const _SlidePlaceholderFactory = _SlidePlaceholderFactory$imported;
const SlideLayout = SlideLayout$imported;
const SlideMaster = SlideMaster$imported;
const Table = Table$imported;
const Emu = Emu$imported;
const Inches = Inches$imported;
void $py;


const DescribeBaseShapeFactory: any = {
  fixtures: {
    factory_fixture: { args: ["request", "parent_"], params: [["p:sp", Shape], ["p:pic", Picture], ["p:pic/p:nvPicPr/p:nvPr/a:videoFile", Movie], ["p:graphicFrame", GraphicFrame], ["p:grpSp", GroupShape], ["p:cxnSp", Connector]], fn: (self: any, request: any, parent_: any): any => {
        let ShapeClass_: any, ShapeCls: any, shape_cxml: any, shape_elm: any, shape_mock: any;
        [shape_cxml, ShapeCls] = $list(request.param);
        shape_elm = element(shape_cxml);
        shape_mock = instanceMock(request, ShapeCls);
        ShapeClass_ = $call(classMock, request, $mod("pptx.shapes.shapetree.%s", ShapeCls.name), $kw({return_value: shape_mock}));
        return [shape_elm, parent_, ShapeClass_, shape_mock];
      } },
    parent_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlideShapes);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_constructs_the_right_shape_for_an_element", args: ["factory_fixture"], fn: (self: any, factory_fixture: any): any => {
        let ShapeClass_: any, parent_: any, shape: any, shape_: any, shape_elm: any;
        [shape_elm, parent_, ShapeClass_, shape_] = $list(factory_fixture);
        shape = $call(BaseShapeFactory, shape_elm, parent_);
        ShapeClass_.assertCalledOnceWith(shape_elm, parent_);
        $assert($is(shape, shape_));
      } },
  ],
};
pytestClass("DescribeBaseShapeFactory", DescribeBaseShapeFactory);

const Describe_BaseShapes: any = {
  fixtures: {
    clone_ph_fixture: { args: ["placeholder_"], fn: (self: any, placeholder_: any): any => {
        let expected_xml: any, shapes: any;
        shapes = $call(SlideShapes, element("p:spTree{a:b=c}"), null);
        expected_xml = xml("p:spTree{a:b=c}/p:sp/(p:nvSpPr/(p:cNvPr{id=1,name=Vertical Chart Placeholder 0},p:cNvSpPr/a:spLocks{noGrp=1},p:nvPr/p:ph{type=chart,idx=42,orient=vert,sz=half}),p:spPr)");
        placeholder_.element = element("p:sp/p:nvSpPr/p:nvPr/p:ph{type=chart,idx=42,orient=vert,sz=half}");
        return [shapes, placeholder_, expected_xml];
      } },
    getitem_fixture: { args: ["BaseShapeFactory_", "shape_"], fn: (self: any, BaseShapeFactory_: any, shape_: any): any => {
        let idx: any, shapes: any, sp: any, spTree: any;
        spTree = element("p:spTree/(p:sp,p:sp)");
        shapes = $call(_BaseShapes, spTree, null);
        idx = 1;
        sp = $get(spTree.xpath("p:sp"), idx);
        return [shapes, idx, BaseShapeFactory_, sp, shape_];
      } },
    getitem_raises_fixture: { args: [], fn: (self: any): any => {
        let shapes: any, spTree: any;
        spTree = element("p:spTree/(p:sp,p:sp)");
        shapes = $call(_BaseShapes, spTree, null);
        return shapes;
      } },
    iter_fixture: { args: ["BaseShapeFactory_"], fn: (self: any, BaseShapeFactory_: any): any => {
        let calls: any, expected_shapes: any, shapes: any, spTree: any, sps: any;
        spTree = element("p:spTree/(p:sp,p:sp)");
        sps = spTree.xpath("p:sp");
        shapes = $call(_BaseShapes, spTree, null);
        expected_shapes = [$call(Shape, null, null), $call(Shape, null, null)];
        calls = [call($get(sps, 0), shapes), call($get(sps, 1), shapes)];
        BaseShapeFactory_.sideEffect = $iter(expected_shapes);
        return [shapes, expected_shapes, BaseShapeFactory_, calls];
      } },
    iter_elms_fixture: { args: [], fn: (self: any): any => {
        let shapes: any, spTree: any, sps: any;
        spTree = element("p:spTree/(p:sp,p:sp)");
        sps = spTree.xpath("p:sp");
        shapes = $call(_BaseShapes, spTree, null);
        return [shapes, sps];
      } },
    len_fixture: { args: [], fn: (self: any): any => {
        let expected_count: any, shapes: any;
        shapes = $call(_BaseShapes, element("p:spTree/(p:spPr,p:sp,p:sp)"), null);
        expected_count = 2;
        return [shapes, expected_count];
      } },
    next_id_fixture: { args: ["request"], params: [["p:spTree/p:nvSpPr", 1], ["p:spTree/p:nvSpPr/p:cNvPr{id=0}", 1], ["p:spTree/p:nvSpPr/p:cNvPr{id=1}", 2], ["p:spTree/p:nvSpPr/p:cNvPr{id=2}", 3], ["p:spTree/p:nvSpPr/(p:cNvPr{id=1},p:cNvPr{id=3})", 4], ["p:spTree/p:nvSpPr/(p:cNvPr{id=foo},p:cNvPr{id=2})", 3], ["p:spTree/p:nvSpPr/(p:cNvPr{id=1fo},p:cNvPr{id=2})", 3], ["p:spTree/p:nvSpPr/(p:cNvPr{id=1},p:cNvPr{id=1},p:cNvPr{id=1},p:cNvPr{id=4})", 5]], fn: (self: any, request: any): any => {
        let expected_value: any, shapes: any, spTree_cxml: any;
        [spTree_cxml, expected_value] = $list(request.param);
        shapes = $call(_BaseShapes, element(spTree_cxml), null);
        return [shapes, expected_value];
      } },
    ph_name_fixture: { args: ["request"], params: [[PP_PLACEHOLDER.OBJECT, 3, ST_Direction.HORZ, "Content Placeholder 2"], [PP_PLACEHOLDER.TABLE, 4, ST_Direction.HORZ, "Table Placeholder 4"], [PP_PLACEHOLDER.TABLE, 7, ST_Direction.VERT, "Vertical Table Placeholder 6"], [PP_PLACEHOLDER.TITLE, 2, ST_Direction.HORZ, "Title 2"]], fn: (self: any, request: any): any => {
        let expected_name: any, orient: any, ph_type: any, shapes: any, spTree: any, sp_id: any;
        [ph_type, sp_id, orient, expected_name] = $list(request.param);
        spTree = element("p:spTree/(p:cNvPr{name=Title 1},p:cNvPr{name=Table Placeholder 3})");
        shapes = $call(SlideShapes, spTree, null);
        return [shapes, ph_type, sp_id, orient, expected_name];
      } },
    turbo_fixture: { args: ["request"], params: [[null, false], [42, true]], fn: (self: any, request: any): any => {
        let cached_max_shape_id: any, expected_value: any, shapes: any;
        [cached_max_shape_id, expected_value] = $list(request.param);
        shapes = $call(_BaseShapes, null, null);
        if ($truthy(cached_max_shape_id)) {
          shapes._cachedMaxShapeId = cached_max_shape_id;
        }
        return [shapes, expected_value];
      } },
    turbo_set_fixture: { args: ["request"], params: [["p:spTree/p:nvSpPr", true], ["p:spTree/p:nvSpPr/p:cNvPr{id=2}", true], ["p:spTree/p:nvSpPr/(p:cNvPr{id=1},p:cNvPr{id=3})", false], ["p:spTree/p:nvSpPr/(p:cNvPr{id=1},p:cNvPr{id=1},p:cNvPr{id=1},p:cNvPr{id=4})", true]], fn: (self: any, request: any): any => {
        let expected_value: any, shapes: any, spTree_cxml: any, value: any;
        [spTree_cxml, value] = $list(request.param);
        shapes = $call(_BaseShapes, element(spTree_cxml), null);
        expected_value = value;
        return [shapes, value, expected_value];
      } },
    BaseShapeFactory_: { args: ["request", "shape_"], fn: (self: any, request: any, shape_: any): any => {
        return $call(functionMock, request, "pptx.shapes.shapetree.BaseShapeFactory", $kw({return_value: shape_, autospec: true}));
      } },
    placeholder_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Shape);
      } },
    shape_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, BaseShape);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_how_many_shapes_it_contains", args: ["len_fixture"], fn: (self: any, len_fixture: any): any => {
        let expected_count: any, shapes: any;
        [shapes, expected_count] = $list(len_fixture);
        $assertEqual($len(shapes), expected_count);
      } },
    { name: "it_can_iterate_over_the_shapes_it_contains", args: ["iter_fixture"], fn: (self: any, iter_fixture: any): any => {
        let BaseShapeFactory_: any, calls: any, expected_shapes: any, shapes: any;
        [shapes, expected_shapes, BaseShapeFactory_, calls] = $list(iter_fixture);
        $assertEqual($list(shapes).map((s: any) => s), expected_shapes);
        $assertEqual(BaseShapeFactory_.callArgsList, calls);
      } },
    { name: "it_iterates_shape_elements_to_help__iter__", args: ["iter_elms_fixture"], fn: (self: any, iter_elms_fixture: any): any => {
        let expected_elms: any, shapes: any;
        [shapes, expected_elms] = $list(iter_elms_fixture);
        $assertEqual($list(shapes._iterMemberElms()).map((e: any) => e), expected_elms);
      } },
    { name: "it_supports_indexed_access", args: ["getitem_fixture"], fn: (self: any, getitem_fixture: any): any => {
        let BaseShapeFactory_: any, idx: any, shape: any, shape_: any, shapes: any, sp: any;
        [shapes, idx, BaseShapeFactory_, sp, shape_] = $list(getitem_fixture);
        shape = $get(shapes, idx);
        BaseShapeFactory_.assertCalledOnceWith(sp, shapes);
        $assert($is(shape, shape_));
      } },
    { name: "it_raises_on_shape_index_out_of_range", args: ["getitem_raises_fixture"], fn: (self: any, getitem_raises_fixture: any): any => {
        let shapes: any;
        shapes = getitem_raises_fixture;
        $raises(IndexError, () => {
          $get(shapes, 2);
        });
      } },
    { name: "it_can_clone_a_placeholder", args: ["clone_ph_fixture"], fn: (self: any, clone_ph_fixture: any): any => {
        let expected_xml: any, placeholder_: any, shapes: any;
        [shapes, placeholder_, expected_xml] = $list(clone_ph_fixture);
        shapes.clonePlaceholder(placeholder_);
        $assertEqual(shapes._element.xml, expected_xml);
      } },
    { name: "it_knows_if_turbo_add_is_enabled", args: ["turbo_fixture"], fn: (self: any, turbo_fixture: any): any => {
        let expected_value: any, shapes: any, turbo_add_enabled: any;
        [shapes, expected_value] = $list(turbo_fixture);
        turbo_add_enabled = shapes.turboAddEnabled;
        $assertEqual(turbo_add_enabled, expected_value);
      } },
    { name: "it_can_change_turbo_add_enabled", args: ["turbo_set_fixture"], fn: (self: any, turbo_set_fixture: any): any => {
        let expected_value: any, shapes: any, value: any;
        [shapes, value, expected_value] = $list(turbo_set_fixture);
        shapes.turboAddEnabled = value;
        $assertEqual(shapes.turboAddEnabled, expected_value);
      } },
    { name: "it_finds_the_next_shape_id_to_help", args: ["next_id_fixture"], fn: (self: any, next_id_fixture: any): any => {
        let expected_value: any, shapes: any;
        [shapes, expected_value] = $list(next_id_fixture);
        $assertEqual(shapes._nextShapeId, expected_value);
      } },
    { name: "it_finds_the_next_placeholder_name_to_help", args: ["ph_name_fixture"], fn: (self: any, ph_name_fixture: any): any => {
        let expected_value: any, orient: any, ph_type: any, shapes: any, sp_id: any;
        [shapes, ph_type, sp_id, orient, expected_value] = $list(ph_name_fixture);
        $assertEqual(shapes._nextPhName(ph_type, sp_id, orient), expected_value);
      } },
  ],
};
pytestClass("Describe_BaseShapes", Describe_BaseShapes);

const Describe_BaseGroupShapes: any = {
  fixtures: {
    add_cht_gr_frm_fixture: { args: [], fn: (self: any): any => {
        let cx: any, cy: any, expected_xml: any, rId: any, shapes: any, x: any, y: any;
        shapes = $call(_BaseGroupShapes, element("p:spTree"), null);
        [rId, x, y, cx, cy] = $list(["rId42", 1, 2, 3, 4]);
        expected_xml = "<p:spTree xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">\n  <p:graphicFrame xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">\n    <p:nvGraphicFramePr>\n      <p:cNvPr id=\"1\" name=\"Chart 0\"/>\n      <p:cNvGraphicFramePr>\n        <a:graphicFrameLocks noGrp=\"1\"/>\n      </p:cNvGraphicFramePr>\n      <p:nvPr/>\n    </p:nvGraphicFramePr>\n    <p:xfrm>\n      <a:off x=\"1\" y=\"2\"/>\n      <a:ext cx=\"3\" cy=\"4\"/>\n    </p:xfrm>\n    <a:graphic>\n      <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">\n        <c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId42\"/>\n      </a:graphicData>\n    </a:graphic>\n  </p:graphicFrame>\n</p:spTree>";
        return [shapes, rId, x, y, cx, cy, expected_xml];
      } },
    add_cxnSp_fixture: { args: ["request"], params: [[1, 2, 3, 5, "p:spPr/(a:xfrm/(a:off{x=1,y=2},a:ext{cx=2,cy=3})"], [8, 3, 4, 9, "p:spPr/(a:xfrm{flipH=1}/(a:off{x=4,y=3},a:ext{cx=4,cy=6})"], [1, 6, 5, 2, "p:spPr/(a:xfrm{flipV=1}/(a:off{x=1,y=2},a:ext{cx=4,cy=4})"], [9, 8, 2, 3, "p:spPr/(a:xfrm{flipH=1,flipV=1}/(a:off{x=2,y=3},a:ext{cx=7,cy=5})"]], fn: (self: any, request: any): any => {
        let begin_x: any, begin_y: any, connector_type: any, end_x: any, end_y: any, expected_xml: any, shapes: any, spPr_cxml: any, tmpl_cxml: any;
        [begin_x, begin_y, end_x, end_y, spPr_cxml] = $list(request.param);
        shapes = $call(_BaseGroupShapes, element("p:spTree"), null);
        connector_type = MSO_CONNECTOR.STRAIGHT;
        tmpl_cxml = "p:cxnSp/(p:nvCxnSpPr/(p:cNvPr{id=1,name=Connector 0},p:cNvCxnSpPr,p:nvPr),%s,a:prstGeom{prst=line}/a:avLst),p:style/(a:lnRef{idx=2}/a:schemeClr{val=accent1},a:fillRef{idx=0}/a:schemeClr{val=accent1},a:effectRef{idx=1}/a:schemeClr{val=accent1},a:fontRef{idx=minor}/a:schemeClr{val=tx1}))";
        expected_xml = xml($mod(tmpl_cxml, spPr_cxml));
        return [shapes, connector_type, begin_x, begin_y, end_x, end_y, expected_xml];
      } },
    add_pic_fixture: { args: ["image_part_", "_next_shape_id_prop_"], fn: (self: any, image_part_: any, _next_shape_id_prop_: any): any => {
        let cx: any, cy: any, expected_xml: any, rId: any, shapes: any, x: any, y: any;
        shapes = $call(_BaseGroupShapes, element("p:spTree"), null);
        [rId, x, y, cx, cy] = $list(["rId24", 10, 11, 12, 13]);
        _next_shape_id_prop_.returnValue = 42;
        image_part_.scale.returnValue = [101, 102];
        image_part_.desc = "sprocket.jpg";
        expected_xml = "<p:spTree xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">\n  <p:pic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n    <p:nvPicPr>\n      <p:cNvPr id=\"42\" name=\"Picture 41\" descr=\"sprocket.jpg\"/>\n      <p:cNvPicPr>\n        <a:picLocks noChangeAspect=\"1\"/>\n      </p:cNvPicPr>\n      <p:nvPr/>\n    </p:nvPicPr>\n    <p:blipFill>\n      <a:blip r:embed=\"rId24\"/>\n      <a:stretch>\n        <a:fillRect/>\n      </a:stretch>\n    </p:blipFill>\n    <p:spPr>\n      <a:xfrm>\n        <a:off x=\"10\" y=\"11\"/>\n        <a:ext cx=\"101\" cy=\"102\"/>\n      </a:xfrm>\n      <a:prstGeom prst=\"rect\">\n        <a:avLst/>\n      </a:prstGeom>\n    </p:spPr>\n  </p:pic>\n</p:spTree>";
        return [shapes, image_part_, rId, x, y, cx, cy, expected_xml];
      } },
    add_sp_fixture: { args: ["autoshape_type_", "_next_shape_id_prop_"], fn: (self: any, autoshape_type_: any, _next_shape_id_prop_: any): any => {
        let cx: any, cy: any, expected_xml: any, shapes: any, x: any, y: any;
        shapes = $call(_BaseGroupShapes, element("p:spTree"), null);
        [x, y, cx, cy] = $list([8, 7, 6, 5]);
        _next_shape_id_prop_.returnValue = 7;
        autoshape_type_.basename = "Rounded Rectangle";
        autoshape_type_.prst = "roundRect";
        expected_xml = "<p:spTree xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">\n  <p:sp xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">\n    <p:nvSpPr>\n      <p:cNvPr id=\"7\" name=\"Rounded Rectangle 6\"/>\n      <p:cNvSpPr/>\n      <p:nvPr/>\n    </p:nvSpPr>\n    <p:spPr>\n      <a:xfrm>\n        <a:off x=\"8\" y=\"7\"/>\n        <a:ext cx=\"6\" cy=\"5\"/>\n      </a:xfrm>\n      <a:prstGeom prst=\"roundRect\">\n        <a:avLst/>\n      </a:prstGeom>\n    </p:spPr>\n    <p:style>\n      <a:lnRef idx=\"1\">\n        <a:schemeClr val=\"accent1\"/>\n      </a:lnRef>\n      <a:fillRef idx=\"3\">\n        <a:schemeClr val=\"accent1\"/>\n      </a:fillRef>\n      <a:effectRef idx=\"2\">\n        <a:schemeClr val=\"accent1\"/>\n      </a:effectRef>\n      <a:fontRef idx=\"minor\">\n        <a:schemeClr val=\"lt1\"/>\n      </a:fontRef>\n    </p:style>\n    <p:txBody>\n      <a:bodyPr rtlCol=\"0\" anchor=\"ctr\"/>\n      <a:lstStyle/>\n      <a:p>\n        <a:pPr algn=\"ctr\"/>\n      </a:p>\n    </p:txBody>\n  </p:sp>\n</p:spTree>";
        return [shapes, autoshape_type_, x, y, cx, cy, expected_xml];
      } },
    add_textbox_sp_fixture: { args: ["_next_shape_id_prop_"], fn: (self: any, _next_shape_id_prop_: any): any => {
        let cx: any, cy: any, expected_xml: any, shapes: any, x: any, y: any;
        shapes = $call(_BaseGroupShapes, element("p:spTree"), null);
        [x, y, cx, cy] = $list([1, 2, 3, 4]);
        _next_shape_id_prop_.returnValue = 6;
        expected_xml = "<p:spTree xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">\n  <p:sp xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">\n    <p:nvSpPr>\n      <p:cNvPr id=\"6\" name=\"TextBox 5\"/>\n      <p:cNvSpPr txBox=\"1\"/>\n      <p:nvPr/>\n    </p:nvSpPr>\n    <p:spPr>\n      <a:xfrm>\n        <a:off x=\"1\" y=\"2\"/>\n        <a:ext cx=\"3\" cy=\"4\"/>\n      </a:xfrm>\n      <a:prstGeom prst=\"rect\">\n        <a:avLst/>\n      </a:prstGeom>\n      <a:noFill/>\n    </p:spPr>\n    <p:txBody>\n      <a:bodyPr wrap=\"none\">\n        <a:spAutoFit/>\n      </a:bodyPr>\n      <a:lstStyle/>\n      <a:p/>\n    </p:txBody>\n  </p:sp>\n</p:spTree>";
        return [shapes, x, y, cx, cy, expected_xml];
      } },
    connector_fixture: { args: ["_add_cxnSp_", "_shape_factory_", "_recalculate_extents_", "connector_"], fn: (self: any, _add_cxnSp_: any, _shape_factory_: any, _recalculate_extents_: any, connector_: any): any => {
        let begin_x: any, begin_y: any, connector_type: any, cxnSp: any, end_x: any, end_y: any, shapes: any;
        shapes = $call(_BaseGroupShapes, element("p:spTree"), null);
        connector_type = MSO_CONNECTOR.STRAIGHT;
        [begin_x, begin_y, end_x, end_y] = $list([1, 2, 3, 4]);
        cxnSp = element("p:cxnSp");
        _add_cxnSp_.returnValue = cxnSp;
        _shape_factory_.returnValue = connector_;
        return [shapes, connector_type, begin_x, begin_y, end_x, end_y, cxnSp, connector_];
      } },
    freeform_fixture: { args: ["request", "FreeformBuilder_new_", "builder_"], params: [[0, 0, 1.0, 1.0, 1.0], [100, 200, 2.0, 2.0, 2.0], [100, 200, [4.2, 2.4], 4.2, 2.4]], fn: (self: any, request: any, FreeformBuilder_new_: any, builder_: any): any => {
        let scale: any, shapes: any, start_x: any, start_y: any, x_scale: any, y_scale: any;
        [start_x, start_y, scale, x_scale, y_scale] = $list(request.param);
        shapes = $call(_BaseGroupShapes, null, null);
        FreeformBuilder_new_.returnValue = builder_;
        return [shapes, start_x, start_y, scale, FreeformBuilder_new_, x_scale, y_scale, builder_];
      } },
    group_fixture: { args: ["CT_GroupShape_add_grpSp_", "_shape_factory_", "group_shape_"], fn: (self: any, CT_GroupShape_add_grpSp_: any, _shape_factory_: any, group_shape_: any): any => {
        let grpSp: any, shapes: any, spTree: any;
        spTree = element("p:spTree{id=2e838acdc755e83113ed03904d2fe081f}");
        grpSp = element("p:grpSp{id=052874e154b48f9bec4266f80913cae38f}");
        shapes = $call(_BaseGroupShapes, spTree, null);
        CT_GroupShape_add_grpSp_.returnValue = grpSp;
        _shape_factory_.returnValue = group_shape_;
        return [shapes, spTree, grpSp, group_shape_];
      } },
    index_fixture: { args: ["request", "shape_"], params: [["p:spTree/(p:sp,p:sp,p:sp)", SlideShapes, 0], ["p:spTree/(p:sp,p:sp,p:sp)", SlideShapes, 1], ["p:spTree/(p:sp,p:sp,p:sp)", SlideShapes, 2], ["p:grpSp/(p:sp,p:sp,p:sp)", GroupShapes, 0], ["p:grpSp/(p:sp,p:sp,p:sp)", GroupShapes, 1], ["p:grpSp/(p:sp,p:sp,p:sp)", GroupShapes, 2]], fn: (self: any, request: any, shape_: any): any => {
        let ShapesCls: any, expected_value: any, grpSp: any, grpSp_cxml: any, idx: any, shapes: any, sps: any;
        [grpSp_cxml, ShapesCls, idx] = $list(request.param);
        grpSp = element(grpSp_cxml);
        sps = grpSp.xpath("p:sp");
        shapes = $call(ShapesCls, grpSp, null);
        shape_.element = $get(sps, idx);
        expected_value = idx;
        return [shapes, shape_, expected_value];
      } },
    index_raises_fixture: { args: ["request", "shape_"], params: [["p:spTree/(p:sp,p:sp,p:sp)", SlideShapes], ["p:grpSp/(p:sp,p:sp,p:sp)", GroupShapes]], fn: (self: any, request: any, shape_: any): any => {
        let ShapesCls: any, grpSp_cxml: any, shapes: any;
        [grpSp_cxml, ShapesCls] = $list(request.param);
        shapes = $call(SlideShapes, element(grpSp_cxml), null);
        shape_.element = element("p:sp");
        return [shapes, shape_];
      } },
    _next_shape_id_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, _BaseGroupShapes, "_next_shape_id");
      } },
    picture_fixture: { args: ["part_prop_", "slide_part_", "image_part_", "_add_pic_from_image_part_", "_recalculate_extents_", "_shape_factory_", "picture_"], fn: (self: any, part_prop_: any, slide_part_: any, image_part_: any, _add_pic_from_image_part_: any, _recalculate_extents_: any, _shape_factory_: any, picture_: any): any => {
        let cx: any, cy: any, image_file: any, pic: any, rId: any, shapes: any, x: any, y: any;
        shapes = $call(_BaseGroupShapes, null, null);
        [image_file, x, y, cx, cy, rId] = $list(["foobar.png", 1, 2, 3, 4, "rId42"]);
        pic = element("p:pic");
        part_prop_.returnValue = slide_part_;
        slide_part_.getOrAddImagePart.returnValue = [image_part_, rId];
        _add_pic_from_image_part_.returnValue = pic;
        _shape_factory_.returnValue = picture_;
        return [shapes, image_file, x, y, cx, cy, image_part_, rId, pic, picture_];
      } },
    shape_fixture: { args: ["AutoShapeType_", "autoshape_type_", "_add_sp_", "_recalculate_extents_", "_shape_factory_", "shape_"], fn: (self: any, AutoShapeType_: any, autoshape_type_: any, _add_sp_: any, _recalculate_extents_: any, _shape_factory_: any, shape_: any): any => {
        let autoshape_type_id: any, cx: any, cy: any, shapes: any, sp: any, x: any, y: any;
        shapes = $call(_BaseGroupShapes, null, null);
        autoshape_type_id = MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE;
        [x, y, cx, cy] = $list([21, 22, 23, 24]);
        sp = element("p:sp");
        AutoShapeType_.returnValue = autoshape_type_;
        _add_sp_.returnValue = sp;
        _shape_factory_.returnValue = shape_;
        return [shapes, autoshape_type_id, x, y, cx, cy, AutoShapeType_, autoshape_type_, sp, shape_];
      } },
    textbox_fixture: { args: ["_add_textbox_sp_", "_recalculate_extents_", "_shape_factory_", "shape_"], fn: (self: any, _add_textbox_sp_: any, _recalculate_extents_: any, _shape_factory_: any, shape_: any): any => {
        let cx: any, cy: any, shapes: any, sp: any, x: any, y: any;
        shapes = $call(_BaseGroupShapes, null, null);
        [x, y, cx, cy] = $list([31, 32, 33, 34]);
        sp = element("p:sp");
        _add_textbox_sp_.returnValue = sp;
        _shape_factory_.returnValue = shape_;
        return [shapes, x, y, cx, cy, sp, shape_];
      } },
    _add_chart_graphicFrame_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, _BaseGroupShapes, "_add_chart_graphicFrame", $kw({autospec: true}));
      } },
    _add_cxnSp_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, _BaseGroupShapes, "_add_cxnSp", $kw({autospec: true}));
      } },
    _add_pic_from_image_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, _BaseGroupShapes, "_add_pic_from_image_part", $kw({autospec: true}));
      } },
    _add_sp_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, _BaseGroupShapes, "_add_sp", $kw({autospec: true}));
      } },
    _add_textbox_sp_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, _BaseGroupShapes, "_add_textbox_sp", $kw({autospec: true}));
      } },
    autoshape_type_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, AutoShapeType);
      } },
    AutoShapeType_: { args: ["request"], fn: (self: any, request: any): any => {
        return classMock(request, "pptx.shapes.shapetree.AutoShapeType");
      } },
    builder_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, FreeformBuilder);
      } },
    chart_data_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, ChartData);
      } },
    connector_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Connector);
      } },
    CT_GroupShape_add_grpSp_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, CT_GroupShape, "add_grpSp", $kw({autospec: true}));
      } },
    FreeformBuilder_new_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, FreeformBuilder, "new", $kw({autospec: false}));
      } },
    graphic_frame_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, GraphicFrame);
      } },
    group_shape_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, GroupShape);
      } },
    image_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, ImagePart);
      } },
    part_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, _BaseGroupShapes, "part");
      } },
    picture_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Picture);
      } },
    _recalculate_extents_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, _BaseGroupShapes, "_recalculate_extents", $kw({autospec: true}));
      } },
    shape_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Shape);
      } },
    _shape_factory_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, _BaseGroupShapes, "_shape_factory", $kw({autospec: true}));
      } },
    slide_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlidePart);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_add_a_chart", args: ["chart_data_", "part_prop_", "slide_part_", "graphic_frame_", "_add_chart_graphicFrame_", "_recalculate_extents_", "_shape_factory_"], fn: (self: any, chart_data_: any, part_prop_: any, slide_part_: any, graphic_frame_: any, _add_chart_graphicFrame_: any, _recalculate_extents_: any, _shape_factory_: any): any => {
        let cx: any, cy: any, graphicFrame: any, graphic_frame: any, shapes: any, x: any, y: any;
        [x, y, cx, cy] = $list([1, 2, 3, 4]);
        part_prop_.returnValue = slide_part_;
        slide_part_.addChartPart.returnValue = "rId42";
        graphicFrame = element("p:graphicFrame");
        _add_chart_graphicFrame_.returnValue = graphicFrame;
        _shape_factory_.returnValue = graphic_frame_;
        shapes = $call(_BaseGroupShapes, null, null);
        graphic_frame = shapes.addChart(XL_CHART_TYPE.PIE, x, y, cx, cy, chart_data_);
        shapes.part.addChartPart.assertCalledOnceWith(XL_CHART_TYPE.PIE, chart_data_);
        _add_chart_graphicFrame_.assertCalledOnceWith(shapes, "rId42", x, y, cx, cy);
        _recalculate_extents_.assertCalledOnceWith(shapes);
        _shape_factory_.assertCalledOnceWith(shapes, graphicFrame);
        $assert($is(graphic_frame, graphic_frame_));
      } },
    { name: "it_can_add_a_connector_shape", args: ["connector_fixture"], fn: (self: any, connector_fixture: any): any => {
        let begin_x: any, begin_y: any, connector: any, connector_: any, connector_type: any, cxnSp_: any, end_x: any, end_y: any, shapes: any;
        [shapes, connector_type, begin_x, begin_y] = $list($slice(connector_fixture, null, 4, null));
        [end_x, end_y, cxnSp_, connector_] = $list($slice(connector_fixture, 4, null, null));
        connector = shapes.addConnector(connector_type, begin_x, begin_y, end_x, end_y);
        shapes._addCxnSp.assertCalledOnceWith(shapes, connector_type, begin_x, begin_y, end_x, end_y);
        shapes._recalculateExtents.assertCalledOnceWith(shapes);
        shapes._shapeFactory.assertCalledOnceWith(shapes, cxnSp_);
        $assert($is(connector, connector_));
      } },
    { name: "it_can_provide_a_freeform_builder", args: ["freeform_fixture"], fn: (self: any, freeform_fixture: any): any => {
        let FreeformBuilder_new_: any, builder: any, builder_: any, scale: any, shapes: any, start_x: any, start_y: any, x_scale: any, y_scale: any;
        [shapes, start_x, start_y, scale] = $list($slice(freeform_fixture, null, 4, null));
        [FreeformBuilder_new_, x_scale, y_scale] = $list($slice(freeform_fixture, 4, 7, null));
        builder_ = $get(freeform_fixture, 7);
        builder = shapes.buildFreeform(start_x, start_y, scale);
        FreeformBuilder_new_.assertCalledOnceWith(shapes, start_x, start_y, x_scale, y_scale);
        $assert($is(builder, builder_));
      } },
    { name: "it_can_add_a_group_shape", args: ["group_fixture"], fn: (self: any, group_fixture: any): any => {
        let group_shape: any, group_shape_: any, grpSp: any, shapes: any, spTree: any;
        [shapes, spTree, grpSp, group_shape_] = $list(group_fixture);
        group_shape = shapes.addGroupShape();
        spTree.addGrpSp.assertCalledOnceWith(spTree);
        shapes._shapeFactory.assertCalledOnceWith(shapes, grpSp);
        $assert($is(group_shape, group_shape_));
      } },
    { name: "it_can_add_an_ole_object", args: ["request", "_next_shape_id_prop_", "_recalculate_extents_", "_shape_factory_"], fn: (self: any, request: any, _next_shape_id_prop_: any, _recalculate_extents_: any, _shape_factory_: any): any => {
        let _OleObjectElementCreator_: any, cx: any, cy: any, graphicFrame: any, ole_object_shape_: any, shape: any, shapes: any, x: any, y: any;
        _next_shape_id_prop_.returnValue = 42;
        graphicFrame = element("p:graphicFrame");
        _OleObjectElementCreator_ = classMock(request, "pptx.shapes.shapetree._OleObjectElementCreator");
        _OleObjectElementCreator_.graphicFrame.returnValue = graphicFrame;
        ole_object_shape_ = instanceMock(request, GraphicFrame);
        _shape_factory_.returnValue = ole_object_shape_;
        [x, y, cx, cy] = $list([1, 2, 3, 4]);
        shapes = $call(_BaseGroupShapes, element("p:spTree"), null);
        shape = shapes.addOleObject("worksheet.xlsx", PROG_ID.XLSX, x, y, cx, cy, "test.xlsx", $call(Inches, 0.5), $call(Inches, 0.75));
        _OleObjectElementCreator_.graphicFrame.assertCalledOnceWith(shapes, 42, "worksheet.xlsx", PROG_ID.XLSX, x, y, cx, cy, "test.xlsx", $call(Inches, 0.5), $call(Inches, 0.75));
        $assert($is($get(shapes._spTree, -1), graphicFrame));
        _recalculate_extents_.assertCalledOnceWith(shapes);
        _shape_factory_.assertCalledOnceWith(shapes, graphicFrame);
        $assert($is(shape, ole_object_shape_));
      } },
    { name: "it_can_add_a_picture", args: ["picture_fixture"], fn: (self: any, picture_fixture: any): any => {
        let cx: any, cy: any, image_file: any, image_part_: any, pic: any, picture: any, picture_: any, rId: any, shapes: any, x: any, y: any;
        [shapes, image_file, x, y, cx, cy] = $list($slice(picture_fixture, null, 6, null));
        [image_part_, rId, pic, picture_] = $list($slice(picture_fixture, 6, null, null));
        picture = shapes.addPicture(image_file, x, y, cx, cy);
        shapes.part.getOrAddImagePart.assertCalledOnceWith(image_file);
        shapes._addPicFromImagePart.assertCalledOnceWith(shapes, image_part_, rId, x, y, cx, cy);
        shapes._recalculateExtents.assertCalledOnceWith(shapes);
        shapes._shapeFactory.assertCalledOnceWith(shapes, pic);
        $assert($is(picture, picture_));
      } },
    { name: "it_can_add_a_shape", args: ["shape_fixture"], fn: (self: any, shape_fixture: any): any => {
        let AutoShapeType_: any, autoshape_type_: any, autoshape_type_id: any, cx: any, cy: any, shape: any, shape_: any, shapes: any, sp: any, x: any, y: any;
        [shapes, autoshape_type_id, x, y, cx, cy] = $list($slice(shape_fixture, null, 6, null));
        [AutoShapeType_, autoshape_type_, sp, shape_] = $list($slice(shape_fixture, 6, null, null));
        shape = shapes.addShape(autoshape_type_id, x, y, cx, cy);
        AutoShapeType_.assertCalledOnceWith(autoshape_type_id);
        shapes._addSp.assertCalledOnceWith(shapes, autoshape_type_, x, y, cx, cy);
        shapes._recalculateExtents.assertCalledOnceWith(shapes);
        shapes._shapeFactory.assertCalledOnceWith(shapes, sp);
        $assert($is(shape, shape_));
      } },
    { name: "it_can_add_a_textbox", args: ["textbox_fixture"], fn: (self: any, textbox_fixture: any): any => {
        let cx: any, cy: any, shape: any, shape_: any, shapes: any, sp: any, x: any, y: any;
        [shapes, x, y, cx, cy, sp, shape_] = $list(textbox_fixture);
        shape = shapes.addTextbox(x, y, cx, cy);
        shapes._addTextboxSp.assertCalledOnceWith(shapes, x, y, cx, cy);
        shapes._recalculateExtents.assertCalledOnceWith(shapes);
        shapes._shapeFactory.assertCalledOnceWith(shapes, sp);
        $assert($is(shape, shape_));
      } },
    { name: "it_knows_the_index_of_each_of_its_shapes", args: ["index_fixture"], fn: (self: any, index_fixture: any): any => {
        let expected_value: any, shape_: any, shapes: any;
        [shapes, shape_, expected_value] = $list(index_fixture);
        $assertEqual($index(shapes, shape_), expected_value);
      } },
    { name: "it_raises_on_index_where_shape_not_found", args: ["index_raises_fixture"], fn: (self: any, index_raises_fixture: any): any => {
        let shape_: any, shapes: any;
        [shapes, shape_] = $list(index_raises_fixture);
        $raises(ValueError, () => {
          $index(shapes, shape_);
        });
      } },
    { name: "it_adds_a_chart_graphicFrame_to_help", args: ["add_cht_gr_frm_fixture"], fn: (self: any, add_cht_gr_frm_fixture: any): any => {
        let cx: any, cy: any, expected_xml: any, graphicFrame: any, rId: any, shapes: any, x: any, y: any;
        [shapes, rId, x, y, cx, cy, expected_xml] = $list(add_cht_gr_frm_fixture);
        graphicFrame = shapes._addChartGraphicFrame(rId, x, y, cx, cy);
        $assertEqual(shapes._element.xml, expected_xml);
        $assert($is(graphicFrame, $get(shapes._element.xpath("p:graphicFrame"), 0)));
      } },
    { name: "it_adds_a_cxnSp_to_help", args: ["add_cxnSp_fixture"], fn: (self: any, add_cxnSp_fixture: any): any => {
        let begin_x: any, begin_y: any, connector_type: any, cxnSp: any, end_x: any, end_y: any, expected_xml: any, shapes: any;
        [shapes, connector_type, begin_x, begin_y] = $list($slice(add_cxnSp_fixture, null, 4, null));
        [end_x, end_y, expected_xml] = $list($slice(add_cxnSp_fixture, 4, null, null));
        cxnSp = shapes._addCxnSp(connector_type, begin_x, begin_y, end_x, end_y);
        $assert($is(cxnSp, $get(shapes._element.xpath("p:cxnSp"), 0)));
        $assertEqual(cxnSp.xml, expected_xml);
      } },
    { name: "it_adds_a_pic_element_to_help", args: ["add_pic_fixture"], fn: (self: any, add_pic_fixture: any): any => {
        let cx: any, cy: any, expected_xml: any, image_part_: any, pic: any, rId: any, shapes: any, x: any, y: any;
        [shapes, image_part_, rId, x, y, cx, cy] = $list($slice(add_pic_fixture, null, 7, null));
        expected_xml = $get(add_pic_fixture, 7);
        pic = shapes._addPicFromImagePart(image_part_, rId, x, y, cx, cy);
        image_part_.scale.assertCalledOnceWith(cx, cy);
        $assertEqual(shapes._element.xml, expected_xml);
        $assert($is(pic, $get(shapes._element.xpath("p:pic"), 0)));
      } },
    { name: "it_adds_an_sp_element_to_help", args: ["add_sp_fixture"], fn: (self: any, add_sp_fixture: any): any => {
        let autoshape_type_: any, cx: any, cy: any, expected_xml: any, shapes: any, sp: any, x: any, y: any;
        [shapes, autoshape_type_, x, y, cx, cy, expected_xml] = $list(add_sp_fixture);
        sp = shapes._addSp(autoshape_type_, x, y, cx, cy);
        $assertEqual(shapes._element.xml, expected_xml);
        $assert($is(sp, $get(shapes._element.xpath("p:sp"), 0)));
      } },
    { name: "it_adds_a_textbox_sp_element_to_help", args: ["add_textbox_sp_fixture"], fn: (self: any, add_textbox_sp_fixture: any): any => {
        let cx: any, cy: any, expected_xml: any, shapes: any, sp: any, x: any, y: any;
        [shapes, x, y, cx, cy, expected_xml] = $list(add_textbox_sp_fixture);
        sp = shapes._addTextboxSp(x, y, cx, cy);
        $assertEqual(shapes._element.xml, expected_xml);
        $assert($is(sp, $get(shapes._element.xpath("p:sp"), 0)));
      } },
  ],
};
pytestClass("Describe_BaseGroupShapes", Describe_BaseGroupShapes);

const DescribeGroupShapes: any = {
  fixtures: {
    recalc_fixture: { args: ["grpSp_"], fn: (self: any, grpSp_: any): any => {
        return $call(GroupShapes, grpSp_, null);
      } },
    grpSp_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, CT_GroupShape);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_recalculates_its_extents_to_help", args: ["recalc_fixture"], fn: (self: any, recalc_fixture: any): any => {
        let shapes: any;
        shapes = recalc_fixture;
        shapes._recalculateExtents();
        shapes._grpSp.recalculateExtents.assertCalledOnceWith();
      } },
  ],
};
pytestClass("DescribeGroupShapes", DescribeGroupShapes);

const DescribeBasePlaceholders: any = {
  fixtures: {
    member_fixture: { args: ["request", "shape_elm_"], params: [true, false], fn: (self: any, request: any, shape_elm_: any): any => {
        let is_ph_shape: any;
        is_ph_shape = request.param;
        shape_elm_.hasPhElm = is_ph_shape;
        return [shape_elm_, is_ph_shape];
      } },
    shape_elm_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, BaseShapeElement);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_contains_only_placeholder_shapes", args: ["member_fixture"], fn: (self: any, member_fixture: any): any => {
        let _is_ph_shape: any, is_ph_shape: any, shape_elm_: any;
        [shape_elm_, is_ph_shape] = $list(member_fixture);
        _is_ph_shape = BasePlaceholders._isMemberElm(shape_elm_);
        $assertEqual(_is_ph_shape, is_ph_shape);
      } },
  ],
};
pytestClass("DescribeBasePlaceholders", DescribeBasePlaceholders);

const DescribeNotesSlidePlaceholders: any = {
  fixtures: {
    factory_fixture: { args: ["_NotesSlideShapeFactory_", "placeholder_"], fn: (self: any, _NotesSlideShapeFactory_: any, placeholder_: any): any => {
        let placeholders: any, sp: any;
        placeholders = $call(NotesSlidePlaceholders, null, null);
        sp = element("p:sp");
        return [placeholders, sp, _NotesSlideShapeFactory_, placeholder_];
      } },
    _NotesSlideShapeFactory_: { args: ["request", "placeholder_"], fn: (self: any, request: any, placeholder_: any): any => {
        return $call(functionMock, request, "pptx.shapes.shapetree._NotesSlideShapeFactory", $kw({return_value: placeholder_, autospec: true}));
      } },
    placeholder_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, NotesSlidePlaceholder);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_brokers_access_to_its_shape_factory", args: ["factory_fixture"], fn: (self: any, factory_fixture: any): any => {
        let _NotesSlideShapeFactory_: any, placeholder: any, placeholder_: any, placeholders: any, sp: any;
        [placeholders, sp, _NotesSlideShapeFactory_, placeholder_] = $list(factory_fixture);
        placeholder = placeholders._shapeFactory(sp);
        _NotesSlideShapeFactory_.assertCalledOnceWith(sp, placeholders);
        $assert($is(placeholder, placeholder_));
      } },
  ],
};
pytestClass("DescribeNotesSlidePlaceholders", DescribeNotesSlidePlaceholders);

const DescribeNotesSlideShapes: any = {
  fixtures: {
    basename_fixture: { args: ["request"], params: [[PP_PLACEHOLDER.BODY, "Notes Placeholder"], [PP_PLACEHOLDER.HEADER, "Header Placeholder"]], fn: (self: any, request: any): any => {
        let expected_value: any, notes_slide_shapes: any, ph_type: any;
        [ph_type, expected_value] = $list(request.param);
        notes_slide_shapes = $call(NotesSlideShapes, null, null);
        return [notes_slide_shapes, ph_type, expected_value];
      } },
    factory_fixture: { args: ["_NotesSlideShapeFactory_", "shape_"], fn: (self: any, _NotesSlideShapeFactory_: any, shape_: any): any => {
        let shapes: any, sp: any;
        shapes = $call(NotesSlideShapes, null, null);
        sp = element("p:sp");
        return [shapes, sp, _NotesSlideShapeFactory_, shape_];
      } },
    _NotesSlideShapeFactory_: { args: ["request", "shape_"], fn: (self: any, request: any, shape_: any): any => {
        return $call(functionMock, request, "pptx.shapes.shapetree._NotesSlideShapeFactory", $kw({return_value: shape_, autospec: true}));
      } },
    shape_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Shape);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_notes_slide_placeholder_basenames", args: ["basename_fixture"], fn: (self: any, basename_fixture: any): any => {
        let expected_value: any, notes_slide_shapes: any, ph_type: any;
        [notes_slide_shapes, ph_type, expected_value] = $list(basename_fixture);
        $assertEqual(notes_slide_shapes.phBasename(ph_type), expected_value);
      } },
    { name: "it_brokers_access_to_its_shape_factory", args: ["factory_fixture"], fn: (self: any, factory_fixture: any): any => {
        let _NotesSlideShapeFactory_: any, shape: any, shape_: any, shapes: any, sp: any;
        [shapes, sp, _NotesSlideShapeFactory_, shape_] = $list(factory_fixture);
        shape = shapes._shapeFactory(sp);
        _NotesSlideShapeFactory_.assertCalledOnceWith(sp, shapes);
        $assert($is(shape, shape_));
      } },
  ],
};
pytestClass("DescribeNotesSlideShapes", DescribeNotesSlideShapes);

const DescribeSlidePlaceholders: any = {
  fixtures: {
    getitem_fixture: { args: ["request", "SlideShapeFactory_", "placeholder_"], params: [["p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph{type=pic,idx=1}", 1, 0], ["p:spTree/p:pic/p:nvPicPr/p:nvPr/p:ph{type=pic,idx=1}", 1, 0], ["p:spTree/(p:sp,p:sp/p:nvSpPr/p:nvPr/p:ph{type=title})", 0, 1], ["p:spTree/(p:sp,p:pic/p:nvPicPr/p:nvPr/p:ph{type=pic,idx=1})", 1, 1], ["p:spTree/(p:sp/p:nvSpPr/p:nvPr/p:ph{type=title},p:sp/p:nvSpPr/p:nvPr/p:ph{type=pic,idx=3})", 3, 1]], fn: (self: any, request: any, SlideShapeFactory_: any, placeholder_: any): any => {
        let idx: any, offset: any, placeholders: any, shape_elm: any, spTree: any, spTree_cxml: any;
        [spTree_cxml, idx, offset] = $list(request.param);
        spTree = element(spTree_cxml);
        placeholders = $call(SlidePlaceholders, spTree, null);
        shape_elm = $get(spTree, offset);
        SlideShapeFactory_.returnValue = placeholder_;
        return [placeholders, idx, SlideShapeFactory_, shape_elm, placeholder_];
      } },
    iter_fixture: { args: ["request", "SlideShapeFactory_", "placeholder_"], params: [["p:spTree/(p:sp/p:nvSpPr/p:nvPr/p:ph{type=body,idx=1},p:sp/p:nvSpPr/p:nvPr/p:ph{type=title},p:pic/p:nvPicPr/p:nvPr/p:ph{type=pic,idx=3})", [1, 0, 2]]], fn: (self: any, request: any, SlideShapeFactory_: any, placeholder_: any): any => {
        let calls: any, placeholders: any, sequence: any, spTree: any, spTree_cxml: any, values: any;
        [spTree_cxml, sequence] = $list(request.param);
        spTree = element(spTree_cxml);
        placeholders = $call(SlidePlaceholders, spTree, null);
        SlideShapeFactory_.returnValue = placeholder_;
        calls = $list(sequence).map((i: any) => call($get(spTree, i), placeholders));
        values = $mul([placeholder_], $len(sequence));
        return [placeholders, SlideShapeFactory_, calls, values];
      } },
    len_fixture: { args: ["request"], params: [["p:spTree", 0], ["p:spTree/(p:sp,p:pic,p:sp)", 0], ["p:spTree/(p:sp,p:sp/p:nvSpPr/p:nvPr/p:ph{type=title},p:pic)", 1], ["p:spTree/(p:sp/p:nvSpPr/p:nvPr/p:ph{type=body,idx=1},p:sp/p:nvSpPr/p:nvPr/p:ph{type=title},p:pic/p:nvPicPr/p:nvPr/p:ph{type=pic,idx=3})", 3]], fn: (self: any, request: any): any => {
        let length: any, placeholders: any, spTree_cxml: any;
        [spTree_cxml, length] = $list(request.param);
        placeholders = $call(SlidePlaceholders, element(spTree_cxml), null);
        return [placeholders, length];
      } },
    placeholder_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _BaseSlidePlaceholder);
      } },
    SlideShapeFactory_: { args: ["request", "placeholder_"], fn: (self: any, request: any, placeholder_: any): any => {
        return $call(functionMock, request, "pptx.shapes.shapetree.SlideShapeFactory", $kw({return_value: placeholder_}));
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_get_a_placeholder_by_idx", args: ["getitem_fixture"], fn: (self: any, getitem_fixture: any): any => {
        let SlideShapeFactory_: any, idx: any, placeholder: any, placeholder_: any, placeholders: any, shape_elm: any;
        [placeholders, idx, SlideShapeFactory_] = $list($slice(getitem_fixture, null, 3, null));
        [shape_elm, placeholder_] = $list($slice(getitem_fixture, 3, null, null));
        placeholder = $get(placeholders, idx);
        SlideShapeFactory_.assertCalledOnceWith(shape_elm, placeholders);
        $assert($is(placeholder, placeholder_));
      } },
    { name: "it_can_iterate_over_its_placeholders", args: ["iter_fixture"], fn: (self: any, iter_fixture: any): any => {
        let SlideShapeFactory_: any, expected_calls: any, expected_values: any, placeholders: any, ps: any;
        [placeholders, SlideShapeFactory_] = $list($slice(iter_fixture, null, 2, null));
        [expected_calls, expected_values] = $list($slice(iter_fixture, 2, null, null));
        ps = $list(placeholders).map((p: any) => p);
        $assertEqual(SlideShapeFactory_.callArgsList, expected_calls);
        $assertEqual(ps, expected_values);
      } },
    { name: "it_knows_how_many_placeholders_it_contains", args: ["len_fixture"], fn: (self: any, len_fixture: any): any => {
        let expected_value: any, placeholders: any;
        [placeholders, expected_value] = $list(len_fixture);
        $assertEqual($len(placeholders), expected_value);
      } },
  ],
};
pytestClass("DescribeSlidePlaceholders", DescribeSlidePlaceholders);

const Describe_SlidePlaceholderFactory: any = {
  fixtures: {
    factory_fixture: { args: ["request", "SlidePlaceholder_", "ChartPlaceholder_", "PicturePlaceholder_", "TablePlaceholder_", "PlaceholderGraphicFrame_", "PlaceholderPicture_", "placeholder_"], params: [["p:sp/p:nvSpPr/p:nvPr/p:ph{type=title}", "sp"], ["p:sp/p:nvSpPr/p:nvPr/p:ph{type=pic,idx=1}", "pph"], ["p:sp/p:nvSpPr/p:nvPr/p:ph{type=clipArt,idx=1}", "pph"], ["p:sp/p:nvSpPr/p:nvPr/p:ph{type=tbl,idx=1}", "tph"], ["p:sp/p:nvSpPr/p:nvPr/p:ph{type=chart,idx=10}", "cph"], ["p:pic/p:nvPicPr/p:nvPr/p:ph{type=pic,idx=1}", "php"], ["p:graphicFrame/p:nvSpPr/p:nvPr/p:ph{type=tbl,idx=2}", "phgf"], ["p:graphicFrame/p:nvSpPr/p:nvPr/p:ph{type=chart,idx=2}", "phgf"], ["p:graphicFrame/p:nvSpPr/p:nvPr/p:ph{type=dgm,idx=2}", "phgf"]], fn: (self: any, request: any, SlidePlaceholder_: any, ChartPlaceholder_: any, PicturePlaceholder_: any, TablePlaceholder_: any, PlaceholderGraphicFrame_: any, PlaceholderPicture_: any, placeholder_: any): any => {
        let Constructor_: any, constructor_key: any, shape_: any, shape_cxml: any, shape_elm: any;
        [shape_cxml, constructor_key] = $list(request.param);
        shape_elm = element(shape_cxml);
        [Constructor_, shape_] = $list($get($d({sp: [SlidePlaceholder_, placeholder_], cph: [ChartPlaceholder_, placeholder_], pph: [PicturePlaceholder_, placeholder_], tph: [TablePlaceholder_, placeholder_], phgf: [PlaceholderGraphicFrame_, placeholder_], php: [PlaceholderPicture_, placeholder_]}), constructor_key));
        return [shape_elm, 42, Constructor_, shape_];
      } },
    ChartPlaceholder_: { args: ["request", "placeholder_"], fn: (self: any, request: any, placeholder_: any): any => {
        return $call(classMock, request, "pptx.shapes.shapetree.ChartPlaceholder", $kw({return_value: placeholder_}));
      } },
    PicturePlaceholder_: { args: ["request", "placeholder_"], fn: (self: any, request: any, placeholder_: any): any => {
        return $call(classMock, request, "pptx.shapes.shapetree.PicturePlaceholder", $kw({return_value: placeholder_}));
      } },
    PlaceholderGraphicFrame_: { args: ["request", "placeholder_"], fn: (self: any, request: any, placeholder_: any): any => {
        return $call(classMock, request, "pptx.shapes.shapetree.PlaceholderGraphicFrame", $kw({return_value: placeholder_}));
      } },
    PlaceholderPicture_: { args: ["request", "placeholder_"], fn: (self: any, request: any, placeholder_: any): any => {
        return $call(classMock, request, "pptx.shapes.shapetree.PlaceholderPicture", $kw({return_value: placeholder_}));
      } },
    placeholder_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _BaseSlidePlaceholder);
      } },
    SlidePlaceholder_: { args: ["request", "placeholder_"], fn: (self: any, request: any, placeholder_: any): any => {
        return $call(classMock, request, "pptx.shapes.shapetree.SlidePlaceholder", $kw({return_value: placeholder_}));
      } },
    TablePlaceholder_: { args: ["request", "placeholder_"], fn: (self: any, request: any, placeholder_: any): any => {
        return $call(classMock, request, "pptx.shapes.shapetree.TablePlaceholder", $kw({return_value: placeholder_}));
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_constructs_the_right_type_of_placeholder", args: ["factory_fixture"], fn: (self: any, factory_fixture: any): any => {
        let Constructor_: any, element: any, parent_: any, placeholder: any, placeholder_: any;
        [element, parent_, Constructor_, placeholder_] = $list(factory_fixture);
        placeholder = $call(_SlidePlaceholderFactory, element, parent_);
        Constructor_.assertCalledOnceWith(element, parent_);
        $assert($is(placeholder, placeholder_));
      } },
  ],
};
pytestClass("Describe_SlidePlaceholderFactory", Describe_SlidePlaceholderFactory);

const DescribeSlideShapeFactory: any = {
  fixtures: {
    factory_fixture: { args: ["request", "_SlidePlaceholderFactory_", "placeholder_", "BaseShapeFactory_", "base_shape_"], params: [["p:sp/p:nvSpPr/p:nvPr/p:ph{type=title}", "spf"], ["p:pic/p:nvSpPr/p:nvPr/p:ph{type=pic,idx=1}", "spf"], ["p:sp", "bsf"], ["p:pic", "bsf"]], fn: (self: any, request: any, _SlidePlaceholderFactory_: any, placeholder_: any, BaseShapeFactory_: any, base_shape_: any): any => {
        let Constructor_: any, shape_: any, shape_cxml: any, shape_elm: any, shape_type: any;
        [shape_cxml, shape_type] = $list(request.param);
        shape_elm = element(shape_cxml);
        [Constructor_, shape_] = $list($get($d({spf: [_SlidePlaceholderFactory_, placeholder_], bsf: [BaseShapeFactory_, base_shape_]}), shape_type));
        return [shape_elm, 42, Constructor_, shape_];
      } },
    BaseShapeFactory_: { args: ["request", "base_shape_"], fn: (self: any, request: any, base_shape_: any): any => {
        return $call(functionMock, request, "pptx.shapes.shapetree.BaseShapeFactory", $kw({return_value: base_shape_}));
      } },
    base_shape_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, BaseShape);
      } },
    placeholder_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _BaseSlidePlaceholder);
      } },
    _SlidePlaceholderFactory_: { args: ["request", "placeholder_"], fn: (self: any, request: any, placeholder_: any): any => {
        return $call(functionMock, request, "pptx.shapes.shapetree._SlidePlaceholderFactory", $kw({return_value: placeholder_}));
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_constructs_the_right_type_of_shape", args: ["factory_fixture"], fn: (self: any, factory_fixture: any): any => {
        let Constructor_: any, element: any, parent_: any, shape: any, shape_: any;
        [element, parent_, Constructor_, shape_] = $list(factory_fixture);
        shape = $call(SlideShapeFactory, element, parent_);
        Constructor_.assertCalledOnceWith(element, parent_);
        $assert($is(shape, shape_));
      } },
  ],
};
pytestClass("DescribeSlideShapeFactory", DescribeSlideShapeFactory);

const DescribeSlideShapes: any = {
  fixtures: {
    add_timing_fixture: { args: ["request"], params: [[0, 1], [1, 2], [3, 1]], fn: (self: any, request: any): any => {
        let after_idx: any, before_idx: any, expected_xml: any, pic: any, shapes: any, sld: any, snippets: any, spTree: any;
        [before_idx, after_idx] = $list(request.param);
        snippets = snippetSeq("timing");
        sld = parseXml($get(snippets, before_idx));
        spTree = $get(sld.xpath(".//p:spTree"), 0);
        shapes = $call(SlideShapes, spTree, null);
        pic = element("p:pic/p:nvPicPr/p:cNvPr{id=42}");
        expected_xml = $get(snippets, after_idx);
        return [shapes, pic, sld, expected_xml];
      } },
    clone_fixture: { args: ["slide_layout_", "clone_placeholder_", "placeholder_"], fn: (self: any, slide_layout_: any, clone_placeholder_: any, placeholder_: any): any => {
        let calls: any, shapes: any;
        shapes = $call(SlideShapes, null, null);
        calls = [call(shapes, placeholder_)];
        slide_layout_.iterCloneablePlaceholders.returnValue = $iter([placeholder_]);
        return [shapes, slide_layout_, calls];
      } },
    factory_fixture: { args: ["SlideShapeFactory_", "shape_"], fn: (self: any, SlideShapeFactory_: any, shape_: any): any => {
        let shapes: any, sp: any;
        shapes = $call(SlideShapes, null, null);
        sp = element("p:sp");
        return [shapes, sp, SlideShapeFactory_, shape_];
      } },
    movie_fixture: { args: ["_MoviePicElementCreator_", "_add_video_timing_", "_shape_factory_", "movie_", "_next_shape_id_prop_"], fn: (self: any, _MoviePicElementCreator_: any, _add_video_timing_: any, _shape_factory_: any, movie_: any, _next_shape_id_prop_: any): any => {
        let cx: any, cy: any, mime_type: any, movie_file: any, movie_pic: any, poster_frame_image: any, shape_id_: any, shapes: any, x: any, y: any;
        shapes = $call(SlideShapes, element("p:spTree"), null);
        [movie_file, x, y, cx, cy] = $list(["foobar.mp4", 1, 2, 3, 4]);
        [poster_frame_image, mime_type] = $list(["foobar.png", "video/mp4"]);
        movie_pic = element("p:pic");
        _MoviePicElementCreator_.newMoviePic.returnValue = movie_pic;
        _shape_factory_.returnValue = movie_;
        shape_id_ = _next_shape_id_prop_.returnValue;
        return [shapes, movie_file, x, y, cx, cy, poster_frame_image, mime_type, shape_id_, _MoviePicElementCreator_, movie_pic, _add_video_timing_, _shape_factory_, movie_];
      } },
    table_fixture: { args: ["table_", "_shape_factory_"], fn: (self: any, table_: any, _shape_factory_: any): any => {
        let cols: any, cx: any, cy: any, expected_xml: any, rows: any, shapes: any, x: any, y: any;
        shapes = $call(SlideShapes, element("p:spTree"), null);
        [rows, cols, x, y, cx, cy] = $list([1, 2, 10, 11, 12, 13]);
        _shape_factory_.returnValue = table_;
        expected_xml = "<p:spTree xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">\n  <p:graphicFrame xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">\n    <p:nvGraphicFramePr>\n      <p:cNvPr id=\"1\" name=\"Table 0\"/>\n      <p:cNvGraphicFramePr>\n        <a:graphicFrameLocks noGrp=\"1\"/>\n      </p:cNvGraphicFramePr>\n      <p:nvPr/>\n    </p:nvGraphicFramePr>\n    <p:xfrm>\n      <a:off x=\"10\" y=\"11\"/>\n      <a:ext cx=\"12\" cy=\"13\"/>\n    </p:xfrm>\n    <a:graphic>\n      <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/table\">\n        <a:tbl>\n          <a:tblPr firstRow=\"1\" bandRow=\"1\">\n            <a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>\n          </a:tblPr>\n          <a:tblGrid>\n            <a:gridCol w=\"6\"/>\n            <a:gridCol w=\"6\"/>\n          </a:tblGrid>\n          <a:tr h=\"13\">\n            <a:tc>\n              <a:txBody>\n                <a:bodyPr/>\n                <a:lstStyle/>\n                <a:p/>\n              </a:txBody>\n              <a:tcPr/>\n            </a:tc>\n            <a:tc>\n              <a:txBody>\n                <a:bodyPr/>\n                <a:lstStyle/>\n                <a:p/>\n              </a:txBody>\n              <a:tcPr/>\n            </a:tc>\n          </a:tr>\n        </a:tbl>\n      </a:graphicData>\n    </a:graphic>\n  </p:graphicFrame>\n</p:spTree>";
        return [shapes, rows, cols, x, y, cx, cy, table_, expected_xml];
      } },
    title_fixture: { args: ["request", "_shape_factory_", "shape_"], params: [["p:spTree/(p:sp,p:sp/p:nvSpPr/p:nvPr/p:ph{idx=0})", true], ["p:spTree/(p:sp,p:sp)", false]], fn: (self: any, request: any, _shape_factory_: any, shape_: any): any => {
        let _shape_: any, calls: any, found: any, shapes: any, spTree: any, spTree_cxml: any;
        [spTree_cxml, found] = $list(request.param);
        spTree = element(spTree_cxml);
        shapes = $call(SlideShapes, spTree, null);
        calls = ($truthy(found) ? [call(shapes, $get(spTree.xpath("p:sp"), 1))] : []);
        _shape_ = ($truthy(found) ? shape_ : null);
        return [shapes, _shape_factory_, calls, _shape_];
      } },
    _add_video_timing_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, SlideShapes, "_add_video_timing", $kw({autospec: true}));
      } },
    clone_placeholder_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, SlideShapes, "clone_placeholder", $kw({autospec: true}));
      } },
    movie_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Movie);
      } },
    _MoviePicElementCreator_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(classMock, request, "pptx.shapes.shapetree._MoviePicElementCreator", $kw({autospec: true}));
      } },
    _next_shape_id_prop_: { args: ["request", "shape_id_"], fn: (self: any, request: any, shape_id_: any): any => {
        return $call(propertyMock, request, SlideShapes, "_next_shape_id", $kw({return_value: shape_id_}));
      } },
    placeholder_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Shape);
      } },
    shape_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Shape);
      } },
    shape_id_: { args: [], fn: (self: any): any => {
        return 42;
      } },
    _shape_factory_: { args: ["request", "shape_"], fn: (self: any, request: any, shape_: any): any => {
        return $call(methodMock, request, SlideShapes, "_shape_factory", $kw({return_value: shape_, autospec: true}));
      } },
    slide_layout_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlideLayout);
      } },
    SlideShapeFactory_: { args: ["request", "shape_"], fn: (self: any, request: any, shape_: any): any => {
        return $call(functionMock, request, "pptx.shapes.shapetree.SlideShapeFactory", $kw({return_value: shape_, autospec: true}));
      } },
    table_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Table);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_access_to_its_shape_factory", args: ["factory_fixture"], fn: (self: any, factory_fixture: any): any => {
        let SlideShapeFactory_: any, shape: any, shape_: any, shapes: any, sp: any;
        [shapes, sp, SlideShapeFactory_, shape_] = $list(factory_fixture);
        shape = shapes._shapeFactory(sp);
        SlideShapeFactory_.assertCalledOnceWith(sp, shapes);
        $assert($is(shape, shape_));
      } },
    { name: "it_can_find_the_title_placeholder", args: ["title_fixture"], fn: (self: any, title_fixture: any): any => {
        let _shape_factory_: any, calls: any, shape_: any, shapes: any, title_placeholder: any;
        [shapes, _shape_factory_, calls, shape_] = $list(title_fixture);
        title_placeholder = shapes.title;
        $assertEqual(_shape_factory_.callArgsList, calls);
        $assert($is(title_placeholder, shape_));
      } },
    { name: "it_can_add_a_movie", args: ["movie_fixture"], fn: (self: any, movie_fixture: any): any => {
        let _MoviePicElementCreator_: any, _add_video_timing_: any, _shape_factory_: any, cx: any, cy: any, mime_type: any, movie: any, movie_: any, movie_file: any, movie_pic: any, poster_frame_image: any, shape_id_: any, shapes: any, x: any, y: any;
        [shapes, movie_file, x, y, cx, cy] = $list($slice(movie_fixture, null, 6, null));
        [poster_frame_image, mime_type, shape_id_] = $list($slice(movie_fixture, 6, 9, null));
        [_MoviePicElementCreator_, movie_pic] = $list($slice(movie_fixture, 9, 11, null));
        [_add_video_timing_, _shape_factory_, movie_] = $list($slice(movie_fixture, 11, null, null));
        movie = shapes.addMovie(movie_file, x, y, cx, cy, poster_frame_image, mime_type);
        _MoviePicElementCreator_.newMoviePic.assertCalledOnceWith(shapes, shape_id_, movie_file, x, y, cx, cy, poster_frame_image, mime_type);
        $assert($is($get(shapes._spTree, -1), movie_pic));
        _add_video_timing_.assertCalledOnceWith(shapes, movie_pic);
        _shape_factory_.assertCalledOnceWith(shapes, movie_pic);
        $assert($is(movie, movie_));
      } },
    { name: "it_can_add_a_table", args: ["table_fixture"], fn: (self: any, table_fixture: any): any => {
        let cols: any, cx: any, cy: any, expected_xml: any, graphicFrame: any, rows: any, shapes: any, table: any, table_: any, x: any, y: any;
        [shapes, rows, cols, x, y, cx, cy, table_, expected_xml] = $list(table_fixture);
        table = shapes.addTable(rows, cols, x, y, cx, cy);
        graphicFrame = $get(shapes._element.xpath("p:graphicFrame"), 0);
        shapes._shapeFactory.assertCalledOnceWith(shapes, graphicFrame);
        $assert($is(table, table_));
        $assertEqual(shapes._element.xml, expected_xml);
      } },
    { name: "it_can_clone_placeholder_shapes_from_a_layout", args: ["clone_fixture"], fn: (self: any, clone_fixture: any): any => {
        let calls: any, shapes: any, slide_layout_: any;
        [shapes, slide_layout_, calls] = $list(clone_fixture);
        shapes.cloneLayoutPlaceholders(slide_layout_);
        $assertEqual(shapes.clonePlaceholder.callArgsList, calls);
      } },
    { name: "it_adds_a_video_timing_to_help", args: ["add_timing_fixture"], fn: (self: any, add_timing_fixture: any): any => {
        let expected_xml: any, pic: any, shapes: any, sld: any;
        [shapes, pic, sld, expected_xml] = $list(add_timing_fixture);
        shapes._addVideoTiming(pic);
        $assertEqual(sld.xml, expected_xml);
      } },
  ],
};
pytestClass("DescribeSlideShapes", DescribeSlideShapes);

const DescribeLayoutShapes: any = {
  fixtures: {
    factory_fixture: { args: ["_LayoutShapeFactory_", "placeholder_"], fn: (self: any, _LayoutShapeFactory_: any, placeholder_: any): any => {
        let shapes: any, sp: any;
        shapes = $call(LayoutShapes, null, null);
        sp = element("p:sp");
        return [shapes, sp, _LayoutShapeFactory_, placeholder_];
      } },
    _LayoutShapeFactory_: { args: ["request", "placeholder_"], fn: (self: any, request: any, placeholder_: any): any => {
        return $call(functionMock, request, "pptx.shapes.shapetree._LayoutShapeFactory", $kw({return_value: placeholder_, autospec: true}));
      } },
    placeholder_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, LayoutPlaceholder);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_access_to_its_shape_factory", args: ["factory_fixture"], fn: (self: any, factory_fixture: any): any => {
        let _LayoutShapeFactory_: any, placeholder: any, placeholder_: any, shapes: any, sp: any;
        [shapes, sp, _LayoutShapeFactory_, placeholder_] = $list(factory_fixture);
        placeholder = shapes._shapeFactory(sp);
        _LayoutShapeFactory_.assertCalledOnceWith(sp, shapes);
        $assert($is(placeholder, placeholder_));
      } },
  ],
};
pytestClass("DescribeLayoutShapes", DescribeLayoutShapes);

const Describe_LayoutShapeFactory: any = {
  fixtures: {
    factory_fixture: { args: ["request", "ph_bldr", "parent_", "_LayoutPlaceholder_", "layout_placeholder_", "BaseShapeFactory_", "base_shape_"], params: ["ph", "sp", "pic"], fn: (self: any, request: any, ph_bldr: any, parent_: any, _LayoutPlaceholder_: any, layout_placeholder_: any, BaseShapeFactory_: any, base_shape_: any): any => {
        let ShapeConstructor_: any, shape_bldr: any, shape_elm: any, shape_mock: any;
        [shape_bldr, ShapeConstructor_, shape_mock] = $list($get($d({ph: [ph_bldr, _LayoutPlaceholder_, layout_placeholder_], sp: [anSp(), BaseShapeFactory_, base_shape_], pic: [aPic(), BaseShapeFactory_, base_shape_]}), request.param));
        shape_elm = shape_bldr.withNsdecls().element;
        return [shape_elm, parent_, ShapeConstructor_, shape_mock];
      } },
    BaseShapeFactory_: { args: ["request", "base_shape_"], fn: (self: any, request: any, base_shape_: any): any => {
        return $call(functionMock, request, "pptx.shapes.shapetree.BaseShapeFactory", $kw({return_value: base_shape_}));
      } },
    base_shape_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, BaseShape);
      } },
    _LayoutPlaceholder_: { args: ["request", "layout_placeholder_"], fn: (self: any, request: any, layout_placeholder_: any): any => {
        return $call(classMock, request, "pptx.shapes.shapetree.LayoutPlaceholder", $kw({return_value: layout_placeholder_}));
      } },
    layout_placeholder_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, LayoutPlaceholder);
      } },
    parent_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, BasePlaceholders);
      } },
    ph_bldr: { args: [], fn: (self: any): any => {
        return anSp().withChild(anNvSpPr().withChild(anNvPr().withChild(aPh().withIdx(1))));
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_constructs_the_right_shape_for_an_element", args: ["factory_fixture"], fn: (self: any, factory_fixture: any): any => {
        let ShapeConstructor_: any, parent_: any, shape: any, shape_: any, shape_elm: any;
        [shape_elm, parent_, ShapeConstructor_, shape_] = $list(factory_fixture);
        shape = $call(_LayoutShapeFactory, shape_elm, parent_);
        ShapeConstructor_.assertCalledOnceWith(shape_elm, parent_);
        $assert($is(shape, shape_));
      } },
  ],
};
pytestClass("Describe_LayoutShapeFactory", Describe_LayoutShapeFactory);

const DescribeLayoutPlaceholders: any = {
  fixtures: {
    default_fixture: { args: ["_iter_"], fn: (self: any, _iter_: any): any => {
        let default$: any, placeholders: any;
        placeholders = $call(LayoutPlaceholders, null, null);
        default$ = "barfoo";
        return [placeholders, default$];
      } },
    factory_fixture: { args: ["_LayoutShapeFactory_", "placeholder_"], fn: (self: any, _LayoutShapeFactory_: any, placeholder_: any): any => {
        let placeholders: any, sp: any;
        placeholders = $call(LayoutPlaceholders, null, null);
        sp = element("p:sp");
        return [placeholders, sp, _LayoutShapeFactory_, placeholder_];
      } },
    get_fixture: { args: ["request", "_iter_", "placeholder_", "placeholder_2_"], params: [0, 1], fn: (self: any, request: any, _iter_: any, placeholder_: any, placeholder_2_: any): any => {
        let _placeholder_: any, idx: any, layout_placeholders: any;
        idx = request.param;
        layout_placeholders = $call(LayoutPlaceholders, null, null);
        _placeholder_ = $get([placeholder_, placeholder_2_], idx);
        [placeholder_.element.phIdx, placeholder_2_.element.phIdx] = $list([0, 1]);
        return [layout_placeholders, idx, _placeholder_];
      } },
    _iter_: { args: ["request", "placeholder_", "placeholder_2_"], fn: (self: any, request: any, placeholder_: any, placeholder_2_: any): any => {
        return $call(methodMock, request, LayoutPlaceholders, "__iter__", $kw({return_value: $iter([placeholder_, placeholder_2_])}));
      } },
    _LayoutShapeFactory_: { args: ["request", "placeholder_"], fn: (self: any, request: any, placeholder_: any): any => {
        return $call(functionMock, request, "pptx.shapes.shapetree._LayoutShapeFactory", $kw({return_value: placeholder_, autospec: true}));
      } },
    placeholder_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, LayoutPlaceholder);
      } },
    placeholder_2_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, LayoutPlaceholder);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_access_to_its_shape_factory", args: ["factory_fixture"], fn: (self: any, factory_fixture: any): any => {
        let _LayoutShapeFactory_: any, placeholder: any, placeholder_: any, placeholders: any, sp: any;
        [placeholders, sp, _LayoutShapeFactory_, placeholder_] = $list(factory_fixture);
        placeholder = placeholders._shapeFactory(sp);
        _LayoutShapeFactory_.assertCalledOnceWith(sp, placeholders);
        $assert($is(placeholder, placeholder_));
      } },
    { name: "it_can_find_a_placeholder_by_idx_value", args: ["get_fixture"], fn: (self: any, get_fixture: any): any => {
        let idx: any, placeholder_: any, placeholders: any;
        [placeholders, idx, placeholder_] = $list(get_fixture);
        $assert($is($dictget(placeholders, idx), placeholder_));
      } },
    { name: "it_returns_default_on_ph_idx_not_found", args: ["default_fixture"], fn: (self: any, default_fixture: any): any => {
        let default$: any, placeholders: any;
        [placeholders, default$] = $list(default_fixture);
        $assert($is($dictget(placeholders, 42, default$), default$));
      } },
  ],
};
pytestClass("DescribeLayoutPlaceholders", DescribeLayoutPlaceholders);

const Describe_MasterShapeFactory: any = {
  fixtures: {
    factory_fixture: { args: ["request", "ph_bldr", "slide_master_", "_MasterPlaceholder_", "master_placeholder_", "BaseShapeFactory_", "base_shape_"], params: ["ph", "sp", "pic"], fn: (self: any, request: any, ph_bldr: any, slide_master_: any, _MasterPlaceholder_: any, master_placeholder_: any, BaseShapeFactory_: any, base_shape_: any): any => {
        let ShapeConstructor_: any, shape_bldr: any, shape_elm: any, shape_mock: any;
        [shape_bldr, ShapeConstructor_, shape_mock] = $list($get($d({ph: [ph_bldr, _MasterPlaceholder_, master_placeholder_], sp: [anSp(), BaseShapeFactory_, base_shape_], pic: [aPic(), BaseShapeFactory_, base_shape_]}), request.param));
        shape_elm = shape_bldr.withNsdecls().element;
        return [shape_elm, slide_master_, ShapeConstructor_, shape_mock];
      } },
    BaseShapeFactory_: { args: ["request", "base_shape_"], fn: (self: any, request: any, base_shape_: any): any => {
        return $call(functionMock, request, "pptx.shapes.shapetree.BaseShapeFactory", $kw({return_value: base_shape_}));
      } },
    base_shape_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, BaseShape);
      } },
    _MasterPlaceholder_: { args: ["request", "master_placeholder_"], fn: (self: any, request: any, master_placeholder_: any): any => {
        return $call(classMock, request, "pptx.shapes.shapetree.MasterPlaceholder", $kw({return_value: master_placeholder_}));
      } },
    master_placeholder_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, MasterPlaceholder);
      } },
    ph_bldr: { args: [], fn: (self: any): any => {
        return anSp().withChild(anNvSpPr().withChild(anNvPr().withChild(aPh().withIdx(1))));
      } },
    slide_master_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlideMaster);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_constructs_a_master_placeholder_for_a_shape_element", args: ["factory_fixture"], fn: (self: any, factory_fixture: any): any => {
        let ShapeConstructor_: any, parent_: any, shape: any, shape_: any, shape_elm: any;
        [shape_elm, parent_, ShapeConstructor_, shape_] = $list(factory_fixture);
        shape = $call(_MasterShapeFactory, shape_elm, parent_);
        ShapeConstructor_.assertCalledOnceWith(shape_elm, parent_);
        $assert($is(shape, shape_));
      } },
  ],
};
pytestClass("Describe_MasterShapeFactory", Describe_MasterShapeFactory);

const DescribeMasterShapes: any = {
  fixtures: {
    factory_fixture: { args: ["_MasterShapeFactory_", "placeholder_"], fn: (self: any, _MasterShapeFactory_: any, placeholder_: any): any => {
        let shapes: any, sp: any;
        shapes = $call(MasterShapes, null, null);
        sp = element("p:sp");
        return [shapes, sp, _MasterShapeFactory_, placeholder_];
      } },
    _MasterShapeFactory_: { args: ["request", "placeholder_"], fn: (self: any, request: any, placeholder_: any): any => {
        return $call(functionMock, request, "pptx.shapes.shapetree._MasterShapeFactory", $kw({return_value: placeholder_, autospec: true}));
      } },
    placeholder_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, MasterPlaceholder);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_access_to_its_shape_factory", args: ["factory_fixture"], fn: (self: any, factory_fixture: any): any => {
        let _MasterShapeFactory_: any, placeholder: any, placeholder_: any, shapes: any, sp: any;
        [shapes, sp, _MasterShapeFactory_, placeholder_] = $list(factory_fixture);
        placeholder = shapes._shapeFactory(sp);
        _MasterShapeFactory_.assertCalledOnceWith(sp, shapes);
        $assert($is(placeholder, placeholder_));
      } },
  ],
};
pytestClass("DescribeMasterShapes", DescribeMasterShapes);

const DescribeMasterPlaceholders: any = {
  fixtures: {
    default_fixture: { args: ["_iter_"], fn: (self: any, _iter_: any): any => {
        let default$: any, placeholders: any;
        placeholders = $call(MasterPlaceholders, null, null);
        default$ = "barfoo";
        return [placeholders, default$];
      } },
    factory_fixture: { args: ["_MasterShapeFactory_", "placeholder_"], fn: (self: any, _MasterShapeFactory_: any, placeholder_: any): any => {
        let placeholders: any, sp: any;
        placeholders = $call(MasterPlaceholders, null, null);
        sp = element("p:sp");
        return [placeholders, sp, _MasterShapeFactory_, placeholder_];
      } },
    get_fixture: { args: ["request", "_iter_", "placeholder_", "placeholder_2_"], params: ["title", "body"], fn: (self: any, request: any, _iter_: any, placeholder_: any, placeholder_2_: any): any => {
        let _placeholder_: any, ph_type: any, placeholders: any;
        ph_type = request.param;
        placeholders = $call(MasterPlaceholders, null, null);
        _placeholder_ = $get($d({title: placeholder_, body: placeholder_2_}), ph_type);
        return [placeholders, ph_type, _placeholder_];
      } },
    _iter_: { args: ["request", "placeholder_", "placeholder_2_"], fn: (self: any, request: any, placeholder_: any, placeholder_2_: any): any => {
        return $call(methodMock, request, MasterPlaceholders, "__iter__", $kw({return_value: $iter([placeholder_, placeholder_2_])}));
      } },
    _MasterShapeFactory_: { args: ["request", "placeholder_"], fn: (self: any, request: any, placeholder_: any): any => {
        return $call(functionMock, request, "pptx.shapes.shapetree._MasterShapeFactory", $kw({return_value: placeholder_, autospec: true}));
      } },
    placeholder_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(instanceMock, request, MasterPlaceholder, $kw({ph_type: "title"}));
      } },
    placeholder_2_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(instanceMock, request, MasterPlaceholder, $kw({ph_type: "body"}));
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_access_to_its_shape_factory", args: ["factory_fixture"], fn: (self: any, factory_fixture: any): any => {
        let _MasterShapeFactory_: any, placeholder: any, placeholder_: any, placeholders: any, sp: any;
        [placeholders, sp, _MasterShapeFactory_, placeholder_] = $list(factory_fixture);
        placeholder = placeholders._shapeFactory(sp);
        _MasterShapeFactory_.assertCalledOnceWith(sp, placeholders);
        $assert($is(placeholder, placeholder_));
      } },
    { name: "it_can_find_a_placeholder_by_type", args: ["get_fixture"], fn: (self: any, get_fixture: any): any => {
        let ph_type: any, placeholder_: any, placeholders: any;
        [placeholders, ph_type, placeholder_] = $list(get_fixture);
        $assert($is($dictget(placeholders, ph_type), placeholder_));
      } },
    { name: "it_returns_default_on_ph_type_not_found", args: ["default_fixture"], fn: (self: any, default_fixture: any): any => {
        let default$: any, placeholders: any;
        [placeholders, default$] = $list(default_fixture);
        $assert($is($dictget(placeholders, 42, default$), default$));
      } },
  ],
};
pytestClass("DescribeMasterPlaceholders", DescribeMasterPlaceholders);

const Describe_MoviePicElementCreator: any = {
  fixtures: {
    media_rId_fixture: { args: ["_video_part_rIds_prop_"], fn: (self: any, _video_part_rIds_prop_: any): any => {
        let expected_value: any, movie_pic_element_creator: any;
        movie_pic_element_creator = $call(_MoviePicElementCreator, null, null, null, null, null, null, null, null, null);
        expected_value = "rId24";
        _video_part_rIds_prop_.returnValue = [expected_value, "rId666"];
        return [movie_pic_element_creator, expected_value];
      } },
    movie_pic_fixture: { args: ["shapes_", "_MoviePicElementCreator_init_", "_pic_prop_", "pic_"], fn: (self: any, shapes_: any, _MoviePicElementCreator_init_: any, _pic_prop_: any, pic_: any): any => {
        let cx: any, cy: any, mime_type: any, movie_file: any, poster_frame_image: any, shape_id: any, x: any, y: any;
        [shape_id, movie_file, x, y, cx, cy] = $list([42, "movie.mp4", 1, 2, 3, 4]);
        [poster_frame_image, mime_type] = $list(["image.png", "video/mp4"]);
        return [shapes_, shape_id, movie_file, x, y, cx, cy, poster_frame_image, mime_type, _MoviePicElementCreator_init_, _pic_prop_, pic_];
      } },
    part_rIds_fixture: { args: ["slide_part_", "video_", "_slide_part_prop_", "_video_prop_"], fn: (self: any, slide_part_: any, video_: any, _slide_part_prop_: any, _video_prop_: any): any => {
        let media_rId: any, movie_pic_element_creator: any, video_rId: any;
        movie_pic_element_creator = $call(_MoviePicElementCreator, null, null, null, null, null, null, null, null, null);
        [media_rId, video_rId] = $list(["rId42", "rId24"]);
        _slide_part_prop_.returnValue = slide_part_;
        slide_part_.getOrAddVideoMediaPart.returnValue = [media_rId, video_rId];
        _video_prop_.returnValue = video_;
        return [movie_pic_element_creator, slide_part_, video_, media_rId, video_rId];
      } },
    pfrm_rId_fixture: { args: ["_slide_part_prop_", "slide_part_", "_poster_frame_image_file_prop_"], fn: (self: any, _slide_part_prop_: any, slide_part_: any, _poster_frame_image_file_prop_: any): any => {
        let expected_value: any, movie_pic_element_creator: any, poster_frame_image_file: any;
        movie_pic_element_creator = $call(_MoviePicElementCreator, null, null, null, null, null, null, null, null, null);
        [poster_frame_image_file, expected_value] = $list(["image.png", "rId42"]);
        _slide_part_prop_.returnValue = slide_part_;
        _poster_frame_image_file_prop_.returnValue = poster_frame_image_file;
        slide_part_.getOrAddImagePart.returnValue = [null, expected_value];
        return [movie_pic_element_creator, slide_part_, poster_frame_image_file, expected_value];
      } },
    pic_fixture: { args: ["new_video_pic_", "pic_", "_shape_name_prop_", "_video_rId_prop_", "_media_rId_prop_", "_poster_frame_rId_prop_"], fn: (self: any, new_video_pic_: any, pic_: any, _shape_name_prop_: any, _video_rId_prop_: any, _media_rId_prop_: any, _poster_frame_rId_prop_: any): any => {
        let cx: any, cy: any, media_rId: any, movie_pic_element_creator: any, poster_frame_rId: any, shape_id: any, shape_name: any, video_rId: any, x: any, y: any;
        [shape_id, x, y, cx, cy] = $list([42, 1, 2, 3, 4]);
        movie_pic_element_creator = $call(_MoviePicElementCreator, null, shape_id, null, x, y, cx, cy, null, null);
        const $t1 = "movie.mp4";
        _shape_name_prop_.returnValue = $t1;
        shape_name = $t1;
        const $t2 = "rId1";
        _video_rId_prop_.returnValue = $t2;
        video_rId = $t2;
        const $t3 = ["rId2"];
        _media_rId_prop_.returnValue = $t3;
        media_rId = $t3;
        const $t4 = "rId3";
        _poster_frame_rId_prop_.returnValue = $t4;
        poster_frame_rId = $t4;
        new_video_pic_.returnValue = pic_;
        return [movie_pic_element_creator, new_video_pic_, shape_id, shape_name, video_rId, media_rId, poster_frame_rId, x, y, cx, cy, pic_];
      } },
    shape_name_fixture: { args: ["_video_prop_", "video_"], fn: (self: any, _video_prop_: any, video_: any): any => {
        let filename: any, movie_pic_element_creator: any;
        movie_pic_element_creator = $call(_MoviePicElementCreator, null, null, null, null, null, null, null, null, null);
        _video_prop_.returnValue = video_;
        const $t5 = "movie.mp4";
        video_.filename = $t5;
        filename = $t5;
        return [movie_pic_element_creator, filename];
      } },
    slide_part_fixture: { args: ["shapes_", "slide_part_"], fn: (self: any, shapes_: any, slide_part_: any): any => {
        let movie_pic_element_creator: any;
        movie_pic_element_creator = $call(_MoviePicElementCreator, shapes_, null, null, null, null, null, null, null, null);
        shapes_.part = slide_part_;
        return [movie_pic_element_creator, slide_part_];
      } },
    video_rId_fixture: { args: ["_video_part_rIds_prop_"], fn: (self: any, _video_part_rIds_prop_: any): any => {
        let expected_value: any, movie_pic_element_creator: any;
        movie_pic_element_creator = $call(_MoviePicElementCreator, null, null, null, null, null, null, null, null, null);
        expected_value = "rId42";
        _video_part_rIds_prop_.returnValue = ["rId666", expected_value];
        return [movie_pic_element_creator, expected_value];
      } },
    video_fixture: { args: ["video_", "from_path_or_file_like_"], fn: (self: any, video_: any, from_path_or_file_like_: any): any => {
        let mime_type: any, movie_file: any, movie_pic_element_creator: any;
        [movie_file, mime_type] = $list(["movie.mp4", "video/mp4"]);
        movie_pic_element_creator = $call(_MoviePicElementCreator, null, null, movie_file, null, null, null, null, null, mime_type);
        from_path_or_file_like_.returnValue = video_;
        return [movie_pic_element_creator, movie_file, mime_type, video_];
      } },
    from_path_or_file_like_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, Video, "from_path_or_file_like", $kw({autospec: false}));
      } },
    _media_rId_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, _MoviePicElementCreator, "_media_rId");
      } },
    _MoviePicElementCreator_init_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(initializerMock, request, _MoviePicElementCreator, $kw({autospec: true}));
      } },
    new_video_pic_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, CT_Picture, "new_video_pic", $kw({autospec: false}));
      } },
    pic_: { args: [], fn: (self: any): any => {
        return element("p:pic");
      } },
    _pic_prop_: { args: ["request", "pic_"], fn: (self: any, request: any, pic_: any): any => {
        return $call(propertyMock, request, _MoviePicElementCreator, "_pic", $kw({return_value: pic_}));
      } },
    _poster_frame_image_file_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, _MoviePicElementCreator, "_poster_frame_image_file");
      } },
    _poster_frame_rId_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, _MoviePicElementCreator, "_poster_frame_rId");
      } },
    _shape_name_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, _MoviePicElementCreator, "_shape_name");
      } },
    shapes_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _BaseShapes);
      } },
    slide_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlidePart);
      } },
    _slide_part_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, _MoviePicElementCreator, "_slide_part");
      } },
    video_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Video);
      } },
    _video_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, _MoviePicElementCreator, "_video");
      } },
    _video_rId_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, _MoviePicElementCreator, "_video_rId");
      } },
    _video_part_rIds_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, _MoviePicElementCreator, "_video_part_rIds");
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_creates_a_new_movie_pic_element", args: ["movie_pic_fixture"], fn: (self: any, movie_pic_fixture: any): any => {
        let _MoviePicElementCreator_init_: any, _pic_prop_: any, cx: any, cy: any, mime_type: any, movie_file: any, pic: any, pic_: any, poster_frame_image: any, shape_id: any, shapes_: any, x: any, y: any;
        [shapes_, shape_id, movie_file, x, y, cx, cy] = $list($slice(movie_pic_fixture, null, 7, null));
        [poster_frame_image, mime_type] = $list($slice(movie_pic_fixture, 7, 9, null));
        [_MoviePicElementCreator_init_, _pic_prop_] = $list($slice(movie_pic_fixture, 9, 11, null));
        pic_ = $get(movie_pic_fixture, 11);
        pic = _MoviePicElementCreator.newMoviePic(shapes_, shape_id, movie_file, x, y, cx, cy, poster_frame_image, mime_type);
        _MoviePicElementCreator_init_.assertCalledOnceWith(ANY, shapes_, shape_id, movie_file, x, y, cx, cy, poster_frame_image, mime_type);
        _pic_prop_.assertCalledOnceWith();
        $assert($is(pic, pic_));
      } },
    { name: "it_creates_a_pic_element", args: ["pic_fixture"], fn: (self: any, pic_fixture: any): any => {
        let cx: any, cy: any, media_rId: any, movie_pic_element_creator: any, new_video_pic_: any, pic: any, pic_: any, poster_frame_rId: any, shape_id: any, shape_name: any, video_rId: any, x: any, y: any;
        [movie_pic_element_creator, new_video_pic_, shape_id] = $list($slice(pic_fixture, null, 3, null));
        [shape_name, video_rId, media_rId, poster_frame_rId] = $list($slice(pic_fixture, 3, 7, null));
        [x, y, cx, cy, pic_] = $list($slice(pic_fixture, 7, null, null));
        pic = movie_pic_element_creator._pic;
        new_video_pic_.assertCalledOnceWith(shape_id, shape_name, video_rId, media_rId, poster_frame_rId, x, y, cx, cy);
        $assert($is(pic, pic_));
      } },
    { name: "it_knows_the_shape_name_to_help", args: ["shape_name_fixture"], fn: (self: any, shape_name_fixture: any): any => {
        let filename: any, movie_pic_element_creator: any, shape_name: any;
        [movie_pic_element_creator, filename] = $list(shape_name_fixture);
        shape_name = movie_pic_element_creator._shapeName;
        $assertEqual(shape_name, filename);
      } },
    { name: "it_constructs_the_video_to_help", args: ["video_fixture"], fn: (self: any, video_fixture: any): any => {
        let mime_type: any, movie_file: any, movie_pic_element_creator: any, video: any, video_: any;
        [movie_pic_element_creator, movie_file] = $list($slice(video_fixture, null, 2, null));
        [mime_type, video_] = $list($slice(video_fixture, 2, null, null));
        video = movie_pic_element_creator._video;
        Video.fromPathOrFileLike.assertCalledOnceWith(movie_file, mime_type);
        $assert($is(video, video_));
      } },
    { name: "it_knows_the_media_rId_to_help", args: ["media_rId_fixture"], fn: (self: any, media_rId_fixture: any): any => {
        let expected_value: any, movie_pic_element_creator: any;
        [movie_pic_element_creator, expected_value] = $list(media_rId_fixture);
        $assertEqual(movie_pic_element_creator._mediaRId, expected_value);
      } },
    { name: "it_knows_the_video_rId_to_help", args: ["video_rId_fixture"], fn: (self: any, video_rId_fixture: any): any => {
        let expected_value: any, movie_pic_element_creator: any;
        [movie_pic_element_creator, expected_value] = $list(video_rId_fixture);
        $assertEqual(movie_pic_element_creator._videoRId, expected_value);
      } },
    { name: "it_adds_the_poster_frame_image_to_help", args: ["pfrm_rId_fixture"], fn: (self: any, pfrm_rId_fixture: any): any => {
        let expected_value: any, movie_pic_element_creator: any, poster_frame_image_file: any, poster_frame_rId: any, slide_part_: any;
        [movie_pic_element_creator, slide_part_] = $list($slice(pfrm_rId_fixture, null, 2, null));
        [poster_frame_image_file, expected_value] = $list($slice(pfrm_rId_fixture, 2, null, null));
        poster_frame_rId = movie_pic_element_creator._posterFrameRId;
        slide_part_.getOrAddImagePart.assertCalledOnceWith(poster_frame_image_file);
        $assertEqual(poster_frame_rId, expected_value);
      } },
    { name: "it_gets_the_poster_frame_image_from_the_specified_path_to_help", args: ["request"], fn: (self: any, request: any): any => {
        let BytesIO_: any, image_file: any, movie_pic_element_creator: any;
        BytesIO_ = classMock(request, "pptx.fileio.BytesIO");
        movie_pic_element_creator = $call(_MoviePicElementCreator, null, null, null, null, null, null, null, "image.png", null);
        image_file = movie_pic_element_creator._posterFrameImageFile;
        BytesIO_.assertNotCalled();
        $assertEqual(image_file, "image.png");
      } },
    { name: "but_it_gets_the_poster_frame_image_from_the_default_bytes_when_None_specified", args: ["request"], fn: (self: any, request: any): any => {
        let BytesIO_: any, image_file: any, movie_pic_element_creator: any, stream_: any;
        stream_ = instanceMock(request, BytesIO);
        BytesIO_ = $call(classMock, request, "pptx.fileio.BytesIO", $kw({return_value: stream_}));
        movie_pic_element_creator = $call(_MoviePicElementCreator, null, null, null, null, null, null, null, null, null);
        image_file = movie_pic_element_creator._posterFrameImageFile;
        BytesIO_.assertCalledOnceWith(SPEAKER_IMAGE_BYTES);
        $assertEqual(image_file, stream_);
      } },
    { name: "it_gets_the_video_part_rIds_to_help", args: ["part_rIds_fixture"], fn: (self: any, part_rIds_fixture: any): any => {
        let media_rId: any, movie_pic_element_creator: any, result: any, slide_part_: any, video_: any, video_rId: any;
        [movie_pic_element_creator, slide_part_] = $list($slice(part_rIds_fixture, null, 2, null));
        [video_, media_rId, video_rId] = $list($slice(part_rIds_fixture, 2, null, null));
        result = movie_pic_element_creator._videoPartRIds;
        slide_part_.getOrAddVideoMediaPart.assertCalledOnceWith(video_);
        $assertEqual(result, [media_rId, video_rId]);
      } },
    { name: "it_gets_the_slide_part_to_help", args: ["slide_part_fixture"], fn: (self: any, slide_part_fixture: any): any => {
        let movie_pic_element_creator: any, slide_part: any, slide_part_: any;
        [movie_pic_element_creator, slide_part_] = $list(slide_part_fixture);
        slide_part = movie_pic_element_creator._slidePart;
        $assert($is(slide_part, slide_part_));
      } },
  ],
};
pytestClass("Describe_MoviePicElementCreator", Describe_MoviePicElementCreator);

const Describe_OleObjectElementCreator: any = {
  fixtures: {
    shapes_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _BaseGroupShapes);
      } },
    slide_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlidePart);
      } },
    _slide_part_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, _OleObjectElementCreator, "_slide_part");
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_a_graphicFrame_interface_method", args: ["request", "shapes_"], fn: (self: any, request: any, shapes_: any): any => {
        let _graphicFrame_prop_: any, _init_: any, cx: any, cy: any, graphicFrame: any, graphicFrame_: any, shape_id: any, x: any, y: any;
        [shape_id, x, y, cx, cy] = $list([42, 1, 2, 3, 4]);
        _init_ = $call(initializerMock, request, _OleObjectElementCreator, $kw({autospec: true}));
        graphicFrame_ = element("p:graphicFrame");
        _graphicFrame_prop_ = $call(propertyMock, request, _OleObjectElementCreator, "_graphicFrame", $kw({return_value: graphicFrame_}));
        graphicFrame = _OleObjectElementCreator.graphicFrame(shapes_, shape_id, "sheet.xlsx", PROG_ID.XLSX, x, y, cx, cy, "icon.png", $call(Inches, 0.5), $call(Inches, 0.75));
        _init_.assertCalledOnceWith(ANY, shapes_, shape_id, "sheet.xlsx", PROG_ID.XLSX, x, y, cx, cy, "icon.png", $call(Inches, 0.5), $call(Inches, 0.75));
        _graphicFrame_prop_.assertCalledOnceWith();
        $assert($is(graphicFrame, graphicFrame_));
      } },
    { name: "it_creates_the_graphicFrame_element", args: ["request"], fn: (self: any, request: any): any => {
        let cx: any, cy: any, element_creator: any, shape_id: any, x: any, y: any;
        [shape_id, x, y, cx, cy] = $list([7, 1, 2, 3, 4]);
        $call(propertyMock, request, _OleObjectElementCreator, "_shape_name", $kw({return_value: "Object 42"}));
        $call(propertyMock, request, _OleObjectElementCreator, "_ole_object_rId", $kw({return_value: "rId42"}));
        $call(propertyMock, request, _OleObjectElementCreator, "_progId", $kw({return_value: "Excel.Sheet.42"}));
        $call(propertyMock, request, _OleObjectElementCreator, "_icon_rId", $kw({return_value: "rId24"}));
        $call(propertyMock, request, _OleObjectElementCreator, "_cx", $kw({return_value: cx}));
        $call(propertyMock, request, _OleObjectElementCreator, "_cy", $kw({return_value: cy}));
        element_creator = $call(_OleObjectElementCreator, null, shape_id, null, null, x, y, cx, cy, null, $call(Inches, 0.5), $call(Inches, 0.75));
        $assertEqual(element_creator._graphicFrame.xml, "<p:graphicFrame xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n  <p:nvGraphicFramePr>\n    <p:cNvPr id=\"7\" name=\"Object 42\"/>\n    <p:cNvGraphicFramePr>\n      <a:graphicFrameLocks noGrp=\"1\"/>\n    </p:cNvGraphicFramePr>\n    <p:nvPr/>\n  </p:nvGraphicFramePr>\n  <p:xfrm>\n    <a:off x=\"1\" y=\"2\"/>\n    <a:ext cx=\"3\" cy=\"4\"/>\n  </p:xfrm>\n  <a:graphic>\n    <a:graphicData uri=\"http://schemas.openxmlformats.org/presentationml/2006/ole\">\n      <p:oleObj showAsIcon=\"1\" r:id=\"rId42\" imgW=\"457200\" imgH=\"685800\" progId=\"Excel.Sheet.42\">\n        <p:embed/>\n        <p:pic>\n          <p:nvPicPr>\n            <p:cNvPr id=\"0\" name=\"\"/>\n            <p:cNvPicPr/>\n            <p:nvPr/>\n          </p:nvPicPr>\n          <p:blipFill>\n            <a:blip r:embed=\"rId24\"/>\n            <a:stretch>\n              <a:fillRect/>\n            </a:stretch>\n          </p:blipFill>\n          <p:spPr>\n            <a:xfrm>\n              <a:off x=\"1\" y=\"2\"/>\n              <a:ext cx=\"3\" cy=\"4\"/>\n            </a:xfrm>\n            <a:prstGeom prst=\"rect\">\n              <a:avLst/>\n            </a:prstGeom>\n          </p:spPr>\n        </p:pic>\n      </p:oleObj>\n    </a:graphicData>\n  </a:graphic>\n</p:graphicFrame>\"\n");
      } },
    { name: "it_determines_the_shape_width_to_help", args: ["cx_arg", "prog_id", "expected_value"], parametrize: [{ names: ["cx_arg", "prog_id", "expected_value"], values: [[$call(Emu, 999999), null, $call(Emu, 999999)], [null, PROG_ID.DOCX, $call(Emu, 965200)], [null, PROG_ID.PPTX, $call(Emu, 965200)], [null, PROG_ID.XLSX, $call(Emu, 965200)], [null, "Foo.Bar.6", $call(Emu, 965200)]] }], fn: (self: any, cx_arg: any, prog_id: any, expected_value: any): any => {
        let element_creator: any;
        element_creator = $call(_OleObjectElementCreator, null, null, null, prog_id, null, null, cx_arg, null, null, null, null);
        $assertEqual(element_creator._cx, expected_value);
      } },
    { name: "it_determines_the_shape_height_to_help", args: ["cy_arg", "prog_id", "expected_value"], parametrize: [{ names: ["cy_arg", "prog_id", "expected_value"], values: [[$call(Emu, 666666), null, $call(Emu, 666666)], [null, PROG_ID.DOCX, $call(Emu, 609600)], [null, PROG_ID.PPTX, $call(Emu, 609600)], [null, PROG_ID.XLSX, $call(Emu, 609600)], [null, "Foo.Bar.6", $call(Emu, 609600)]] }], fn: (self: any, cy_arg: any, prog_id: any, expected_value: any): any => {
        let element_creator: any;
        element_creator = $call(_OleObjectElementCreator, null, null, null, prog_id, null, null, null, cy_arg, null, null, null);
        $assertEqual(element_creator._cy, expected_value);
      } },
    { name: "it_determines_the_icon_height_to_help", args: ["icon_height_arg", "expected_value"], parametrize: [{ names: ["icon_height_arg", "expected_value"], values: [[$call(Emu, 666666), $call(Emu, 666666)], [null, $call(Emu, 609600)]] }], fn: (self: any, icon_height_arg: any, expected_value: any): any => {
        let element_creator: any;
        element_creator = $call(_OleObjectElementCreator, null, null, null, null, null, null, null, null, null, null, icon_height_arg);
        $assertEqual(element_creator._iconHeight, expected_value);
      } },
    { name: "it_resolves_the_icon_image_file_to_help", args: ["icon_file_arg", "prog_id", "expected_value"], parametrize: [{ names: ["icon_file_arg", "prog_id", "expected_value"], values: [["user-icon.png", PROG_ID.XLSX, "user-icon.png"], [null, "Foo.Bar.18", "generic-icon.emf"], [null, PROG_ID.DOCX, "docx-icon.emf"], [null, PROG_ID.PPTX, "pptx-icon.emf"], [null, PROG_ID.XLSX, "xlsx-icon.emf"]] }], fn: (self: any, icon_file_arg: any, prog_id: any, expected_value: any): any => {
        let element_creator: any;
        element_creator = $call(_OleObjectElementCreator, null, null, null, prog_id, null, null, null, null, icon_file_arg, null, null);
        // the bundled icons are opened from memory (BytesIO of the template bytes), not from a path
        const $icon = element_creator._iconImageFile;
        $assert(typeof $icon === "string" ? $endswith($icon, expected_value) : $eq($icon.getvalue(), templateBytes(expected_value)));
      } },
    { name: "it_adds_and_relates_the_icon_image_part_to_help", args: ["request", "_slide_part_prop_", "slide_part_"], fn: (self: any, request: any, _slide_part_prop_: any, slide_part_: any): any => {
        let element_creator: any, rId: any;
        $call(propertyMock, request, _OleObjectElementCreator, "_icon_image_file", $kw({return_value: "obj-icon.emf"}));
        slide_part_.getOrAddImagePart.returnValue = [null, "rId16"];
        _slide_part_prop_.returnValue = slide_part_;
        element_creator = $call(_OleObjectElementCreator, null, null, null, null, null, null, null, null, null, null, null);
        rId = element_creator._iconRId;
        slide_part_.getOrAddImagePart.assertCalledOnceWith("obj-icon.emf");
        $assertEqual(rId, "rId16");
      } },
    { name: "it_determines_the_icon_width_to_help", args: ["icon_width_arg", "expected_value"], parametrize: [{ names: ["icon_width_arg", "expected_value"], values: [[$call(Emu, 666666), $call(Emu, 666666)], [null, $call(Emu, 965200)]] }], fn: (self: any, icon_width_arg: any, expected_value: any): any => {
        let element_creator: any;
        element_creator = $call(_OleObjectElementCreator, null, null, null, null, null, null, null, null, null, icon_width_arg, null);
        $assertEqual(element_creator._iconWidth, expected_value);
      } },
    { name: "it_adds_and_relates_the_ole_object_part_to_help", args: ["request", "_slide_part_prop_", "slide_part_"], fn: (self: any, request: any, _slide_part_prop_: any, slide_part_: any): any => {
        let element_creator: any, ole_object_file: any, rId: any;
        ole_object_file = "workbook.xlsx";
        slide_part_.addEmbeddedOleObjectPart.returnValue = "rId14";
        _slide_part_prop_.returnValue = slide_part_;
        element_creator = $call(_OleObjectElementCreator, null, null, ole_object_file, PROG_ID.DOCX, null, null, null, null, null, null, null);
        rId = element_creator._oleObjectRId;
        slide_part_.addEmbeddedOleObjectPart.assertCalledOnceWith(PROG_ID.DOCX, ole_object_file);
        $assertEqual(rId, "rId14");
      } },
    { name: "it_resolves_the_progId_str_to_help", args: ["prog_id_arg", "expected_value"], parametrize: [{ names: ["prog_id_arg", "expected_value"], values: [[PROG_ID.DOCX, "Word.Document.12"], [PROG_ID.PPTX, "PowerPoint.Show.12"], [PROG_ID.XLSX, "Excel.Sheet.12"], ["Something.Else.42", "Something.Else.42"]] }], fn: (self: any, prog_id_arg: any, expected_value: any): any => {
        let element_creator: any;
        element_creator = $call(_OleObjectElementCreator, null, null, null, prog_id_arg, null, null, null, null, null, null, null);
        $assertEqual(element_creator._progId, expected_value);
      } },
    { name: "it_computes_the_shape_name_to_help", args: [], fn: (self: any): any => {
        let element_creator: any, shape_id: any;
        shape_id = 42;
        element_creator = $call(_OleObjectElementCreator, null, shape_id, null, null, null, null, null, null, null, null, null);
        $assertEqual(element_creator._shapeName, "Object 41");
      } },
    { name: "it_provides_access_to_the_slide_part_to_help", args: ["shapes_", "slide_part_"], fn: (self: any, shapes_: any, slide_part_: any): any => {
        let element_creator: any;
        shapes_.part = slide_part_;
        element_creator = $call(_OleObjectElementCreator, shapes_, null, null, null, null, null, null, null, null, null, null);
        $assert($is(element_creator._slidePart, slide_part_));
      } },
  ],
};
pytestClass("Describe_OleObjectElementCreator", Describe_OleObjectElementCreator);

const Describe_NotesSlideShapeFactory: any = {
  fixtures: {
    factory_fixture: { args: ["request", "shapes_", "shape_"], params: [["p:sp", "BaseShapeFactory"], ["p:sp/p:nvSpPr/p:nvPr/p:ph", "NotesSlidePlaceholder"]], fn: (self: any, request: any, shapes_: any, shape_: any): any => {
        let ShapeConstructor_: any, shape_cls_name: any, shape_elm: any, sp_cxml: any;
        [sp_cxml, shape_cls_name] = $list(request.param);
        shape_elm = element(sp_cxml);
        ShapeConstructor_ = $call(classMock, request, $mod("pptx.shapes.shapetree.%s", shape_cls_name), $kw({return_value: shape_}));
        return [shape_elm, shapes_, ShapeConstructor_, shape_];
      } },
    shape_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, BaseShape);
      } },
    shapes_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _BaseShapes);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_constructs_the_right_shape_for_an_element", args: ["factory_fixture"], fn: (self: any, factory_fixture: any): any => {
        let ShapeConstructor_: any, parent_: any, shape: any, shape_: any, shape_elm: any;
        [shape_elm, parent_, ShapeConstructor_, shape_] = $list(factory_fixture);
        shape = $call(_NotesSlideShapeFactory, shape_elm, parent_);
        ShapeConstructor_.assertCalledOnceWith(shape_elm, parent_);
        $assert($is(shape, shape_));
      } },
  ],
};
pytestClass("Describe_NotesSlideShapeFactory", Describe_NotesSlideShapeFactory);

