// Machine-translated from python-pptx tests/shapes/test_placeholder.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $call, $d, $get, $getattr, $is, $isinstance, $kw, $list, $mod, $raises, $setattr, $truthy, classMock, instanceMock, methodMock, propertyMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { ChartData as ChartData$imported } from "../../src/chart/data.js";
import { XL_CHART_TYPE as XCT$imported } from "../../src/enum/chart.js";
import { MSO_SHAPE_TYPE as MSO_SHAPE_TYPE$imported, PP_PLACEHOLDER as PP_PLACEHOLDER$imported } from "../../src/enum/shapes.js";
import { ST_Direction as ST_Direction$imported, ST_PlaceholderSize as ST_PlaceholderSize$imported } from "../../src/oxml/shapes/shared.js";
import { ImagePart as ImagePart$imported } from "../../src/parts/image.js";
import { NotesSlidePart as NotesSlidePart$imported, SlidePart as SlidePart$imported } from "../../src/parts/slide.js";
import { BasePlaceholder as BasePlaceholder$imported, ChartPlaceholder as ChartPlaceholder$imported, LayoutPlaceholder as LayoutPlaceholder$imported, MasterPlaceholder as MasterPlaceholder$imported, NotesSlidePlaceholder as NotesSlidePlaceholder$imported, PicturePlaceholder as PicturePlaceholder$imported, PlaceholderGraphicFrame as PlaceholderGraphicFrame$imported, PlaceholderPicture as PlaceholderPicture$imported, TablePlaceholder as TablePlaceholder$imported, _BaseSlidePlaceholder as _BaseSlidePlaceholder$imported, _InheritsDimensions as _InheritsDimensions$imported } from "../../src/shapes/placeholder.js";
import { NotesSlidePlaceholders as NotesSlidePlaceholders$imported } from "../../src/shapes/shapetree.js";
import { NotesMaster as NotesMaster$imported, SlideLayout as SlideLayout$imported, SlideMaster as SlideMaster$imported } from "../../src/slide.js";
import { aGraphicFrame, aPh, anNvGraphicFramePr, anNvPicPr, anNvPr, anNvSpPr, anSp } from "../oxml/unitdata/shape.js";
import { element, xml } from "../unitutil/cxml.js";
import { snippetSeq } from "../unitutil/file.js";
// python binds imported names at import time; patches made later don't change them
const ChartData = ChartData$imported;
const XCT = XCT$imported;
const MSO_SHAPE_TYPE = MSO_SHAPE_TYPE$imported;
const PP_PLACEHOLDER = PP_PLACEHOLDER$imported;
const ST_Direction = ST_Direction$imported;
const ST_PlaceholderSize = ST_PlaceholderSize$imported;
const ImagePart = ImagePart$imported;
const NotesSlidePart = NotesSlidePart$imported;
const SlidePart = SlidePart$imported;
const BasePlaceholder = BasePlaceholder$imported;
const ChartPlaceholder = ChartPlaceholder$imported;
const LayoutPlaceholder = LayoutPlaceholder$imported;
const MasterPlaceholder = MasterPlaceholder$imported;
const NotesSlidePlaceholder = NotesSlidePlaceholder$imported;
const PicturePlaceholder = PicturePlaceholder$imported;
const PlaceholderGraphicFrame = PlaceholderGraphicFrame$imported;
const PlaceholderPicture = PlaceholderPicture$imported;
const TablePlaceholder = TablePlaceholder$imported;
const _BaseSlidePlaceholder = _BaseSlidePlaceholder$imported;
const _InheritsDimensions = _InheritsDimensions$imported;
const NotesSlidePlaceholders = NotesSlidePlaceholders$imported;
const NotesMaster = NotesMaster$imported;
const SlideLayout = SlideLayout$imported;
const SlideMaster = SlideMaster$imported;
void $py;


const Describe_BaseSlidePlaceholder: any = {
  fixtures: {
    base_ph_fixture: { args: ["request", "part_prop_", "slide_layout_", "layout_placeholder_"], fn: (self: any, request: any, part_prop_: any, slide_layout_: any, layout_placeholder_: any): any => {
        let placeholder: any, sp_cxml: any;
        sp_cxml = "p:sp/p:nvSpPr/p:nvPr/p:ph{type=pic,idx=1}";
        placeholder = $call(_BaseSlidePlaceholder, element(sp_cxml), null);
        part_prop_.returnValue.slideLayout = slide_layout_;
        slide_layout_.placeholders.get.returnValue = layout_placeholder_;
        return [placeholder, slide_layout_, 1, layout_placeholder_];
      } },
    base_val_fixture: { args: ["request", "_base_placeholder_prop_", "layout_placeholder_"], params: [[true, 42], [false, null]], fn: (self: any, request: any, _base_placeholder_prop_: any, layout_placeholder_: any): any => {
        let expected_value: any, has_base_placeholder: any, placeholder: any;
        [has_base_placeholder, expected_value] = $list(request.param);
        placeholder = $call(_BaseSlidePlaceholder, null, null);
        if ($truthy(has_base_placeholder)) {
          layout_placeholder_.width = expected_value;
          _base_placeholder_prop_.returnValue = layout_placeholder_;
        } else {
          _base_placeholder_prop_.returnValue = null;
        }
        return [placeholder, "width", expected_value];
      } },
    dim_set_fixture: { args: ["request"], params: [["p:sp/p:spPr/a:xfrm", "left", 1, "p:sp/p:spPr/a:xfrm/a:off{x=1,y=0}"], ["p:sp/p:spPr/a:xfrm", "top", 2, "p:sp/p:spPr/a:xfrm/a:off{x=0,y=2}"], ["p:sp/p:spPr/a:xfrm", "width", 3, "p:sp/p:spPr/a:xfrm/a:ext{cx=3,cy=0}"], ["p:sp/p:spPr/a:xfrm", "height", 4, "p:sp/p:spPr/a:xfrm/a:ext{cx=0,cy=4}"]], fn: (self: any, request: any): any => {
        let expected_cxml: any, expected_xml: any, placeholder: any, prop_name: any, sp_cxml: any, value: any;
        [sp_cxml, prop_name, value, expected_cxml] = $list(request.param);
        placeholder = $call(_BaseSlidePlaceholder, element(sp_cxml), null);
        expected_xml = xml(expected_cxml);
        return [placeholder, prop_name, value, expected_xml];
      } },
    override_fixture: { args: ["request"], params: [["left", "p:sp/p:spPr/a:xfrm/a:off{x=12}", 12], ["top", "p:sp/p:spPr/a:xfrm/a:off{y=34}", 34], ["width", "p:sp/p:spPr/a:xfrm/a:ext{cx=56}", 56], ["height", "p:sp/p:spPr/a:xfrm/a:ext{cy=78}", 78]], fn: (self: any, request: any): any => {
        let prop_name: any, slide_placeholder: any, sp_cxml: any, value: any;
        [prop_name, sp_cxml, value] = $list(request.param);
        slide_placeholder = $call(_BaseSlidePlaceholder, element(sp_cxml), null);
        return [slide_placeholder, prop_name, value];
      } },
    replace_fixture: { args: [], fn: (self: any): any => {
        let expected_xml: any, pic: any, placeholder: any, spTree: any;
        spTree = element("p:spTree/p:sp/p:nvSpPr/p:nvPr/p:ph{type=pic,idx=10}");
        pic = element("p:pic/p:nvPicPr/p:nvPr");
        placeholder = $call(_BaseSlidePlaceholder, $get(spTree, 0), null);
        expected_xml = xml("p:spTree/p:pic/p:nvPicPr/p:nvPr/p:ph{type=pic,idx=10}");
        return [placeholder, pic, spTree, expected_xml];
      } },
    _base_placeholder_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, _BaseSlidePlaceholder, "_base_placeholder");
      } },
    layout_placeholder_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, LayoutPlaceholder);
      } },
    part_prop_: { args: ["request", "slide_part_"], fn: (self: any, request: any, slide_part_: any): any => {
        return $call(propertyMock, request, _BaseSlidePlaceholder, "part", $kw({return_value: slide_part_}));
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
    { name: "it_knows_its_shape_type", args: [], fn: (self: any): any => {
        let placeholder: any;
        placeholder = $call(_BaseSlidePlaceholder, null, null);
        $assertEqual(placeholder.shapeType, MSO_SHAPE_TYPE.PLACEHOLDER);
      } },
    { name: "it_provides_override_dimensions_when_present", args: ["override_fixture"], fn: (self: any, override_fixture: any): any => {
        let expected_value: any, placeholder: any, prop_name: any;
        [placeholder, prop_name, expected_value] = $list(override_fixture);
        $assertEqual($getattr(placeholder, prop_name), expected_value);
      } },
    { name: "it_provides_inherited_dims_when_no_override", args: ["request", "prop_name"], parametrize: [{ names: ["prop_name"], values: ["left", "top", "width", "height"] }], fn: (self: any, request: any, prop_name: any): any => {
        let placeholder: any, value: any;
        $call(methodMock, request, _BaseSlidePlaceholder, "_inherited_value", $kw({return_value: 42}));
        placeholder = $call(_BaseSlidePlaceholder, element("p:sp/p:spPr"), null);
        value = $getattr(placeholder, prop_name);
        placeholder._inheritedValue.assertCalledOnceWith(placeholder, prop_name);
        $assertEqual(value, 42);
      } },
    { name: "it_gets_an_inherited_dim_value_to_help", args: ["base_val_fixture"], fn: (self: any, base_val_fixture: any): any => {
        let attr_name: any, expected_value: any, placeholder: any, value: any;
        [placeholder, attr_name, expected_value] = $list(base_val_fixture);
        value = placeholder._inheritedValue(attr_name);
        $assertEqual(value, expected_value);
      } },
    { name: "it_finds_its_base_placeholder_to_help", args: ["base_ph_fixture"], fn: (self: any, base_ph_fixture: any): any => {
        let base_placeholder: any, idx: any, layout_: any, layout_placeholder_: any, placeholder: any;
        [placeholder, layout_, idx, layout_placeholder_] = $list(base_ph_fixture);
        base_placeholder = placeholder._basePlaceholder;
        layout_.placeholders.get.assertCalledOnceWith(idx);
        $assert($is(base_placeholder, layout_placeholder_));
      } },
    { name: "it_can_override_inherited_dimensions", args: ["dim_set_fixture"], fn: (self: any, dim_set_fixture: any): any => {
        let expected_xml: any, placeholder: any, prop_name: any, value: any;
        [placeholder, prop_name, value, expected_xml] = $list(dim_set_fixture);
        $setattr(placeholder, prop_name, value);
        $assertEqual(placeholder._element.xml, expected_xml);
      } },
    { name: "it_replaces_a_placeholder_element_to_help", args: ["replace_fixture"], fn: (self: any, replace_fixture: any): any => {
        let element: any, expected_xml: any, placeholder: any, spTree: any;
        [placeholder, element, spTree, expected_xml] = $list(replace_fixture);
        placeholder._replacePlaceholderWith(element);
        $assertEqual(spTree.xml, expected_xml);
        $assert($is(placeholder._element, null));
      } },
  ],
};
pytestClass("Describe_BaseSlidePlaceholder", Describe_BaseSlidePlaceholder);

const DescribeBasePlaceholder: any = {
  fixtures: {
    idx_fixture: { args: ["request"], params: [["sp", null, "title", 0], ["sp", 0, "title", 0], ["sp", 3, "body", 3], ["pic", 6, "pic", 6], ["graphicFrame", 9, "tbl", 9]], fn: (self: any, request: any): any => {
        let expected_idx: any, idx: any, ph_type: any, placeholder: any, shape_elm: any, tagname: any;
        [tagname, idx, ph_type, expected_idx] = $list(request.param);
        shape_elm = self.shapeElmFactory(tagname, ph_type, idx);
        placeholder = $call(BasePlaceholder, shape_elm, null);
        return [placeholder, expected_idx];
      } },
    orient_fixture: { args: ["request"], params: [[null, ST_Direction.HORZ], [ST_Direction.VERT, ST_Direction.VERT]], fn: (self: any, request: any): any => {
        let expected_orient: any, orient: any, ph_bldr: any, placeholder: any, shape_elm: any;
        [orient, expected_orient] = $list(request.param);
        ph_bldr = aPh();
        if ($truthy(!$is(orient, null))) {
          ph_bldr.withOrient(orient);
        }
        shape_elm = anSp().withNsdecls("p").withChild(anNvSpPr().withChild(anNvPr().withChild(ph_bldr))).element;
        placeholder = $call(BasePlaceholder, shape_elm, null);
        return [placeholder, expected_orient];
      } },
    orient_raise_fixture: { args: [], fn: (self: any): any => {
        let shape: any, shape_elm: any;
        shape_elm = anSp().withNsdecls("p").withChild(anNvSpPr().withChild(anNvPr())).element;
        shape = $call(BasePlaceholder, shape_elm, null);
        return shape;
      } },
    sz_fixture: { args: ["request"], params: [[null, ST_PlaceholderSize.FULL], [ST_PlaceholderSize.HALF, ST_PlaceholderSize.HALF]], fn: (self: any, request: any): any => {
        let expected_sz: any, ph_bldr: any, placeholder: any, shape_elm: any, sz: any;
        [sz, expected_sz] = $list(request.param);
        ph_bldr = aPh();
        if ($truthy(!$is(sz, null))) {
          ph_bldr.withSz(sz);
        }
        shape_elm = anSp().withNsdecls("p").withChild(anNvSpPr().withChild(anNvPr().withChild(ph_bldr))).element;
        placeholder = $call(BasePlaceholder, shape_elm, null);
        return [placeholder, expected_sz];
      } },
    type_fixture: { args: ["request"], params: [["sp", null, 1, PP_PLACEHOLDER.OBJECT], ["sp", "title", 0, PP_PLACEHOLDER.TITLE], ["pic", "pic", 6, PP_PLACEHOLDER.PICTURE], ["graphicFrame", "tbl", 9, PP_PLACEHOLDER.TABLE]], fn: (self: any, request: any): any => {
        let expected_ph_type: any, idx: any, ph_type: any, placeholder: any, shape_elm: any, tagname: any;
        [tagname, ph_type, idx, expected_ph_type] = $list(request.param);
        shape_elm = self.shapeElmFactory(tagname, ph_type, idx);
        placeholder = $call(BasePlaceholder, shape_elm, null);
        return [placeholder, expected_ph_type];
      } },
  },
  helpers: {
    shapeElmFactory: (self: any, tagname: any, ph_type: any, idx: any): any => {
        let nvXxPr_bldr: any, ph_bldr: any, root_bldr: any;
        [root_bldr, nvXxPr_bldr] = $list($get($d({sp: [anSp().withNsdecls("p"), anNvSpPr()], pic: [anSp().withNsdecls("p"), anNvPicPr()], graphicFrame: [aGraphicFrame().withNsdecls("p"), anNvGraphicFramePr()]}), tagname));
        ph_bldr = $get($d({[null]: aPh().withIdx(idx), obj: aPh().withIdx(idx), title: aPh().withType("title"), body: aPh().withType("body").withIdx(idx), pic: aPh().withType("pic").withIdx(idx), tbl: aPh().withType("tbl").withIdx(idx)}), ph_type);
        return root_bldr.withChild(nvXxPr_bldr.withChild(anNvPr().withChild(ph_bldr))).element;
      },
  },
  tests: [
    { name: "it_knows_its_idx_value", args: ["idx_fixture"], fn: (self: any, idx_fixture: any): any => {
        let idx: any, placeholder: any;
        [placeholder, idx] = $list(idx_fixture);
        $assertEqual(placeholder.idx, idx);
      } },
    { name: "it_knows_its_orient_value", args: ["orient_fixture"], fn: (self: any, orient_fixture: any): any => {
        let orient: any, placeholder: any;
        [placeholder, orient] = $list(orient_fixture);
        $assertEqual(placeholder.orient, orient);
      } },
    { name: "it_raises_on_ph_orient_when_not_a_placeholder", args: ["orient_raise_fixture"], fn: (self: any, orient_raise_fixture: any): any => {
        let shape: any;
        shape = orient_raise_fixture;
        $raises(ValueError, () => {
          shape.orient;
        });
      } },
    { name: "it_knows_its_sz_value", args: ["sz_fixture"], fn: (self: any, sz_fixture: any): any => {
        let placeholder: any, sz: any;
        [placeholder, sz] = $list(sz_fixture);
        $assertEqual(placeholder.sz, sz);
      } },
    { name: "it_knows_its_placeholder_type", args: ["type_fixture"], fn: (self: any, type_fixture: any): any => {
        let ph_type: any, placeholder: any;
        [placeholder, ph_type] = $list(type_fixture);
        $assertEqual(placeholder.phType, ph_type);
      } },
  ],
};
pytestClass("DescribeBasePlaceholder", DescribeBasePlaceholder);

const DescribeChartPlaceholder: any = {
  fixtures: {
    new_fixture: { args: [], fn: (self: any): any => {
        let chart_ph: any, cx: any, cy: any, expected_xml: any, rId: any, sp_cxml: any, x: any, y: any;
        sp_cxml = "p:sp/p:nvSpPr/p:cNvPr{id=4,name=bar}";
        chart_ph = $call(ChartPlaceholder, element(sp_cxml), null);
        [rId, x, y, cx, cy] = $list(["rId42", 1, 2, 3, 4]);
        expected_xml = $get(snippetSeq("placeholders"), 1);
        return [chart_ph, rId, x, y, cx, cy, expected_xml];
      } },
    part_prop_: { args: ["request", "slide_"], fn: (self: any, request: any, slide_: any): any => {
        return $call(propertyMock, request, ChartPlaceholder, "part", $kw({return_value: slide_}));
      } },
    slide_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlidePart);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_insert_a_chart_into_itself", args: ["request", "part_prop_"], fn: (self: any, request: any, part_prop_: any): any => {
        let PlaceholderGraphicFrame_: any, _new_chart_graphicFrame_: any, _replace_placeholder_with_: any, chart_data_: any, chart_ph: any, graphicFrame: any, ph_graphic_frame: any, placeholder_graphic_frame_: any, slide_part_: any;
        slide_part_ = instanceMock(request, SlidePart);
        slide_part_.addChartPart.returnValue = "rId6";
        part_prop_.returnValue = slide_part_;
        graphicFrame = element("p:graphicFrame");
        _new_chart_graphicFrame_ = $call(methodMock, request, ChartPlaceholder, "_new_chart_graphicFrame", $kw({return_value: graphicFrame}));
        _replace_placeholder_with_ = methodMock(request, ChartPlaceholder, "_replace_placeholder_with");
        placeholder_graphic_frame_ = instanceMock(request, PlaceholderGraphicFrame);
        PlaceholderGraphicFrame_ = $call(classMock, request, "pptx.shapes.placeholder.PlaceholderGraphicFrame", $kw({return_value: placeholder_graphic_frame_}));
        chart_data_ = instanceMock(request, ChartData);
        chart_ph = $call(ChartPlaceholder, element("p:sp/p:spPr/a:xfrm/(a:off{x=1,y=2},a:ext{cx=3,cy=4})"), "parent");
        ph_graphic_frame = chart_ph.insertChart(XCT.PIE, chart_data_);
        slide_part_.addChartPart.assertCalledOnceWith(XCT.PIE, chart_data_);
        _new_chart_graphicFrame_.assertCalledOnceWith(chart_ph, "rId6", 1, 2, 3, 4);
        _replace_placeholder_with_.assertCalledOnceWith(chart_ph, graphicFrame);
        PlaceholderGraphicFrame_.assertCalledOnceWith(graphicFrame, chart_ph._parent);
        $assert($is(ph_graphic_frame, placeholder_graphic_frame_));
      } },
    { name: "it_creates_a_graphicFrame_element_to_help", args: ["new_fixture"], fn: (self: any, new_fixture: any): any => {
        let chart_ph: any, cx: any, cy: any, expected_xml: any, graphicFrame: any, rId: any, x: any, y: any;
        [chart_ph, rId, x, y, cx, cy, expected_xml] = $list(new_fixture);
        graphicFrame = chart_ph._newChartGraphicFrame(rId, x, y, cx, cy);
        $assertEqual(graphicFrame.xml, expected_xml);
      } },
  ],
};
pytestClass("DescribeChartPlaceholder", DescribeChartPlaceholder);

const DescribeLayoutPlaceholder: any = {
  fixtures: {
    base_ph_fixture: { args: ["request", "slide_master_", "master_placeholder_", "part_prop_"], params: [["p:sp/p:nvSpPr/p:nvPr/p:ph{type=body}", PP_PLACEHOLDER.BODY], ["p:sp/p:nvSpPr/p:nvPr/p:ph{type=tbl}", PP_PLACEHOLDER.BODY], ["p:sp/p:nvSpPr/p:nvPr/p:ph{type=title}", PP_PLACEHOLDER.TITLE]], fn: (self: any, request: any, slide_master_: any, master_placeholder_: any, part_prop_: any): any => {
        let layout_placeholder: any, mstr_ph_type: any, sp: any, sp_cxml: any;
        [sp_cxml, mstr_ph_type] = $list(request.param);
        sp = element(sp_cxml);
        layout_placeholder = $call(LayoutPlaceholder, sp, null);
        part_prop_.returnValue.slideMaster = slide_master_;
        slide_master_.placeholders.get.returnValue = master_placeholder_;
        return [layout_placeholder, slide_master_, mstr_ph_type, master_placeholder_];
      } },
    master_placeholder_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, MasterPlaceholder);
      } },
    part_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, LayoutPlaceholder, "part");
      } },
    slide_master_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlideMaster);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_uses_InheritsDimensions_mixin", args: [], fn: (self: any): any => {
        let layout_placeholder: any;
        layout_placeholder = $call(LayoutPlaceholder, null, null);
        $assert($isinstance(layout_placeholder, _InheritsDimensions));
      } },
    { name: "it_finds_its_base_placeholder_to_help", args: ["base_ph_fixture"], fn: (self: any, base_ph_fixture: any): any => {
        let layout_placeholder: any, master_: any, master_placeholder: any, master_placeholder_: any, mstr_ph_type: any;
        [layout_placeholder, master_, mstr_ph_type, master_placeholder_] = $list(base_ph_fixture);
        master_placeholder = layout_placeholder._basePlaceholder;
        master_.placeholders.get.assertCalledOnceWith(mstr_ph_type, null);
        $assert($is(master_placeholder, master_placeholder_));
      } },
  ],
};
pytestClass("DescribeLayoutPlaceholder", DescribeLayoutPlaceholder);

const DescribeNotesSlidePlaceholder: any = {
  fixtures: {
    base_ph_fixture: { args: ["request", "notes_master_", "master_placeholder_", "part_prop_", "notes_slide_part_", "notes_slide_placeholders_"], fn: (self: any, request: any, notes_master_: any, master_placeholder_: any, part_prop_: any, notes_slide_part_: any, notes_slide_placeholders_: any): any => {
        let placeholder: any, sp_cxml: any;
        sp_cxml = "p:sp/p:nvSpPr/p:nvPr/p:ph{type=body}";
        placeholder = $call(NotesSlidePlaceholder, element(sp_cxml), null);
        notes_slide_part_.notesMaster = notes_master_;
        notes_master_.placeholders = notes_slide_placeholders_;
        notes_slide_placeholders_.get.returnValue = master_placeholder_;
        return [placeholder, notes_master_, PP_PLACEHOLDER.BODY, master_placeholder_];
      } },
    notes_slide_placeholders_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, NotesSlidePlaceholders);
      } },
    master_placeholder_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, MasterPlaceholder);
      } },
    notes_master_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, NotesMaster);
      } },
    notes_slide_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, NotesSlidePart);
      } },
    part_prop_: { args: ["request", "notes_slide_part_"], fn: (self: any, request: any, notes_slide_part_: any): any => {
        return $call(propertyMock, request, NotesSlidePlaceholder, "part", $kw({return_value: notes_slide_part_}));
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_finds_its_base_placeholder_to_help", args: ["base_ph_fixture"], fn: (self: any, base_ph_fixture: any): any => {
        let base_placeholder: any, master_placeholder_: any, notes_master_: any, ph_type: any, placeholder: any;
        [placeholder, notes_master_, ph_type, master_placeholder_] = $list(base_ph_fixture);
        base_placeholder = placeholder._basePlaceholder;
        notes_master_.placeholders.get.assertCalledOnceWith(ph_type);
        $assert($is(base_placeholder, master_placeholder_));
      } },
  ],
};
pytestClass("DescribeNotesSlidePlaceholder", DescribeNotesSlidePlaceholder);

const DescribePicturePlaceholder: any = {
  fixtures: {
    get_or_add_fixture: { args: ["part_prop_", "image_part_"], fn: (self: any, part_prop_: any, image_part_: any): any => {
        let desc: any, expected_value: any, image_file: any, image_size: any, placeholder: any, rId: any;
        placeholder = $call(PicturePlaceholder, null, null);
        [image_file, rId, desc, image_size] = $list(["f.png", "rId6", "desc", [42, 24]]);
        part_prop_.returnValue.getOrAddImagePart.returnValue = [image_part_, rId];
        [image_part_.desc, image_part_._pxSize] = $list([desc, image_size]);
        expected_value = [rId, desc, image_size];
        return [placeholder, image_file, expected_value];
      } },
    image_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, ImagePart);
      } },
    part_prop_: { args: ["request", "slide_"], fn: (self: any, request: any, slide_: any): any => {
        return $call(propertyMock, request, PicturePlaceholder, "part", $kw({return_value: slide_}));
      } },
    slide_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlidePart);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_insert_a_picture_into_itself", args: ["request"], fn: (self: any, request: any): any => {
        let PlaceholderPicture_: any, _new_placeholder_pic_: any, _replace_placeholder_with_: any, pic: any, picture_ph: any, placeholder_picture: any, placeholder_picture_: any;
        pic = element("p:pic");
        _new_placeholder_pic_ = $call(methodMock, request, PicturePlaceholder, "_new_placeholder_pic", $kw({return_value: pic}));
        _replace_placeholder_with_ = methodMock(request, PicturePlaceholder, "_replace_placeholder_with");
        placeholder_picture_ = instanceMock(request, PlaceholderPicture);
        PlaceholderPicture_ = $call(classMock, request, "pptx.shapes.placeholder.PlaceholderPicture", $kw({return_value: placeholder_picture_}));
        picture_ph = $call(PicturePlaceholder, null, "parent");
        placeholder_picture = picture_ph.insertPicture("foobar.png");
        _new_placeholder_pic_.assertCalledOnceWith(picture_ph, "foobar.png");
        _replace_placeholder_with_.assertCalledOnceWith(picture_ph, pic);
        PlaceholderPicture_.assertCalledOnceWith(pic, picture_ph._parent);
        $assert($is(placeholder_picture, placeholder_picture_));
      } },
    { name: "it_creates_a_pic_element_to_help", args: ["request", "image_size", "crop_attr_names"], parametrize: [{ names: ["image_size", "crop_attr_names"], values: [[[444, 333], ["l", "r"]], [[333, 444], ["t", "b"]]] }], fn: (self: any, request: any, image_size: any, crop_attr_names: any): any => {
        let _get_or_add_image_: any, pic: any, picture_ph: any;
        _get_or_add_image_ = $call(methodMock, request, PicturePlaceholder, "_get_or_add_image", $kw({return_value: [42, "bar", image_size]}));
        picture_ph = $call(PicturePlaceholder, element("p:sp/(p:nvSpPr/p:cNvPr{id=2,name=foo},p:spPr/a:xfrm/a:ext{cx=99,cy=99})"), null);
        pic = picture_ph._newPlaceholderPic("foobar.png");
        _get_or_add_image_.assertCalledOnceWith(picture_ph, "foobar.png");
        $assertEqual(pic.xml, xml($mod("p:pic/(p:nvPicPr/(p:cNvPr{id=2,name=foo,descr=bar},p:cNvPicPr/a:picLocks{noGrp=1,noChangeAspect=1},p:nvPr),p:blipFill/(a:blip{r:embed=42},a:srcRect{%s=12500,%s=12500},a:stretch/a:fillRect),p:spPr)", crop_attr_names)));
      } },
    { name: "it_adds_an_image_to_help", args: ["get_or_add_fixture"], fn: (self: any, get_or_add_fixture: any): any => {
        let expected_value: any, image_file: any, placeholder: any, value: any;
        [placeholder, image_file, expected_value] = $list(get_or_add_fixture);
        value = placeholder._getOrAddImage(image_file);
        placeholder.part.getOrAddImagePart.assertCalledOnceWith(image_file);
        $assertEqual(value, expected_value);
      } },
  ],
};
pytestClass("DescribePicturePlaceholder", DescribePicturePlaceholder);

const DescribeTablePlaceholder: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_can_insert_a_table_into_itself", args: ["request"], fn: (self: any, request: any): any => {
        let PlaceholderGraphicFrame_: any, _new_placeholder_table_: any, _replace_placeholder_with_: any, graphicFrame: any, ph_graphic_frame: any, placeholder_graphic_frame_: any, table_ph: any;
        graphicFrame = element("p:graphicFrame");
        _new_placeholder_table_ = $call(methodMock, request, TablePlaceholder, "_new_placeholder_table", $kw({return_value: graphicFrame}));
        _replace_placeholder_with_ = methodMock(request, TablePlaceholder, "_replace_placeholder_with");
        placeholder_graphic_frame_ = instanceMock(request, PlaceholderGraphicFrame);
        PlaceholderGraphicFrame_ = $call(classMock, request, "pptx.shapes.placeholder.PlaceholderGraphicFrame", $kw({return_value: placeholder_graphic_frame_}));
        table_ph = $call(TablePlaceholder, null, "parent");
        ph_graphic_frame = table_ph.insertTable(4, 2);
        _new_placeholder_table_.assertCalledOnceWith(table_ph, 4, 2);
        _replace_placeholder_with_.assertCalledOnceWith(table_ph, graphicFrame);
        PlaceholderGraphicFrame_.assertCalledOnceWith(graphicFrame, table_ph._parent);
        $assert($is(ph_graphic_frame, placeholder_graphic_frame_));
      } },
    { name: "it_creates_a_graphicFrame_element_to_help", args: [], fn: (self: any): any => {
        let graphicFrame: any, table_ph: any;
        table_ph = $call(TablePlaceholder, element("p:sp/(p:nvSpPr/p:cNvPr{id=2,name=foo},p:spPr/a:xfrm/(a:off{x=1,y=2},a:ext{cx=3,cy=4}))"), null);
        graphicFrame = table_ph._newPlaceholderTable(1, 1);
        $assertEqual(graphicFrame.xml, $get(snippetSeq("placeholders"), 0));
      } },
  ],
};
pytestClass("DescribeTablePlaceholder", DescribeTablePlaceholder);

