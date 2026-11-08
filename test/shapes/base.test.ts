// Machine-translated from python-pptx tests/shapes/test_base.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $call, $cast, $get, $is, $kw, $list, $raises, $t1, $truthy, classMock, instanceMock, looseMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { ActionSetting as ActionSetting$imported } from "../../src/action.js";
import { ShadowFormat as ShadowFormat$imported } from "../../src/dml/effect.js";
import { PP_PLACEHOLDER as PP_PLACEHOLDER$imported } from "../../src/enum/shapes.js";
import { BaseShapeElement as BaseShapeElement$imported } from "../../src/oxml/shapes/shared.js";
import { CT_TextBody as CT_TextBody$imported } from "../../src/oxml/text.js";
import { Shape as Shape$imported } from "../../src/shapes/autoshape.js";
import { BaseShape as BaseShape$imported, _PlaceholderFormat as _PlaceholderFormat$imported } from "../../src/shapes/base.js";
import { GraphicFrame as GraphicFrame$imported } from "../../src/shapes/graphfrm.js";
import { Subshape as Subshape$imported } from "../../src/shapes/index.js";
import { Picture as Picture$imported } from "../../src/shapes/picture.js";
import { BaseShapeFactory as BaseShapeFactory$imported, SlideShapes as SlideShapes$imported } from "../../src/shapes/shapetree.js";
import { aCNvPr, aCxnSp, aGraphicFrame, aGrpSp, aGrpSpPr, aPXfrm, aPic, anExt, anNvSpPr, anOff, anSp, anSpPr, anXfrm } from "../oxml/unitdata/shape.js";
import { element, xml } from "../unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const ActionSetting = ActionSetting$imported;
const ShadowFormat = ShadowFormat$imported;
const PP_PLACEHOLDER = PP_PLACEHOLDER$imported;
const BaseShapeElement = BaseShapeElement$imported;
const CT_TextBody = CT_TextBody$imported;
const Shape = Shape$imported;
const BaseShape = BaseShape$imported;
const _PlaceholderFormat = _PlaceholderFormat$imported;
const GraphicFrame = GraphicFrame$imported;
const Subshape = Subshape$imported;
const Picture = Picture$imported;
const BaseShapeFactory = BaseShapeFactory$imported;
const SlideShapes = SlideShapes$imported;
void $py;



const DescribeBaseShape: any = {
  fixtures: {
    provides_part: { args: [], fn: (self: any): any => {
        class FakeProvidesPart {
          get part(): any {
            throw NotImplementedError;
          }
        }
        return $call(FakeProvidesPart);
      } },
    click_action_fixture: { args: ["request", "ActionSetting_", "action_setting_"], params: ["p:sp/p:nvSpPr/p:cNvPr", "p:grpSp/p:nvGrpSpPr/p:cNvPr", "p:graphicFrame/p:nvGraphicFramePr/p:cNvPr", "p:cxnSp/p:nvCxnSpPr/p:cNvPr", "p:pic/p:nvPicPr/p:cNvPr"], fn: (self: any, request: any, ActionSetting_: any, action_setting_: any): any => {
        let cNvPr: any, shape: any, sp: any, sp_cxml: any;
        sp_cxml = request.param;
        sp = element(sp_cxml);
        cNvPr = $get(sp.xpath("//p:cNvPr"), 0);
        shape = $call(BaseShape, sp, null);
        return [shape, ActionSetting_, cNvPr, action_setting_];
      } },
    dimensions_get_fixture: { args: ["request", "width", "height"], params: [["sp", false], ["sp_with_ext", true], ["pic", false], ["pic_with_ext", true], ["graphicFrame", false], ["graphicFrame_with_ext", true], ["grpSp", false], ["grpSp_with_ext", true], ["cxnSp", false], ["cxnSp_with_ext", true]], fn: (self: any, request: any, width: any, height: any): any => {
        let expect_values: any, shape: any, shape_elm: any, shape_elm_fixt_name: any;
        [shape_elm_fixt_name, expect_values] = $list(request.param);
        shape_elm = request.getfixturevalue(shape_elm_fixt_name);
        shape = $call(BaseShape, shape_elm, null);
        if ($truthy(!$truthy(expect_values))) {
          const $t1 = null;
          width = $t1;
          height = $t1;
        }
        return [shape, width, height];
      } },
    dimensions_set_fixture: { args: ["request", "width", "height"], params: [["sp", "sp_with_ext"], ["pic", "pic_with_ext"], ["graphicFrame", "graphicFrame_with_ext"], ["grpSp", "grpSp_with_ext"], ["cxnSp", "cxnSp_with_ext"]], fn: (self: any, request: any, width: any, height: any): any => {
        let expected_elm_fixt_name: any, expected_xml: any, shape: any, start_elm: any, start_elm_fixt_name: any;
        [start_elm_fixt_name, expected_elm_fixt_name] = $list(request.param);
        start_elm = request.getfixturevalue(start_elm_fixt_name);
        shape = $call(BaseShape, start_elm, null);
        expected_xml = request.getfixturevalue(expected_elm_fixt_name).xml;
        return [shape, width, height, expected_xml];
      } },
    id_fixture: { args: ["request"], params: [["p:sp/p:nvSpPr/p:cNvPr{id=1}", 1], ["p:cxnSp/p:nvCxnSpPr/p:cNvPr{id=2}", 2], ["p:graphicFrame/p:nvGraphicFramePr/p:cNvPr{id=3}", 3], ["p:grpSp/p:nvGrpSpPr/p:cNvPr{id=4}", 4], ["p:pic/p:nvPicPr/p:cNvPr{id=5}", 5]], fn: (self: any, request: any): any => {
        let expected_value: any, shape: any, xSp_cxml: any;
        [xSp_cxml, expected_value] = $list(request.param);
        shape = $call(BaseShape, element(xSp_cxml), null);
        return [shape, expected_value];
      } },
    is_placeholder_fixture: { args: ["request", "shape_elm_", "txBody_"], params: [true, false], fn: (self: any, request: any, shape_elm_: any, txBody_: any): any => {
        let is_placeholder: any, shape: any;
        is_placeholder = request.param;
        shape_elm_.hasPhElm = is_placeholder;
        shape = $call(BaseShape, shape_elm_, null);
        return [shape, is_placeholder];
      } },
    name_get_fixture: { args: ["shape_name"], fn: (self: any, shape_name: any): any => {
        let shape: any, shape_elm: any;
        shape_elm = anSp().withNsdecls().withChild(anNvSpPr().withChild(aCNvPr().withName(shape_name))).element;
        shape = $call(BaseShape, shape_elm, null);
        return [shape, shape_name];
      } },
    name_set_fixture: { args: ["request"], params: [["p:sp/p:nvSpPr/p:cNvPr{id=1,name=foo}", Shape, "Shape1", "p:sp/p:nvSpPr/p:cNvPr{id=1,name=Shape1}"], ["p:grpSp/p:nvGrpSpPr/p:cNvPr{id=2,name=bar}", BaseShape, "Shape2", "p:grpSp/p:nvGrpSpPr/p:cNvPr{id=2,name=Shape2}"], ["p:graphicFrame/p:nvGraphicFramePr/p:cNvPr{id=3,name=baz}", GraphicFrame, "Shape3", "p:graphicFrame/p:nvGraphicFramePr/p:cNvPr{id=3,name=Shape3}"], ["p:cxnSp/p:nvCxnSpPr/p:cNvPr{id=4,name=boo}", BaseShape, "Shape4", "p:cxnSp/p:nvCxnSpPr/p:cNvPr{id=4,name=Shape4}"], ["p:pic/p:nvPicPr/p:cNvPr{id=5,name=far}", Picture, "Shape5", "p:pic/p:nvPicPr/p:cNvPr{id=5,name=Shape5}"]], fn: (self: any, request: any): any => {
        let ShapeCls: any, expected_xSp_cxml: any, expected_xml: any, new_value: any, shape: any, xSp_cxml: any;
        [xSp_cxml, ShapeCls, new_value, expected_xSp_cxml] = $list(request.param);
        shape = $call(ShapeCls, element(xSp_cxml), null);
        expected_xml = xml(expected_xSp_cxml);
        return [shape, new_value, expected_xml];
      } },
    part_fixture: { args: ["shapes_"], fn: (self: any, shapes_: any): any => {
        let parent_: any, shape: any;
        parent_ = shapes_;
        shape = $call(BaseShape, null, parent_);
        return [shape, parent_];
      } },
    phfmt_fixture: { args: ["_PlaceholderFormat_", "placeholder_format_"], fn: (self: any, _PlaceholderFormat_: any, placeholder_format_: any): any => {
        let ph: any, shape: any, sp: any;
        sp = element("p:sp/p:nvSpPr/p:nvPr/p:ph");
        ph = $get(sp.xpath("//p:ph"), 0);
        shape = $call(BaseShape, sp, null);
        return [shape, _PlaceholderFormat_, placeholder_format_, ph];
      } },
    phfmt_raise_fixture: { args: [], fn: (self: any): any => {
        return $call(BaseShape, element("p:sp/p:nvSpPr/p:nvPr"), null);
      } },
    position_set_fixture: { args: ["request", "left", "top"], params: [["sp", "sp_with_off"], ["pic", "pic_with_off"], ["graphicFrame", "graphicFrame_with_off"], ["grpSp", "grpSp_with_off"], ["cxnSp", "cxnSp_with_off"]], fn: (self: any, request: any, left: any, top: any): any => {
        let expected_elm_fixt_name: any, expected_xml: any, shape: any, start_elm: any, start_elm_fixt_name: any;
        [start_elm_fixt_name, expected_elm_fixt_name] = $list(request.param);
        start_elm = request.getfixturevalue(start_elm_fixt_name);
        shape = $call(BaseShape, start_elm, null);
        expected_xml = request.getfixturevalue(expected_elm_fixt_name).xml;
        return [shape, left, top, expected_xml];
      } },
    rotation_get_fixture: { args: ["request"], params: [["p:sp/p:spPr", 0.0], ["p:sp/p:spPr/a:xfrm{rot=60000}", 1.0], ["p:sp/p:spPr/a:xfrm{rot=2545200}", 42.42], ["p:sp/p:spPr/a:xfrm{rot=-60000}", 359.0], ["p:grpSp/p:grpSpPr/a:xfrm{rot=2545200}", 42.42]], fn: (self: any, request: any): any => {
        let expected_value: any, shape: any, xSp_cxml: any;
        [xSp_cxml, expected_value] = $list(request.param);
        shape = $call(BaseShapeFactory, element(xSp_cxml), null);
        return [shape, expected_value];
      } },
    rotation_set_fixture: { args: ["request"], params: [["p:sp/p:spPr/a:xfrm", 1.0, "p:sp/p:spPr/a:xfrm{rot=60000}"], ["p:sp/p:spPr/a:xfrm{rot=60000}", 0.0, "p:sp/p:spPr/a:xfrm"], ["p:sp/p:spPr/a:xfrm{rot=60000}", -420.0, "p:sp/p:spPr/a:xfrm{rot=18000000}"], ["p:grpSp/p:grpSpPr/a:xfrm", 1.0, "p:grpSp/p:grpSpPr/a:xfrm{rot=60000}"]], fn: (self: any, request: any): any => {
        let expected_xSp_cxml: any, expected_xml: any, new_value: any, shape: any, xSp_cxml: any;
        [xSp_cxml, new_value, expected_xSp_cxml] = $list(request.param);
        shape = $call(BaseShapeFactory, element(xSp_cxml), null);
        expected_xml = xml(expected_xSp_cxml);
        return [shape, new_value, expected_xml];
      } },
    shadow_fixture: { args: ["request", "ShadowFormat_", "shadow_"], params: ["p:sp/p:spPr", "p:cxnSp/p:spPr", "p:pic/p:spPr"], fn: (self: any, request: any, ShadowFormat_: any, shadow_: any): any => {
        let shape: any, sp: any, spPr: any, sp_cxml: any;
        sp_cxml = request.param;
        sp = element(sp_cxml);
        spPr = $get(sp.xpath("//p:spPr"), 0);
        ShadowFormat_.returnValue = shadow_;
        shape = $call(BaseShape, sp, null);
        return [shape, ShadowFormat_, spPr, shadow_];
      } },
    ActionSetting_: { args: ["request", "action_setting_"], fn: (self: any, request: any, action_setting_: any): any => {
        return $call(classMock, request, "pptx.shapes.base.ActionSetting", $kw({return_value: action_setting_}));
      } },
    action_setting_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, ActionSetting);
      } },
    cxnSp: { args: [], fn: (self: any): any => {
        return aCxnSp().withNsdecls().withChild(anSpPr()).element;
      } },
    cxnSp_with_ext: { args: ["width", "height"], fn: (self: any, width: any, height: any): any => {
        return aCxnSp().withNsdecls().withChild(anSpPr().withChild(anXfrm().withChild(anExt().withCx(width).withCy(height)))).element;
      } },
    cxnSp_with_off: { args: ["left", "top"], fn: (self: any, left: any, top: any): any => {
        return aCxnSp().withNsdecls().withChild(anSpPr().withChild(anXfrm().withChild(anOff().withX(left).withY(top)))).element;
      } },
    graphicFrame: { args: [], fn: (self: any): any => {
        return aGraphicFrame().withNsdecls().withChild(aPXfrm()).element;
      } },
    graphicFrame_with_ext: { args: ["width", "height"], fn: (self: any, width: any, height: any): any => {
        return aGraphicFrame().withNsdecls().withChild(aPXfrm().withChild(anExt().withCx(width).withCy(height))).element;
      } },
    graphicFrame_with_off: { args: ["left", "top"], fn: (self: any, left: any, top: any): any => {
        return aGraphicFrame().withNsdecls().withChild(aPXfrm().withChild(anOff().withX(left).withY(top))).element;
      } },
    grpSp: { args: [], fn: (self: any): any => {
        return aGrpSp().withNsdecls("p", "a").withChild(aGrpSpPr()).element;
      } },
    grpSp_with_ext: { args: ["width", "height"], fn: (self: any, width: any, height: any): any => {
        return aGrpSp().withNsdecls("p", "a").withChild(aGrpSpPr().withChild(anXfrm().withChild(anExt().withCx(width).withCy(height)))).element;
      } },
    grpSp_with_off: { args: ["left", "top"], fn: (self: any, left: any, top: any): any => {
        return aGrpSp().withNsdecls("p", "a").withChild(aGrpSpPr().withChild(anXfrm().withChild(anOff().withX(left).withY(top)))).element;
      } },
    height: { args: [], fn: (self: any): any => {
        return 654;
      } },
    left: { args: [], fn: (self: any): any => {
        return 123;
      } },
    pic: { args: [], fn: (self: any): any => {
        return aPic().withNsdecls().withChild(anSpPr()).element;
      } },
    pic_with_off: { args: ["left", "top"], fn: (self: any, left: any, top: any): any => {
        return aPic().withNsdecls().withChild(anSpPr().withChild(anXfrm().withChild(anOff().withX(left).withY(top)))).element;
      } },
    pic_with_ext: { args: ["width", "height"], fn: (self: any, width: any, height: any): any => {
        return aPic().withNsdecls().withChild(anSpPr().withChild(anXfrm().withChild(anExt().withCx(width).withCy(height)))).element;
      } },
    _PlaceholderFormat_: { args: ["request", "placeholder_format_"], fn: (self: any, request: any, placeholder_format_: any): any => {
        return $call(classMock, request, "pptx.shapes.base._PlaceholderFormat", $kw({return_value: placeholder_format_}));
      } },
    placeholder_format_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _PlaceholderFormat);
      } },
    shadow_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, ShadowFormat);
      } },
    ShadowFormat_: { args: ["request"], fn: (self: any, request: any): any => {
        return classMock(request, "pptx.shapes.base.ShadowFormat");
      } },
    shape_elm_: { args: ["request", "shape_id", "shape_name", "txBody_"], fn: (self: any, request: any, shape_id: any, shape_name: any, txBody_: any): any => {
        return $call(instanceMock, request, BaseShapeElement, $kw({shape_id: shape_id, shape_name: shape_name, txBody: txBody_}));
      } },
    shape_id: { args: [], fn: (self: any): any => {
        return 42;
      } },
    shape_name: { args: [], fn: (self: any): any => {
        return "Foobar 41";
      } },
    shapes_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlideShapes);
      } },
    sp: { args: [], fn: (self: any): any => {
        return anSp().withNsdecls().withChild(anSpPr()).element;
      } },
    sp_with_ext: { args: ["width", "height"], fn: (self: any, width: any, height: any): any => {
        return anSp().withNsdecls().withChild(anSpPr().withChild(anXfrm().withChild(anExt().withCx(width).withCy(height)))).element;
      } },
    sp_with_off: { args: ["left", "top"], fn: (self: any, left: any, top: any): any => {
        return anSp().withNsdecls().withChild(anSpPr().withChild(anXfrm().withChild(anOff().withX(left).withY(top)))).element;
      } },
    top: { args: [], fn: (self: any): any => {
        return 456;
      } },
    txBody_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, CT_TextBody);
      } },
    width: { args: [], fn: (self: any): any => {
        return 321;
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_access_to_its_click_action", args: ["click_action_fixture"], fn: (self: any, click_action_fixture: any): any => {
        let ActionSetting_: any, cNvPr: any, click_action: any, click_action_: any, shape: any;
        [shape, ActionSetting_, cNvPr, click_action_] = $list(click_action_fixture);
        click_action = shape.clickAction;
        ActionSetting_.assertCalledOnceWith(cNvPr, shape);
        $assert($is(click_action, click_action_));
      } },
    { name: "it_knows_its_shape_id", args: ["id_fixture"], fn: (self: any, id_fixture: any): any => {
        let expected_value: any, shape: any;
        [shape, expected_value] = $list(id_fixture);
        $assertEqual(shape.shapeId, expected_value);
      } },
    { name: "it_knows_its_name", args: ["name_get_fixture"], fn: (self: any, name_get_fixture: any): any => {
        let name: any, shape: any;
        [shape, name] = $list(name_get_fixture);
        $assertEqual(shape.name, name);
      } },
    { name: "it_can_change_its_name", args: ["name_set_fixture"], fn: (self: any, name_set_fixture: any): any => {
        let expected_xml: any, new_value: any, shape: any;
        [shape, new_value, expected_xml] = $list(name_set_fixture);
        shape.name = new_value;
        $assertEqual(shape._element.xml, expected_xml);
      } },
    { name: "it_has_a_position", args: ["shape_cxml", "expected_x", "expected_y", "provides_part"], parametrize: [{ names: ["shape_cxml", "expected_x", "expected_y"], values: [["p:cxnSp/p:spPr", null, null], ["p:cxnSp/p:spPr/a:xfrm", null, null], ["p:cxnSp/p:spPr/a:xfrm/a:off{x=123,y=456}", 123, 456], ["p:graphicFrame/p:xfrm", null, null], ["p:graphicFrame/p:xfrm/a:off{x=123,y=456}", 123, 456], ["p:grpSp/p:grpSpPr", null, null], ["p:grpSp/p:grpSpPr/a:xfrm/a:off{x=123,y=456}", 123, 456], ["p:pic/p:spPr", null, null], ["p:pic/p:spPr/a:xfrm", null, null], ["p:pic/p:spPr/a:xfrm/a:off{x=123,y=456}", 123, 456], ["p:sp/p:spPr", null, null], ["p:sp/p:spPr/a:xfrm", null, null], ["p:sp/p:spPr/a:xfrm/a:off{x=123,y=456}", 123, 456]] }], fn: (self: any, shape_cxml: any, expected_x: any, expected_y: any, provides_part: any): any => {
        let shape: any, shape_elm: any;
        shape_elm = $cast("ShapeElement", element(shape_cxml));
        shape = $call(BaseShape, shape_elm, provides_part);
        $assertEqual(shape.left, expected_x);
        $assertEqual(shape.top, expected_y);
      } },
    { name: "it_can_change_its_position", args: ["position_set_fixture"], fn: (self: any, position_set_fixture: any): any => {
        let expected_xml: any, left: any, shape: any, top: any;
        [shape, left, top, expected_xml] = $list(position_set_fixture);
        shape.left = left;
        shape.top = top;
        $assertEqual(shape._element.xml, expected_xml);
      } },
    { name: "it_has_dimensions", args: ["dimensions_get_fixture"], fn: (self: any, dimensions_get_fixture: any): any => {
        let expected_height: any, expected_width: any, shape: any;
        [shape, expected_width, expected_height] = $list(dimensions_get_fixture);
        $assertEqual(shape.width, expected_width);
        $assertEqual(shape.height, expected_height);
      } },
    { name: "it_can_change_its_dimensions", args: ["dimensions_set_fixture"], fn: (self: any, dimensions_set_fixture: any): any => {
        let expected_xml: any, height: any, shape: any, width: any;
        [shape, width, height, expected_xml] = $list(dimensions_set_fixture);
        shape.width = width;
        shape.height = height;
        $assertEqual(shape._element.xml, expected_xml);
      } },
    { name: "it_knows_its_rotation_angle", args: ["rotation_get_fixture"], fn: (self: any, rotation_get_fixture: any): any => {
        let expected_value: any, shape: any;
        [shape, expected_value] = $list(rotation_get_fixture);
        $assertEqual(shape.rotation, expected_value);
      } },
    { name: "it_can_change_its_rotation_angle", args: ["rotation_set_fixture"], fn: (self: any, rotation_set_fixture: any): any => {
        let expected_xml: any, new_value: any, shape: any;
        [shape, new_value, expected_xml] = $list(rotation_set_fixture);
        shape.rotation = new_value;
        $assertEqual(shape._element.xml, expected_xml);
      } },
    { name: "it_provides_access_to_its_shadow", args: ["shadow_fixture"], fn: (self: any, shadow_fixture: any): any => {
        let ShadowFormat_: any, shadow: any, shadow_: any, shape: any, spPr: any;
        [shape, ShadowFormat_, spPr, shadow_] = $list(shadow_fixture);
        shadow = shape.shadow;
        ShadowFormat_.assertCalledOnceWith(spPr);
        $assert($is(shadow, shadow_));
      } },
    { name: "it_knows_the_part_it_belongs_to", args: ["part_fixture"], fn: (self: any, part_fixture: any): any => {
        let parent_: any, part: any, shape: any;
        [shape, parent_] = $list(part_fixture);
        part = shape.part;
        $assert($is(part, parent_.part));
      } },
    { name: "it_knows_it_doesnt_have_a_text_frame", args: [], fn: (self: any): any => {
        let shape: any;
        shape = $call(BaseShape, null, null);
        $assert($is(shape.hasTextFrame, false));
      } },
    { name: "it_knows_whether_it_is_a_placeholder", args: ["is_placeholder_fixture"], fn: (self: any, is_placeholder_fixture: any): any => {
        let is_placeholder: any, shape: any;
        [shape, is_placeholder] = $list(is_placeholder_fixture);
        $assert($is(shape.isPlaceholder, is_placeholder));
      } },
    { name: "it_provides_access_to_its_placeholder_format", args: ["phfmt_fixture"], fn: (self: any, phfmt_fixture: any): any => {
        let _PlaceholderFormat_: any, ph: any, placeholder_format: any, placeholder_format_: any, shape: any;
        [shape, _PlaceholderFormat_, placeholder_format_, ph] = $list(phfmt_fixture);
        placeholder_format = shape.placeholderFormat;
        _PlaceholderFormat_.assertCalledOnceWith(ph);
        $assert($is(placeholder_format, placeholder_format_));
      } },
    { name: "it_raises_when_shape_is_not_a_placeholder", args: ["phfmt_raise_fixture"], fn: (self: any, phfmt_raise_fixture: any): any => {
        let shape: any;
        shape = phfmt_raise_fixture;
        $raises(ValueError, () => {
          shape.placeholderFormat;
        });
      } },
    { name: "it_knows_it_doesnt_contain_a_chart", args: [], fn: (self: any): any => {
        let shape: any;
        shape = $call(BaseShape, null, null);
        $assert($is(shape.hasChart, false));
      } },
    { name: "it_knows_it_doesnt_contain_a_table", args: [], fn: (self: any): any => {
        let shape: any;
        shape = $call(BaseShape, null, null);
        $assert($is(shape.hasTable, false));
      } },
  ],
};
pytestClass("DescribeBaseShape", DescribeBaseShape);

const DescribeSubshape: any = {
  fixtures: {
    subshape_with_parent_: { args: ["request"], fn: (self: any, request: any): any => {
        let parent_: any, subshape: any;
        parent_ = $call(looseMock, request, $kw({name: "parent_"}));
        subshape = $call(Subshape, parent_);
        return [subshape, parent_];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_the_part_it_belongs_to", args: ["subshape_with_parent_"], fn: (self: any, subshape_with_parent_: any): any => {
        let parent_: any, part: any, subshape: any;
        [subshape, parent_] = $list(subshape_with_parent_);
        part = subshape.part;
        $assert($is(part, parent_.part));
      } },
  ],
};
pytestClass("DescribeSubshape", DescribeSubshape);

const Describe_PlaceholderFormat: any = {
  fixtures: {
    idx_get_fixture: { args: ["request"], params: [["p:ph", 0], ["p:ph{idx=42}", 42]], fn: (self: any, request: any): any => {
        let expected_value: any, ph_cxml: any, placeholder_format: any;
        [ph_cxml, expected_value] = $list(request.param);
        placeholder_format = $call(_PlaceholderFormat, element(ph_cxml));
        return [placeholder_format, expected_value];
      } },
    type_get_fixture: { args: ["request"], params: [["p:ph", PP_PLACEHOLDER.OBJECT], ["p:ph{type=pic}", PP_PLACEHOLDER.PICTURE]], fn: (self: any, request: any): any => {
        let expected_value: any, ph_cxml: any, placeholder_format: any;
        [ph_cxml, expected_value] = $list(request.param);
        placeholder_format = $call(_PlaceholderFormat, element(ph_cxml));
        return [placeholder_format, expected_value];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_idx", args: ["idx_get_fixture"], fn: (self: any, idx_get_fixture: any): any => {
        let expected_value: any, placeholder_format: any;
        [placeholder_format, expected_value] = $list(idx_get_fixture);
        $assertEqual(placeholder_format.idx, expected_value);
      } },
    { name: "it_knows_its_type", args: ["type_get_fixture"], fn: (self: any, type_get_fixture: any): any => {
        let expected_value: any, placeholder_format: any;
        [placeholder_format, expected_value] = $list(type_get_fixture);
        $assertEqual(placeholder_format.type, expected_value);
      } },
  ],
};
pytestClass("Describe_PlaceholderFormat", Describe_PlaceholderFormat);

