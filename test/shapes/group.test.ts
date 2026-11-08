// Machine-translated from python-pptx tests/shapes/test_group.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $call, $get, $is, $isinstance, $kw, $list, $raises, classMock, initializerMock, instanceMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { ShadowFormat as ShadowFormat$imported } from "../../src/dml/effect.js";
import { MSO_SHAPE_TYPE as MSO_SHAPE_TYPE$imported } from "../../src/enum/shapes.js";
import { GroupShape as GroupShape$imported } from "../../src/shapes/group.js";
import { GroupShapes } from "../../src/shapes/shapetree.js";
import { element } from "../unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const ShadowFormat = ShadowFormat$imported;
const MSO_SHAPE_TYPE = MSO_SHAPE_TYPE$imported;
const GroupShape = GroupShape$imported;
void $py;


const DescribeGroupShape: any = {
  fixtures: {
    click_action_fixture: { args: [], fn: (self: any): any => {
        return $call(GroupShape, null, null);
      } },
    shape_type_fixture: { args: [], fn: (self: any): any => {
        return $call(GroupShape, null, null);
      } },
    shapes_fixture: { args: ["GroupShapes_init_"], fn: (self: any, GroupShapes_init_: any): any => {
        let group: any, grpSp: any;
        grpSp = element("p:grpSp");
        group = $call(GroupShape, grpSp, null);
        return [group, GroupShapes_init_, grpSp];
      } },
    GroupShapes_init_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(initializerMock, request, GroupShapes, $kw({autospec: true}));
      } },
    shadow_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, ShadowFormat);
      } },
    ShadowFormat_: { args: ["request"], fn: (self: any, request: any): any => {
        return classMock(request, "pptx.shapes.group.ShadowFormat");
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_raises_on_access_click_action", args: ["click_action_fixture"], fn: (self: any, click_action_fixture: any): any => {
        let group: any;
        group = click_action_fixture;
        $raises(TypeError, () => {
          group.clickAction;
        });
      } },
    { name: "it_provides_access_to_its_shadow", args: ["ShadowFormat_", "shadow_"], fn: (self: any, ShadowFormat_: any, shadow_: any): any => {
        let group_shape: any, grpSp: any, grpSpPr: any, shadow: any;
        grpSp = element("p:grpSp/p:grpSpPr");
        grpSpPr = $get(grpSp.xpath("//p:grpSpPr"), 0);
        ShadowFormat_.returnValue = shadow_;
        group_shape = $call(GroupShape, grpSp, null);
        shadow = group_shape.shadow;
        ShadowFormat_.assertCalledOnceWith(grpSpPr);
        $assert($is(shadow, shadow_));
      } },
    { name: "it_knows_its_shape_type", args: ["shape_type_fixture"], fn: (self: any, shape_type_fixture: any): any => {
        let group: any;
        group = shape_type_fixture;
        $assertEqual(group.shapeType, MSO_SHAPE_TYPE.GROUP);
      } },
    { name: "it_provides_access_to_its_sub_shapes", args: ["shapes_fixture"], fn: (self: any, shapes_fixture: any): any => {
        let GroupShapes_init_: any, group: any, grpSp: any, shapes: any;
        [group, GroupShapes_init_, grpSp] = $list(shapes_fixture);
        shapes = group.shapes;
        GroupShapes_init_.assertCalledOnceWith(shapes, grpSp, group);
        $assert($isinstance(shapes, GroupShapes));
      } },
  ],
};
pytestClass("DescribeGroupShape", DescribeGroupShape);

