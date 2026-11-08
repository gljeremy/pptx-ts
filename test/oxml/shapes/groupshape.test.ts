// Machine-translated from python-pptx tests/oxml/shapes/test_groupshape.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $assert, $assertEqual, $get, $is, $list, $slice, $t1, call, classMock, instanceMock, methodMock, propertyMock, pytestClass } from "../../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../../src/exc.js";
import { defineChildren } from "../../../src/oxml/xmlchemy.js";
import { CT_Shape as CT_Shape$imported } from "../../../src/oxml/shapes/autoshape.js";
import { CT_GraphicalObjectFrame as CT_GraphicalObjectFrame$imported } from "../../../src/oxml/shapes/graphfrm.js";
import { CT_GroupShape as CT_GroupShape$imported } from "../../../src/oxml/shapes/groupshape.js";
import { CT_Picture as CT_Picture$imported } from "../../../src/oxml/shapes/picture.js";
import { element, xml } from "../../unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const CT_Shape = CT_Shape$imported;
const CT_GraphicalObjectFrame = CT_GraphicalObjectFrame$imported;
const CT_GroupShape = CT_GroupShape$imported;
const CT_Picture = CT_Picture$imported;
void $py;


const DescribeCT_GroupShape: any = {
  fixtures: {
    add_autoshape_fixt: { args: ["spTree", "CT_Shape_", "insert_element_before_", "sp_"], fn: (self: any, spTree: any, CT_Shape_: any, insert_element_before_: any, sp_: any): any => {
        let cx: any, cy: any, id_: any, name: any, prst: any, x: any, y: any;
        [id_, name, prst] = $list([42, "name", "prst"]);
        [x, y, cx, cy] = $list([9, 8, 7, 6]);
        return [spTree, id_, name, prst, x, y, cx, cy, CT_Shape_, insert_element_before_, sp_];
      } },
    add_grpSp_fixture: { args: [], fn: (self: any): any => {
        let expected_grpSp_xml: any, expected_xml: any, spTree: any;
        spTree = element("p:spTree{a:b=c,r:s=t}");
        expected_grpSp_xml = "<p:grpSp xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n  <p:nvGrpSpPr>\n    <p:cNvPr id=\"1\" name=\"Group 0\"/>\n    <p:cNvGrpSpPr/>\n    <p:nvPr/>\n  </p:nvGrpSpPr>\n  <p:grpSpPr>\n    <a:xfrm>\n      <a:off x=\"0\" y=\"0\"/>\n      <a:ext cx=\"0\" cy=\"0\"/>\n      <a:chOff x=\"0\" y=\"0\"/>\n      <a:chExt cx=\"0\" cy=\"0\"/>\n    </a:xfrm>\n  </p:grpSpPr>\n</p:grpSp>\n";
        expected_xml = xml("p:spTree{a:b=c,r:s=t}/p:grpSp/(p:nvGrpSpPr/(p:cNvPr{id=1,name=Group 0},p:cNvGrpSpPr,p:nvPr),p:grpSpPr/a:xfrm/(a:off{x=0,y=0},a:ext{cx=0,cy=0},a:chOff{x=0,y=0},a:chExt{cx=0,cy=0}))");
        return [spTree, expected_grpSp_xml, expected_xml];
      } },
    add_pic_fixt: { args: ["spTree", "CT_Picture_", "insert_element_before_", "pic_"], fn: (self: any, spTree: any, CT_Picture_: any, insert_element_before_: any, pic_: any): any => {
        let cx: any, cy: any, desc: any, id_: any, name: any, rId: any, x: any, y: any;
        [id_, name, desc, rId] = $list([42, "name", "desc", "rId6"]);
        [x, y, cx, cy] = $list([6, 7, 8, 9]);
        return [spTree, id_, name, desc, rId, x, y, cx, cy, CT_Picture_, insert_element_before_, pic_];
      } },
    add_placeholder_fixt: { args: ["spTree", "CT_Shape_", "insert_element_before_", "sp_"], fn: (self: any, spTree: any, CT_Shape_: any, insert_element_before_: any, sp_: any): any => {
        let id_: any, idx: any, name: any, orient: any, ph_type: any, sz: any;
        [id_, name, ph_type] = $list([42, "name", "type"]);
        [orient, sz, idx] = $list(["orient", "sz", 24]);
        return [spTree, id_, name, ph_type, orient, sz, idx, CT_Shape_, insert_element_before_, sp_];
      } },
    add_table_fixt: { args: ["spTree", "CT_GraphicalObjectFrame_", "insert_element_before_", "graphicFrame_"], fn: (self: any, spTree: any, CT_GraphicalObjectFrame_: any, insert_element_before_: any, graphicFrame_: any): any => {
        let cols: any, cx: any, cy: any, id_: any, name: any, new_table_graphicFrame_: any, rows: any, x: any, y: any;
        [id_, name, rows, cols] = $list([42, "name", 12, 23]);
        [x, y, cx, cy] = $list([5, 4, 3, 2]);
        new_table_graphicFrame_ = CT_GraphicalObjectFrame_.newTableGraphicFrame;
        return [spTree, id_, name, rows, cols, x, y, cx, cy, new_table_graphicFrame_, insert_element_before_, graphicFrame_];
      } },
    add_textbox_fixt: { args: ["spTree", "CT_Shape_", "insert_element_before_", "sp_"], fn: (self: any, spTree: any, CT_Shape_: any, insert_element_before_: any, sp_: any): any => {
        let cx: any, cy: any, id_: any, name: any, x: any, y: any;
        [id_, name] = $list([42, "name"]);
        [x, y, cx, cy] = $list([3, 4, 5, 6]);
        return [spTree, id_, name, x, y, cx, cy, CT_Shape_, insert_element_before_, sp_];
      } },
    child_exts_fixture: { args: ["request"], params: [["p:grpSp", [0, 0, 0, 0]], ["p:grpSp/p:sp/p:spPr/a:xfrm/(a:off{x=1,y=2},a:ext{cx=3,cy=4})", [1, 2, 3, 4]], ["p:grpSp/(p:sp/p:spPr/a:xfrm/(a:off{x=1,y=2},a:ext{cx=3,cy=4}),p:sp/p:spPr/a:xfrm/(a:off{x=10,y=20},a:ext{cx=30,cy=40}))", [1, 2, 39, 58]], ["p:grpSp/(p:sp/p:spPr/a:xfrm/(a:off{x=100,y=50},a:ext{cx=25,cy=25}),p:sp/p:spPr/a:xfrm/(a:off{x=50,y=100},a:ext{cx=25,cy=25}),p:sp/p:spPr/a:xfrm/(a:off{x=150,y=75},a:ext{cx=25,cy=25}))", [50, 50, 125, 75]]], fn: (self: any, request: any): any => {
        let expected_values: any, xSp: any, xSp_cxml: any;
        [xSp_cxml, expected_values] = $list(request.param);
        xSp = element(xSp_cxml);
        return [xSp, expected_values];
      } },
    recalc_fixture: { args: ["request", "_child_extents_prop_", "getparent_", "grpSp_"], params: [["p:spTree", null, [], "p:spTree"], ["p:grpSp/p:grpSpPr/a:xfrm", [1, 2, 3, 4], [call()], "p:grpSp/p:grpSpPr/a:xfrm/(a:off{x=1,y=2},a:ext{cx=3,cy=4},a:chOff{x=1,y=2},a:chExt{cx=3,cy=4})"]], fn: (self: any, request: any, _child_extents_prop_: any, getparent_: any, grpSp_: any): any => {
        let calls: any, expected_cxml: any, expected_xml: any, extents: any, parent_sp: any, xSp: any, xSp_cxml: any;
        [xSp_cxml, extents, calls, expected_cxml] = $list(request.param);
        xSp = element(xSp_cxml);
        _child_extents_prop_.returnValue = extents;
        expected_xml = xml(expected_cxml);
        const $t1 = grpSp_;
        getparent_.returnValue = $t1;
        parent_sp = $t1;
        return [xSp, expected_xml, parent_sp, calls];
      } },
    _child_extents_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, CT_GroupShape, "_child_extents");
      } },
    CT_GraphicalObjectFrame_: { args: ["request", "graphicFrame_"], fn: (self: any, request: any, graphicFrame_: any): any => {
        let CT_GraphicalObjectFrame_: any;
        CT_GraphicalObjectFrame_ = classMock(request, "pptx.oxml.shapes.groupshape.CT_GraphicalObjectFrame");
        CT_GraphicalObjectFrame_.newTableGraphicFrame.returnValue = graphicFrame_;
        return CT_GraphicalObjectFrame_;
      } },
    CT_Picture_: { args: ["request", "pic_"], fn: (self: any, request: any, pic_: any): any => {
        let CT_Picture_: any;
        CT_Picture_ = classMock(request, "pptx.oxml.shapes.groupshape.CT_Picture");
        CT_Picture_.newPic.returnValue = pic_;
        return CT_Picture_;
      } },
    CT_Shape_: { args: ["request", "sp_"], fn: (self: any, request: any, sp_: any): any => {
        let CT_Shape_: any;
        CT_Shape_ = classMock(request, "pptx.oxml.shapes.groupshape.CT_Shape");
        CT_Shape_.newAutoshapeSp.returnValue = sp_;
        CT_Shape_.newPlaceholderSp.returnValue = sp_;
        CT_Shape_.newTextboxSp.returnValue = sp_;
        return CT_Shape_;
      } },
    getparent_: { args: ["request"], fn: (self: any, request: any): any => {
        return methodMock(request, CT_GroupShape, "getparent");
      } },
    graphicFrame_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, CT_GraphicalObjectFrame);
      } },
    grpSp_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, CT_GroupShape);
      } },
    insert_element_before_: { args: ["request"], fn: (self: any, request: any): any => {
        return methodMock(request, CT_GroupShape, "insert_element_before");
      } },
    pic_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, CT_Picture);
      } },
    sp_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, CT_Shape);
      } },
    spTree: { args: [], fn: (self: any): any => {
        return element("p:spTree");
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_add_a_graphicFrame_element_containing_a_table", args: ["add_table_fixt"], fn: (self: any, add_table_fixt: any): any => {
        let cols: any, cx: any, cy: any, graphicFrame: any, graphicFrame_: any, id_: any, insert_element_before_: any, name: any, new_table_graphicFrame_: any, rows: any, spTree: any, x: any, y: any;
        [spTree, id_, name, rows, cols, x, y, cx, cy] = $list($slice(add_table_fixt, null, 9, null));
        new_table_graphicFrame_ = $get(add_table_fixt, 9);
        [insert_element_before_, graphicFrame_] = $list($slice(add_table_fixt, 10, null, null));
        graphicFrame = spTree.addTable(id_, name, rows, cols, x, y, cx, cy);
        new_table_graphicFrame_.assertCalledOnceWith(id_, name, rows, cols, x, y, cx, cy);
        insert_element_before_.assertCalledOnceWith(spTree, graphicFrame_, "p:extLst");
        $assert($is(graphicFrame, graphicFrame_));
      } },
    { name: "it_can_add_a_grpSp_element", args: ["add_grpSp_fixture"], fn: (self: any, add_grpSp_fixture: any): any => {
        let expected_grpSp_xml: any, expected_xml: any, grpSp: any, spTree: any;
        [spTree, expected_grpSp_xml, expected_xml] = $list(add_grpSp_fixture);
        grpSp = spTree.addGrpSp();
        $assertEqual(grpSp.xml, expected_grpSp_xml);
        $assertEqual(spTree.xml, expected_xml);
      } },
    { name: "it_can_add_a_pic_element_representing_a_picture", args: ["add_pic_fixt"], fn: (self: any, add_pic_fixt: any): any => {
        let CT_Picture_: any, cx: any, cy: any, desc: any, id_: any, insert_element_before_: any, name: any, pic: any, pic_: any, rId: any, spTree: any, x: any, y: any;
        [spTree, id_, name, desc, rId, x, y, cx, cy] = $list($slice(add_pic_fixt, null, 9, null));
        [CT_Picture_, insert_element_before_, pic_] = $list($slice(add_pic_fixt, 9, null, null));
        pic = spTree.addPic(id_, name, desc, rId, x, y, cx, cy);
        CT_Picture_.newPic.assertCalledOnceWith(id_, name, desc, rId, x, y, cx, cy);
        insert_element_before_.assertCalledOnceWith(spTree, pic_, "p:extLst");
        $assert($is(pic, pic_));
      } },
    { name: "it_can_add_an_sp_element_for_a_placeholder", args: ["add_placeholder_fixt"], fn: (self: any, add_placeholder_fixt: any): any => {
        let CT_Shape_: any, id_: any, idx: any, insert_element_before_: any, name: any, orient: any, ph_type: any, sp: any, spTree: any, sp_: any, sz: any;
        [spTree, id_, name, ph_type, orient, sz, idx] = $list($slice(add_placeholder_fixt, null, 7, null));
        [CT_Shape_, insert_element_before_, sp_] = $list($slice(add_placeholder_fixt, 7, null, null));
        sp = spTree.addPlaceholder(id_, name, ph_type, orient, sz, idx);
        CT_Shape_.newPlaceholderSp.assertCalledOnceWith(id_, name, ph_type, orient, sz, idx);
        insert_element_before_.assertCalledOnceWith(spTree, sp_, "p:extLst");
        $assert($is(sp, sp_));
      } },
    { name: "it_can_add_an_sp_element_for_an_autoshape", args: ["add_autoshape_fixt"], fn: (self: any, add_autoshape_fixt: any): any => {
        let CT_Shape_: any, cx: any, cy: any, id_: any, insert_element_before_: any, name: any, prst: any, sp: any, spTree: any, sp_: any, x: any, y: any;
        [spTree, id_, name, prst, x, y, cx, cy] = $list($slice(add_autoshape_fixt, null, 8, null));
        [CT_Shape_, insert_element_before_, sp_] = $list($slice(add_autoshape_fixt, 8, null, null));
        sp = spTree.addAutoshape(id_, name, prst, x, y, cx, cy);
        CT_Shape_.newAutoshapeSp.assertCalledOnceWith(id_, name, prst, x, y, cx, cy);
        insert_element_before_.assertCalledOnceWith(spTree, sp_, "p:extLst");
        $assert($is(sp, sp_));
      } },
    { name: "it_can_add_a_textbox_sp_element", args: ["add_textbox_fixt"], fn: (self: any, add_textbox_fixt: any): any => {
        let CT_Shape_: any, cx: any, cy: any, id_: any, insert_element_before_: any, name: any, sp: any, spTree: any, sp_: any, x: any, y: any;
        [spTree, id_, name, x, y, cx, cy, CT_Shape_] = $list($slice(add_textbox_fixt, null, 8, null));
        [insert_element_before_, sp_] = $list($slice(add_textbox_fixt, 8, null, null));
        sp = spTree.addTextbox(id_, name, x, y, cx, cy);
        CT_Shape_.newTextboxSp.assertCalledOnceWith(id_, name, x, y, cx, cy);
        insert_element_before_.assertCalledOnceWith(spTree, sp_, "p:extLst");
        $assert($is(sp, sp_));
      } },
    { name: "it_can_recalculate_its_pos_and_size", args: ["recalc_fixture"], fn: (self: any, recalc_fixture: any): any => {
        let calls: any, expected_xml: any, parent_sp: any, xSp: any;
        [xSp, expected_xml, parent_sp, calls] = $list(recalc_fixture);
        xSp.recalculateExtents();
        $assertEqual(xSp.xml, expected_xml);
        $assertEqual(parent_sp.recalculateExtents.callArgsList, calls);
      } },
    { name: "it_calculates_its_child_extents_to_help", args: ["child_exts_fixture"], fn: (self: any, child_exts_fixture: any): any => {
        let cx: any, cy: any, expected_values: any, x: any, xSp: any, y: any;
        [xSp, expected_values] = $list(child_exts_fixture);
        [x, y, cx, cy] = $list(xSp._childExtents);
        $assertEqual([x, y, cx, cy], expected_values);
      } },
  ],
};
pytestClass("DescribeCT_GroupShape", DescribeCT_GroupShape);

