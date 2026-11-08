// Machine-translated from python-pptx tests/shapes/test_connector.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $call, $isinstance, $kw, $list, $mod, instanceMock, methodMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { LineFormat as LineFormat$imported } from "../../src/dml/line.js";
import { MSO_SHAPE_TYPE as MSO_SHAPE_TYPE$imported } from "../../src/enum/shapes.js";
import { BaseShape as BaseShape$imported } from "../../src/shapes/base.js";
import { Connector as Connector$imported } from "../../src/shapes/connector.js";
import { Emu as Emu$imported } from "../../src/util.js";
import { element, xml } from "../unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const LineFormat = LineFormat$imported;
const MSO_SHAPE_TYPE = MSO_SHAPE_TYPE$imported;
const BaseShape = BaseShape$imported;
const Connector = Connector$imported;
const Emu = Emu$imported;
void $py;


const DescribeConnector: any = {
  fixtures: {
    begin_conn_fixture: { args: ["_connect_begin_to_", "_move_begin_to_cxn_"], fn: (self: any, _connect_begin_to_: any, _move_begin_to_cxn_: any): any => {
        let connector: any, cxn_idx: any, shape: any;
        connector = $call(Connector, null, null);
        [shape, cxn_idx] = $list([42, 24]);
        return [connector, shape, cxn_idx];
      } },
    begin_x_get_fixture: { args: ["request"], params: [[42, 24, false, 42], [24, 42, true, 66]], fn: (self: any, request: any): any => {
        let connector: any, cx: any, cxnSp: any, expected_value: any, flipH: any, x: any;
        [x, cx, flipH, expected_value] = $list(request.param);
        cxnSp = element($mod("p:cxnSp/p:spPr/a:xfrm{flipH=%d}/(a:off{x=%d,y=6},a:ext{cx=%d,cy=32})", [flipH, x, cx]));
        connector = $call(Connector, cxnSp, null);
        return [connector, expected_value];
      } },
    begin_x_set_fixture: { args: ["request"], params: [["a:xfrm/(a:off{x=10,y=1},a:ext{cx=10,cy=1})", 5, "a:xfrm/(a:off{x=5,y=1},a:ext{cx=15,cy=1})"], ["a:xfrm/(a:off{x=10,y=1},a:ext{cx=10,cy=1})", 15, "a:xfrm/(a:off{x=15,y=1},a:ext{cx=5,cy=1})"], ["a:xfrm/(a:off{x=10,y=1},a:ext{cx=10,cy=1})", 25, "a:xfrm{flipH=1}/(a:off{x=20,y=1},a:ext{cx=5,cy=1})"], ["a:xfrm{flipH=1}/(a:off{x=10,y=1},a:ext{cx=10,cy=1})", 25, "a:xfrm{flipH=1}/(a:off{x=10,y=1},a:ext{cx=15,cy=1})"], ["a:xfrm{flipH=1}/(a:off{x=10,y=1},a:ext{cx=10,cy=1})", 15, "a:xfrm{flipH=1}/(a:off{x=10,y=1},a:ext{cx=5,cy=1})"], ["a:xfrm{flipH=1}/(a:off{x=10,y=1},a:ext{cx=10,cy=1})", 5, "a:xfrm/(a:off{x=5,y=1},a:ext{cx=5,cy=1})"]], fn: (self: any, request: any): any => {
        let connector: any, cxnSp: any, expected_cxml: any, expected_xml: any, new_x: any, tmpl: any, xfrm_cxml: any;
        [xfrm_cxml, new_x, expected_cxml] = $list(request.param);
        tmpl = "p:cxnSp/p:spPr/%s";
        cxnSp = element($mod(tmpl, xfrm_cxml));
        expected_xml = xml($mod(tmpl, expected_cxml));
        connector = $call(Connector, cxnSp, null);
        return [connector, new_x, expected_xml];
      } },
    begin_y_get_fixture: { args: ["request"], params: [[40, 60, false, 40], [50, 42, true, 92]], fn: (self: any, request: any): any => {
        let connector: any, cxnSp: any, cy: any, expected_value: any, flipV: any, y: any;
        [y, cy, flipV, expected_value] = $list(request.param);
        cxnSp = element($mod("p:cxnSp/p:spPr/a:xfrm{flipV=%d}/(a:off{x=6,y=%d},a:ext{cx=32,cy=%d})", [flipV, y, cy]));
        connector = $call(Connector, cxnSp, null);
        return [connector, expected_value];
      } },
    begin_y_set_fixture: { args: ["request"], params: [["a:xfrm/(a:off{x=1,y=10},a:ext{cx=1,cy=10})", 5, "a:xfrm/(a:off{x=1,y=5},a:ext{cx=1,cy=15})"], ["a:xfrm/(a:off{x=1,y=10},a:ext{cx=1,cy=10})", 15, "a:xfrm/(a:off{x=1,y=15},a:ext{cx=1,cy=5})"], ["a:xfrm/(a:off{x=1,y=10},a:ext{cx=1,cy=10})", 25, "a:xfrm{flipV=1}/(a:off{x=1,y=20},a:ext{cx=1,cy=5})"], ["a:xfrm{flipV=1}/(a:off{x=1,y=10},a:ext{cx=1,cy=10})", 30, "a:xfrm{flipV=1}/(a:off{x=1,y=10},a:ext{cx=1,cy=20})"], ["a:xfrm{flipV=1}/(a:off{x=1,y=10},a:ext{cx=1,cy=10})", 15, "a:xfrm{flipV=1}/(a:off{x=1,y=10},a:ext{cx=1,cy=5})"], ["a:xfrm{flipV=1}/(a:off{x=1,y=10},a:ext{cx=1,cy=10})", 5, "a:xfrm/(a:off{x=1,y=5},a:ext{cx=1,cy=5})"]], fn: (self: any, request: any): any => {
        let connector: any, cxnSp: any, expected_cxml: any, expected_xml: any, new_y: any, tmpl: any, xfrm_cxml: any;
        [xfrm_cxml, new_y, expected_cxml] = $list(request.param);
        tmpl = "p:cxnSp/p:spPr/%s";
        cxnSp = element($mod(tmpl, xfrm_cxml));
        expected_xml = xml($mod(tmpl, expected_cxml));
        connector = $call(Connector, cxnSp, null);
        return [connector, new_y, expected_xml];
      } },
    connect_begin_fixture: { args: ["request", "shape_"], params: [["p:cxnSp{a:b=c}/p:nvCxnSpPr/p:cNvCxnSpPr", "p:cxnSp{a:b=c}/p:nvCxnSpPr/p:cNvCxnSpPr/a:stCxn{id=42,idx=3}"], ["p:cxnSp/p:nvCxnSpPr/p:cNvCxnSpPr/a:stCxn{id=66,idx=6}", "p:cxnSp/p:nvCxnSpPr/p:cNvCxnSpPr/a:stCxn{id=42,idx=3}"]], fn: (self: any, request: any, shape_: any): any => {
        let connector: any, cxnSp: any, cxnSp_cxml: any, cxn_idx: any, expected_cxml: any, expected_xml: any;
        [cxnSp_cxml, expected_cxml] = $list(request.param);
        cxnSp = element(cxnSp_cxml);
        connector = $call(Connector, cxnSp, null);
        [shape_.shapeId, cxn_idx] = $list([42, 3]);
        expected_xml = xml(expected_cxml);
        return [connector, shape_, cxn_idx, expected_xml];
      } },
    connect_end_fixture: { args: ["request", "shape_"], params: [["p:cxnSp{a:b=c}/p:nvCxnSpPr/p:cNvCxnSpPr", "p:cxnSp{a:b=c}/p:nvCxnSpPr/p:cNvCxnSpPr/a:endCxn{id=24,idx=2}"], ["p:cxnSp/p:nvCxnSpPr/p:cNvCxnSpPr/a:endCxn{id=66,idx=6}", "p:cxnSp/p:nvCxnSpPr/p:cNvCxnSpPr/a:endCxn{id=24,idx=2}"]], fn: (self: any, request: any, shape_: any): any => {
        let connector: any, cxnSp: any, cxnSp_cxml: any, cxn_idx: any, expected_cxml: any, expected_xml: any;
        [cxnSp_cxml, expected_cxml] = $list(request.param);
        cxnSp = element(cxnSp_cxml);
        connector = $call(Connector, cxnSp, null);
        [shape_.shapeId, cxn_idx] = $list([24, 2]);
        expected_xml = xml(expected_cxml);
        return [connector, shape_, cxn_idx, expected_xml];
      } },
    end_conn_fixture: { args: ["_connect_end_to_", "_move_end_to_cxn_"], fn: (self: any, _connect_end_to_: any, _move_end_to_cxn_: any): any => {
        let connector: any, cxn_idx: any, shape: any;
        connector = $call(Connector, null, null);
        [shape, cxn_idx] = $list([42, 24]);
        return [connector, shape, cxn_idx];
      } },
    end_x_get_fixture: { args: ["request"], params: [[21, 32, false, 53], [43, 54, true, 43]], fn: (self: any, request: any): any => {
        let connector: any, cx: any, cxnSp: any, expected_value: any, flipH: any, x: any;
        [x, cx, flipH, expected_value] = $list(request.param);
        cxnSp = element($mod("p:cxnSp/p:spPr/a:xfrm{flipH=%d}/(a:off{x=%d,y=6},a:ext{cx=%d,cy=60})", [flipH, x, cx]));
        connector = $call(Connector, cxnSp, null);
        return [connector, expected_value];
      } },
    end_x_set_fixture: { args: ["request"], params: [["a:xfrm/(a:off{x=10,y=1},a:ext{cx=10,cy=1})", 32, "a:xfrm/(a:off{x=10,y=1},a:ext{cx=22,cy=1})"], ["a:xfrm/(a:off{x=10,y=1},a:ext{cx=10,cy=1})", 15, "a:xfrm/(a:off{x=10,y=1},a:ext{cx=5,cy=1})"], ["a:xfrm/(a:off{x=10,y=1},a:ext{cx=10,cy=1})", 5, "a:xfrm{flipH=1}/(a:off{x=5,y=1},a:ext{cx=5,cy=1})"], ["a:xfrm{flipH=1}/(a:off{x=10,y=1},a:ext{cx=10,cy=1})", 5, "a:xfrm{flipH=1}/(a:off{x=5,y=1},a:ext{cx=15,cy=1})"], ["a:xfrm{flipH=1}/(a:off{x=10,y=1},a:ext{cx=10,cy=1})", 15, "a:xfrm{flipH=1}/(a:off{x=15,y=1},a:ext{cx=5,cy=1})"], ["a:xfrm{flipH=1}/(a:off{x=10,y=1},a:ext{cx=10,cy=1})", 28, "a:xfrm/(a:off{x=20,y=1},a:ext{cx=8,cy=1})"]], fn: (self: any, request: any): any => {
        let connector: any, cxnSp: any, expected_cxml: any, expected_xml: any, new_x: any, tmpl: any, xfrm_cxml: any;
        [xfrm_cxml, new_x, expected_cxml] = $list(request.param);
        tmpl = "p:cxnSp/p:spPr/%s";
        cxnSp = element($mod(tmpl, xfrm_cxml));
        expected_xml = xml($mod(tmpl, expected_cxml));
        connector = $call(Connector, cxnSp, null);
        return [connector, new_x, expected_xml];
      } },
    end_y_get_fixture: { args: ["request"], params: [[31, 42, false, 73], [53, 14, true, 53]], fn: (self: any, request: any): any => {
        let connector: any, cxnSp: any, cy: any, expected_value: any, flipV: any, y: any;
        [y, cy, flipV, expected_value] = $list(request.param);
        cxnSp = element($mod("p:cxnSp/p:spPr/a:xfrm{flipV=%d}/(a:off{x=6,y=%d},a:ext{cx=32,cy=%d})", [flipV, y, cy]));
        connector = $call(Connector, cxnSp, null);
        return [connector, expected_value];
      } },
    end_y_set_fixture: { args: ["request"], params: [["a:xfrm/(a:off{x=1,y=10},a:ext{cx=1,cy=10})", 28, "a:xfrm/(a:off{x=1,y=10},a:ext{cx=1,cy=18})"], ["a:xfrm/(a:off{x=1,y=10},a:ext{cx=1,cy=10})", 13, "a:xfrm/(a:off{x=1,y=10},a:ext{cx=1,cy=3})"], ["a:xfrm/(a:off{x=1,y=10},a:ext{cx=1,cy=10})", 4, "a:xfrm{flipV=1}/(a:off{x=1,y=4},a:ext{cx=1,cy=6})"], ["a:xfrm{flipV=1}/(a:off{x=1,y=10},a:ext{cx=1,cy=10})", 6, "a:xfrm{flipV=1}/(a:off{x=1,y=6},a:ext{cx=1,cy=14})"], ["a:xfrm{flipV=1}/(a:off{x=1,y=10},a:ext{cx=1,cy=10})", 12, "a:xfrm{flipV=1}/(a:off{x=1,y=12},a:ext{cx=1,cy=8})"], ["a:xfrm{flipV=1}/(a:off{x=1,y=10},a:ext{cx=1,cy=10})", 27, "a:xfrm/(a:off{x=1,y=20},a:ext{cx=1,cy=7})"]], fn: (self: any, request: any): any => {
        let connector: any, cxnSp: any, expected_cxml: any, expected_xml: any, new_y: any, tmpl: any, xfrm_cxml: any;
        [xfrm_cxml, new_y, expected_cxml] = $list(request.param);
        tmpl = "p:cxnSp/p:spPr/%s";
        cxnSp = element($mod(tmpl, xfrm_cxml));
        expected_xml = xml($mod(tmpl, expected_cxml));
        connector = $call(Connector, cxnSp, null);
        return [connector, new_y, expected_xml];
      } },
    move_begin_fixture: { args: ["request", "shape_"], params: [[0, "p:cxnSp/p:spPr/a:xfrm/(a:off{x=25,y=15},a:ext{cx=74,cy=123})"], [1, "p:cxnSp/p:spPr/a:xfrm/(a:off{x=10,y=33},a:ext{cx=89,cy=105})"], [2, "p:cxnSp/p:spPr/a:xfrm/(a:off{x=25,y=51},a:ext{cx=74,cy=87})"], [3, "p:cxnSp/p:spPr/a:xfrm/(a:off{x=40,y=33},a:ext{cx=59,cy=105})"]], fn: (self: any, request: any, shape_: any): any => {
        let connector: any, cxnSp: any, cxn_idx: any, expected_cxml: any, expected_xml: any;
        [cxn_idx, expected_cxml] = $list(request.param);
        cxnSp = element("p:cxnSp/p:spPr/a:xfrm/(a:off{x=66,y=99},a:ext{cx=33,cy=39})");
        connector = $call(Connector, cxnSp, null);
        [shape_.left, shape_.top, shape_.width, shape_.height] = $list([10, 15, 30, 36]);
        expected_xml = xml(expected_cxml);
        return [connector, shape_, cxn_idx, expected_xml];
      } },
    move_end_fixture: { args: ["request", "shape_"], params: [[0, "p:cxnSp/p:spPr/a:xfrm/(a:off{x=10,y=15},a:ext{cx=50,cy=10})"], [1, "p:cxnSp/p:spPr/a:xfrm/(a:off{x=10,y=15},a:ext{cx=40,cy=19})"], [2, "p:cxnSp/p:spPr/a:xfrm/(a:off{x=10,y=15},a:ext{cx=50,cy=28})"], [3, "p:cxnSp/p:spPr/a:xfrm/(a:off{x=10,y=15},a:ext{cx=60,cy=19})"]], fn: (self: any, request: any, shape_: any): any => {
        let connector: any, cxnSp: any, cxn_idx: any, expected_cxml: any, expected_xml: any;
        [cxn_idx, expected_cxml] = $list(request.param);
        cxnSp = element("p:cxnSp/p:spPr/a:xfrm/(a:off{x=10,y=15},a:ext{cx=10,cy=5})");
        connector = $call(Connector, cxnSp, null);
        [shape_.left, shape_.top, shape_.width, shape_.height] = $list([50, 25, 20, 18]);
        expected_xml = xml(expected_cxml);
        return [connector, shape_, cxn_idx, expected_xml];
      } },
    _connect_begin_to_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, Connector, "_connect_begin_to", $kw({autospec: true}));
      } },
    _connect_end_to_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, Connector, "_connect_end_to", $kw({autospec: true}));
      } },
    _move_begin_to_cxn_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, Connector, "_move_begin_to_cxn", $kw({autospec: true}));
      } },
    _move_end_to_cxn_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, Connector, "_move_end_to_cxn", $kw({autospec: true}));
      } },
    shape_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, BaseShape);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_begin_point_x_location", args: ["begin_x_get_fixture"], fn: (self: any, begin_x_get_fixture: any): any => {
        let begin_x: any, connector: any, expected_value: any;
        [connector, expected_value] = $list(begin_x_get_fixture);
        begin_x = connector.beginX;
        $assert($isinstance(begin_x, Emu));
        $assertEqual(connector.beginX, expected_value);
      } },
    { name: "it_can_change_its_begin_point_x_location", args: ["begin_x_set_fixture"], fn: (self: any, begin_x_set_fixture: any): any => {
        let connector: any, expected_xml: any, new_x: any;
        [connector, new_x, expected_xml] = $list(begin_x_set_fixture);
        connector.beginX = new_x;
        $assertEqual(connector._element.xml, expected_xml);
      } },
    { name: "it_knows_its_begin_point_y_location", args: ["begin_y_get_fixture"], fn: (self: any, begin_y_get_fixture: any): any => {
        let begin_y: any, connector: any, expected_value: any;
        [connector, expected_value] = $list(begin_y_get_fixture);
        begin_y = connector.beginY;
        $assert($isinstance(begin_y, Emu));
        $assertEqual(connector.beginY, expected_value);
      } },
    { name: "it_can_change_its_begin_point_y_location", args: ["begin_y_set_fixture"], fn: (self: any, begin_y_set_fixture: any): any => {
        let connector: any, expected_xml: any, new_y: any;
        [connector, new_y, expected_xml] = $list(begin_y_set_fixture);
        connector.beginY = new_y;
        $assertEqual(connector._element.xml, expected_xml);
      } },
    { name: "it_knows_its_end_point_x_location", args: ["end_x_get_fixture"], fn: (self: any, end_x_get_fixture: any): any => {
        let connector: any, end_x: any, expected_value: any;
        [connector, expected_value] = $list(end_x_get_fixture);
        end_x = connector.endX;
        $assert($isinstance(end_x, Emu));
        $assertEqual(connector.endX, expected_value);
      } },
    { name: "it_can_change_its_end_point_x_location", args: ["end_x_set_fixture"], fn: (self: any, end_x_set_fixture: any): any => {
        let connector: any, expected_xml: any, new_x: any;
        [connector, new_x, expected_xml] = $list(end_x_set_fixture);
        connector.endX = new_x;
        $assertEqual(connector._element.xml, expected_xml);
      } },
    { name: "it_knows_its_end_point_y_location", args: ["end_y_get_fixture"], fn: (self: any, end_y_get_fixture: any): any => {
        let connector: any, end_y: any, expected_value: any;
        [connector, expected_value] = $list(end_y_get_fixture);
        end_y = connector.endY;
        $assert($isinstance(end_y, Emu));
        $assertEqual(connector.endY, expected_value);
      } },
    { name: "it_can_change_its_end_point_y_location", args: ["end_y_set_fixture"], fn: (self: any, end_y_set_fixture: any): any => {
        let connector: any, expected_xml: any, new_y: any;
        [connector, new_y, expected_xml] = $list(end_y_set_fixture);
        connector.endY = new_y;
        $assertEqual(connector._element.xml, expected_xml);
      } },
    { name: "it_can_connect_its_begin_point_to_a_shape", args: ["begin_conn_fixture"], fn: (self: any, begin_conn_fixture: any): any => {
        let connector: any, cxn_idx: any, shape: any;
        [connector, shape, cxn_idx] = $list(begin_conn_fixture);
        connector.beginConnect(shape, cxn_idx);
        connector._connectBeginTo.assertCalledOnceWith(connector, shape, cxn_idx);
        connector._moveBeginToCxn.assertCalledOnceWith(connector, shape, cxn_idx);
      } },
    { name: "it_can_connect_its_end_point_to_a_shape", args: ["end_conn_fixture"], fn: (self: any, end_conn_fixture: any): any => {
        let connector: any, cxn_idx: any, shape: any;
        [connector, shape, cxn_idx] = $list(end_conn_fixture);
        connector.endConnect(shape, cxn_idx);
        connector._connectEndTo.assertCalledOnceWith(connector, shape, cxn_idx);
        connector._moveEndToCxn.assertCalledOnceWith(connector, shape, cxn_idx);
      } },
    { name: "it_provides_access_to_its_line_format", args: [], fn: (self: any): any => {
        let connector: any, line: any;
        connector = $call(Connector, element("p:cxnSp/p:spPr"), null);
        line = connector.line;
        $assert($isinstance(line, LineFormat));
        line.width = 91440;
        $assertEqual(line.width, 91440);
      } },
    { name: "it_knows_its_shape_type", args: [], fn: (self: any): any => {
        $assertEqual($call(Connector, null, null).shapeType, MSO_SHAPE_TYPE.LINE);
      } },
    { name: "it_connects_its_begin_point_to_help", args: ["connect_begin_fixture"], fn: (self: any, connect_begin_fixture: any): any => {
        let connector: any, cxn_idx: any, expected_xml: any, shape: any;
        [connector, shape, cxn_idx, expected_xml] = $list(connect_begin_fixture);
        connector._connectBeginTo(shape, cxn_idx);
        $assertEqual(connector._element.xml, expected_xml);
      } },
    { name: "it_connects_its_end_point_to_help", args: ["connect_end_fixture"], fn: (self: any, connect_end_fixture: any): any => {
        let connector: any, cxn_idx: any, expected_xml: any, shape: any;
        [connector, shape, cxn_idx, expected_xml] = $list(connect_end_fixture);
        connector._connectEndTo(shape, cxn_idx);
        $assertEqual(connector._element.xml, expected_xml);
      } },
    { name: "it_moves_its_begin_point_to_help", args: ["move_begin_fixture"], fn: (self: any, move_begin_fixture: any): any => {
        let connector: any, cxn_idx: any, expected_xml: any, shape: any;
        [connector, shape, cxn_idx, expected_xml] = $list(move_begin_fixture);
        connector._moveBeginToCxn(shape, cxn_idx);
        $assertEqual(connector._element.xml, expected_xml);
      } },
    { name: "it_moves_its_end_point_to_help", args: ["move_end_fixture"], fn: (self: any, move_end_fixture: any): any => {
        let connector: any, cxn_idx: any, expected_xml: any, shape: any;
        [connector, shape, cxn_idx, expected_xml] = $list(move_end_fixture);
        connector._moveEndToCxn(shape, cxn_idx);
        $assertEqual(connector._element.xml, expected_xml);
      } },
  ],
};
pytestClass("DescribeConnector", DescribeConnector);

