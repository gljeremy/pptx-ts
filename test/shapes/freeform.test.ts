// Machine-translated from python-pptx tests/shapes/test_freeform.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $append, $assert, $assertEqual, $call, $extend, $get, $is, $isinstance, $kw, $list, $truthy, call, initializerMock, instanceMock, methodMock, propertyMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { Shape as Shape$imported } from "../../src/shapes/autoshape.js";
import { FreeformBuilder, _BaseDrawingOperation as _BaseDrawingOperation$imported, _Close, _LineSegment, _MoveTo } from "../../src/shapes/freeform.js";
import { SlideShapes as SlideShapes$imported } from "../../src/shapes/shapetree.js";
import { Emu as Emu$imported, Mm as Mm$imported } from "../../src/util.js";
import { element, xml } from "../unitutil/cxml.js";
import { snippetSeq } from "../unitutil/file.js";
// python binds imported names at import time; patches made later don't change them
const Shape = Shape$imported;
const _BaseDrawingOperation = _BaseDrawingOperation$imported;
const SlideShapes = SlideShapes$imported;
const Emu = Emu$imported;
const Mm = Mm$imported;
void $py;


const DescribeFreeformBuilder: any = {
  fixtures: {
    dx_fixture: { args: ["request"], params: [[0, [1, null, 2, 3], 3], [6, [1, null, 2, 3], 5], [50, [150, -5, null, 100], 155]], fn: (self: any, request: any): any => {
        let builder: any, drawing_ops: any, expected_value: any, start_x: any, xs: any;
        [start_x, xs, expected_value] = $list(request.param);
        drawing_ops = [];
        for (const x of $list(xs)) {
          if ($truthy($is(x, null))) {
            $append(drawing_ops, $call(_Close));
          } else {
            $append(drawing_ops, $call(_BaseDrawingOperation, null, x, null));
          }
        }
        builder = $call(FreeformBuilder, null, start_x, null, null, null);
        $extend(builder._drawingOperations, drawing_ops);
        return [builder, expected_value];
      } },
    dy_fixture: { args: ["request"], params: [[0, [1, null, 2, 3], 3], [2, [1, null, 2, 3], 2], [32, [160, -8, null, 101], 168]], fn: (self: any, request: any): any => {
        let builder: any, drawing_ops: any, expected_value: any, start_y: any, ys: any;
        [start_y, ys, expected_value] = $list(request.param);
        drawing_ops = [];
        for (const y of $list(ys)) {
          if ($truthy($is(y, null))) {
            $append(drawing_ops, $call(_Close));
          } else {
            $append(drawing_ops, $call(_BaseDrawingOperation, null, null, y));
          }
        }
        builder = $call(FreeformBuilder, null, null, start_y, null, null);
        $extend(builder._drawingOperations, drawing_ops);
        return [builder, expected_value];
      } },
    left_fixture: { args: ["request", "shape_offset_x_prop_"], params: [[0, 1.0, 0], [4, 10.0, 40], [914400, 914.3, 836035920]], fn: (self: any, request: any, shape_offset_x_prop_: any): any => {
        let builder: any, expected_value: any, offset_x: any, x_scale: any;
        [offset_x, x_scale, expected_value] = $list(request.param);
        shape_offset_x_prop_.returnValue = offset_x;
        builder = $call(FreeformBuilder, null, null, null, x_scale, null);
        return [builder, expected_value];
      } },
    local_fixture: { args: ["shape_offset_x_prop_", "shape_offset_y_prop_"], fn: (self: any, shape_offset_x_prop_: any, shape_offset_y_prop_: any): any => {
        let builder: any, expected_value: any, local_x: any, local_y: any;
        [local_x, local_y] = $list([123, 456]);
        shape_offset_x_prop_.returnValue = 23;
        shape_offset_y_prop_.returnValue = 156;
        builder = $call(FreeformBuilder, null, null, null, null, null);
        expected_value = [100, 300];
        return [builder, local_x, local_y, expected_value];
      } },
    sp_fixture: { args: ["_left_prop_", "_top_prop_", "_width_prop_", "_height_prop_"], fn: (self: any, _left_prop_: any, _top_prop_: any, _width_prop_: any, _height_prop_: any): any => {
        let builder: any, expected_xml: any, origin_x: any, origin_y: any, shapes: any, spTree: any;
        [origin_x, origin_y] = $list([42, 24]);
        spTree = element("p:spTree");
        shapes = $call(SlideShapes, spTree, null);
        [_left_prop_.returnValue, _top_prop_.returnValue] = $list([12, 34]);
        [_width_prop_.returnValue, _height_prop_.returnValue] = $list([56, 78]);
        builder = $call(FreeformBuilder, shapes, null, null, null, null);
        expected_xml = $get(snippetSeq("freeform"), 0);
        return [builder, origin_x, origin_y, spTree, expected_xml];
      } },
    top_fixture: { args: ["request", "shape_offset_y_prop_"], params: [[0, 11.0, 0], [100, 10.36, 1036], [914242, 943.1, 862221630]], fn: (self: any, request: any, shape_offset_y_prop_: any): any => {
        let builder: any, expected_value: any, offset_y: any, y_scale: any;
        [offset_y, y_scale, expected_value] = $list(request.param);
        shape_offset_y_prop_.returnValue = offset_y;
        builder = $call(FreeformBuilder, null, null, null, null, y_scale);
        return [builder, expected_value];
      } },
    width_fixture: { args: ["request", "_dx_prop_"], params: [[0, 1.0, 0], [42, 10.0, 420], [914400, 914.4, 836127360]], fn: (self: any, request: any, _dx_prop_: any): any => {
        let builder: any, dx: any, expected_value: any, x_scale: any;
        [dx, x_scale, expected_value] = $list(request.param);
        _dx_prop_.returnValue = dx;
        builder = $call(FreeformBuilder, null, null, null, x_scale, null);
        return [builder, expected_value];
      } },
    _add_freeform_sp_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, FreeformBuilder, "_add_freeform_sp", $kw({autospec: true}));
      } },
    apply_operation_to_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, _LineSegment, "apply_operation_to", $kw({autospec: true}));
      } },
    close_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _Close);
      } },
    _Close_new_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, _Close, "new", $kw({autospec: false}));
      } },
    _dx_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, FreeformBuilder, "_dx");
      } },
    _dy_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, FreeformBuilder, "_dy");
      } },
    _height_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, FreeformBuilder, "_height");
      } },
    _init_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(initializerMock, request, FreeformBuilder, $kw({autospec: true}));
      } },
    _left_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, FreeformBuilder, "_left");
      } },
    line_segment_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _LineSegment);
      } },
    _LineSegment_new_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, _LineSegment, "new", $kw({autospec: false}));
      } },
    move_to_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _MoveTo);
      } },
    _MoveTo_new_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, _MoveTo, "new", $kw({autospec: false}));
      } },
    shape_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Shape);
      } },
    shape_offset_x_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, FreeformBuilder, "shape_offset_x");
      } },
    shape_offset_y_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, FreeformBuilder, "shape_offset_y");
      } },
    shapes_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SlideShapes);
      } },
    _start_path_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, FreeformBuilder, "_start_path", $kw({autospec: true}));
      } },
    _top_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, FreeformBuilder, "_top");
      } },
    _width_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, FreeformBuilder, "_width");
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_a_constructor", args: ["shapes_", "_init_"], fn: (self: any, shapes_: any, _init_: any): any => {
        let builder: any, start_x: any, start_x_int: any, start_y: any, start_y_int: any, x_scale: any, y_scale: any;
        [start_x, start_y, x_scale, y_scale] = $list([99.56, 200.49, 4.2, 2.4]);
        [start_x_int, start_y_int] = $list([100, 200]);
        builder = FreeformBuilder.new(shapes_, start_x, start_y, x_scale, y_scale);
        _init_.assertCalledOnceWith(builder, shapes_, start_x_int, start_y_int, x_scale, y_scale);
        $assert($isinstance(builder, FreeformBuilder));
      } },
    { name: "it_can_add_straight_line_segments", args: ["request", "close"], parametrize: [{ names: ["close"], values: [true, false] }], fn: (self: any, request: any, close: any): any => {
        let _add_close_: any, _add_line_segment_: any, builder: any, return_value: any;
        _add_line_segment_ = methodMock(request, FreeformBuilder, "_add_line_segment");
        _add_close_ = methodMock(request, FreeformBuilder, "_add_close");
        builder = $call(FreeformBuilder, null, null, null, null, null);
        return_value = builder.addLineSegments([[1, 2], [3, 4], [5, 6]], close);
        $assertEqual(_add_line_segment_.callArgsList, [call(builder, 1, 2), call(builder, 3, 4), call(builder, 5, 6)]);
        $assertEqual(_add_close_.callArgsList, ($truthy(close) ? [call(builder)] : []));
        $assert($is(return_value, builder));
      } },
    { name: "it_can_move_the_pen_location", args: ["_MoveTo_new_", "move_to_"], fn: (self: any, _MoveTo_new_: any, move_to_: any): any => {
        let builder: any, return_value: any, x: any, y: any;
        [x, y] = $list([42, 24]);
        _MoveTo_new_.returnValue = move_to_;
        builder = $call(FreeformBuilder, null, null, null, null, null);
        return_value = builder.moveTo(x, y);
        _MoveTo_new_.assertCalledOnceWith(builder, x, y);
        $assert($is($get(builder._drawingOperations, -1), move_to_));
        $assert($is(return_value, builder));
      } },
    { name: "it_can_build_the_specified_freeform_shape", args: ["shapes_", "apply_operation_to_", "_add_freeform_sp_", "_start_path_", "shape_"], fn: (self: any, shapes_: any, apply_operation_to_: any, _add_freeform_sp_: any, _start_path_: any, shape_: any): any => {
        let builder: any, calls: any, drawing_ops: any, origin_x: any, origin_y: any, path: any, shape: any, sp: any;
        [origin_x, origin_y] = $list([$call(Mm, 42), $call(Mm, 24)]);
        [sp, path] = $list([element("p:sp"), element("a:path")]);
        drawing_ops = [$call(_LineSegment, null, null, null), $call(_LineSegment, null, null, null)];
        shapes_._shapeFactory.returnValue = shape_;
        _add_freeform_sp_.returnValue = sp;
        _start_path_.returnValue = path;
        builder = $call(FreeformBuilder, shapes_, null, null, null, null);
        $extend(builder._drawingOperations, drawing_ops);
        calls = [call($get(drawing_ops, 0), path), call($get(drawing_ops, 1), path)];
        shape = builder.convertToShape(origin_x, origin_y);
        _add_freeform_sp_.assertCalledOnceWith(builder, origin_x, origin_y);
        _start_path_.assertCalledOnceWith(builder, sp);
        $assertEqual(apply_operation_to_.callArgsList, calls);
        shapes_._shapeFactory.assertCalledOnceWith(sp);
        $assert($is(shape, shape_));
      } },
    { name: "it_knows_the_shape_x_offset", args: ["start_x", "xs", "expected_value"], parametrize: [{ names: ["start_x", "xs", "expected_value"], values: [[$call(Mm, 0), [1, null, 2, 3], $call(Mm, 0)], [$call(Mm, 6), [1, null, 2, 3], $call(Mm, 1)], [$call(Mm, 50), [150, -5, null, 100], $call(Mm, -5)]] }], fn: (self: any, start_x: any, xs: any, expected_value: any): any => {
        let builder: any, drawing_ops: any;
        builder = $call(FreeformBuilder, null, start_x, null, null, null);
        drawing_ops = $list(xs).map((x: any) => ($truthy($is(x, null)) ? $call(_Close) : $call(_LineSegment, builder, $call(Mm, x), $call(Mm, 0))));
        $extend(builder._drawingOperations, drawing_ops);
        $assertEqual(builder.shapeOffsetX, expected_value);
      } },
    { name: "it_knows_the_shape_y_offset", args: ["start_y", "ys", "expected_value"], parametrize: [{ names: ["start_y", "ys", "expected_value"], values: [[$call(Mm, 0), [2, null, 6, 8], $call(Mm, 0)], [$call(Mm, 4), [2, null, 6, 8], $call(Mm, 2)], [$call(Mm, 19), [213, -22, null, 100], $call(Mm, -22)]] }], fn: (self: any, start_y: any, ys: any, expected_value: any): any => {
        let builder: any, drawing_ops: any;
        builder = $call(FreeformBuilder, null, null, start_y, null, null);
        drawing_ops = $list(ys).map((y: any) => ($truthy($is(y, null)) ? $call(_Close) : $call(_LineSegment, builder, $call(Mm, 0), $call(Mm, y))));
        $extend(builder._drawingOperations, drawing_ops);
        $assertEqual(builder.shapeOffsetY, expected_value);
      } },
    { name: "it_adds_a_freeform_sp_to_help", args: ["_left_prop_", "_top_prop_", "_width_prop_", "_height_prop_"], fn: (self: any, _left_prop_: any, _top_prop_: any, _width_prop_: any, _height_prop_: any): any => {
        let builder: any, expected_xml: any, origin_x: any, origin_y: any, shapes: any, sp: any, spTree: any;
        [origin_x, origin_y] = $list([$call(Emu, 42), $call(Emu, 24)]);
        spTree = element("p:spTree");
        shapes = $call(SlideShapes, spTree, null);
        [_left_prop_.returnValue, _top_prop_.returnValue] = $list([$call(Emu, 12), $call(Emu, 34)]);
        [_width_prop_.returnValue, _height_prop_.returnValue] = $list([56, 78]);
        builder = $call(FreeformBuilder, shapes, null, null, null, null);
        expected_xml = $get(snippetSeq("freeform"), 0);
        sp = builder._addFreeformSp(origin_x, origin_y);
        $assertEqual(spTree.xml, expected_xml);
        $assert($is(sp, $get(spTree.xpath("p:sp"), 0)));
      } },
    { name: "it_adds_a_line_segment_to_help", args: ["_LineSegment_new_", "line_segment_"], fn: (self: any, _LineSegment_new_: any, line_segment_: any): any => {
        let builder: any, x: any, y: any;
        [x, y] = $list([4, 2]);
        _LineSegment_new_.returnValue = line_segment_;
        builder = $call(FreeformBuilder, null, null, null, null, null);
        builder._addLineSegment(x, y);
        _LineSegment_new_.assertCalledOnceWith(builder, x, y);
        $assertEqual(builder._drawingOperations, [line_segment_]);
      } },
    { name: "it_closes_a_contour_to_help", args: ["_Close_new_", "close_"], fn: (self: any, _Close_new_: any, close_: any): any => {
        let builder: any;
        _Close_new_.returnValue = close_;
        builder = $call(FreeformBuilder, null, null, null, null, null);
        builder._addClose();
        _Close_new_.assertCalledOnceWith();
        $assertEqual(builder._drawingOperations, [close_]);
      } },
    { name: "it_knows_the_freeform_left_extent_to_help", args: ["left_fixture"], fn: (self: any, left_fixture: any): any => {
        let builder: any, expected_value: any, left: any;
        [builder, expected_value] = $list(left_fixture);
        left = builder._left;
        $assertEqual(left, expected_value);
      } },
    { name: "it_knows_the_freeform_top_extent_to_help", args: ["top_fixture"], fn: (self: any, top_fixture: any): any => {
        let builder: any, expected_value: any, top: any;
        [builder, expected_value] = $list(top_fixture);
        top = builder._top;
        $assertEqual(top, expected_value);
      } },
    { name: "it_knows_the_freeform_width_to_help", args: ["width_fixture"], fn: (self: any, width_fixture: any): any => {
        let builder: any, expected_value: any, width: any;
        [builder, expected_value] = $list(width_fixture);
        width = builder._width;
        $assertEqual(width, expected_value);
      } },
    { name: "it_knows_the_freeform_height_to_help", args: ["dy", "y_scale", "expected_value", "_dy_prop_"], parametrize: [{ names: ["dy", "y_scale", "expected_value"], values: [[0, 2.0, 0], [24, 10.0, 240], [914400, 314.1, 287213040]] }], fn: (self: any, dy: any, y_scale: any, expected_value: any, _dy_prop_: any): any => {
        let builder: any, height: any;
        _dy_prop_.returnValue = dy;
        builder = $call(FreeformBuilder, null, null, null, null, y_scale);
        height = builder._height;
        $assertEqual(height, expected_value);
      } },
    { name: "it_knows_the_local_coordinate_width_to_help", args: ["dx_fixture"], fn: (self: any, dx_fixture: any): any => {
        let builder: any, dx: any, expected_value: any;
        [builder, expected_value] = $list(dx_fixture);
        dx = builder._dx;
        $assertEqual(dx, expected_value);
      } },
    { name: "it_knows_the_local_coordinate_height_to_help", args: ["dy_fixture"], fn: (self: any, dy_fixture: any): any => {
        let builder: any, dy: any, expected_value: any;
        [builder, expected_value] = $list(dy_fixture);
        dy = builder._dy;
        $assertEqual(dy, expected_value);
      } },
    { name: "it_can_start_a_new_path_to_help", args: ["request", "_dx_prop_", "_dy_prop_"], fn: (self: any, request: any, _dx_prop_: any, _dy_prop_: any): any => {
        let _local_to_shape_: any, builder: any, path: any, sp: any, start_x: any, start_y: any;
        _local_to_shape_ = $call(methodMock, request, FreeformBuilder, "_local_to_shape", $kw({return_value: [101, 202]}));
        sp = element("p:sp/p:spPr/a:custGeom");
        [start_x, start_y] = $list([42, 24]);
        [_dx_prop_.returnValue, _dy_prop_.returnValue] = $list([1001, 2002]);
        builder = $call(FreeformBuilder, null, start_x, start_y, null, null);
        path = builder._startPath(sp);
        _local_to_shape_.assertCalledOnceWith(builder, start_x, start_y);
        $assertEqual(sp.xml, xml("p:sp/p:spPr/a:custGeom/a:pathLst/a:path{w=1001,h=2002}/a:moveTo/a:pt{x=101,y=202}"));
        $assert($is(path, $get(sp.xpath(".//a:path"), -1)));
      } },
    { name: "it_translates_local_to_shape_coordinates_to_help", args: ["local_fixture"], fn: (self: any, local_fixture: any): any => {
        let builder: any, expected_value: any, local_x: any, local_y: any, shape_x_y: any;
        [builder, local_x, local_y, expected_value] = $list(local_fixture);
        shape_x_y = builder._localToShape(local_x, local_y);
        $assertEqual(shape_x_y, expected_value);
      } },
  ],
};
pytestClass("DescribeFreeformBuilder", DescribeFreeformBuilder);

const Describe_BaseDrawingOperation: any = {
  fixtures: {
    x_fixture: { args: [], fn: (self: any): any => {
        let drawing_operation: any, expected_value: any, x: any;
        x = 42;
        drawing_operation = $call(_BaseDrawingOperation, null, x, null);
        expected_value = x;
        return [drawing_operation, expected_value];
      } },
    y_fixture: { args: [], fn: (self: any): any => {
        let drawing_operation: any, expected_value: any, y: any;
        y = 24;
        drawing_operation = $call(_BaseDrawingOperation, null, null, y);
        expected_value = y;
        return [drawing_operation, expected_value];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_x_coordinate", args: ["x_fixture"], fn: (self: any, x_fixture: any): any => {
        let drawing_operation: any, expected_value: any, x: any;
        [drawing_operation, expected_value] = $list(x_fixture);
        x = drawing_operation.x;
        $assertEqual(x, expected_value);
      } },
    { name: "it_knows_its_y_coordinate", args: ["y_fixture"], fn: (self: any, y_fixture: any): any => {
        let drawing_operation: any, expected_value: any, y: any;
        [drawing_operation, expected_value] = $list(y_fixture);
        y = drawing_operation.y;
        $assertEqual(y, expected_value);
      } },
  ],
};
pytestClass("Describe_BaseDrawingOperation", Describe_BaseDrawingOperation);

const Describe_Close: any = {
  fixtures: {
    new_fixture: { args: ["_init_"], fn: (self: any, _init_: any): any => {
        return _init_;
      } },
    apply_fixture: { args: [], fn: (self: any): any => {
        let close: any, expected_xml: any, path: any;
        path = element("a:path");
        close = $call(_Close);
        expected_xml = xml("a:path/a:close");
        return [close, path, expected_xml];
      } },
    _init_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(initializerMock, request, _Close, $kw({autospec: true}));
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_a_constructor", args: ["new_fixture"], fn: (self: any, new_fixture: any): any => {
        let _init_: any, close: any;
        _init_ = new_fixture;
        close = _Close.new();
        _init_.assertCalledOnceWith();
        $assert($isinstance(close, _Close));
      } },
    { name: "it_can_add_close_a_contour", args: ["apply_fixture"], fn: (self: any, apply_fixture: any): any => {
        let close: any, close_elm: any, expected_xml: any, path: any;
        [close, path, expected_xml] = $list(apply_fixture);
        close_elm = close.applyOperationTo(path);
        $assertEqual(path.xml, expected_xml);
        $assert($is(close_elm, $get(path.xpath("a:close"), -1)));
      } },
  ],
};
pytestClass("Describe_Close", Describe_Close);

const Describe_LineSegment: any = {
  fixtures: {
    apply_fixture: { args: ["builder_"], fn: (self: any, builder_: any): any => {
        let expected_xml: any, line_segment: any, path: any, x: any, y: any;
        [x, y] = $list([420, 240]);
        path = element("a:path");
        [builder_.shapeOffsetX, builder_.shapeOffsetY] = $list([100, 200]);
        line_segment = $call(_LineSegment, builder_, x, y);
        expected_xml = xml("a:path/a:lnTo/a:pt{x=320,y=40}");
        return [line_segment, path, expected_xml];
      } },
    new_fixture: { args: ["builder_", "_init_"], fn: (self: any, builder_: any, _init_: any): any => {
        let x: any, x_int: any, y: any, y_int: any;
        [x, y, x_int, y_int] = $list([99.51, 200.49, 100, 200]);
        return [builder_, x, y, _init_, x_int, y_int];
      } },
    builder_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, FreeformBuilder);
      } },
    _init_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(initializerMock, request, _LineSegment, $kw({autospec: true}));
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_a_constructor", args: ["new_fixture"], fn: (self: any, new_fixture: any): any => {
        let _init_: any, builder_: any, line_segment: any, x: any, x_int: any, y: any, y_int: any;
        [builder_, x, y, _init_, x_int, y_int] = $list(new_fixture);
        line_segment = _LineSegment.new(builder_, x, y);
        _init_.assertCalledOnceWith(line_segment, builder_, x_int, y_int);
        $assert($isinstance(line_segment, _LineSegment));
      } },
    { name: "it_can_add_its_line_segment_to_a_path", args: ["apply_fixture"], fn: (self: any, apply_fixture: any): any => {
        let expected_xml: any, line_segment: any, lnTo: any, path: any;
        [line_segment, path, expected_xml] = $list(apply_fixture);
        lnTo = line_segment.applyOperationTo(path);
        $assertEqual(path.xml, expected_xml);
        $assert($is(lnTo, $get(path.xpath("a:lnTo"), -1)));
      } },
  ],
};
pytestClass("Describe_LineSegment", Describe_LineSegment);

const Describe_MoveTo: any = {
  fixtures: {
    apply_fixture: { args: ["builder_"], fn: (self: any, builder_: any): any => {
        let expected_xml: any, move_to: any, path: any, x: any, y: any;
        [x, y] = $list([120, 340]);
        path = element("a:path");
        [builder_.shapeOffsetX, builder_.shapeOffsetY] = $list([100, 200]);
        move_to = $call(_MoveTo, builder_, x, y);
        expected_xml = xml("a:path/a:moveTo/a:pt{x=20,y=140}");
        return [move_to, path, expected_xml];
      } },
    new_fixture: { args: ["builder_", "_init_"], fn: (self: any, builder_: any, _init_: any): any => {
        let x: any, x_int: any, y: any, y_int: any;
        [x, y, x_int, y_int] = $list([99.51, 200.49, 100, 200]);
        return [builder_, x, y, _init_, x_int, y_int];
      } },
    builder_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, FreeformBuilder);
      } },
    _init_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(initializerMock, request, _MoveTo, $kw({autospec: true}));
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_a_constructor", args: ["new_fixture"], fn: (self: any, new_fixture: any): any => {
        let _init_: any, builder_: any, move_to: any, x: any, x_int: any, y: any, y_int: any;
        [builder_, x, y, _init_, x_int, y_int] = $list(new_fixture);
        move_to = _MoveTo.new(builder_, x, y);
        _init_.assertCalledOnceWith(move_to, builder_, x_int, y_int);
        $assert($isinstance(move_to, _MoveTo));
      } },
    { name: "it_can_add_its_move_to_a_path", args: ["apply_fixture"], fn: (self: any, apply_fixture: any): any => {
        let expected_xml: any, moveTo: any, move_to: any, path: any;
        [move_to, path, expected_xml] = $list(apply_fixture);
        moveTo = move_to.applyOperationTo(path);
        $assertEqual(path.xml, expected_xml);
        $assert($is(moveTo, $get(path.xpath("a:moveTo"), -1)));
      } },
  ],
};
pytestClass("Describe_MoveTo", Describe_MoveTo);

