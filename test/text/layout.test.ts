// Machine-translated from python-pptx tests/text/test_layout.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $all, $append, $assert, $assertEqual, $call, $cmp, $eq, $extend, $is, $kw, $list, $range, $slice, $truthy, $zip, ANY, call, classMock, functionMock, initializerMock, instanceMock, methodMock, propertyMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { TextFitter, _BinarySearchTree as _BinarySearchTree$imported, _Line as _Line$imported, _LineSource as _LineSource$imported } from "../../src/text/layout.js";
// python binds imported names at import time; patches made later don't change them
const _BinarySearchTree = _BinarySearchTree$imported;
const _Line = _Line$imported;
const _LineSource = _LineSource$imported;
void $py;


const DescribeTextFitter: any = {
  fixtures: {
    _best_fit_fixture: { args: ["_BinarySearchTree_", "_fits_inside_predicate_"], fn: (self: any, _BinarySearchTree_: any, _fits_inside_predicate_: any): any => {
        let font_size_: any, max_size: any, predicate_: any, sizes_: any, text_fitter: any;
        text_fitter = $call(TextFitter, null, [null, null], null);
        max_size = 42;
        sizes_ = _BinarySearchTree_.fromOrderedSequence.returnValue;
        predicate_ = _fits_inside_predicate_.returnValue;
        font_size_ = sizes_.findMax.returnValue;
        return [text_fitter, max_size, _BinarySearchTree_, sizes_, predicate_, font_size_];
      } },
    fits_cx_pred_fixture: { args: ["request", "_rendered_size_"], params: [[49, true], [50, true], [51, false]], fn: (self: any, request: any, _rendered_size_: any): any => {
        let expected_value: any, line: any, point_size: any, rendered_width: any, text_fitter: any;
        [rendered_width, expected_value] = $list(request.param);
        text_fitter = $call(TextFitter, null, [50, null], "foobar.ttf");
        [point_size, line] = $list([12, $call(_Line, "foobar", null)]);
        _rendered_size_.returnValue = [rendered_width, null];
        return [text_fitter, point_size, line, _rendered_size_, expected_value];
      } },
    _BinarySearchTree_: { args: ["request"], fn: (self: any, request: any): any => {
        return classMock(request, "pptx.text.layout._BinarySearchTree");
      } },
    _fits_inside_predicate_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, TextFitter, "_fits_inside_predicate");
      } },
    line_source_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _LineSource);
      } },
    _rendered_size_: { args: ["request"], fn: (self: any, request: any): any => {
        return functionMock(request, "pptx.text.layout._rendered_size");
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_determine_the_best_fit_font_size", args: ["request", "line_source_"], fn: (self: any, request: any, line_source_: any): any => {
        let _LineSource_: any, _best_fit_font_size_: any, _init_: any, extents: any, font_size: any, max_size: any;
        _LineSource_ = $call(classMock, request, "pptx.text.layout._LineSource", $kw({return_value: line_source_}));
        _init_ = initializerMock(request, TextFitter);
        _best_fit_font_size_ = $call(methodMock, request, TextFitter, "_best_fit_font_size", $kw({return_value: 36}));
        [extents, max_size] = $list([[19, 20], 42]);
        font_size = TextFitter.bestFitFontSize("Foobar", extents, max_size, "foobar.ttf");
        _LineSource_.assertCalledOnceWith("Foobar");
        _init_.assertCalledOnceWith(ANY, line_source_, extents, "foobar.ttf");
        _best_fit_font_size_.assertCalledOnceWith(ANY, max_size);
        $assertEqual(font_size, 36);
      } },
    { name: "it_finds_best_fit_font_size_to_help_best_fit", args: ["_best_fit_fixture"], fn: (self: any, _best_fit_fixture: any): any => {
        let _BinarySearchTree_: any, font_size: any, font_size_: any, max_size: any, predicate_: any, sizes_: any, text_fitter: any;
        [text_fitter, max_size, _BinarySearchTree_] = $list($slice(_best_fit_fixture, null, 3, null));
        [sizes_, predicate_, font_size_] = $list($slice(_best_fit_fixture, 3, null, null));
        font_size = text_fitter._bestFitFontSize(max_size);
        _BinarySearchTree_.fromOrderedSequence.assertCalledOnceWith($range(1, (max_size + 1)));
        sizes_.findMax.assertCalledOnceWith(predicate_);
        $assert($is(font_size, font_size_));
      } },
    { name: "it_provides_a_fits_inside_predicate_fn", args: ["request", "line_source_", "_rendered_size_", "extents", "point_size", "text_lines", "expected_value"], parametrize: [{ names: ["extents", "point_size", "text_lines", "expected_value"], values: [[[66, 99], 6, ["foo", "bar"], false], [[66, 100], 6, ["foo", "bar"], true], [[66, 101], 6, ["foo", "bar"], true]] }], fn: (self: any, request: any, line_source_: any, _rendered_size_: any, extents: any, point_size: any, text_lines: any, expected_value: any): any => {
        let _wrap_lines_: any, predicate: any, result: any, text_fitter: any;
        _wrap_lines_ = $call(methodMock, request, TextFitter, "_wrap_lines", $kw({return_value: text_lines}));
        _rendered_size_.returnValue = [null, 50];
        text_fitter = $call(TextFitter, line_source_, extents, "foobar.ttf");
        predicate = text_fitter._fitsInsidePredicate;
        result = predicate(point_size);
        _wrap_lines_.assertCalledOnceWith(text_fitter, line_source_, point_size);
        _rendered_size_.assertCalledOnceWith("Ty", point_size, text_fitter._fontFile);
        $assert($is(result, expected_value));
      } },
    { name: "it_provides_a_fits_in_width_predicate_fn", args: ["fits_cx_pred_fixture"], fn: (self: any, fits_cx_pred_fixture: any): any => {
        let _rendered_size_: any, expected_value: any, line: any, point_size: any, predicate: any, result: any, text_fitter: any;
        [text_fitter, point_size, line] = $list($slice(fits_cx_pred_fixture, null, 3, null));
        [_rendered_size_, expected_value] = $list($slice(fits_cx_pred_fixture, 3, null, null));
        predicate = text_fitter._fitsInWidthPredicate(point_size);
        result = predicate(line);
        _rendered_size_.assertCalledOnceWith(line.text, point_size, text_fitter._fontFile);
        $assert($is(result, expected_value));
      } },
    { name: "it_wraps_lines_to_help_best_fit", args: ["request"], fn: (self: any, request: any): any => {
        let _break_line_: any, line_source: any, remainder: any, text_fitter: any;
        [line_source, remainder] = $list([$call(_LineSource, "foo bar"), $call(_LineSource, "bar")]);
        _break_line_ = $call(methodMock, request, TextFitter, "_break_line", $kw({side_effect: [["foo", remainder], ["bar", $call(_LineSource, "")]]}));
        text_fitter = $call(TextFitter, null, [null, null], null);
        text_fitter._wrapLines(line_source, 21);
        $assertEqual(_break_line_.callArgsList, [call(text_fitter, line_source, 21), call(text_fitter, remainder, 21)]);
      } },
    { name: "it_breaks_off_a_line_to_help_wrap", args: ["request", "line_source_", "_BinarySearchTree_"], fn: (self: any, request: any, line_source_: any, _BinarySearchTree_: any): any => {
        let _fits_in_width_predicate_: any, bst_: any, max_value_: any, predicate_: any, text_fitter: any, value: any;
        bst_ = instanceMock(request, _BinarySearchTree);
        _fits_in_width_predicate_ = methodMock(request, TextFitter, "_fits_in_width_predicate");
        _BinarySearchTree_.fromOrderedSequence.returnValue = bst_;
        predicate_ = _fits_in_width_predicate_.returnValue;
        max_value_ = bst_.findMax.returnValue;
        text_fitter = $call(TextFitter, null, [null, null], null);
        value = text_fitter._breakLine(line_source_, 21);
        _BinarySearchTree_.fromOrderedSequence.assertCalledOnceWith(line_source_);
        text_fitter._fitsInWidthPredicate.assertCalledOnceWith(text_fitter, 21);
        bst_.findMax.assertCalledOnceWith(predicate_);
        $assert($is(value, max_value_));
      } },
  ],
};
pytestClass("DescribeTextFitter", DescribeTextFitter);

const Describe_BinarySearchTree: any = {
  fixtures: {
    max_fixture: { args: ["request"], params: [[$range(10), ((n: any) => $cmp(n, "<", 6.5)), 6], [$range(10), ((n: any) => $cmp(n, ">", 9.9)), null], [$range(10), ((n: any) => $cmp(n, "<", 0.0)), null]], fn: (self: any, request: any): any => {
        let bst: any, expected_value: any, predicate: any, seq: any;
        [seq, predicate, expected_value] = $list(request.param);
        bst = _BinarySearchTree.fromOrderedSequence(seq);
        return [bst, predicate, expected_value];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_construct_from_an_ordered_sequence", args: [], fn: (self: any): any => {
        let bst: any;
        bst = _BinarySearchTree.fromOrderedSequence($range(10));
        const in_order = (node: any): any => {
          let result: any;
          result = [];
          if ($truthy($is(node, null))) {
            return result;
          }
          $extend(result, in_order(node._lesser));
          $append(result, node.value);
          $extend(result, in_order(node._greater));
          return result;
        };
        $assertEqual(bst.value, 9);
        $assertEqual(bst._lesser.value, 4);
        $assert($is(bst._greater, null));
        $assertEqual(in_order(bst), $list($range(10)));
      } },
    { name: "it_can_find_the_max_value_satisfying_a_predicate", args: ["max_fixture"], fn: (self: any, max_fixture: any): any => {
        let bst: any, expected_value: any, predicate: any;
        [bst, predicate, expected_value] = $list(max_fixture);
        $assertEqual(bst.findMax(predicate), expected_value);
      } },
  ],
};
pytestClass("Describe_BinarySearchTree", Describe_BinarySearchTree);

const Describe_LineSource: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_generates_text_remainder_pairs", args: [], fn: (self: any): any => {
        let expected: any, line_source: any;
        line_source = $call(_LineSource, "foo bar baz");
        expected = [["foo", $call(_LineSource, "bar baz")], ["foo bar", $call(_LineSource, "baz")], ["foo bar baz", $call(_LineSource, "")]];
        $assert($all($list($zip(expected, line_source)).map(([a, b]: any) => $eq(a, b))));
      } },
  ],
};
pytestClass("Describe_LineSource", Describe_LineSource);

