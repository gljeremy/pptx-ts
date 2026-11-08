// Machine-translated from python-pptx tests/chart/test_axis.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $call, $get, $getattr, $is, $kw, $list, $raises, classMock, instanceMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { AxisTitle as AxisTitle$imported, CategoryAxis as CategoryAxis$imported, DateAxis as DateAxis$imported, MajorGridlines as MajorGridlines$imported, TickLabels as TickLabels$imported, ValueAxis as ValueAxis$imported, _BaseAxis as _BaseAxis$imported } from "../../src/chart/axis.js";
import { ChartFormat as ChartFormat$imported } from "../../src/dml/chtfmt.js";
import { XL_AXIS_CROSSES as XL_AXIS_CROSSES$imported, XL_CATEGORY_TYPE as XL_CATEGORY_TYPE$imported, XL_TICK_LABEL_POSITION as XL_TICK_LBL_POS$imported, XL_TICK_MARK as XL_TICK_MARK$imported } from "../../src/enum/chart.js";
import { Font as Font$imported } from "../../src/text/text.js";
import { element, xml } from "../unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const AxisTitle = AxisTitle$imported;
const CategoryAxis = CategoryAxis$imported;
const DateAxis = DateAxis$imported;
const MajorGridlines = MajorGridlines$imported;
const TickLabels = TickLabels$imported;
const ValueAxis = ValueAxis$imported;
const _BaseAxis = _BaseAxis$imported;
const ChartFormat = ChartFormat$imported;
const XL_AXIS_CROSSES = XL_AXIS_CROSSES$imported;
const XL_CATEGORY_TYPE = XL_CATEGORY_TYPE$imported;
const XL_TICK_LBL_POS = XL_TICK_LBL_POS$imported;
const XL_TICK_MARK = XL_TICK_MARK$imported;
const Font = Font$imported;
void $py;


const Describe_BaseAxis: any = {
  fixtures: {
    format_fixture: { args: ["request", "ChartFormat_", "format_"], params: ["c:catAx", "c:dateAx", "c:valAx"], fn: (self: any, request: any, ChartFormat_: any, format_: any): any => {
        let axis: any, xAx_cxml: any;
        xAx_cxml = request.param;
        axis = $call(_BaseAxis, element(xAx_cxml));
        return [axis, ChartFormat_, format_];
      } },
    has_title_get_fixture: { args: ["request"], params: [["c:catAx", false], ["c:catAx/c:title", true], ["c:dateAx", false], ["c:dateAx/c:title", true], ["c:valAx", false], ["c:valAx/c:title", true]], fn: (self: any, request: any): any => {
        let axis: any, expected_value: any, xAx_cxml: any;
        [xAx_cxml, expected_value] = $list(request.param);
        axis = $call(_BaseAxis, element(xAx_cxml));
        return [axis, expected_value];
      } },
    has_title_set_fixture: { args: ["request"], params: [["c:catAx", true, "c:catAx/c:title/(c:layout,c:overlay{val=0})"], ["c:catAx/c:title", true, "c:catAx/c:title"], ["c:catAx/c:title", false, "c:catAx"], ["c:catAx", false, "c:catAx"], ["c:dateAx", true, "c:dateAx/c:title/(c:layout,c:overlay{val=0})"], ["c:dateAx/c:title", true, "c:dateAx/c:title"], ["c:dateAx/c:title", false, "c:dateAx"], ["c:dateAx", false, "c:dateAx"], ["c:valAx", true, "c:valAx/c:title/(c:layout,c:overlay{val=0})"], ["c:valAx/c:title", true, "c:valAx/c:title"], ["c:valAx/c:title", false, "c:valAx"], ["c:valAx", false, "c:valAx"]], fn: (self: any, request: any): any => {
        let axis: any, expected_xAx_cxml: any, expected_xml: any, new_value: any, xAx_cxml: any;
        [xAx_cxml, new_value, expected_xAx_cxml] = $list(request.param);
        axis = $call(_BaseAxis, element(xAx_cxml));
        expected_xml = xml(expected_xAx_cxml);
        return [axis, new_value, expected_xml];
      } },
    maj_grdlns_fixture: { args: ["request", "MajorGridlines_", "major_gridlines_"], params: ["c:catAx", "c:dateAx", "c:valAx"], fn: (self: any, request: any, MajorGridlines_: any, major_gridlines_: any): any => {
        let axis: any, xAx: any, xAx_cxml: any;
        xAx_cxml = request.param;
        xAx = element(xAx_cxml);
        axis = $call(_BaseAxis, xAx);
        return [axis, MajorGridlines_, xAx, major_gridlines_];
      } },
    major_gridlines_get_fixture: { args: ["request"], params: [["c:catAx", false], ["c:catAx/c:majorGridlines", true], ["c:dateAx", false], ["c:dateAx/c:majorGridlines", true], ["c:valAx", false], ["c:valAx/c:majorGridlines", true]], fn: (self: any, request: any): any => {
        let base_axis: any, expected_value: any, xAx_cxml: any;
        [xAx_cxml, expected_value] = $list(request.param);
        base_axis = $call(_BaseAxis, element(xAx_cxml));
        return [base_axis, expected_value];
      } },
    major_gridlines_set_fixture: { args: ["request"], params: [["c:catAx", true, "c:catAx/c:majorGridlines"], ["c:catAx/c:majorGridlines", true, "c:catAx/c:majorGridlines"], ["c:catAx/c:majorGridlines", false, "c:catAx"], ["c:catAx", false, "c:catAx"], ["c:dateAx", true, "c:dateAx/c:majorGridlines"], ["c:dateAx/c:majorGridlines", true, "c:dateAx/c:majorGridlines"], ["c:dateAx/c:majorGridlines", false, "c:dateAx"], ["c:dateAx", false, "c:dateAx"], ["c:valAx", true, "c:valAx/c:majorGridlines"], ["c:valAx/c:majorGridlines", true, "c:valAx/c:majorGridlines"], ["c:valAx/c:majorGridlines", false, "c:valAx"], ["c:valAx", false, "c:valAx"]], fn: (self: any, request: any): any => {
        let base_axis: any, expected_xAx_cxml: any, expected_xml: any, new_value: any, xAx_cxml: any;
        [xAx_cxml, new_value, expected_xAx_cxml] = $list(request.param);
        base_axis = $call(_BaseAxis, element(xAx_cxml));
        expected_xml = xml(expected_xAx_cxml);
        return [base_axis, new_value, expected_xml];
      } },
    major_tick_get_fixture: { args: ["request"], params: [["c:catAx", XL_TICK_MARK.CROSS], ["c:catAx/c:majorTickMark", XL_TICK_MARK.CROSS], ["c:catAx/c:majorTickMark{val=out}", XL_TICK_MARK.OUTSIDE], ["c:dateAx", XL_TICK_MARK.CROSS], ["c:dateAx/c:majorTickMark", XL_TICK_MARK.CROSS], ["c:dateAx/c:majorTickMark{val=out}", XL_TICK_MARK.OUTSIDE], ["c:valAx", XL_TICK_MARK.CROSS], ["c:valAx/c:majorTickMark", XL_TICK_MARK.CROSS], ["c:valAx/c:majorTickMark{val=in}", XL_TICK_MARK.INSIDE]], fn: (self: any, request: any): any => {
        let axis: any, expected_value: any, xAx_cxml: any;
        [xAx_cxml, expected_value] = $list(request.param);
        axis = $call(_BaseAxis, element(xAx_cxml));
        return [axis, expected_value];
      } },
    major_tick_set_fixture: { args: ["request"], params: [["c:catAx", XL_TICK_MARK.INSIDE, "c:catAx/c:majorTickMark{val=in}"], ["c:catAx/c:majorTickMark{val=in}", XL_TICK_MARK.OUTSIDE, "c:catAx/c:majorTickMark{val=out}"], ["c:catAx/c:majorTickMark{val=out}", XL_TICK_MARK.CROSS, "c:catAx"], ["c:catAx", XL_TICK_MARK.CROSS, "c:catAx"], ["c:catAx/c:majorTickMark{val=cross}", XL_TICK_MARK.CROSS, "c:catAx"], ["c:dateAx", XL_TICK_MARK.INSIDE, "c:dateAx/c:majorTickMark{val=in}"], ["c:dateAx/c:majorTickMark{val=in}", XL_TICK_MARK.OUTSIDE, "c:dateAx/c:majorTickMark{val=out}"], ["c:dateAx/c:majorTickMark{val=out}", XL_TICK_MARK.CROSS, "c:dateAx"], ["c:dateAx", XL_TICK_MARK.CROSS, "c:dateAx"], ["c:dateAx/c:majorTickMark{val=cross}", XL_TICK_MARK.CROSS, "c:dateAx"], ["c:valAx", XL_TICK_MARK.INSIDE, "c:valAx/c:majorTickMark{val=in}"], ["c:valAx/c:majorTickMark{val=in}", XL_TICK_MARK.OUTSIDE, "c:valAx/c:majorTickMark{val=out}"], ["c:valAx/c:majorTickMark{val=out}", XL_TICK_MARK.CROSS, "c:valAx"], ["c:valAx", XL_TICK_MARK.CROSS, "c:valAx"], ["c:valAx/c:majorTickMark{val=cross}", XL_TICK_MARK.CROSS, "c:valAx"]], fn: (self: any, request: any): any => {
        let axis: any, expected_xAx_cxml: any, expected_xml: any, new_value: any, xAx_cxml: any;
        [xAx_cxml, new_value, expected_xAx_cxml] = $list(request.param);
        axis = $call(_BaseAxis, element(xAx_cxml));
        expected_xml = xml(expected_xAx_cxml);
        return [axis, new_value, expected_xml];
      } },
    maximum_scale_get_fixture: { args: ["request"], params: [["c:catAx/c:scaling", null], ["c:catAx/c:scaling/c:max{val=12.34}", 12.34], ["c:dateAx/c:scaling", null], ["c:dateAx/c:scaling/c:max{val=42.24}", 42.24], ["c:valAx/c:scaling", null], ["c:valAx/c:scaling/c:max{val=23.45}", 23.45]], fn: (self: any, request: any): any => {
        let axis: any, expected_value: any, xAx_cxml: any;
        [xAx_cxml, expected_value] = $list(request.param);
        axis = $call(_BaseAxis, element(xAx_cxml));
        return [axis, expected_value];
      } },
    maximum_scale_set_fixture: { args: ["request"], params: [["c:catAx/c:scaling", 34.56, "c:catAx/c:scaling/c:max{val=34.56}"], ["c:catAx/c:scaling/c:max{val=34.56}", 42.42, "c:catAx/c:scaling/c:max{val=42.42}"], ["c:catAx/c:scaling/c:max{val=42.42}", null, "c:catAx/c:scaling"], ["c:catAx/c:scaling", null, "c:catAx/c:scaling"], ["c:dateAx/c:scaling", 45.67, "c:dateAx/c:scaling/c:max{val=45.67}"], ["c:dateAx/c:scaling/c:max{val=45.67}", 42.42, "c:dateAx/c:scaling/c:max{val=42.42}"], ["c:dateAx/c:scaling/c:max{val=42.42}", null, "c:dateAx/c:scaling"], ["c:dateAx/c:scaling", null, "c:dateAx/c:scaling"], ["c:valAx/c:scaling", 56.78, "c:valAx/c:scaling/c:max{val=56.78}"], ["c:valAx/c:scaling/c:max{val=56.78}", 42.42, "c:valAx/c:scaling/c:max{val=42.42}"], ["c:valAx/c:scaling/c:max{val=42.42}", null, "c:valAx/c:scaling"], ["c:valAx/c:scaling", null, "c:valAx/c:scaling"]], fn: (self: any, request: any): any => {
        let axis: any, expected_xAx_cxml: any, expected_xml: any, new_value: any, xAx_cxml: any;
        [xAx_cxml, new_value, expected_xAx_cxml] = $list(request.param);
        axis = $call(_BaseAxis, element(xAx_cxml));
        expected_xml = xml(expected_xAx_cxml);
        return [axis, new_value, expected_xml];
      } },
    minimum_scale_get_fixture: { args: ["request"], params: [["c:catAx/c:scaling", null], ["c:catAx/c:scaling/c:min{val=12.34}", 12.34], ["c:dateAx/c:scaling", null], ["c:dateAx/c:scaling/c:min{val=42.24}", 42.24], ["c:valAx/c:scaling", null], ["c:valAx/c:scaling/c:min{val=23.45}", 23.45]], fn: (self: any, request: any): any => {
        let axis: any, expected_value: any, xAx_cxml: any;
        [xAx_cxml, expected_value] = $list(request.param);
        axis = $call(_BaseAxis, element(xAx_cxml));
        return [axis, expected_value];
      } },
    minimum_scale_set_fixture: { args: ["request"], params: [["c:catAx/c:scaling", 34.56, "c:catAx/c:scaling/c:min{val=34.56}"], ["c:catAx/c:scaling/c:min{val=34.56}", 42.42, "c:catAx/c:scaling/c:min{val=42.42}"], ["c:catAx/c:scaling/c:min{val=42.42}", null, "c:catAx/c:scaling"], ["c:catAx/c:scaling", null, "c:catAx/c:scaling"], ["c:dateAx/c:scaling", 45.67, "c:dateAx/c:scaling/c:min{val=45.67}"], ["c:dateAx/c:scaling/c:min{val=45.67}", 42.42, "c:dateAx/c:scaling/c:min{val=42.42}"], ["c:dateAx/c:scaling/c:min{val=42.42}", null, "c:dateAx/c:scaling"], ["c:dateAx/c:scaling", null, "c:dateAx/c:scaling"], ["c:valAx/c:scaling", 56.78, "c:valAx/c:scaling/c:min{val=56.78}"], ["c:valAx/c:scaling/c:min{val=56.78}", 42.42, "c:valAx/c:scaling/c:min{val=42.42}"], ["c:valAx/c:scaling/c:min{val=42.42}", null, "c:valAx/c:scaling"], ["c:valAx/c:scaling", null, "c:valAx/c:scaling"]], fn: (self: any, request: any): any => {
        let axis: any, expected_xAx_cxml: any, expected_xml: any, new_value: any, xAx_cxml: any;
        [xAx_cxml, new_value, expected_xAx_cxml] = $list(request.param);
        axis = $call(_BaseAxis, element(xAx_cxml));
        expected_xml = xml(expected_xAx_cxml);
        return [axis, new_value, expected_xml];
      } },
    minor_gridlines_get_fixture: { args: ["request"], params: [["c:catAx", false], ["c:catAx/c:minorGridlines", true], ["c:dateAx", false], ["c:dateAx/c:minorGridlines", true], ["c:valAx", false], ["c:valAx/c:minorGridlines", true]], fn: (self: any, request: any): any => {
        let base_axis: any, expected_value: any, xAx_cxml: any;
        [xAx_cxml, expected_value] = $list(request.param);
        base_axis = $call(_BaseAxis, element(xAx_cxml));
        return [base_axis, expected_value];
      } },
    minor_gridlines_set_fixture: { args: ["request"], params: [["c:catAx", true, "c:catAx/c:minorGridlines"], ["c:catAx/c:minorGridlines", true, "c:catAx/c:minorGridlines"], ["c:catAx/c:minorGridlines", false, "c:catAx"], ["c:catAx", false, "c:catAx"], ["c:dateAx", true, "c:dateAx/c:minorGridlines"], ["c:dateAx/c:minorGridlines", true, "c:dateAx/c:minorGridlines"], ["c:dateAx/c:minorGridlines", false, "c:dateAx"], ["c:dateAx", false, "c:dateAx"], ["c:valAx", true, "c:valAx/c:minorGridlines"], ["c:valAx/c:minorGridlines", true, "c:valAx/c:minorGridlines"], ["c:valAx/c:minorGridlines", false, "c:valAx"], ["c:valAx", false, "c:valAx"]], fn: (self: any, request: any): any => {
        let base_axis: any, expected_xAx_cxml: any, expected_xml: any, new_value: any, xAx_cxml: any;
        [xAx_cxml, new_value, expected_xAx_cxml] = $list(request.param);
        base_axis = $call(_BaseAxis, element(xAx_cxml));
        expected_xml = xml(expected_xAx_cxml);
        return [base_axis, new_value, expected_xml];
      } },
    minor_tick_get_fixture: { args: ["request"], params: [["c:catAx", XL_TICK_MARK.CROSS], ["c:catAx/c:minorTickMark", XL_TICK_MARK.CROSS], ["c:catAx/c:minorTickMark{val=out}", XL_TICK_MARK.OUTSIDE], ["c:dateAx", XL_TICK_MARK.CROSS], ["c:dateAx/c:minorTickMark", XL_TICK_MARK.CROSS], ["c:dateAx/c:minorTickMark{val=out}", XL_TICK_MARK.OUTSIDE], ["c:valAx", XL_TICK_MARK.CROSS], ["c:valAx/c:minorTickMark", XL_TICK_MARK.CROSS], ["c:valAx/c:minorTickMark{val=in}", XL_TICK_MARK.INSIDE]], fn: (self: any, request: any): any => {
        let axis: any, expected_value: any, xAx_cxml: any;
        [xAx_cxml, expected_value] = $list(request.param);
        axis = $call(_BaseAxis, element(xAx_cxml));
        return [axis, expected_value];
      } },
    minor_tick_set_fixture: { args: ["request"], params: [["c:catAx", XL_TICK_MARK.INSIDE, "c:catAx/c:minorTickMark{val=in}"], ["c:catAx/c:minorTickMark{val=in}", XL_TICK_MARK.OUTSIDE, "c:catAx/c:minorTickMark{val=out}"], ["c:catAx/c:minorTickMark{val=out}", XL_TICK_MARK.CROSS, "c:catAx"], ["c:catAx", XL_TICK_MARK.CROSS, "c:catAx"], ["c:catAx/c:minorTickMark{val=cross}", XL_TICK_MARK.CROSS, "c:catAx"], ["c:dateAx", XL_TICK_MARK.INSIDE, "c:dateAx/c:minorTickMark{val=in}"], ["c:dateAx/c:minorTickMark{val=in}", XL_TICK_MARK.OUTSIDE, "c:dateAx/c:minorTickMark{val=out}"], ["c:dateAx/c:minorTickMark{val=out}", XL_TICK_MARK.CROSS, "c:dateAx"], ["c:dateAx", XL_TICK_MARK.CROSS, "c:dateAx"], ["c:dateAx/c:minorTickMark{val=cross}", XL_TICK_MARK.CROSS, "c:dateAx"], ["c:valAx", XL_TICK_MARK.INSIDE, "c:valAx/c:minorTickMark{val=in}"], ["c:valAx/c:minorTickMark{val=in}", XL_TICK_MARK.OUTSIDE, "c:valAx/c:minorTickMark{val=out}"], ["c:valAx/c:minorTickMark{val=out}", XL_TICK_MARK.CROSS, "c:valAx"], ["c:valAx", XL_TICK_MARK.CROSS, "c:valAx"], ["c:valAx/c:minorTickMark{val=cross}", XL_TICK_MARK.CROSS, "c:valAx"]], fn: (self: any, request: any): any => {
        let axis: any, expected_xAx_cxml: any, expected_xml: any, new_value: any, xAx_cxml: any;
        [xAx_cxml, new_value, expected_xAx_cxml] = $list(request.param);
        axis = $call(_BaseAxis, element(xAx_cxml));
        expected_xml = xml(expected_xAx_cxml);
        return [axis, new_value, expected_xml];
      } },
    reverse_order_get_fixture: { args: ["request"], params: [["c:catAx/c:scaling", false], ["c:valAx/c:scaling/c:orientation", false], ["c:catAx/c:scaling/c:orientation{val=minMax}", false], ["c:valAx/c:scaling/c:orientation{val=maxMin}", true]], fn: (self: any, request: any): any => {
        let expected_value: any, xAx_cxml: any;
        [xAx_cxml, expected_value] = $list(request.param);
        return [element(xAx_cxml), expected_value];
      } },
    reverse_order_set_fixture: { args: ["request"], params: [["c:catAx/c:scaling", false, "c:catAx/c:scaling"], ["c:catAx/c:scaling", true, "c:catAx/c:scaling/c:orientation{val=maxMin}"], ["c:valAx/c:scaling/c:orientation", false, "c:valAx/c:scaling"], ["c:valAx/c:scaling/c:orientation", true, "c:valAx/c:scaling/c:orientation{val=maxMin}"], ["c:dateAx/c:scaling/c:orientation{val=minMax}", false, "c:dateAx/c:scaling"], ["c:dateAx/c:scaling/c:orientation{val=minMax}", true, "c:dateAx/c:scaling/c:orientation{val=maxMin}"], ["c:catAx/c:scaling/c:orientation{val=maxMin}", false, "c:catAx/c:scaling"], ["c:catAx/c:scaling/c:orientation{val=maxMin}", true, "c:catAx/c:scaling/c:orientation{val=maxMin}"]], fn: (self: any, request: any): any => {
        let expected_xAx_cxml: any, expected_xml: any, new_value: any, xAx: any, xAx_cxml: any;
        [xAx_cxml, new_value, expected_xAx_cxml] = $list(request.param);
        [xAx, expected_xml] = $list([element(xAx_cxml), xml(expected_xAx_cxml)]);
        return [xAx, new_value, expected_xml];
      } },
    tick_labels_fixture: { args: ["request", "TickLabels_", "tick_labels_"], params: ["c:catAx", "c:dateAx", "c:valAx"], fn: (self: any, request: any, TickLabels_: any, tick_labels_: any): any => {
        let axis: any, xAx: any, xAx_cxml: any;
        xAx_cxml = request.param;
        xAx = element(xAx_cxml);
        axis = $call(_BaseAxis, xAx);
        return [axis, tick_labels_, TickLabels_, xAx];
      } },
    tick_lbl_pos_get_fixture: { args: ["request"], params: [["c:catAx", XL_TICK_LBL_POS.NEXT_TO_AXIS], ["c:catAx/c:tickLblPos", XL_TICK_LBL_POS.NEXT_TO_AXIS], ["c:catAx/c:tickLblPos{val=high}", XL_TICK_LBL_POS.HIGH], ["c:dateAx", XL_TICK_LBL_POS.NEXT_TO_AXIS], ["c:dateAx/c:tickLblPos", XL_TICK_LBL_POS.NEXT_TO_AXIS], ["c:dateAx/c:tickLblPos{val=low}", XL_TICK_LBL_POS.LOW], ["c:valAx", XL_TICK_LBL_POS.NEXT_TO_AXIS], ["c:valAx/c:tickLblPos", XL_TICK_LBL_POS.NEXT_TO_AXIS], ["c:valAx/c:tickLblPos{val=none}", XL_TICK_LBL_POS.NONE]], fn: (self: any, request: any): any => {
        let axis: any, expected_value: any, xAx_cxml: any;
        [xAx_cxml, expected_value] = $list(request.param);
        axis = $call(_BaseAxis, element(xAx_cxml));
        return [axis, expected_value];
      } },
    tick_lbl_pos_set_fixture: { args: ["request"], params: [["c:catAx", XL_TICK_LBL_POS.HIGH, "c:catAx/c:tickLblPos{val=high}"], ["c:catAx/c:tickLblPos{val=high}", XL_TICK_LBL_POS.LOW, "c:catAx/c:tickLblPos{val=low}"], ["c:catAx/c:tickLblPos{val=low}", null, "c:catAx/c:tickLblPos"], ["c:catAx", null, "c:catAx/c:tickLblPos"], ["c:dateAx", XL_TICK_LBL_POS.NEXT_TO_AXIS, "c:dateAx/c:tickLblPos{val=nextTo}"], ["c:dateAx/c:tickLblPos{val=nextTo}", XL_TICK_LBL_POS.NONE, "c:dateAx/c:tickLblPos{val=none}"], ["c:dateAx/c:tickLblPos{val=none}", null, "c:dateAx/c:tickLblPos"], ["c:valAx", XL_TICK_LBL_POS.HIGH, "c:valAx/c:tickLblPos{val=high}"], ["c:valAx/c:tickLblPos{val=high}", XL_TICK_LBL_POS.LOW, "c:valAx/c:tickLblPos{val=low}"], ["c:valAx/c:tickLblPos{val=low}", null, "c:valAx/c:tickLblPos"]], fn: (self: any, request: any): any => {
        let axis: any, expected_xAx_cxml: any, expected_xml: any, new_value: any, xAx_cxml: any;
        [xAx_cxml, new_value, expected_xAx_cxml] = $list(request.param);
        axis = $call(_BaseAxis, element(xAx_cxml));
        expected_xml = xml(expected_xAx_cxml);
        return [axis, new_value, expected_xml];
      } },
    title_fixture: { args: ["request", "AxisTitle_", "axis_title_"], params: ["c:catAx", "c:dateAx", "c:valAx"], fn: (self: any, request: any, AxisTitle_: any, axis_title_: any): any => {
        let axis: any, xAx_cxml: any;
        xAx_cxml = request.param;
        axis = $call(_BaseAxis, element(xAx_cxml));
        return [axis, AxisTitle_, axis_title_];
      } },
    visible_get_fixture: { args: ["request"], params: [["c:catAx", false], ["c:catAx/c:delete", false], ["c:catAx/c:delete{val=0}", true], ["c:catAx/c:delete{val=1}", false], ["c:catAx/c:delete{val=false}", true], ["c:dateAx", false], ["c:dateAx/c:delete", false], ["c:dateAx/c:delete{val=0}", true], ["c:dateAx/c:delete{val=1}", false], ["c:dateAx/c:delete{val=false}", true], ["c:valAx", false], ["c:valAx/c:delete", false], ["c:valAx/c:delete{val=0}", true], ["c:valAx/c:delete{val=1}", false], ["c:valAx/c:delete{val=false}", true]], fn: (self: any, request: any): any => {
        let axis: any, expected_bool_value: any, xAx_cxml: any;
        [xAx_cxml, expected_bool_value] = $list(request.param);
        axis = $call(_BaseAxis, element(xAx_cxml));
        return [axis, expected_bool_value];
      } },
    visible_set_fixture: { args: ["request"], params: [["c:catAx", false, "c:catAx/c:delete"], ["c:catAx/c:delete", true, "c:catAx/c:delete{val=0}"], ["c:catAx/c:delete{val=1}", true, "c:catAx/c:delete{val=0}"], ["c:catAx/c:delete{val=0}", false, "c:catAx/c:delete"], ["c:catAx", true, "c:catAx/c:delete{val=0}"], ["c:dateAx", false, "c:dateAx/c:delete"], ["c:dateAx/c:delete", true, "c:dateAx/c:delete{val=0}"], ["c:dateAx/c:delete{val=0}", false, "c:dateAx/c:delete"], ["c:dateAx", true, "c:dateAx/c:delete{val=0}"], ["c:valAx/c:delete", true, "c:valAx/c:delete{val=0}"], ["c:valAx/c:delete{val=1}", false, "c:valAx/c:delete"], ["c:valAx/c:delete{val=0}", true, "c:valAx/c:delete{val=0}"]], fn: (self: any, request: any): any => {
        let axis: any, expected_xAx_cxml: any, expected_xml: any, new_value: any, xAx_cxml: any;
        [xAx_cxml, new_value, expected_xAx_cxml] = $list(request.param);
        axis = $call(_BaseAxis, element(xAx_cxml));
        expected_xml = xml(expected_xAx_cxml);
        return [axis, new_value, expected_xml];
      } },
    AxisTitle_: { args: ["request", "axis_title_"], fn: (self: any, request: any, axis_title_: any): any => {
        return $call(classMock, request, "pptx.chart.axis.AxisTitle", $kw({return_value: axis_title_}));
      } },
    axis_title_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, AxisTitle);
      } },
    ChartFormat_: { args: ["request", "format_"], fn: (self: any, request: any, format_: any): any => {
        return $call(classMock, request, "pptx.chart.axis.ChartFormat", $kw({return_value: format_}));
      } },
    format_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, ChartFormat);
      } },
    MajorGridlines_: { args: ["request", "major_gridlines_"], fn: (self: any, request: any, major_gridlines_: any): any => {
        return $call(classMock, request, "pptx.chart.axis.MajorGridlines", $kw({return_value: major_gridlines_}));
      } },
    major_gridlines_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, MajorGridlines);
      } },
    TickLabels_: { args: ["request", "tick_labels_"], fn: (self: any, request: any, tick_labels_: any): any => {
        return $call(classMock, request, "pptx.chart.axis.TickLabels", $kw({return_value: tick_labels_}));
      } },
    tick_labels_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, TickLabels);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_access_to_its_title", args: ["title_fixture"], fn: (self: any, title_fixture: any): any => {
        let AxisTitle_: any, axis: any, axis_title: any, axis_title_: any;
        [axis, AxisTitle_, axis_title_] = $list(title_fixture);
        axis_title = axis.axisTitle;
        AxisTitle_.assertCalledOnceWith(axis._element.title);
        $assert($is(axis_title, axis_title_));
      } },
    { name: "it_provides_access_to_its_format", args: ["format_fixture"], fn: (self: any, format_fixture: any): any => {
        let ChartFormat_: any, axis: any, format: any, format_: any;
        [axis, ChartFormat_, format_] = $list(format_fixture);
        format = axis.format;
        ChartFormat_.assertCalledOnceWith(axis._xAx);
        $assert($is(format, format_));
      } },
    { name: "it_knows_whether_it_has_major_gridlines", args: ["major_gridlines_get_fixture"], fn: (self: any, major_gridlines_get_fixture: any): any => {
        let base_axis: any, expected_value: any;
        [base_axis, expected_value] = $list(major_gridlines_get_fixture);
        $assert($is(base_axis.hasMajorGridlines, expected_value));
      } },
    { name: "it_can_change_whether_it_has_major_gridlines", args: ["major_gridlines_set_fixture"], fn: (self: any, major_gridlines_set_fixture: any): any => {
        let base_axis: any, expected_xml: any, new_value: any;
        [base_axis, new_value, expected_xml] = $list(major_gridlines_set_fixture);
        base_axis.hasMajorGridlines = new_value;
        $assertEqual(base_axis._element.xml, expected_xml);
      } },
    { name: "it_knows_whether_it_has_minor_gridlines", args: ["minor_gridlines_get_fixture"], fn: (self: any, minor_gridlines_get_fixture: any): any => {
        let base_axis: any, expected_value: any;
        [base_axis, expected_value] = $list(minor_gridlines_get_fixture);
        $assert($is(base_axis.hasMinorGridlines, expected_value));
      } },
    { name: "it_can_change_whether_it_has_minor_gridlines", args: ["minor_gridlines_set_fixture"], fn: (self: any, minor_gridlines_set_fixture: any): any => {
        let base_axis: any, expected_xml: any, new_value: any;
        [base_axis, new_value, expected_xml] = $list(minor_gridlines_set_fixture);
        base_axis.hasMinorGridlines = new_value;
        $assertEqual(base_axis._element.xml, expected_xml);
      } },
    { name: "it_knows_whether_it_has_a_title", args: ["has_title_get_fixture"], fn: (self: any, has_title_get_fixture: any): any => {
        let axis: any, expected_value: any;
        [axis, expected_value] = $list(has_title_get_fixture);
        $assert($is(axis.hasTitle, expected_value));
      } },
    { name: "it_can_change_whether_it_has_a_title", args: ["has_title_set_fixture"], fn: (self: any, has_title_set_fixture: any): any => {
        let axis: any, expected_xml: any, new_value: any;
        [axis, new_value, expected_xml] = $list(has_title_set_fixture);
        axis.hasTitle = new_value;
        $assertEqual(axis._element.xml, expected_xml);
      } },
    { name: "it_provides_access_to_its_major_gridlines", args: ["maj_grdlns_fixture"], fn: (self: any, maj_grdlns_fixture: any): any => {
        let MajorGridlines_: any, axis: any, major_gridlines: any, major_gridlines_: any, xAx: any;
        [axis, MajorGridlines_, xAx, major_gridlines_] = $list(maj_grdlns_fixture);
        major_gridlines = axis.majorGridlines;
        MajorGridlines_.assertCalledOnceWith(xAx);
        $assert($is(major_gridlines, major_gridlines_));
      } },
    { name: "it_knows_its_major_tick_setting", args: ["major_tick_get_fixture"], fn: (self: any, major_tick_get_fixture: any): any => {
        let axis: any, expected_value: any;
        [axis, expected_value] = $list(major_tick_get_fixture);
        $assertEqual(axis.majorTickMark, expected_value);
      } },
    { name: "it_can_change_its_major_tick_mark", args: ["major_tick_set_fixture"], fn: (self: any, major_tick_set_fixture: any): any => {
        let axis: any, expected_xml: any, new_value: any;
        [axis, new_value, expected_xml] = $list(major_tick_set_fixture);
        axis.majorTickMark = new_value;
        $assertEqual(axis._element.xml, expected_xml);
      } },
    { name: "it_knows_its_maximum_scale", args: ["maximum_scale_get_fixture"], fn: (self: any, maximum_scale_get_fixture: any): any => {
        let axis: any, expected_value: any;
        [axis, expected_value] = $list(maximum_scale_get_fixture);
        $assertEqual(axis.maximumScale, expected_value);
      } },
    { name: "it_can_change_its_maximum_scale", args: ["maximum_scale_set_fixture"], fn: (self: any, maximum_scale_set_fixture: any): any => {
        let axis: any, expected_xml: any, new_value: any;
        [axis, new_value, expected_xml] = $list(maximum_scale_set_fixture);
        axis.maximumScale = new_value;
        $assertEqual(axis._element.xml, expected_xml);
      } },
    { name: "it_knows_its_minimum_scale", args: ["minimum_scale_get_fixture"], fn: (self: any, minimum_scale_get_fixture: any): any => {
        let axis: any, expected_value: any;
        [axis, expected_value] = $list(minimum_scale_get_fixture);
        $assertEqual(axis.minimumScale, expected_value);
      } },
    { name: "it_can_change_its_minimum_scale", args: ["minimum_scale_set_fixture"], fn: (self: any, minimum_scale_set_fixture: any): any => {
        let axis: any, expected_xml: any, new_value: any;
        [axis, new_value, expected_xml] = $list(minimum_scale_set_fixture);
        axis.minimumScale = new_value;
        $assertEqual(axis._element.xml, expected_xml);
      } },
    { name: "it_knows_its_minor_tick_setting", args: ["minor_tick_get_fixture"], fn: (self: any, minor_tick_get_fixture: any): any => {
        let axis: any, expected_value: any;
        [axis, expected_value] = $list(minor_tick_get_fixture);
        $assertEqual(axis.minorTickMark, expected_value);
      } },
    { name: "it_can_change_its_minor_tick_mark", args: ["minor_tick_set_fixture"], fn: (self: any, minor_tick_set_fixture: any): any => {
        let axis: any, expected_xml: any, new_value: any;
        [axis, new_value, expected_xml] = $list(minor_tick_set_fixture);
        axis.minorTickMark = new_value;
        $assertEqual(axis._element.xml, expected_xml);
      } },
    { name: "it_knows_whether_it_renders_in_reverse_order", args: ["reverse_order_get_fixture"], fn: (self: any, reverse_order_get_fixture: any): any => {
        let expected_value: any, xAx: any;
        [xAx, expected_value] = $list(reverse_order_get_fixture);
        $assertEqual($call(_BaseAxis, xAx).reverseOrder, expected_value);
      } },
    { name: "it_can_change_whether_it_renders_in_reverse_order", args: ["reverse_order_set_fixture"], fn: (self: any, reverse_order_set_fixture: any): any => {
        let axis: any, expected_xml: any, new_value: any, xAx: any;
        [xAx, new_value, expected_xml] = $list(reverse_order_set_fixture);
        axis = $call(_BaseAxis, xAx);
        axis.reverseOrder = new_value;
        $assertEqual(axis._element.xml, expected_xml);
      } },
    { name: "it_knows_its_tick_label_position", args: ["tick_lbl_pos_get_fixture"], fn: (self: any, tick_lbl_pos_get_fixture: any): any => {
        let axis: any, expected_value: any;
        [axis, expected_value] = $list(tick_lbl_pos_get_fixture);
        $assertEqual(axis.tickLabelPosition, expected_value);
      } },
    { name: "it_can_change_its_tick_label_position", args: ["tick_lbl_pos_set_fixture"], fn: (self: any, tick_lbl_pos_set_fixture: any): any => {
        let axis: any, expected_xml: any, new_value: any;
        [axis, new_value, expected_xml] = $list(tick_lbl_pos_set_fixture);
        axis.tickLabelPosition = new_value;
        $assertEqual(axis._element.xml, expected_xml);
      } },
    { name: "it_provides_access_to_the_tick_labels", args: ["tick_labels_fixture"], fn: (self: any, tick_labels_fixture: any): any => {
        let TickLabels_: any, axis: any, tick_labels: any, tick_labels_: any, xAx: any;
        [axis, tick_labels_, TickLabels_, xAx] = $list(tick_labels_fixture);
        tick_labels = axis.tickLabels;
        TickLabels_.assertCalledOnceWith(xAx);
        $assert($is(tick_labels, tick_labels_));
      } },
    { name: "it_knows_whether_it_is_visible", args: ["visible_get_fixture"], fn: (self: any, visible_get_fixture: any): any => {
        let axis: any, expected_bool_value: any;
        [axis, expected_bool_value] = $list(visible_get_fixture);
        $assert($is(axis.visible, expected_bool_value));
      } },
    { name: "it_can_change_whether_it_is_visible", args: ["visible_set_fixture"], fn: (self: any, visible_set_fixture: any): any => {
        let axis: any, expected_xml: any, new_value: any;
        [axis, new_value, expected_xml] = $list(visible_set_fixture);
        axis.visible = new_value;
        $assertEqual(axis._element.xml, expected_xml);
      } },
    { name: "but_it_raises_on_assign_non_bool_to_visible", args: [], fn: (self: any): any => {
        let axis: any;
        axis = $call(_BaseAxis, null);
        $raises(ValueError, () => {
          axis.visible = "foobar";
        });
      } },
  ],
};
pytestClass("Describe_BaseAxis", Describe_BaseAxis);

const DescribeAxisTitle: any = {
  fixtures: {
    format_fixture: { args: ["request", "ChartFormat_", "format_"], fn: (self: any, request: any, ChartFormat_: any, format_: any): any => {
        let axis_title: any;
        axis_title = $call(AxisTitle, element("c:title"));
        return [axis_title, ChartFormat_, format_];
      } },
    has_tf_get_fixture: { args: ["request"], params: [["c:title", false], ["c:title/c:tx", false], ["c:title/c:tx/c:strRef", false], ["c:title/c:tx/c:rich", true]], fn: (self: any, request: any): any => {
        let axis_title: any, expected_value: any, title_cxml: any;
        [title_cxml, expected_value] = $list(request.param);
        axis_title = $call(AxisTitle, element(title_cxml));
        return [axis_title, expected_value];
      } },
    has_tf_set_fixture: { args: ["request"], params: [["c:title{a:b=c}", true, "c:title{a:b=c}/c:tx/c:rich/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr)"], ["c:title{a:b=c}/c:tx", true, "c:title{a:b=c}/c:tx/c:rich/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr)"], ["c:title{a:b=c}/c:tx/c:strRef", true, "c:title{a:b=c}/c:tx/c:rich/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr)"], ["c:title/c:tx/c:rich", true, "c:title/c:tx/c:rich"], ["c:title", false, "c:title"], ["c:title/c:tx", false, "c:title"], ["c:title/c:tx/c:rich", false, "c:title"], ["c:title/c:tx/c:strRef", false, "c:title"]], fn: (self: any, request: any): any => {
        let axis_title: any, expected_cxml: any, expected_xml: any, title_cxml: any, value: any;
        [title_cxml, value, expected_cxml] = $list(request.param);
        axis_title = $call(AxisTitle, element(title_cxml));
        expected_xml = xml(expected_cxml);
        return [axis_title, value, expected_xml];
      } },
    text_frame_fixture: { args: ["request", "TextFrame_"], fn: (self: any, request: any, TextFrame_: any): any => {
        let axis_title: any, text_frame_: any;
        axis_title = $call(AxisTitle, element("c:title"));
        text_frame_ = TextFrame_.returnValue;
        return [axis_title, TextFrame_, text_frame_];
      } },
    ChartFormat_: { args: ["request", "format_"], fn: (self: any, request: any, format_: any): any => {
        return $call(classMock, request, "pptx.chart.axis.ChartFormat", $kw({return_value: format_}));
      } },
    format_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, ChartFormat);
      } },
    TextFrame_: { args: ["request"], fn: (self: any, request: any): any => {
        return classMock(request, "pptx.chart.axis.TextFrame");
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_whether_it_has_a_text_frame", args: ["has_tf_get_fixture"], fn: (self: any, has_tf_get_fixture: any): any => {
        let axis_title: any, expected_value: any, value: any;
        [axis_title, expected_value] = $list(has_tf_get_fixture);
        value = axis_title.hasTextFrame;
        $assert($is(value, expected_value));
      } },
    { name: "it_can_change_whether_it_has_a_text_frame", args: ["has_tf_set_fixture"], fn: (self: any, has_tf_set_fixture: any): any => {
        let axis_title: any, expected_xml: any, value: any;
        [axis_title, value, expected_xml] = $list(has_tf_set_fixture);
        axis_title.hasTextFrame = value;
        $assertEqual(axis_title._element.xml, expected_xml);
      } },
    { name: "it_provides_access_to_its_format", args: ["format_fixture"], fn: (self: any, format_fixture: any): any => {
        let ChartFormat_: any, axis_title: any, format: any, format_: any;
        [axis_title, ChartFormat_, format_] = $list(format_fixture);
        format = axis_title.format;
        ChartFormat_.assertCalledOnceWith(axis_title._element);
        $assert($is(format, format_));
      } },
    { name: "it_provides_access_to_its_text_frame", args: ["text_frame_fixture"], fn: (self: any, text_frame_fixture: any): any => {
        let TextFrame_: any, axis_title: any, text_frame: any, text_frame_: any;
        [axis_title, TextFrame_, text_frame_] = $list(text_frame_fixture);
        text_frame = axis_title.textFrame;
        TextFrame_.assertCalledOnceWith(axis_title._element.tx.rich, axis_title);
        $assert($is(text_frame, text_frame_));
      } },
  ],
};
pytestClass("DescribeAxisTitle", DescribeAxisTitle);

const DescribeCategoryAxis: any = {
  fixtures: {
    cat_type_get_fixture: { args: [], fn: (self: any): any => {
        let category_axis: any, expected_value: any;
        category_axis = $call(CategoryAxis, null);
        expected_value = XL_CATEGORY_TYPE.CATEGORY_SCALE;
        return [category_axis, expected_value];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_category_type", args: ["cat_type_get_fixture"], fn: (self: any, cat_type_get_fixture: any): any => {
        let category_axis: any, expected_value: any;
        [category_axis, expected_value] = $list(cat_type_get_fixture);
        $assert($is(category_axis.categoryType, expected_value));
      } },
  ],
};
pytestClass("DescribeCategoryAxis", DescribeCategoryAxis);

const DescribeDateAxis: any = {
  fixtures: {
    cat_type_get_fixture: { args: [], fn: (self: any): any => {
        let date_axis: any, expected_value: any;
        date_axis = $call(DateAxis, null);
        expected_value = XL_CATEGORY_TYPE.TIME_SCALE;
        return [date_axis, expected_value];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_category_type", args: ["cat_type_get_fixture"], fn: (self: any, cat_type_get_fixture: any): any => {
        let date_axis: any, expected_value: any;
        [date_axis, expected_value] = $list(cat_type_get_fixture);
        $assert($is(date_axis.categoryType, expected_value));
      } },
  ],
};
pytestClass("DescribeDateAxis", DescribeDateAxis);

const DescribeMajorGridlines: any = {
  fixtures: {
    format_fixture: { args: ["request", "ChartFormat_", "format_"], params: [["c:valAx", "c:valAx/c:majorGridlines"], ["c:catAx/c:majorGridlines", "c:catAx/c:majorGridlines"]], fn: (self: any, request: any, ChartFormat_: any, format_: any): any => {
        let expected_cxml: any, expected_xml: any, gridlines: any, xAx_cxml: any;
        [xAx_cxml, expected_cxml] = $list(request.param);
        gridlines = $call(MajorGridlines, element(xAx_cxml));
        expected_xml = xml(expected_cxml);
        return [gridlines, expected_xml, ChartFormat_, format_];
      } },
    ChartFormat_: { args: ["request", "format_"], fn: (self: any, request: any, format_: any): any => {
        return $call(classMock, request, "pptx.chart.axis.ChartFormat", $kw({return_value: format_}));
      } },
    format_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, ChartFormat);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_access_to_its_format", args: ["format_fixture"], fn: (self: any, format_fixture: any): any => {
        let ChartFormat_: any, expected_xml: any, format: any, format_: any, gridlines: any;
        [gridlines, expected_xml, ChartFormat_, format_] = $list(format_fixture);
        format = gridlines.format;
        $assertEqual(gridlines._xAx.xml, expected_xml);
        ChartFormat_.assertCalledOnceWith($get(gridlines._xAx.xpath("c:majorGridlines"), 0));
        $assert($is(format, format_));
      } },
  ],
};
pytestClass("DescribeMajorGridlines", DescribeMajorGridlines);

const DescribeTickLabels: any = {
  fixtures: {
    font_fixture: { args: ["Font_", "font_"], fn: (self: any, Font_: any, font_: any): any => {
        let catAx: any, defRPr: any, tick_labels: any;
        catAx = element("c:catAx/c:txPr/a:p/a:pPr/a:defRPr");
        defRPr = $get(catAx.xpath(".//a:defRPr"), 0);
        tick_labels = $call(TickLabels, catAx);
        return [tick_labels, Font_, defRPr, font_];
      } },
    number_format_get_fixture: { args: ["request"], params: [["c:catAx", "General"], ["c:valAx/c:numFmt{formatCode=General}", "General"]], fn: (self: any, request: any): any => {
        let expected_value: any, tick_labels: any, xAx_cxml: any;
        [xAx_cxml, expected_value] = $list(request.param);
        tick_labels = $call(TickLabels, element(xAx_cxml));
        return [tick_labels, expected_value];
      } },
    number_format_set_fixture: { args: ["request"], params: [["c:catAx", "General", "c:catAx/c:numFmt{formatCode=General,sourceLinked=0}"], ["c:valAx/c:numFmt{formatCode=General}", "00.00", "c:valAx/c:numFmt{formatCode=00.00,sourceLinked=0}"]], fn: (self: any, request: any): any => {
        let expected_xAx_cxml: any, expected_xml: any, new_value: any, tick_labels: any, xAx_cxml: any;
        [xAx_cxml, new_value, expected_xAx_cxml] = $list(request.param);
        tick_labels = $call(TickLabels, element(xAx_cxml));
        expected_xml = xml(expected_xAx_cxml);
        return [tick_labels, new_value, expected_xml];
      } },
    number_format_is_linked_get_fixture: { args: ["request"], params: [["c:catAx", false], ["c:valAx/c:numFmt", true], ["c:valAx/c:numFmt{sourceLinked=0}", false], ["c:catAx/c:numFmt{sourceLinked=1}", true]], fn: (self: any, request: any): any => {
        let expected_value: any, tick_labels: any, xAx_cxml: any;
        [xAx_cxml, expected_value] = $list(request.param);
        tick_labels = $call(TickLabels, element(xAx_cxml));
        return [tick_labels, expected_value];
      } },
    number_format_is_linked_set_fixture: { args: ["request"], params: [["c:valAx", true, "c:valAx/c:numFmt{sourceLinked=1}"], ["c:catAx", false, "c:catAx/c:numFmt{sourceLinked=0}"], ["c:valAx", null, "c:valAx/c:numFmt"], ["c:catAx/c:numFmt", true, "c:catAx/c:numFmt{sourceLinked=1}"], ["c:valAx/c:numFmt{sourceLinked=1}", false, "c:valAx/c:numFmt{sourceLinked=0}"]], fn: (self: any, request: any): any => {
        let expected_xAx_cxml: any, expected_xml: any, new_value: any, tick_labels: any, xAx_cxml: any;
        [xAx_cxml, new_value, expected_xAx_cxml] = $list(request.param);
        tick_labels = $call(TickLabels, element(xAx_cxml));
        expected_xml = xml(expected_xAx_cxml);
        return [tick_labels, new_value, expected_xml];
      } },
    offset_get_fixture: { args: ["request"], params: [["c:catAx", 100], ["c:catAx/c:lblOffset", 100], ["c:catAx/c:lblOffset{val=420}", 420], ["c:catAx/c:lblOffset{val=004}", 4], ["c:catAx/c:lblOffset{val=42%}", 42], ["c:catAx/c:lblOffset{val=02%}", 2]], fn: (self: any, request: any): any => {
        let catAx_cxml: any, expected_value: any, tick_labels: any;
        [catAx_cxml, expected_value] = $list(request.param);
        tick_labels = $call(TickLabels, element(catAx_cxml));
        return [tick_labels, expected_value];
      } },
    offset_set_fixture: { args: ["request"], params: [["c:catAx", 420, "c:catAx/c:lblOffset{val=420}"], ["c:catAx/c:lblOffset{val=420}", 100, "c:catAx"]], fn: (self: any, request: any): any => {
        let catAx_cxml: any, expected_catAx_cxml: any, expected_xml: any, new_value: any, tick_labels: any;
        [catAx_cxml, new_value, expected_catAx_cxml] = $list(request.param);
        tick_labels = $call(TickLabels, element(catAx_cxml));
        expected_xml = xml(expected_catAx_cxml);
        return [tick_labels, new_value, expected_xml];
      } },
    txPr_fixture: { args: ["request"], params: [["c:valAx{a:b=c}", "c:valAx{a:b=c}/c:txPr/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr)"], ["c:valAx{a:b=c}/c:txPr/(a:bodyPr,a:p)", "c:valAx{a:b=c}/c:txPr/(a:bodyPr,a:p/a:pPr/a:defRPr)"], ["c:valAx{a:b=c}/c:txPr/(a:bodyPr,a:p/a:pPr)", "c:valAx{a:b=c}/c:txPr/(a:bodyPr,a:p/a:pPr/a:defRPr)"]], fn: (self: any, request: any): any => {
        let expected_cxml: any, expected_xml: any, tick_labels: any, xAx_cxml: any;
        [xAx_cxml, expected_cxml] = $list(request.param);
        tick_labels = $call(TickLabels, element(xAx_cxml));
        expected_xml = xml(expected_cxml);
        return [tick_labels, expected_xml];
      } },
    Font_: { args: ["request", "font_"], fn: (self: any, request: any, font_: any): any => {
        return $call(classMock, request, "pptx.chart.axis.Font", $kw({return_value: font_}));
      } },
    font_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Font);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_access_to_its_font", args: ["font_fixture"], fn: (self: any, font_fixture: any): any => {
        let Font_: any, defRPr: any, font: any, font_: any, tick_labels: any;
        [tick_labels, Font_, defRPr, font_] = $list(font_fixture);
        font = tick_labels.font;
        Font_.assertCalledOnceWith(defRPr);
        $assert($is(font, font_));
      } },
    { name: "it_adds_a_txPr_to_help_font", args: ["txPr_fixture"], fn: (self: any, txPr_fixture: any): any => {
        let expected_xml: any, tick_labels: any;
        [tick_labels, expected_xml] = $list(txPr_fixture);
        tick_labels.font;
        $assertEqual(tick_labels._element.xml, expected_xml);
      } },
    { name: "it_knows_its_number_format", args: ["number_format_get_fixture"], fn: (self: any, number_format_get_fixture: any): any => {
        let expected_value: any, tick_labels: any;
        [tick_labels, expected_value] = $list(number_format_get_fixture);
        $assertEqual(tick_labels.numberFormat, expected_value);
      } },
    { name: "it_can_change_its_number_format", args: ["number_format_set_fixture"], fn: (self: any, number_format_set_fixture: any): any => {
        let expected_xml: any, new_value: any, tick_labels: any;
        [tick_labels, new_value, expected_xml] = $list(number_format_set_fixture);
        tick_labels.numberFormat = new_value;
        $assertEqual(tick_labels._element.xml, expected_xml);
      } },
    { name: "it_knows_whether_its_number_format_is_linked", args: ["number_format_is_linked_get_fixture"], fn: (self: any, number_format_is_linked_get_fixture: any): any => {
        let expected_value: any, tick_labels: any;
        [tick_labels, expected_value] = $list(number_format_is_linked_get_fixture);
        $assert($is(tick_labels.numberFormatIsLinked, expected_value));
      } },
    { name: "it_can_change_whether_its_number_format_is_linked", args: ["number_format_is_linked_set_fixture"], fn: (self: any, number_format_is_linked_set_fixture: any): any => {
        let expected_xml: any, new_value: any, tick_labels: any;
        [tick_labels, new_value, expected_xml] = $list(number_format_is_linked_set_fixture);
        tick_labels.numberFormatIsLinked = new_value;
        $assertEqual(tick_labels._element.xml, expected_xml);
      } },
    { name: "it_knows_its_offset", args: ["offset_get_fixture"], fn: (self: any, offset_get_fixture: any): any => {
        let expected_value: any, tick_labels: any;
        [tick_labels, expected_value] = $list(offset_get_fixture);
        $assertEqual(tick_labels.offset, expected_value);
      } },
    { name: "it_can_change_its_offset", args: ["offset_set_fixture"], fn: (self: any, offset_set_fixture: any): any => {
        let expected_xml: any, new_value: any, tick_labels: any;
        [tick_labels, new_value, expected_xml] = $list(offset_set_fixture);
        tick_labels.offset = new_value;
        $assertEqual(tick_labels._element.xml, expected_xml);
      } },
  ],
};
pytestClass("DescribeTickLabels", DescribeTickLabels);

const DescribeValueAxis: any = {
  fixtures: {
    crosses_at_get_fixture: { args: ["request"], params: [["c:plotArea/(c:valAx/c:axId{val=42},c:valAx/c:crossAx{val=42})", null], ["c:plotArea/(c:catAx/(c:axId{val=42},c:crossesAt{val=2.4}),c:valAx/c:crossAx{val=42})", 2.4], ["c:plotArea/(c:dateAx/(c:axId{val=42},c:crossesAt{val=-1.2}),c:valAx/c:crossAx{val=42})", -1.2]], fn: (self: any, request: any): any => {
        let expected_value: any, plotArea: any, plotArea_cxml: any, valAx: any, value_axis: any;
        [plotArea_cxml, expected_value] = $list(request.param);
        plotArea = element(plotArea_cxml);
        valAx = $get(plotArea.xpath("c:valAx[c:crossAx/@val=\"42\"]"), 0);
        value_axis = $call(ValueAxis, valAx);
        return [value_axis, expected_value];
      } },
    crosses_at_set_fixture: { args: ["request"], params: [["c:plotArea/(c:valAx/c:axId{val=42},c:valAx/c:crossAx{val=42})", 2.4, "c:plotArea/(c:valAx/(c:axId{val=42},c:crossesAt{val=2.4}),c:valAx/c:crossAx{val=42})"], ["c:plotArea/(c:catAx/(c:axId{val=42},c:crosses{val=min}),c:valAx/c:crossAx{val=42})", 1.5, "c:plotArea/(c:catAx/(c:axId{val=42},c:crossesAt{val=1.5}),c:valAx/c:crossAx{val=42})"], ["c:plotArea/(c:dateAx/(c:axId{val=42},c:crossesAt{val=2.4}),c:valAx/c:crossAx{val=42})", 1.5, "c:plotArea/(c:dateAx/(c:axId{val=42},c:crossesAt{val=1.5}),c:valAx/c:crossAx{val=42})"], ["c:plotArea/(c:catAx/(c:axId{val=42},c:crossesAt{val=1.5}),c:valAx/c:crossAx{val=42})", null, "c:plotArea/(c:catAx/(c:axId{val=42}),c:valAx/c:crossAx{val=42})"]], fn: (self: any, request: any): any => {
        let expected_cxml: any, expected_xml: any, new_value: any, plotArea: any, plotArea_cxml: any, valAx: any, value_axis: any;
        [plotArea_cxml, new_value, expected_cxml] = $list(request.param);
        plotArea = element(plotArea_cxml);
        valAx = $get(plotArea.xpath("c:valAx[c:crossAx/@val=\"42\"]"), 0);
        value_axis = $call(ValueAxis, valAx);
        expected_xml = xml(expected_cxml);
        return [value_axis, new_value, plotArea, expected_xml];
      } },
    crosses_get_fixture: { args: ["request"], params: [["c:plotArea/(c:valAx/c:axId{val=42},c:valAx/c:crossAx{val=42})", "CUSTOM"], ["c:plotArea/(c:catAx/(c:axId{val=42},c:crosses{val=autoZero}),c:valAx/c:crossAx{val=42})", "AUTOMATIC"], ["c:plotArea/(c:valAx/(c:axId{val=42},c:crosses{val=min}),c:valAx/c:crossAx{val=42})", "MINIMUM"]], fn: (self: any, request: any): any => {
        let cxml: any, expected_value: any, member: any, valAx: any, value_axis: any;
        [cxml, member] = $list(request.param);
        valAx = $get(element(cxml).xpath("c:valAx[c:crossAx/@val=\"42\"]"), 0);
        value_axis = $call(ValueAxis, valAx);
        expected_value = $getattr(XL_AXIS_CROSSES, member);
        return [value_axis, expected_value];
      } },
    crosses_set_fixture: { args: ["request"], params: [["c:plotArea/(c:valAx/(c:axId{val=42},c:crossesAt{val=2.4}),c:valAx/c:crossAx{val=42})", "AUTOMATIC", "c:plotArea/(c:valAx/(c:axId{val=42},c:crosses{val=autoZero}),c:valAx/c:crossAx{val=42})"], ["c:plotArea/(c:catAx/(c:axId{val=42},c:crosses{val=autoZero}),c:valAx/c:crossAx{val=42})", "MINIMUM", "c:plotArea/(c:catAx/(c:axId{val=42},c:crosses{val=min}),c:valAx/c:crossAx{val=42})"], ["c:plotArea/(c:valAx/(c:axId{val=42},c:crosses{val=min}),c:valAx/c:crossAx{val=42})", "CUSTOM", "c:plotArea/(c:valAx/(c:axId{val=42},c:crossesAt{val=0.0}),c:valAx/c:crossAx{val=42})"], ["c:plotArea/(c:catAx/(c:axId{val=42},c:crossesAt{val=2.4}),c:valAx/c:crossAx{val=42})", "CUSTOM", "c:plotArea/(c:catAx/(c:axId{val=42},c:crossesAt{val=2.4}),c:valAx/c:crossAx{val=42})"]], fn: (self: any, request: any): any => {
        let expected_cxml: any, expected_xml: any, member: any, new_value: any, plotArea: any, plotArea_cxml: any, valAx: any, value_axis: any;
        [plotArea_cxml, member, expected_cxml] = $list(request.param);
        plotArea = element(plotArea_cxml);
        valAx = $get(plotArea.xpath("c:valAx[c:crossAx/@val=\"42\"]"), 0);
        value_axis = $call(ValueAxis, valAx);
        new_value = $getattr(XL_AXIS_CROSSES, member);
        expected_xml = xml(expected_cxml);
        return [value_axis, new_value, plotArea, expected_xml];
      } },
    major_unit_get_fixture: { args: ["request"], params: [["c:valAx", null], ["c:valAx/c:majorUnit{val=4.2}", 4.2]], fn: (self: any, request: any): any => {
        let expected_value: any, valAx_cxml: any, value_axis: any;
        [valAx_cxml, expected_value] = $list(request.param);
        value_axis = $call(ValueAxis, element(valAx_cxml));
        return [value_axis, expected_value];
      } },
    major_unit_set_fixture: { args: ["request"], params: [["c:valAx", 42, "c:valAx/c:majorUnit{val=42.0}"], ["c:valAx", null, "c:valAx"], ["c:valAx/c:majorUnit{val=42.0}", 24.0, "c:valAx/c:majorUnit{val=24.0}"], ["c:valAx/c:majorUnit{val=42.0}", null, "c:valAx"]], fn: (self: any, request: any): any => {
        let expected_valAx_cxml: any, expected_xml: any, new_value: any, valAx_cxml: any, value_axis: any;
        [valAx_cxml, new_value, expected_valAx_cxml] = $list(request.param);
        value_axis = $call(ValueAxis, element(valAx_cxml));
        expected_xml = xml(expected_valAx_cxml);
        return [value_axis, new_value, expected_xml];
      } },
    minor_unit_get_fixture: { args: ["request"], params: [["c:valAx", null], ["c:valAx/c:minorUnit{val=2.4}", 2.4]], fn: (self: any, request: any): any => {
        let expected_value: any, valAx_cxml: any, value_axis: any;
        [valAx_cxml, expected_value] = $list(request.param);
        value_axis = $call(ValueAxis, element(valAx_cxml));
        return [value_axis, expected_value];
      } },
    minor_unit_set_fixture: { args: ["request"], params: [["c:valAx", 36, "c:valAx/c:minorUnit{val=36.0}"], ["c:valAx", null, "c:valAx"], ["c:valAx/c:minorUnit{val=36.0}", 12.6, "c:valAx/c:minorUnit{val=12.6}"], ["c:valAx/c:minorUnit{val=36.0}", null, "c:valAx"]], fn: (self: any, request: any): any => {
        let expected_valAx_cxml: any, expected_xml: any, new_value: any, valAx_cxml: any, value_axis: any;
        [valAx_cxml, new_value, expected_valAx_cxml] = $list(request.param);
        value_axis = $call(ValueAxis, element(valAx_cxml));
        expected_xml = xml(expected_valAx_cxml);
        return [value_axis, new_value, expected_xml];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_the_other_axis_crossing_type", args: ["crosses_get_fixture"], fn: (self: any, crosses_get_fixture: any): any => {
        let expected_value: any, value_axis: any;
        [value_axis, expected_value] = $list(crosses_get_fixture);
        $assertEqual(value_axis.crosses, expected_value);
      } },
    { name: "it_can_change_the_other_axis_crossing_type", args: ["crosses_set_fixture"], fn: (self: any, crosses_set_fixture: any): any => {
        let expected_xml: any, new_value: any, plotArea: any, value_axis: any;
        [value_axis, new_value, plotArea, expected_xml] = $list(crosses_set_fixture);
        value_axis.crosses = new_value;
        $assertEqual(plotArea.xml, expected_xml);
      } },
    { name: "it_knows_the_other_axis_crossing_value", args: ["crosses_at_get_fixture"], fn: (self: any, crosses_at_get_fixture: any): any => {
        let expected_value: any, value_axis: any;
        [value_axis, expected_value] = $list(crosses_at_get_fixture);
        $assertEqual(value_axis.crossesAt, expected_value);
      } },
    { name: "it_can_change_the_other_axis_crossing_value", args: ["crosses_at_set_fixture"], fn: (self: any, crosses_at_set_fixture: any): any => {
        let expected_xml: any, new_value: any, plotArea: any, value_axis: any;
        [value_axis, new_value, plotArea, expected_xml] = $list(crosses_at_set_fixture);
        value_axis.crossesAt = new_value;
        $assertEqual(plotArea.xml, expected_xml);
      } },
    { name: "it_knows_its_major_unit", args: ["major_unit_get_fixture"], fn: (self: any, major_unit_get_fixture: any): any => {
        let expected_value: any, value_axis: any;
        [value_axis, expected_value] = $list(major_unit_get_fixture);
        $assertEqual(value_axis.majorUnit, expected_value);
      } },
    { name: "it_can_change_its_major_unit", args: ["major_unit_set_fixture"], fn: (self: any, major_unit_set_fixture: any): any => {
        let expected_xml: any, new_value: any, value_axis: any;
        [value_axis, new_value, expected_xml] = $list(major_unit_set_fixture);
        value_axis.majorUnit = new_value;
        $assertEqual(value_axis._element.xml, expected_xml);
      } },
    { name: "it_knows_its_minor_unit", args: ["minor_unit_get_fixture"], fn: (self: any, minor_unit_get_fixture: any): any => {
        let expected_value: any, value_axis: any;
        [value_axis, expected_value] = $list(minor_unit_get_fixture);
        $assertEqual(value_axis.minorUnit, expected_value);
      } },
    { name: "it_can_change_its_minor_unit", args: ["minor_unit_set_fixture"], fn: (self: any, minor_unit_set_fixture: any): any => {
        let expected_xml: any, new_value: any, value_axis: any;
        [value_axis, new_value, expected_xml] = $list(minor_unit_set_fixture);
        value_axis.minorUnit = new_value;
        $assertEqual(value_axis._element.xml, expected_xml);
      } },
  ],
};
pytestClass("DescribeValueAxis", DescribeValueAxis);

