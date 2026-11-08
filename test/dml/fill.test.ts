// Machine-translated from python-pptx tests/dml/test_fill.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $add, $assert, $assertEqual, $call, $get, $is, $isinstance, $kw, $len, $list, $raises, $slice, $t1, $t2, $truthy, classMock, instanceMock, methodMock, propertyMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { ColorFormat as ColorFormat$imported } from "../../src/dml/color.js";
import { FillFormat as FillFormat$imported, _BlipFill as _BlipFill$imported, _Fill as _Fill$imported, _GradFill as _GradFill$imported, _GradientStop as _GradientStop$imported, _GradientStops as _GradientStops$imported, _GrpFill as _GrpFill$imported, _NoFill as _NoFill$imported, _NoneFill as _NoneFill$imported, _PattFill as _PattFill$imported, _SolidFill as _SolidFill$imported } from "../../src/dml/fill.js";
import { MSO_FILL as MSO_FILL$imported, MSO_PATTERN as MSO_PATTERN$imported } from "../../src/enum/dml.js";
import { CT_GradientStopList as CT_GradientStopList$imported } from "../../src/oxml/dml/fill.js";
import { element, xml } from "../unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const ColorFormat = ColorFormat$imported;
const FillFormat = FillFormat$imported;
const _BlipFill = _BlipFill$imported;
const _Fill = _Fill$imported;
const _GradFill = _GradFill$imported;
const _GradientStop = _GradientStop$imported;
const _GradientStops = _GradientStops$imported;
const _GrpFill = _GrpFill$imported;
const _NoFill = _NoFill$imported;
const _NoneFill = _NoneFill$imported;
const _PattFill = _PattFill$imported;
const _SolidFill = _SolidFill$imported;
const MSO_FILL = MSO_FILL$imported;
const MSO_PATTERN = MSO_PATTERN$imported;
const CT_GradientStopList = CT_GradientStopList$imported;
void $py;


const DescribeFillFormat: any = {
  fixtures: {
    back_color_fixture: { args: ["fill_", "color_"], fn: (self: any, fill_: any, color_: any): any => {
        let fill_format: any;
        fill_.backColor = color_;
        fill_format = $call(FillFormat, null, fill_);
        return [fill_format, color_];
      } },
    background_fixture: { args: ["request", "no_fill_"], params: [["p:spPr{a:b=c}", "p:spPr{a:b=c}/a:noFill"], ["a:tcPr/a:blipFill", "a:tcPr/a:noFill"], ["a:rPr/a:gradFill", "a:rPr/a:noFill"], ["a:tcPr/a:grpFill", "a:tcPr/a:noFill"], ["a:defRPr/a:noFill", "a:defRPr/a:noFill"], ["a:endParaRPr/a:pattFill", "a:endParaRPr/a:noFill"], ["a:rPr/a:solidFill", "a:rPr/a:noFill"]], fn: (self: any, request: any, no_fill_: any): any => {
        let _NoFill_: any, cxml: any, expected_cxml: any, expected_xml: any, fill: any;
        [cxml, expected_cxml] = $list(request.param);
        fill = FillFormat.fromFillParent(element(cxml));
        _NoFill_ = $call(classMock, request, "pptx.dml.fill._NoFill", $kw({return_value: no_fill_, autospec: true}));
        expected_xml = xml(expected_cxml);
        return [fill, _NoFill_, expected_xml, no_fill_];
      } },
    fill_type_fixture: { args: ["fill_"], fn: (self: any, fill_: any): any => {
        let expected_value: any, fill: any;
        const $t1 = 42;
        expected_value = $t1;
        fill_.type = $t1;
        fill = $call(FillFormat, null, fill_);
        return [fill, expected_value];
      } },
    fore_color_fixture: { args: ["fill_", "color_"], fn: (self: any, fill_: any, color_: any): any => {
        let fill_format: any;
        fill_.foreColor = color_;
        fill_format = $call(FillFormat, null, fill_);
        return [fill_format, color_];
      } },
    gradient_fixture: { args: ["request", "grad_fill_"], params: [["p:spPr{a:b=c}", "p:spPr{a:b=c}/a:gradFill"], ["p:bgPr/a:noFill", "p:bgPr/a:gradFill"], ["a:tcPr/a:blipFill", "a:tcPr/a:gradFill"], ["a:rPr/a:grpFill", "a:rPr/a:gradFill"], ["a:tcPr/a:noFill", "a:tcPr/a:gradFill"], ["a:defRPr/a:solidFill", "a:defRPr/a:gradFill"], ["a:endParaRPr/a:blipFill", "a:endParaRPr/a:gradFill"]], fn: (self: any, request: any, grad_fill_: any): any => {
        let _GradFill_: any, cxml: any, expected_cxml: any, expected_xml: any, fill: any;
        [cxml, expected_cxml] = $list(request.param);
        fill = FillFormat.fromFillParent(element(cxml));
        _GradFill_ = $call(classMock, request, "pptx.dml.fill._GradFill", $kw({return_value: grad_fill_, autospec: true}));
        expected_xml = xml($add(expected_cxml, "{rotWithShape=1}/(a:gsLst/(a:gs{pos=0}/a:schemeClr{val=accent1}/(a:tint{val=100000},a:shade{val=100000},a:satMod{val=130000}),a:gs{pos=100000}/a:schemeClr{val=accent1}/(a:tint{val=50000},a:shade{val=100000},a:satMod{val=350000})),a:lin{scaled=0})"));
        return [fill, _GradFill_, expected_xml, grad_fill_];
      } },
    pattern_get_fixture: { args: ["fill_"], fn: (self: any, fill_: any): any => {
        let expected_value: any, fill: any;
        const $t2 = MSO_PATTERN.WAVE;
        expected_value = $t2;
        fill_.pattern = $t2;
        fill = $call(FillFormat, null, fill_);
        return [fill, expected_value];
      } },
    pattern_set_fixture: { args: ["fill_"], fn: (self: any, fill_: any): any => {
        let fill: any, pattern: any;
        pattern = MSO_PATTERN.DIVOT;
        fill = $call(FillFormat, null, fill_);
        return [fill, pattern];
      } },
    patterned_fixture: { args: ["request", "patt_fill_"], params: [["p:spPr{a:b=c}", "p:spPr{a:b=c}/a:pattFill"], ["a:tcPr/a:blipFill", "a:tcPr/a:pattFill"], ["a:rPr/a:gradFill", "a:rPr/a:pattFill"], ["a:tcPr/a:grpFill", "a:tcPr/a:pattFill"], ["a:defRPr/a:noFill", "a:defRPr/a:pattFill"], ["a:endParaRPr/a:pattFill", "a:endParaRPr/a:pattFill"], ["a:rPr/a:solidFill", "a:rPr/a:pattFill"]], fn: (self: any, request: any, patt_fill_: any): any => {
        let _PattFill_: any, cxml: any, expected_cxml: any, expected_xml: any, fill: any;
        [cxml, expected_cxml] = $list(request.param);
        fill = FillFormat.fromFillParent(element(cxml));
        _PattFill_ = $call(classMock, request, "pptx.dml.fill._PattFill", $kw({return_value: patt_fill_, autospec: true}));
        expected_xml = xml(expected_cxml);
        return [fill, _PattFill_, expected_xml, patt_fill_];
      } },
    solid_fixture: { args: ["request", "solid_fill_"], params: [["p:spPr{a:b=c}", "p:spPr{a:b=c}/a:solidFill"], ["a:tcPr/a:blipFill", "a:tcPr/a:solidFill"], ["a:rPr/a:gradFill", "a:rPr/a:solidFill"], ["a:tcPr/a:grpFill", "a:tcPr/a:solidFill"], ["a:defRPr/a:noFill", "a:defRPr/a:solidFill"], ["a:rPr/a:pattFill", "a:rPr/a:solidFill"], ["a:endParaRPr/a:solidFill", "a:endParaRPr/a:solidFill"]], fn: (self: any, request: any, solid_fill_: any): any => {
        let _SolidFill_: any, cxml: any, expected_cxml: any, expected_xml: any, fill: any;
        [cxml, expected_cxml] = $list(request.param);
        fill = FillFormat.fromFillParent(element(cxml));
        _SolidFill_ = $call(classMock, request, "pptx.dml.fill._SolidFill", $kw({return_value: solid_fill_, autospec: true}));
        expected_xml = xml(expected_cxml);
        return [fill, _SolidFill_, expected_xml, solid_fill_];
      } },
    color_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, ColorFormat);
      } },
    fill_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _Fill);
      } },
    grad_fill_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _GradFill);
      } },
    gradient_stops_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _GradientStops);
      } },
    no_fill_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _NoFill);
      } },
    patt_fill_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _PattFill);
      } },
    solid_fill_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _SolidFill);
      } },
    type_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, FillFormat, "type");
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_set_the_fill_type_to_no_fill", args: ["background_fixture"], fn: (self: any, background_fixture: any): any => {
        let _NoFill_: any, expected_xml: any, fill: any, no_fill_: any;
        [fill, _NoFill_, expected_xml, no_fill_] = $list(background_fixture);
        fill.background();
        $assertEqual(fill._xPr.xml, expected_xml);
        _NoFill_.assertCalledOnceWith(fill._xPr.egFillProperties);
        $assert($is(fill._fill, no_fill_));
      } },
    { name: "it_can_set_the_fill_type_to_gradient", args: ["gradient_fixture"], fn: (self: any, gradient_fixture: any): any => {
        let _GradFill_: any, expected_xml: any, fill: any, grad_fill_: any;
        [fill, _GradFill_, expected_xml, grad_fill_] = $list(gradient_fixture);
        fill.gradient();
        $assertEqual(fill._xPr.xml, expected_xml);
        _GradFill_.assertCalledOnceWith(fill._xPr.egFillProperties);
        $assert($is(fill._fill, grad_fill_));
      } },
    { name: "it_can_set_the_fill_type_to_solid", args: ["solid_fixture"], fn: (self: any, solid_fixture: any): any => {
        let _SolidFill_: any, expected_xml: any, fill: any, solid_fill_: any;
        [fill, _SolidFill_, expected_xml, solid_fill_] = $list(solid_fixture);
        fill.solid();
        $assertEqual(fill._xPr.xml, expected_xml);
        _SolidFill_.assertCalledOnceWith(fill._xPr.egFillProperties);
        $assert($is(fill._fill, solid_fill_));
      } },
    { name: "it_can_set_the_fill_type_to_patterned", args: ["patterned_fixture"], fn: (self: any, patterned_fixture: any): any => {
        let _PattFill_: any, expected_xml: any, fill: any, patt_fill_: any;
        [fill, _PattFill_, expected_xml, patt_fill_] = $list(patterned_fixture);
        fill.patterned();
        $assertEqual(fill._xPr.xml, expected_xml);
        _PattFill_.assertCalledOnceWith(fill._xPr.egFillProperties);
        $assert($is(fill._fill, patt_fill_));
      } },
    { name: "it_provides_access_to_its_foreground_color", args: ["fore_color_fixture"], fn: (self: any, fore_color_fixture: any): any => {
        let color: any, color_: any, fill: any;
        [fill, color_] = $list(fore_color_fixture);
        color = fill.foreColor;
        $assert($is(color, color_));
      } },
    { name: "it_provides_access_to_its_background_color", args: ["back_color_fixture"], fn: (self: any, back_color_fixture: any): any => {
        let color: any, color_: any, fill: any;
        [fill, color_] = $list(back_color_fixture);
        color = fill.backColor;
        $assert($is(color, color_));
      } },
    { name: "it_knows_its_fill_type", args: ["fill_type_fixture"], fn: (self: any, fill_type_fixture: any): any => {
        let expected_value: any, fill: any, fill_type: any;
        [fill, expected_value] = $list(fill_type_fixture);
        fill_type = fill.type;
        $assertEqual(fill_type, expected_value);
      } },
    { name: "it_knows_the_angle_of_a_linear_gradient", args: ["grad_fill_", "type_prop_"], fn: (self: any, grad_fill_: any, type_prop_: any): any => {
        let angle: any, fill: any;
        grad_fill_.gradientAngle = 42.0;
        type_prop_.returnValue = MSO_FILL.GRADIENT;
        fill = $call(FillFormat, null, grad_fill_);
        angle = fill.gradientAngle;
        $assertEqual(angle, 42.0);
      } },
    { name: "it_can_change_the_angle_of_a_linear_gradient", args: ["grad_fill_", "type_prop_"], fn: (self: any, grad_fill_: any, type_prop_: any): any => {
        let fill: any;
        type_prop_.returnValue = MSO_FILL.GRADIENT;
        fill = $call(FillFormat, null, grad_fill_);
        fill.gradientAngle = 42.24;
        $assertEqual(grad_fill_.gradientAngle, 42.24);
      } },
    { name: "it_provides_access_to_the_gradient_stops", args: ["type_prop_", "grad_fill_", "gradient_stops_"], fn: (self: any, type_prop_: any, grad_fill_: any, gradient_stops_: any): any => {
        let fill: any, gradient_stops: any;
        type_prop_.returnValue = MSO_FILL.GRADIENT;
        grad_fill_.gradientStops = gradient_stops_;
        fill = $call(FillFormat, null, grad_fill_);
        gradient_stops = fill.gradientStops;
        $assert($is(gradient_stops, gradient_stops_));
      } },
    { name: "it_raises_on_non_gradient_fill", args: ["grad_fill_", "type_prop_"], fn: (self: any, grad_fill_: any, type_prop_: any): any => {
        let fill: any;
        type_prop_.returnValue = null;
        fill = $call(FillFormat, null, grad_fill_);
        $raises(TypeError, () => {
          fill.gradientAngle;
        });
        $raises(TypeError, () => {
          fill.gradientAngle = 123.4;
        });
        $raises(TypeError, () => {
          fill.gradientStops;
        });
      } },
    { name: "it_knows_its_pattern", args: ["pattern_get_fixture"], fn: (self: any, pattern_get_fixture: any): any => {
        let expected_value: any, fill: any, pattern: any;
        [fill, expected_value] = $list(pattern_get_fixture);
        pattern = fill.pattern;
        $assertEqual(pattern, expected_value);
      } },
    { name: "it_can_change_its_pattern", args: ["pattern_set_fixture"], fn: (self: any, pattern_set_fixture: any): any => {
        let fill: any, pattern: any;
        [fill, pattern] = $list(pattern_set_fixture);
        fill.pattern = pattern;
        $assert($is(fill.pattern, pattern));
      } },
  ],
};
pytestClass("DescribeFillFormat", DescribeFillFormat);

const Describe_Fill: any = {
  fixtures: {
    back_raise_fixture: { args: [], fn: (self: any): any => {
        let exception_type: any, fill: any;
        fill = $call(_Fill, "foobar");
        exception_type = TypeError;
        return [fill, exception_type];
      } },
    fore_raise_fixture: { args: [], fn: (self: any): any => {
        let exception_type: any, fill: any;
        fill = $call(_Fill, "foobar");
        exception_type = TypeError;
        return [fill, exception_type];
      } },
    pattern_raise_fixture: { args: [], fn: (self: any): any => {
        let exception_type: any, fill: any;
        fill = $call(_Fill, "barfoo");
        exception_type = TypeError;
        return [fill, exception_type];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_raises_on_fore_color_access", args: ["fore_raise_fixture"], fn: (self: any, fore_raise_fixture: any): any => {
        let exception_type: any, fill: any;
        [fill, exception_type] = $list(fore_raise_fixture);
        $raises(exception_type, () => {
          fill.foreColor;
        });
      } },
    { name: "it_raises_on_back_color_access", args: ["back_raise_fixture"], fn: (self: any, back_raise_fixture: any): any => {
        let exception_type: any, fill: any;
        [fill, exception_type] = $list(back_raise_fixture);
        $raises(exception_type, () => {
          fill.backColor;
        });
      } },
    { name: "it_raises_on_pattern_access", args: ["pattern_raise_fixture"], fn: (self: any, pattern_raise_fixture: any): any => {
        let exception_type: any, fill: any;
        [fill, exception_type] = $list(pattern_raise_fixture);
        $raises(exception_type, () => {
          fill.pattern;
        });
      } },
  ],
};
pytestClass("Describe_Fill", Describe_Fill);

const Describe_BlipFill: any = {
  fixtures: {
    fill_type_fixture: { args: [], fn: (self: any): any => {
        let blip_fill: any, expected_value: any, xFill: any;
        xFill = element("a:blipFill");
        blip_fill = $call(_BlipFill, xFill);
        expected_value = MSO_FILL.PICTURE;
        return [blip_fill, expected_value];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_fill_type", args: ["fill_type_fixture"], fn: (self: any, fill_type_fixture: any): any => {
        let blip_fill: any, expected_value: any, fill_type: any;
        [blip_fill, expected_value] = $list(fill_type_fixture);
        fill_type = blip_fill.type;
        $assertEqual(fill_type, expected_value);
      } },
  ],
};
pytestClass("Describe_BlipFill", Describe_BlipFill);

const Describe_GradFill: any = {
  fixtures: {
    angle_fixture: { args: ["request"], params: [["a:gradFill", null], ["a:gradFill/a:lin{ang=0}", 0.0], ["a:gradFill/a:lin{ang=2730000}", 314.5], ["a:gradFill/a:lin{ang=16200000}", 90.0]], fn: (self: any, request: any): any => {
        let cxml: any, expected_value: any, grad_fill: any;
        [cxml, expected_value] = $list(request.param);
        grad_fill = $call(_GradFill, element(cxml));
        return [grad_fill, expected_value];
      } },
    angle_set_fixture: { args: ["request"], params: [[301.2, "a:gradFill/a:lin{ang=3528000}"], [31.22, "a:gradFill/a:lin{ang=19726800}"], [0, "a:gradFill/a:lin{ang=0}"], [-460.0, "a:gradFill/a:lin{ang=6000000}"]], fn: (self: any, request: any): any => {
        let angle: any, expected_cxml: any, expected_xml: any, grad_fill: any;
        [angle, expected_cxml] = $list(request.param);
        grad_fill = $call(_GradFill, element("a:gradFill/a:lin{ang=0}"));
        expected_xml = xml(expected_cxml);
        return [grad_fill, angle, expected_xml];
      } },
    fill_type_fixture: { args: [], fn: (self: any): any => {
        let expected_value: any, grad_fill: any, xFill: any;
        xFill = element("a:gradFill");
        grad_fill = $call(_GradFill, xFill);
        expected_value = MSO_FILL.GRADIENT;
        return [grad_fill, expected_value];
      } },
    stops_fixture: { args: ["request", "_GradientStops_", "gradient_stops_"], params: [["a:gradFill", "a:gradFill", true], ["a:gradFill/a:gsLst", "a:gradFill/a:gsLst", false]], fn: (self: any, request: any, _GradientStops_: any, gradient_stops_: any): any => {
        let add_gsLst: any, expected_cxml: any, expected_xml: any, gradFill_cxml: any, grad_fill: any;
        [gradFill_cxml, expected_cxml, add_gsLst] = $list(request.param);
        _GradientStops_.returnValue = gradient_stops_;
        grad_fill = $call(_GradFill, element(gradFill_cxml));
        if ($truthy(add_gsLst)) {
          expected_xml = xml($add(expected_cxml, "/a:gsLst/(a:gs{pos=0}/a:schemeClr{val=accent1}/(a:tint{val=100000},a:shade{val=100000},a:satMod{val=130000}),a:gs{pos=100000}/a:schemeClr{val=accent1}/(a:tint{val=50000},a:shade{val=100000},a:satMod{val=350000}))"));
        } else {
          expected_xml = xml(expected_cxml);
        }
        return [grad_fill, _GradientStops_, expected_xml, gradient_stops_];
      } },
    _GradientStops_: { args: ["request"], fn: (self: any, request: any): any => {
        return classMock(request, "pptx.dml.fill._GradientStops");
      } },
    gradient_stops_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _GradientStops);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_the_gradient_angle", args: ["angle_fixture"], fn: (self: any, angle_fixture: any): any => {
        let angle: any, expected_value: any, grad_fill: any;
        [grad_fill, expected_value] = $list(angle_fixture);
        angle = grad_fill.gradientAngle;
        $assertEqual(angle, expected_value);
      } },
    { name: "it_can_change_the_gradient_angle", args: ["angle_set_fixture"], fn: (self: any, angle_set_fixture: any): any => {
        let angle: any, expected_xml: any, grad_fill: any;
        [grad_fill, angle, expected_xml] = $list(angle_set_fixture);
        grad_fill.gradientAngle = angle;
        $assertEqual(grad_fill._gradFill.xml, expected_xml);
      } },
    { name: "it_provides_access_to_the_gradient_stops", args: ["stops_fixture"], fn: (self: any, stops_fixture: any): any => {
        let _GradientStops_: any, expected_xml: any, gradFill: any, grad_fill: any, gradient_stops: any, gradient_stops_: any;
        [grad_fill, _GradientStops_] = $list($slice(stops_fixture, null, 2, null));
        [expected_xml, gradient_stops_] = $list($slice(stops_fixture, 2, null, null));
        gradient_stops = grad_fill.gradientStops;
        gradFill = grad_fill._gradFill;
        _GradientStops_.assertCalledOnceWith($get(gradFill, 0));
        $assertEqual(gradFill.xml, expected_xml);
        $assert($is(gradient_stops, gradient_stops_));
      } },
    { name: "it_raises_on_non_linear_gradient", args: [], fn: (self: any): any => {
        let gradFill: any, grad_fill: any;
        gradFill = element("a:gradFill/a:path");
        grad_fill = $call(_GradFill, gradFill);
        $raises(ValueError, () => {
          grad_fill.gradientAngle;
        });
        $raises(ValueError, () => {
          grad_fill.gradientAngle = 43.21;
        });
      } },
    { name: "it_knows_its_fill_type", args: ["fill_type_fixture"], fn: (self: any, fill_type_fixture: any): any => {
        let expected_value: any, fill_type: any, grad_fill: any;
        [grad_fill, expected_value] = $list(fill_type_fixture);
        fill_type = grad_fill.type;
        $assertEqual(fill_type, expected_value);
      } },
  ],
};
pytestClass("Describe_GradFill", Describe_GradFill);

const Describe_GrpFill: any = {
  fixtures: {
    fill_type_fixture: { args: [], fn: (self: any): any => {
        let expected_value: any, grp_fill: any, xFill: any;
        xFill = element("a:grpFill");
        grp_fill = $call(_GrpFill, xFill);
        expected_value = MSO_FILL.GROUP;
        return [grp_fill, expected_value];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_fill_type", args: ["fill_type_fixture"], fn: (self: any, fill_type_fixture: any): any => {
        let expected_value: any, fill_type: any, grp_fill: any;
        [grp_fill, expected_value] = $list(fill_type_fixture);
        fill_type = grp_fill.type;
        $assertEqual(fill_type, expected_value);
      } },
  ],
};
pytestClass("Describe_GrpFill", Describe_GrpFill);

const Describe_NoFill: any = {
  fixtures: {
    fill_type_fixture: { args: [], fn: (self: any): any => {
        let expected_value: any, no_fill: any, xFill: any;
        xFill = element("a:noFill");
        no_fill = $call(_NoFill, xFill);
        expected_value = MSO_FILL.BACKGROUND;
        return [no_fill, expected_value];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_fill_type", args: ["fill_type_fixture"], fn: (self: any, fill_type_fixture: any): any => {
        let expected_value: any, fill_type: any, no_fill: any;
        [no_fill, expected_value] = $list(fill_type_fixture);
        fill_type = no_fill.type;
        $assertEqual(fill_type, expected_value);
      } },
  ],
};
pytestClass("Describe_NoFill", Describe_NoFill);

const Describe_NoneFill: any = {
  fixtures: {
    fill_type_fixture: { args: [], fn: (self: any): any => {
        let expected_value: any, none_fill: any;
        none_fill = $call(_NoneFill, null);
        expected_value = null;
        return [none_fill, expected_value];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_fill_type", args: ["fill_type_fixture"], fn: (self: any, fill_type_fixture: any): any => {
        let expected_value: any, fill_type: any, none_fill: any;
        [none_fill, expected_value] = $list(fill_type_fixture);
        fill_type = none_fill.type;
        $assertEqual(fill_type, expected_value);
      } },
  ],
};
pytestClass("Describe_NoneFill", Describe_NoneFill);

const Describe_PattFill: any = {
  fixtures: {
    back_color_fixture: { args: ["request", "color_", "ColorFormat_from_colorchoice_parent_"], params: [["a:pattFill", "a:pattFill/a:bgClr/a:srgbClr{val=FFFFFF}"], ["a:pattFill/a:bgClr", "a:pattFill/a:bgClr"], ["a:pattFill/a:bgClr/a:schemeClr{val=accent1}", "a:pattFill/a:bgClr/a:schemeClr{val=accent1}"]], fn: (self: any, request: any, color_: any, ColorFormat_from_colorchoice_parent_: any): any => {
        let expected_cxml: any, expected_xml: any, pattFill: any, pattFill_cxml: any, patt_fill: any;
        [pattFill_cxml, expected_cxml] = $list(request.param);
        pattFill = element(pattFill_cxml);
        expected_xml = xml(expected_cxml);
        ColorFormat_from_colorchoice_parent_.returnValue = color_;
        patt_fill = $call(_PattFill, pattFill);
        return [patt_fill, pattFill, expected_xml, color_, ColorFormat_from_colorchoice_parent_];
      } },
    fore_color_fixture: { args: ["request", "color_", "ColorFormat_from_colorchoice_parent_"], params: [["a:pattFill", "a:pattFill/a:fgClr/a:srgbClr{val=000000}"], ["a:pattFill/a:fgClr", "a:pattFill/a:fgClr"], ["a:pattFill/a:fgClr/a:schemeClr{val=accent2}", "a:pattFill/a:fgClr/a:schemeClr{val=accent2}"]], fn: (self: any, request: any, color_: any, ColorFormat_from_colorchoice_parent_: any): any => {
        let expected_cxml: any, expected_xml: any, pattFill: any, pattFill_cxml: any, patt_fill: any;
        [pattFill_cxml, expected_cxml] = $list(request.param);
        pattFill = element(pattFill_cxml);
        expected_xml = xml(expected_cxml);
        ColorFormat_from_colorchoice_parent_.returnValue = color_;
        patt_fill = $call(_PattFill, pattFill);
        return [patt_fill, pattFill, expected_xml, color_, ColorFormat_from_colorchoice_parent_];
      } },
    fill_type_fixture: { args: [], fn: (self: any): any => {
        let expected_value: any, patt_fill: any;
        patt_fill = $call(_PattFill, element("a:pattFill"));
        expected_value = MSO_FILL.PATTERNED;
        return [patt_fill, expected_value];
      } },
    pattern_get_fixture: { args: ["request"], params: [["a:pattFill", null], ["a:pattFill{prst=diagCross}", MSO_PATTERN.DIAGONAL_CROSS], ["a:pattFill{prst=wave}", MSO_PATTERN.WAVE]], fn: (self: any, request: any): any => {
        let expected_value: any, pattFill: any, pattFill_cxml: any, patt_fill: any;
        [pattFill_cxml, expected_value] = $list(request.param);
        pattFill = element(pattFill_cxml);
        patt_fill = $call(_PattFill, pattFill);
        return [patt_fill, expected_value];
      } },
    pattern_set_fixture: { args: ["request"], params: [["a:pattFill", MSO_PATTERN.WAVE, "a:pattFill{prst=wave}"], ["a:pattFill{prst=wave}", MSO_PATTERN.DIVOT, "a:pattFill{prst=divot}"], ["a:pattFill{prst=divot}", null, "a:pattFill"]], fn: (self: any, request: any): any => {
        let expected_cxml: any, expected_xml: any, pattFill: any, pattFill_cxml: any, patt_fill: any, pattern: any;
        [pattFill_cxml, pattern, expected_cxml] = $list(request.param);
        pattFill = element(pattFill_cxml);
        expected_xml = xml(expected_cxml);
        patt_fill = $call(_PattFill, pattFill);
        return [patt_fill, pattern, pattFill, expected_xml];
      } },
    ColorFormat_from_colorchoice_parent_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, ColorFormat, "from_colorchoice_parent", $kw({autospec: false}));
      } },
    color_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, ColorFormat);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_fill_type", args: ["fill_type_fixture"], fn: (self: any, fill_type_fixture: any): any => {
        let expected_value: any, fill_type: any, patt_fill: any;
        [patt_fill, expected_value] = $list(fill_type_fixture);
        fill_type = patt_fill.type;
        $assertEqual(fill_type, expected_value);
      } },
    { name: "it_knows_its_pattern", args: ["pattern_get_fixture"], fn: (self: any, pattern_get_fixture: any): any => {
        let expected_value: any, patt_fill: any, pattern: any;
        [patt_fill, expected_value] = $list(pattern_get_fixture);
        pattern = patt_fill.pattern;
        $assertEqual(pattern, expected_value);
      } },
    { name: "it_can_change_its_pattern", args: ["pattern_set_fixture"], fn: (self: any, pattern_set_fixture: any): any => {
        let expected_xml: any, pattFill: any, patt_fill: any, pattern: any;
        [patt_fill, pattern, pattFill, expected_xml] = $list(pattern_set_fixture);
        patt_fill.pattern = pattern;
        $assertEqual(pattFill.xml, expected_xml);
      } },
    { name: "it_provides_access_to_its_foreground_color", args: ["fore_color_fixture"], fn: (self: any, fore_color_fixture: any): any => {
        let ColorFormat_from_colorchoice_parent_: any, color: any, color_: any, expected_xml: any, pattFill: any, patt_fill: any;
        [patt_fill, pattFill, expected_xml, color_] = $list($slice(fore_color_fixture, null, 4, null));
        ColorFormat_from_colorchoice_parent_ = $get(fore_color_fixture, 4);
        color = patt_fill.foreColor;
        $assertEqual(pattFill.xml, expected_xml);
        ColorFormat_from_colorchoice_parent_.assertCalledOnceWith(pattFill.fgClr);
        $assert($is(color, color_));
      } },
    { name: "it_provides_access_to_its_background_color", args: ["back_color_fixture"], fn: (self: any, back_color_fixture: any): any => {
        let ColorFormat_from_colorchoice_parent_: any, color: any, color_: any, expected_xml: any, pattFill: any, patt_fill: any;
        [patt_fill, pattFill, expected_xml, color_] = $list($slice(back_color_fixture, null, 4, null));
        ColorFormat_from_colorchoice_parent_ = $get(back_color_fixture, 4);
        color = patt_fill.backColor;
        $assertEqual(pattFill.xml, expected_xml);
        ColorFormat_from_colorchoice_parent_.assertCalledOnceWith(pattFill.bgClr);
        $assert($is(color, color_));
      } },
  ],
};
pytestClass("Describe_PattFill", Describe_PattFill);

const Describe_SolidFill: any = {
  fixtures: {
    fill_type_fixture: { args: [], fn: (self: any): any => {
        let expected_value: any, solid_fill: any;
        solid_fill = $call(_SolidFill, element("a:solidFill"));
        expected_value = MSO_FILL.SOLID;
        return [solid_fill, expected_value];
      } },
    fore_color_fixture: { args: ["ColorFormat_from_colorchoice_parent_", "color_"], fn: (self: any, ColorFormat_from_colorchoice_parent_: any, color_: any): any => {
        let solidFill: any, solid_fill: any;
        ColorFormat_from_colorchoice_parent_.returnValue = color_;
        solidFill = element("a:solidFill");
        solid_fill = $call(_SolidFill, solidFill);
        return [solid_fill, solidFill, color_, ColorFormat_from_colorchoice_parent_];
      } },
    ColorFormat_from_colorchoice_parent_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, ColorFormat, "from_colorchoice_parent", $kw({autospec: false}));
      } },
    color_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, ColorFormat);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_fill_type", args: ["fill_type_fixture"], fn: (self: any, fill_type_fixture: any): any => {
        let expected_value: any, fill_type: any, solid_fill: any;
        [solid_fill, expected_value] = $list(fill_type_fixture);
        fill_type = solid_fill.type;
        $assertEqual(fill_type, expected_value);
      } },
    { name: "it_provides_access_to_its_foreground_color", args: ["fore_color_fixture"], fn: (self: any, fore_color_fixture: any): any => {
        let ColorFormat_from_colorchoice_parent_: any, color: any, color_: any, solidFill: any, solid_fill: any;
        [solid_fill, solidFill, color_] = $list($slice(fore_color_fixture, null, 3, null));
        ColorFormat_from_colorchoice_parent_ = $get(fore_color_fixture, 3);
        color = solid_fill.foreColor;
        ColorFormat_from_colorchoice_parent_.assertCalledOnceWith(solidFill);
        $assert($is(color, color_));
      } },
  ],
};
pytestClass("Describe_SolidFill", Describe_SolidFill);

const Describe_GradientStops: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_access_to_its_stops", args: [], fn: (self: any): any => {
        let stops: any;
        stops = $call(_GradientStops, CT_GradientStopList.newGsLst());
        $assertEqual($len(stops), 2);
        for (const stop of $list(stops)) {
          $assert($isinstance(stop, _GradientStop));
        }
      } },
  ],
};
pytestClass("Describe_GradientStops", Describe_GradientStops);

const Describe_GradientStop: any = {
  fixtures: {
    pos_get_fixture: { args: ["request"], params: [["a:gs{pos=0}", 0.0], ["a:gs{pos=42240}", 0.4224], ["a:gs{pos=100000}", 1.0]], fn: (self: any, request: any): any => {
        let expected_value: any, gs_cxml: any, stop: any;
        [gs_cxml, expected_value] = $list(request.param);
        stop = $call(_GradientStop, element(gs_cxml));
        return [stop, expected_value];
      } },
    pos_set_fixture: { args: ["request"], params: [["a:gs{pos=0}", 0.42, 0.42], ["a:gs{pos=42000}", 1.0, 1.0], ["a:gs{pos=100000}", 0.0, 0.0]], fn: (self: any, request: any): any => {
        let expected_value: any, gs_cxml: any, new_value: any, stop: any;
        [gs_cxml, new_value, expected_value] = $list(request.param);
        stop = $call(_GradientStop, element(gs_cxml));
        return [stop, new_value, expected_value];
      } },
    raises_fixture: { args: ["request"], params: [-0.42, 1.001], fn: (self: any, request: any): any => {
        let out_of_range_value: any, stop: any;
        out_of_range_value = request.param;
        stop = $call(_GradientStop, element("a:gs{pos=50000}"));
        return [stop, out_of_range_value];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_access_to_its_color", args: ["request"], fn: (self: any, request: any): any => {
        let ColorFormat_: any, color: any, color_: any, gs: any, stop: any;
        gs = element("a:gs");
        ColorFormat_ = classMock(request, "pptx.dml.fill.ColorFormat");
        color_ = instanceMock(request, ColorFormat);
        ColorFormat_.fromColorchoiceParent.returnValue = color_;
        stop = $call(_GradientStop, gs);
        color = stop.color;
        ColorFormat_.fromColorchoiceParent.assertCalledOnceWith(gs);
        $assert($is(color, color_));
      } },
    { name: "it_knows_its_position", args: ["pos_get_fixture"], fn: (self: any, pos_get_fixture: any): any => {
        let expected_value: any, position: any, stop: any;
        [stop, expected_value] = $list(pos_get_fixture);
        position = stop.position;
        $assertEqual(position, expected_value);
      } },
    { name: "it_can_change_its_position", args: ["pos_set_fixture"], fn: (self: any, pos_set_fixture: any): any => {
        let expected_value: any, new_value: any, stop: any;
        [stop, new_value, expected_value] = $list(pos_set_fixture);
        stop.position = new_value;
        $assertEqual(stop.position, expected_value);
      } },
    { name: "it_raises_on_position_out_of_range", args: ["raises_fixture"], fn: (self: any, raises_fixture: any): any => {
        let out_of_range_value: any, stop: any;
        [stop, out_of_range_value] = $list(raises_fixture);
        $raises(ValueError, () => {
          stop.position = out_of_range_value;
        });
      } },
  ],
};
pytestClass("Describe_GradientStop", Describe_GradientStop);

