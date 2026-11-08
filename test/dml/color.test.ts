// Machine-translated from python-pptx tests/dml/test_color.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assertEqual, $call, $d, $eq, $get, $is, $list, $raises, $str, $truthy, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { ColorFormat as ColorFormat$imported, RGBColor as RGBColor$imported } from "../../src/dml/color.js";
import { MSO_COLOR_TYPE as MSO_COLOR_TYPE$imported, MSO_THEME_COLOR as MSO_THEME_COLOR$imported } from "../../src/enum/dml.js";
import { aLumMod, aLumOff, aPrstClr, aSchemeClr, aSolidFill, aSysClr, anHslClr, anScrgbClr, anSrgbClr } from "../oxml/unitdata/dml.js";
// python binds imported names at import time; patches made later don't change them
const ColorFormat = ColorFormat$imported;
const RGBColor = RGBColor$imported;
const MSO_COLOR_TYPE = MSO_COLOR_TYPE$imported;
const MSO_THEME_COLOR = MSO_THEME_COLOR$imported;
void $py;


const DescribeColorFormat: any = {
  fixtures: {
    color_format_having_none_color_type: { args: [], fn: (self: any): any => {
        let color_format: any, solidFill: any;
        solidFill = aSolidFill().withNsdecls().element;
        color_format = ColorFormat.fromColorchoiceParent(solidFill);
        return color_format;
      } },
    color_format_with_brightness: { args: ["request"], params: ["hsl", "prst", "scheme", "scrgb", "srgb", "sys"], fn: (self: any, request: any): any => {
        let color_format: any, exp_brightness: any, lumMod: any, lumOff: any, mapping: any, solidFill: any, xClr_bldr: any, xClr_bldr_fn: any;
        mapping = $d({hsl: [anHslClr, 55000, 45000, 0.45], prst: [aPrstClr, null, null, 0.0], scheme: [aSchemeClr, 15000, null, -0.85], scrgb: [anScrgbClr, 15000, 85000, 0.85], srgb: [anSrgbClr, null, null, 0.0], sys: [aSysClr, 23000, null, -0.77]});
        [xClr_bldr_fn, lumMod, lumOff, exp_brightness] = $list($get(mapping, request.param));
        xClr_bldr = xClr_bldr_fn();
        if ($truthy(!$is(lumMod, null))) {
          xClr_bldr.withChild(aLumMod().withVal(lumMod));
        }
        if ($truthy(!$is(lumOff, null))) {
          xClr_bldr.withChild(aLumOff().withVal(lumOff));
        }
        solidFill = aSolidFill().withNsdecls().withChild(xClr_bldr).element;
        color_format = ColorFormat.fromColorchoiceParent(solidFill);
        return [color_format, exp_brightness];
      } },
    color_type_fixture_: { args: ["request"], params: ["none", "hsl", "prst", "scheme", "scrgb", "srgb", "sys"], fn: (self: any, request: any): any => {
        let color_format: any, color_type: any, mapping: any, solidFill: any, solidFill_bldr: any, xClr_bldr_fn: any;
        mapping = $d({none: [null, null], hsl: [anHslClr, MSO_COLOR_TYPE.HSL], prst: [aPrstClr, MSO_COLOR_TYPE.PRESET], srgb: [anSrgbClr, MSO_COLOR_TYPE.RGB], scheme: [aSchemeClr, MSO_COLOR_TYPE.SCHEME], scrgb: [anScrgbClr, MSO_COLOR_TYPE.SCRGB], sys: [aSysClr, MSO_COLOR_TYPE.SYSTEM]});
        [xClr_bldr_fn, color_type] = $list($get(mapping, request.param));
        solidFill_bldr = aSolidFill().withNsdecls();
        if ($truthy(!$is(xClr_bldr_fn, null))) {
          solidFill_bldr.withChild(xClr_bldr_fn());
        }
        solidFill = solidFill_bldr.element;
        color_format = ColorFormat.fromColorchoiceParent(solidFill);
        return [color_format, color_type];
      } },
    _NoneColor_color_format: { args: [], fn: (self: any): any => {
        let color_format: any, solidFill: any;
        solidFill = aSolidFill().withNsdecls().element;
        color_format = ColorFormat.fromColorchoiceParent(solidFill);
        return color_format;
      } },
    rgb_color_format: { args: [], fn: (self: any): any => {
        let color_format: any, solidFill: any, srgbClr_bldr: any;
        srgbClr_bldr = anSrgbClr().withVal("123456");
        solidFill = aSolidFill().withNsdecls().withChild(srgbClr_bldr).element;
        color_format = ColorFormat.fromColorchoiceParent(solidFill);
        return color_format;
      } },
    rgb_raise_fixture_: { args: ["request"], params: ["none", "hsl", "prst", "scheme", "scrgb", "sys"], fn: (self: any, request: any): any => {
        let color_format: any, exception_type: any, mapping: any, solidFill: any, solidFill_bldr: any, xClr_bldr_fn: any;
        mapping = $d({none: [null, AttributeError], hsl: [anHslClr, AttributeError], prst: [aPrstClr, AttributeError], scheme: [aSchemeClr, AttributeError], scrgb: [anScrgbClr, AttributeError], sys: [aSysClr, AttributeError]});
        [xClr_bldr_fn, exception_type] = $list($get(mapping, request.param));
        solidFill_bldr = aSolidFill().withNsdecls();
        if ($truthy(!$is(xClr_bldr_fn, null))) {
          solidFill_bldr.withChild(xClr_bldr_fn());
        }
        solidFill = solidFill_bldr.element;
        color_format = ColorFormat.fromColorchoiceParent(solidFill);
        return [color_format, exception_type];
      } },
    set_brightness_fixture_: { args: ["request"], params: ["0 to 0", "0 to -0.4", "0.15 to 0.25", "0.15 to -0.15", "-0.25 to 0.4", "-0.3 to -0.4", "-0.4 to 0"], fn: (self: any, request: any): any => {
        let brightness: any, color_format: any, expected_xml: any, mapping: any, mod_in: any, mod_out: any, off_in: any, off_out: any, solidFill: any, xClr_bldr: any, xClr_bldr_fn: any;
        mapping = $d({"0 to 0": [anSrgbClr, null, null, 0, null, null], "0 to -0.4": [anHslClr, null, null, -0.4, 60000, null], "0.15 to 0.25": [aPrstClr, 85000, 15000, 0.25, 75000, 25000], "0.15 to -0.15": [aSchemeClr, 85000, 15000, -0.15, 85000, null], "-0.25 to 0.4": [anScrgbClr, 75000, null, 0.4, 60000, 40000], "-0.3 to -0.4": [anSrgbClr, 70000, null, -0.4, 60000, null], "-0.4 to 0": [aSysClr, 60000, null, 0, null, null]});
        [xClr_bldr_fn, mod_in, off_in, brightness, mod_out, off_out] = $list($get(mapping, request.param));
        xClr_bldr = xClr_bldr_fn();
        if ($truthy(!$is(mod_in, null))) {
          xClr_bldr.withChild(aLumMod().withVal(mod_in));
        }
        if ($truthy(!$is(off_in, null))) {
          xClr_bldr.withChild(aLumOff().withVal(off_in));
        }
        solidFill = aSolidFill().withNsdecls().withChild(xClr_bldr).element;
        color_format = ColorFormat.fromColorchoiceParent(solidFill);
        xClr_bldr = xClr_bldr_fn();
        if ($truthy(!$is(mod_out, null))) {
          xClr_bldr.withChild(aLumMod().withVal(mod_out));
        }
        if ($truthy(!$is(off_out, null))) {
          xClr_bldr.withChild(aLumOff().withVal(off_out));
        }
        expected_xml = aSolidFill().withNsdecls().withChild(xClr_bldr).xml();
        return [color_format, brightness, expected_xml];
      } },
    set_rgb_fixture_: { args: ["request"], params: ["none", "hsl", "prst", "scheme", "scrgb", "srgb", "sys"], fn: (self: any, request: any): any => {
        let color_format: any, expected_xml: any, mapping: any, rgb_color: any, solidFill: any, solidFill_bldr: any, xClr_bldr_fn: any;
        mapping = $d({none: null, hsl: anHslClr, prst: aPrstClr, scheme: aSchemeClr, scrgb: anScrgbClr, srgb: anSrgbClr, sys: aSysClr});
        xClr_bldr_fn = $get(mapping, request.param);
        solidFill_bldr = aSolidFill().withNsdecls();
        if ($truthy(!$is(xClr_bldr_fn, null))) {
          solidFill_bldr.withChild(xClr_bldr_fn());
        }
        solidFill = solidFill_bldr.element;
        color_format = ColorFormat.fromColorchoiceParent(solidFill);
        rgb_color = $call(RGBColor, 18, 52, 86);
        expected_xml = aSolidFill().withNsdecls().withChild(anSrgbClr().withVal("123456")).xml();
        return [color_format, rgb_color, expected_xml];
      } },
    set_theme_color_fixture_: { args: ["request"], params: ["none", "hsl", "prst", "scheme", "scrgb", "srgb", "sys"], fn: (self: any, request: any): any => {
        let color_format: any, expected_xml: any, mapping: any, solidFill: any, solidFill_bldr: any, theme_color: any, xClr_bldr_fn: any;
        mapping = $d({none: null, hsl: anHslClr, prst: aPrstClr, scheme: aSchemeClr, scrgb: anScrgbClr, srgb: anSrgbClr, sys: aSysClr});
        xClr_bldr_fn = $get(mapping, request.param);
        solidFill_bldr = aSolidFill().withNsdecls();
        if ($truthy(!$is(xClr_bldr_fn, null))) {
          solidFill_bldr.withChild(xClr_bldr_fn());
        }
        solidFill = solidFill_bldr.element;
        color_format = ColorFormat.fromColorchoiceParent(solidFill);
        theme_color = MSO_THEME_COLOR.ACCENT_6;
        expected_xml = aSolidFill().withNsdecls().withChild(aSchemeClr().withVal("accent6")).xml();
        return [color_format, theme_color, expected_xml];
      } },
    get_theme_color_fixture_: { args: ["request"], params: ["hsl", "prst", "scheme", "scrgb", "srgb", "sys"], fn: (self: any, request: any): any => {
        let color_format: any, mapping: any, solidFill: any, theme_color: any, xClr_bldr: any, xClr_bldr_fn: any;
        mapping = $d({hsl: [anHslClr, MSO_THEME_COLOR.NOT_THEME_COLOR], prst: [aPrstClr, MSO_THEME_COLOR.NOT_THEME_COLOR], scheme: [aSchemeClr, MSO_THEME_COLOR.ACCENT_1], scrgb: [anScrgbClr, MSO_THEME_COLOR.NOT_THEME_COLOR], srgb: [anSrgbClr, MSO_THEME_COLOR.NOT_THEME_COLOR], sys: [aSysClr, MSO_THEME_COLOR.NOT_THEME_COLOR]});
        [xClr_bldr_fn, theme_color] = $list($get(mapping, request.param));
        xClr_bldr = xClr_bldr_fn();
        if ($truthy(!$eq(theme_color, MSO_THEME_COLOR.NOT_THEME_COLOR))) {
          xClr_bldr.withVal("accent1");
        }
        solidFill = aSolidFill().withNsdecls().withChild(xClr_bldr).element;
        color_format = ColorFormat.fromColorchoiceParent(solidFill);
        return [color_format, theme_color];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_the_type_of_its_color", args: ["color_type_fixture_"], fn: (self: any, color_type_fixture_: any): any => {
        let color_format: any, color_type: any;
        [color_format, color_type] = $list(color_type_fixture_);
        $assertEqual(color_format.type, color_type);
      } },
    { name: "it_knows_the_RGB_value_of_an_RGB_color", args: ["rgb_color_format"], fn: (self: any, rgb_color_format: any): any => {
        let color_format: any;
        color_format = rgb_color_format;
        $assertEqual(color_format.rgb, $call(RGBColor, 18, 52, 86));
      } },
    { name: "it_raises_on_rgb_get_for_colors_other_than_rgb", args: ["rgb_raise_fixture_"], fn: (self: any, rgb_raise_fixture_: any): any => {
        let color_format: any, exception_type: any;
        [color_format, exception_type] = $list(rgb_raise_fixture_);
        $raises(exception_type, () => {
          color_format.rgb;
        });
      } },
    { name: "it_knows_the_theme_color_of_a_theme_color", args: ["get_theme_color_fixture_"], fn: (self: any, get_theme_color_fixture_: any): any => {
        let color_format: any, theme_color: any;
        [color_format, theme_color] = $list(get_theme_color_fixture_);
        $assertEqual(color_format.themeColor, theme_color);
      } },
    { name: "it_raises_on_theme_color_get_for_NoneColor", args: ["_NoneColor_color_format"], fn: (self: any, _NoneColor_color_format: any): any => {
        $raises(AttributeError, () => {
          _NoneColor_color_format.themeColor;
        });
      } },
    { name: "it_knows_its_brightness_adjustment", args: ["color_format_with_brightness"], fn: (self: any, color_format_with_brightness: any): any => {
        let color_format: any, expected_brightness: any;
        [color_format, expected_brightness] = $list(color_format_with_brightness);
        $assertEqual(color_format.brightness, expected_brightness);
      } },
    { name: "it_can_set_itself_to_an_RGB_color", args: ["set_rgb_fixture_"], fn: (self: any, set_rgb_fixture_: any): any => {
        let color_format: any, expected_xml: any, rgb_color: any;
        [color_format, rgb_color, expected_xml] = $list(set_rgb_fixture_);
        color_format.rgb = rgb_color;
        $assertEqual(color_format._xFill.xml, expected_xml);
      } },
    { name: "it_raises_on_assign_non_RGBColor_type_to_rgb", args: ["rgb_color_format"], fn: (self: any, rgb_color_format: any): any => {
        let color_format: any;
        color_format = rgb_color_format;
        $raises(ValueError, () => {
          color_format.rgb = [18, 52, 86];
        });
      } },
    { name: "it_can_set_itself_to_a_theme_color", args: ["set_theme_color_fixture_"], fn: (self: any, set_theme_color_fixture_: any): any => {
        let color_format: any, expected_xml: any, theme_color: any;
        [color_format, theme_color, expected_xml] = $list(set_theme_color_fixture_);
        color_format.themeColor = theme_color;
        $assertEqual(color_format._xFill.xml, expected_xml);
      } },
    { name: "it_can_set_its_brightness_adjustment", args: ["set_brightness_fixture_"], fn: (self: any, set_brightness_fixture_: any): any => {
        let brightness: any, color_format: any, expected_xml: any;
        [color_format, brightness, expected_xml] = $list(set_brightness_fixture_);
        color_format.brightness = brightness;
        $assertEqual(color_format._xFill.xml, expected_xml);
      } },
    { name: "it_raises_on_attempt_to_set_brightness_out_of_range", args: ["rgb_color_format"], fn: (self: any, rgb_color_format: any): any => {
        $raises(ValueError, () => {
          rgb_color_format.brightness = 1.1;
        });
        $raises(ValueError, () => {
          rgb_color_format.brightness = -1.1;
        });
      } },
    { name: "it_raises_on_attempt_to_set_brightness_on_None_color_type", args: ["color_format_having_none_color_type"], fn: (self: any, color_format_having_none_color_type: any): any => {
        let color_format: any;
        color_format = color_format_having_none_color_type;
        $raises(ValueError, () => {
          color_format.brightness = 0.5;
        });
      } },
  ],
};
pytestClass("DescribeColorFormat", DescribeColorFormat);

const DescribeRGBColor: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_is_natively_constructed_using_three_ints_0_to_255", args: [], fn: (self: any): any => {
        $call(RGBColor, 18, 52, 86);
        $raises(ValueError, () => {
          $call(RGBColor, "12", "34", "56");
        });
        $raises(ValueError, () => {
          $call(RGBColor, -1, 34, 56);
        });
        $raises(ValueError, () => {
          $call(RGBColor, 12, 256, 56);
        });
      } },
    { name: "it_can_construct_from_a_hex_string_rgb_value", args: [], fn: (self: any): any => {
        let rgb: any;
        rgb = RGBColor.fromString("123456");
        $assertEqual(rgb, $call(RGBColor, 18, 52, 86));
      } },
    { name: "it_can_provide_a_hex_string_rgb_value", args: [], fn: (self: any): any => {
        $assertEqual($str($call(RGBColor, 18, 52, 86)), "123456");
      } },
  ],
};
pytestClass("DescribeRGBColor", DescribeRGBColor);

