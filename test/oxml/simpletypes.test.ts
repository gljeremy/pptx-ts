// Machine-translated from python-pptx tests/oxml/test_simpletypes.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $abs, $assert, $assertEqual, $call, $cmp, $is, $isinstance, $kw, $list, $mod, $raises, $truthy, instanceMock, methodMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { BaseIntType as BaseIntType$imported, BaseSimpleType as BaseSimpleType$imported, ST_Coordinate as ST_Coordinate$imported, ST_HexColorRGB as ST_HexColorRGB$imported, ST_Percentage as ST_Percentage$imported } from "../../src/oxml/simpletypes.js";
// python binds imported names at import time; patches made later don't change them
const BaseIntType = BaseIntType$imported;
const BaseSimpleType = BaseSimpleType$imported;
const ST_Coordinate = ST_Coordinate$imported;
const ST_HexColorRGB = ST_HexColorRGB$imported;
const ST_Percentage = ST_Percentage$imported;
void $py;


const DescribeBaseSimpleType: any = {
  fixtures: {
    to_xml_fixture: { args: ["request", "py_value_", "str_value_", "convert_to_xml_", "validate_"], fn: (self: any, request: any, py_value_: any, str_value_: any, convert_to_xml_: any, validate_: any): any => {
        return [ST_SimpleType, py_value_, str_value_];
      } },
    valid_int_fixture: { args: ["request"], params: [[42, null], [0, null], [-42, null], ["42", TypeError], [null, TypeError], [42.42, TypeError]], fn: (self: any, request: any): any => {
        let expected_exception: any, value: any;
        [value, expected_exception] = $list(request.param);
        return [value, expected_exception];
      } },
    valid_str_fixture: { args: ["request"], params: [["foobar", null], ["", null], [" foo ", null], [["foo"], TypeError], [42, TypeError], [null, TypeError], [42.42, TypeError]], fn: (self: any, request: any): any => {
        let expected_exception: any, value: any;
        [value, expected_exception] = $list(request.param);
        return [value, expected_exception];
      } },
    convert_from_xml_: { args: ["request", "py_value_"], fn: (self: any, request: any, py_value_: any): any => {
        return $call(methodMock, request, ST_SimpleType, "convert_from_xml", $kw({autospec: false, return_value: py_value_}));
      } },
    convert_to_xml_: { args: ["request", "str_value_"], fn: (self: any, request: any, str_value_: any): any => {
        return $call(methodMock, request, ST_SimpleType, "convert_to_xml", $kw({autospec: false, return_value: str_value_}));
      } },
    py_value_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Number);
      } },
    str_value_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, String);
      } },
    validate_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, ST_SimpleType, "validate", $kw({autospec: false}));
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_convert_an_XML_attribute_value_to_a_python_type", args: ["str_value_", "py_value_", "convert_from_xml_"], fn: (self: any, str_value_: any, py_value_: any, convert_from_xml_: any): any => {
        let py_value: any;
        py_value = ST_SimpleType.fromXml(str_value_);
        ST_SimpleType.convertFromXml.assertCalledOnceWith(str_value_);
        $assert($is(py_value, py_value_));
      } },
    { name: "it_can_convert_python_value_to_string", args: ["to_xml_fixture"], fn: (self: any, to_xml_fixture: any): any => {
        let SimpleType: any, py_value_: any, str_value: any, str_value_: any;
        [SimpleType, py_value_, str_value_] = $list(to_xml_fixture);
        str_value = SimpleType.toXml(py_value_);
        SimpleType.validate.assertCalledOnceWith(py_value_);
        SimpleType.convertToXml.assertCalledOnceWith(py_value_);
        $assert($is(str_value, str_value_));
      } },
    { name: "it_can_validate_a_value_as_a_python_int", args: ["valid_int_fixture"], fn: (self: any, valid_int_fixture: any): any => {
        let expected_exception: any, value: any;
        [value, expected_exception] = $list(valid_int_fixture);
        if ($truthy($is(expected_exception, null))) {
          BaseSimpleType.validateInt(value);
        } else {
          $raises(expected_exception, () => {
            BaseSimpleType.validateInt(value);
          });
        }
      } },
    { name: "it_can_validate_a_value_as_a_python_string", args: ["valid_str_fixture"], fn: (self: any, valid_str_fixture: any): any => {
        let expected_exception: any, value: any;
        [value, expected_exception] = $list(valid_str_fixture);
        if ($truthy($is(expected_exception, null))) {
          BaseSimpleType.validateString(value);
        } else {
          $raises(expected_exception, () => {
            BaseSimpleType.validateString(value);
          });
        }
      } },
  ],
};
pytestClass("DescribeBaseSimpleType", DescribeBaseSimpleType);

const DescribeBaseIntType: any = {
  fixtures: {
    from_xml_fixture: { args: ["request"], params: [["42", 42, null], ["-42", -42, null], ["-0042", -42, null], ["", null, ValueError], ["foo", null, ValueError], ["42.42", null, ValueError], ["0x0a3", null, ValueError], [null, null, TypeError]], fn: (self: any, request: any): any => {
        let expected_exception: any, expected_value: any, str_value: any;
        [str_value, expected_value, expected_exception] = $list(request.param);
        return [str_value, expected_value, expected_exception];
      } },
    to_xml_fixture: { args: ["request"], params: [[42, "42"], [-42, "-42"], [42, "42"]], fn: (self: any, request: any): any => {
        let expected_str_value: any, value: any;
        [value, expected_str_value] = $list(request.param);
        return [value, expected_str_value];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_convert_a_string_to_an_int", args: ["from_xml_fixture"], fn: (self: any, from_xml_fixture: any): any => {
        let expected_exception: any, expected_value: any, str_value: any, value: any;
        [str_value, expected_value, expected_exception] = $list(from_xml_fixture);
        if ($truthy($is(expected_exception, null))) {
          value = BaseIntType.convertFromXml(str_value);
          $assertEqual(value, expected_value);
        } else {
          $raises(expected_exception, () => {
            BaseIntType.convertFromXml(str_value);
          });
        }
      } },
    { name: "it_can_convert_an_int_to_a_string", args: ["to_xml_fixture"], fn: (self: any, to_xml_fixture: any): any => {
        let expected_str_value: any, str_value: any, value: any;
        [value, expected_str_value] = $list(to_xml_fixture);
        str_value = BaseIntType.convertToXml(value);
        $assertEqual(str_value, expected_str_value);
      } },
  ],
};
pytestClass("DescribeBaseIntType", DescribeBaseIntType);

const DescribeST_Coordinate: any = {
  fixtures: {
    univ_meas_fixture: { args: ["request"], params: [["1.2in", 1097280], ["42mm", 1512000], ["0024cm", 8640000], ["-42pt", -533400], ["-036.214pc", -5519014], ["0pi", 0]], fn: (self: any, request: any): any => {
        let expected_value: any, str_value: any;
        [str_value, expected_value] = $list(request.param);
        return [str_value, expected_value];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_convert_from_ST_UniversalMeasure", args: ["univ_meas_fixture"], fn: (self: any, univ_meas_fixture: any): any => {
        let expected_value: any, str_value: any;
        [str_value, expected_value] = $list(univ_meas_fixture);
        $assertEqual(ST_Coordinate.convertFromXml(str_value), expected_value);
      } },
  ],
};
pytestClass("DescribeST_Coordinate", DescribeST_Coordinate);

const DescribeST_HexColorRGB: any = {
  fixtures: {
    to_xml_fixture: { args: ["request"], params: [["deadbf", "DEADBF"], ["012345", "012345"], ["0a1b3c", "0A1B3C"]], fn: (self: any, request: any): any => {
        let expected_value: any, value: any;
        [value, expected_value] = $list(request.param);
        return [value, expected_value];
      } },
    valid_fixture: { args: ["request"], params: [["012345", null], ["ABCDEF", null], ["deadbf", null], ["0A1B3C", null], [null, TypeError], [123456, TypeError], ["F00BAR", ValueError], ["F00b", ValueError]], fn: (self: any, request: any): any => {
        let exception: any, str_value: any;
        [str_value, exception] = $list(request.param);
        return [str_value, exception];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_validate_a_hex_RGB_string", args: ["valid_fixture"], fn: (self: any, valid_fixture: any): any => {
        let exception: any, str_value: any;
        [str_value, exception] = $list(valid_fixture);
        if ($truthy($is(exception, null))) {
          try {
            ST_HexColorRGB.validate(str_value);
          } catch (_e: any) {
            if ($isinstance(_e, ValueError)) {
              throw new Error($mod("string '%s' did not validate", str_value));
            } else {
              throw _e;
            }
          }
        } else {
          $raises(exception, () => {
            ST_HexColorRGB.validate(str_value);
          });
        }
      } },
    { name: "it_converts_RGB_string_to_upper_case", args: ["to_xml_fixture"], fn: (self: any, to_xml_fixture: any): any => {
        let _value: any, expected_value: any, value: any;
        [value, expected_value] = $list(to_xml_fixture);
        _value = ST_HexColorRGB.convertToXml(value);
        $assertEqual(_value, expected_value);
      } },
  ],
};
pytestClass("DescribeST_HexColorRGB", DescribeST_HexColorRGB);

const DescribeST_Percentage: any = {
  fixtures: {
    percent_fixture: { args: ["request"], params: [["12.34%", 0.1234], ["42%", 0.42], ["024%", 0.24], ["-42%", -0.42], ["-036.214%", -0.36214], ["0%", 0.0]], fn: (self: any, request: any): any => {
        let expected_value: any, str_value: any;
        [str_value, expected_value] = $list(request.param);
        return [str_value, expected_value];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_convert_from_xml_percent_literals", args: ["percent_fixture"], fn: (self: any, percent_fixture: any): any => {
        let expected_value: any, str_value: any, value: any;
        [str_value, expected_value] = $list(percent_fixture);
        value = ST_Percentage.convertFromXml(str_value);
        $assert($cmp($abs((value - expected_value)), "<", 1e-06));
      } },
  ],
};
pytestClass("DescribeST_Percentage", DescribeST_Percentage);

export class ST_SimpleType extends BaseSimpleType {
  static convertFromXml(str_value: any): any {
    return 666;
  }
  static convertToXml(value: any): any {
    return "666";
  }
  static validate(value: any): any {
  }
}

