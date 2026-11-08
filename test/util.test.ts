// Machine-translated from python-pptx tests/test_util.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "./pyshim.js";
import { $assertEqual, $call, $getattr, $list, pytestClass } from "./pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../src/exc.js";
import { defineChildren } from "../src/oxml/xmlchemy.js";
import { Centipoints as Centipoints$imported, Cm as Cm$imported, Emu as Emu$imported, Inches as Inches$imported, Length as Length$imported, Mm as Mm$imported, Pt as Pt$imported } from "../src/util.js";
// python binds imported names at import time; patches made later don't change them
const Centipoints = Centipoints$imported;
const Cm = Cm$imported;
const Emu = Emu$imported;
const Inches = Inches$imported;
const Length = Length$imported;
const Mm = Mm$imported;
const Pt = Pt$imported;
void $py;


const DescribeLength: any = {
  fixtures: {
    construct_fixture: { args: ["request"], params: [[Length, 914400, 914400], [Inches, 1.1, 1005840], [Centipoints, 12.5, 1587], [Cm, 2.53, 910799], [Emu, 9144.9, 9144], [Mm, 13.8, 496800], [Pt, 24.5, 311150]], fn: (self: any, request: any): any => {
        let UnitCls: any, emu: any, units_val: any;
        [UnitCls, units_val, emu] = $list(request.param);
        return [UnitCls, units_val, emu];
      } },
    units_fixture: { args: ["request"], params: [[914400, "inches", 1.0], [914400, "centipoints", 7200.0], [914400, "cm", 2.54], [914400, "emu", 914400], [914400, "mm", 25.4], [914400, "pt", 72.0]], fn: (self: any, request: any): any => {
        let emu: any, expected_length_in_units: any, units_prop_name: any;
        [emu, units_prop_name, expected_length_in_units] = $list(request.param);
        return [emu, units_prop_name, expected_length_in_units];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_construct_from_convenient_units", args: ["construct_fixture"], fn: (self: any, construct_fixture: any): any => {
        let UnitCls: any, emu: any, length: any, units_val: any;
        [UnitCls, units_val, emu] = $list(construct_fixture);
        length = $call(UnitCls, units_val);
        $assertEqual(length, emu);
      } },
    { name: "it_can_self_convert_to_convenient_units", args: ["units_fixture"], fn: (self: any, units_fixture: any): any => {
        let emu: any, expected_length_in_units: any, length: any, length_in_units: any, units_prop_name: any;
        [emu, units_prop_name, expected_length_in_units] = $list(units_fixture);
        length = $call(Length, emu);
        length_in_units = $getattr(length, units_prop_name);
        $assertEqual(length_in_units, expected_length_in_units);
      } },
  ],
};
pytestClass("DescribeLength", DescribeLength);

