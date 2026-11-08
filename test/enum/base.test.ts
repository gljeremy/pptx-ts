// Machine-translated from python-pptx tests/enum/test_base.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $call, $is, $raises, $repr, $str, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { PP_ACTION as PP_ACTION$imported, PP_ACTION_TYPE as PP_ACTION_TYPE$imported } from "../../src/enum/action.js";
import { MSO_LINE_DASH_STYLE as MSO_LINE_DASH_STYLE$imported } from "../../src/enum/dml.js";
// python binds imported names at import time; patches made later don't change them
const PP_ACTION = PP_ACTION$imported;
const PP_ACTION_TYPE = PP_ACTION_TYPE$imported;
const MSO_LINE_DASH_STYLE = MSO_LINE_DASH_STYLE$imported;
void $py;


const DescribeBaseEnum: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_produces_members_each_equivalent_to_an_integer_value", args: [], fn: (self: any): any => {
        $assertEqual(PP_ACTION_TYPE.END_SHOW, 6);
        $assertEqual(PP_ACTION_TYPE.NONE, 0);
      } },
    { name: "but_member_reprs_are_a_str_indicating_the_enum_and_member_name", args: [], fn: (self: any): any => {
        $assertEqual($repr(PP_ACTION_TYPE.END_SHOW), "<PP_ACTION_TYPE.END_SHOW: 6>");
        $assertEqual($repr(PP_ACTION_TYPE.RUN_MACRO), "<PP_ACTION_TYPE.RUN_MACRO: 8>");
      } },
    { name: "and_member_str_values_are_a_str_indicating_the_member_name", args: [], fn: (self: any): any => {
        $assertEqual($str(PP_ACTION_TYPE.FIRST_SLIDE), "FIRST_SLIDE (3)");
        $assertEqual($str(PP_ACTION_TYPE.HYPERLINK), "HYPERLINK (7)");
      } },
    { name: "it_provides_a_docstring_for_each_member", args: [], fn: (self: any): any => {
        // python docstring assertion omitted (no __doc__ in JS)
        // python docstring assertion omitted (no __doc__ in JS)
      } },
    { name: "it_can_look_up_a_member_by_its_value", args: [], fn: (self: any): any => {
        $assertEqual($call(PP_ACTION_TYPE, 10), PP_ACTION_TYPE.NAMED_SLIDE_SHOW);
        $assertEqual($call(PP_ACTION_TYPE, 101), PP_ACTION_TYPE.NAMED_SLIDE);
      } },
    { name: "but_it_raises_when_no_member_has_that_value", args: [], fn: (self: any): any => {
        $raises(ValueError, () => {
          $call(PP_ACTION_TYPE, 42);
        }, "42 is not a valid PP_ACTION_TYPE");
      } },
    { name: "it_knows_its_name", args: [], fn: (self: any): any => {
        $assertEqual(PP_ACTION_TYPE.NEXT_SLIDE.name, "NEXT_SLIDE");
        $assertEqual(PP_ACTION_TYPE.NONE.name, "NONE");
      } },
    { name: "it_can_be_referred_to_by_a_convenience_alias_if_defined", args: [], fn: (self: any): any => {
        $assert($is(PP_ACTION_TYPE.OPEN_FILE, PP_ACTION.OPEN_FILE));
      } },
  ],
};
pytestClass("DescribeBaseEnum", DescribeBaseEnum);

const DescribeBaseXmlEnum: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_can_look_up_a_member_by_its_corresponding_XML_attribute_value", args: [], fn: (self: any): any => {
        $assertEqual(MSO_LINE_DASH_STYLE.fromXml("dash"), MSO_LINE_DASH_STYLE.DASH);
        $assertEqual(MSO_LINE_DASH_STYLE.fromXml("dashDot"), MSO_LINE_DASH_STYLE.DASH_DOT);
      } },
    { name: "but_it_raises_on_an_attribute_value_that_is_not_regitstered", args: [], fn: (self: any): any => {
        $raises(ValueError, () => {
          MSO_LINE_DASH_STYLE.fromXml("wavy");
        }, "MSO_LINE_DASH_STYLE has no XML mapping for 'wavy'");
      } },
    { name: "and_the_empty_string_never_maps_to_a_member", args: [], fn: (self: any): any => {
        $raises(ValueError, () => {
          MSO_LINE_DASH_STYLE.fromXml("");
        }, "MSO_LINE_DASH_STYLE has no XML mapping for ''");
      } },
    { name: "it_knows_the_XML_attribute_value_for_each_member_that_has_one", args: [], fn: (self: any): any => {
        $assertEqual(MSO_LINE_DASH_STYLE.toXml(MSO_LINE_DASH_STYLE.SOLID), "solid");
      } },
    { name: "and_it_looks_up_the_member_by_int_value_before_mapping_when_provided_that_way", args: [], fn: (self: any): any => {
        $assertEqual(MSO_LINE_DASH_STYLE.toXml(3), "sysDot");
      } },
    { name: "but_it_raises_when_no_member_has_the_provided_int_value", args: [], fn: (self: any): any => {
        $raises(ValueError, () => {
          MSO_LINE_DASH_STYLE.toXml(42);
        }, "42 is not a valid MSO_LINE_DASH_STYLE");
      } },
    { name: "and_it_raises_when_the_member_has_no_XML_value", args: [], fn: (self: any): any => {
        $raises(ValueError, () => {
          MSO_LINE_DASH_STYLE.toXml(-2);
        }, "MSO_LINE_DASH_STYLE.DASH_STYLE_MIXED has no XML r");
      } },
  ],
};
pytestClass("DescribeBaseXmlEnum", DescribeBaseXmlEnum);

