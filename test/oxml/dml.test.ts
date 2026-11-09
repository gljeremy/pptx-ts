// Machine-translated from python-pptx tests/oxml/test_dml.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $append, $assert, $assertEqual, $is, $isinstance, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { MSO_THEME_COLOR as MSO_THEME_COLOR$imported } from "../../src/enum/dml.js";
import { CT_Percentage as CT_Percentage$imported, CT_SRgbColor as CT_SRgbColor$imported, CT_SchemeColor as CT_SchemeColor$imported } from "../../src/oxml/dml/color.js";
import { qn as qn$imported } from "../../src/oxml/ns.js";
import { aLumMod, aLumOff, aSchemeClr, anSrgbClr } from "../oxml/unitdata/dml.js";
// python binds imported names at import time; patches made later don't change them
const MSO_THEME_COLOR = MSO_THEME_COLOR$imported;
const CT_Percentage = CT_Percentage$imported;
const CT_SRgbColor = CT_SRgbColor$imported;
const CT_SchemeColor = CT_SchemeColor$imported;
const qn = qn$imported;
void $py;


const Describe_BaseColorElement: any = {
  fixtures: {
    lumMod: { args: [], fn: (self: any): any => {
        return aLumMod().withNsdecls().element;
      } },
    lumOff: { args: [], fn: (self: any): any => {
        return aLumOff().withNsdecls().element;
      } },
    schemeClr: { args: [], fn: (self: any): any => {
        return aSchemeClr().withNsdecls().element;
      } },
    schemeClr_xml: { args: [], fn: (self: any): any => {
        return aSchemeClr().withNsdecls().xml();
      } },
    schemeClr_with_lumMod: { args: ["lumMod"], fn: (self: any, lumMod: any): any => {
        let schemeClr: any;
        schemeClr = aSchemeClr().withNsdecls().element;
        $append(schemeClr, lumMod);
        return schemeClr;
      } },
    schemeClr_with_lumMod_xml: { args: [], fn: (self: any): any => {
        let lumMod_bldr: any;
        lumMod_bldr = aLumMod().withVal(75000);
        return aSchemeClr().withNsdecls().withChild(lumMod_bldr).xml();
      } },
    schemeClr_with_lumOff_xml: { args: [], fn: (self: any): any => {
        let lumOff_bldr: any;
        lumOff_bldr = aLumOff().withVal(40000);
        return aSchemeClr().withNsdecls().withChild(lumOff_bldr).xml();
      } },
    schemeClr_with_lumOff: { args: ["lumOff"], fn: (self: any, lumOff: any): any => {
        let schemeClr: any;
        schemeClr = aSchemeClr().withNsdecls().element;
        $append(schemeClr, lumOff);
        return schemeClr;
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_get_the_lumMod_child_element_if_there_is_one", args: ["schemeClr", "schemeClr_with_lumMod", "lumMod"], fn: (self: any, schemeClr: any, schemeClr_with_lumMod: any, lumMod: any): any => {
        $assert($is(schemeClr.lumMod, null));
        $assert($is(schemeClr_with_lumMod.lumMod, lumMod));
      } },
    { name: "it_can_get_the_lumOff_child_element_if_there_is_one", args: ["schemeClr", "schemeClr_with_lumOff", "lumOff"], fn: (self: any, schemeClr: any, schemeClr_with_lumOff: any, lumOff: any): any => {
        $assert($is(schemeClr.lumOff, null));
        $assert($is(schemeClr_with_lumOff.lumOff, lumOff));
      } },
    { name: "it_can_remove_existing_lumMod_and_lumOff_child_elements", args: ["schemeClr_with_lumMod", "schemeClr_with_lumOff", "schemeClr_xml"], fn: (self: any, schemeClr_with_lumMod: any, schemeClr_with_lumOff: any, schemeClr_xml: any): any => {
        schemeClr_with_lumMod.clearLum();
        schemeClr_with_lumOff.clearLum();
        $assertEqual(schemeClr_with_lumMod.xml, schemeClr_xml);
        $assertEqual(schemeClr_with_lumOff.xml, schemeClr_xml);
      } },
    { name: "it_can_add_a_lumMod_child_element", args: ["schemeClr", "schemeClr_with_lumMod_xml"], fn: (self: any, schemeClr: any, schemeClr_with_lumMod_xml: any): any => {
        let lumMod: any;
        lumMod = schemeClr.addLumMod(0.75);
        $assertEqual(schemeClr.xml, schemeClr_with_lumMod_xml);
        $assertEqual(schemeClr.find(qn("a:lumMod")), lumMod);
      } },
    { name: "it_can_add_a_lumOff_child_element", args: ["schemeClr", "schemeClr_with_lumOff_xml"], fn: (self: any, schemeClr: any, schemeClr_with_lumOff_xml: any): any => {
        let lumOff: any;
        lumOff = schemeClr.addLumOff(0.4);
        $assertEqual(schemeClr.xml, schemeClr_with_lumOff_xml);
        $assertEqual(schemeClr.find(qn("a:lumOff")), lumOff);
      } },
  ],
};
pytestClass("Describe_BaseColorElement", Describe_BaseColorElement);

const DescribeCT_Percentage: any = {
  fixtures: {
    ct_percentage: { args: [], fn: (self: any): any => {
        return aLumMod().withNsdecls().withVal("99999").element;
      } },
    lumMod: { args: [], fn: (self: any): any => {
        return aLumMod().withNsdecls().withVal("33333").element;
      } },
    lumOff: { args: [], fn: (self: any): any => {
        return aLumOff().withNsdecls().withVal("66666").element;
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_is_used_by_the_parser_for_a_lumOff_element", args: ["lumOff"], fn: (self: any, lumOff: any): any => {
        $assert($isinstance(lumOff, CT_Percentage));
      } },
    { name: "it_is_used_by_the_parser_for_a_lumMod_element", args: ["lumMod"], fn: (self: any, lumMod: any): any => {
        $assert($isinstance(lumMod, CT_Percentage));
      } },
    { name: "it_knows_the_percentage_value", args: ["ct_percentage"], fn: (self: any, ct_percentage: any): any => {
        $assertEqual(ct_percentage.val, 0.99999);
      } },
  ],
};
pytestClass("DescribeCT_Percentage", DescribeCT_Percentage);

const DescribeCT_SchemeColor: any = {
  fixtures: {
    schemeClr: { args: [], fn: (self: any): any => {
        return aSchemeClr().withNsdecls().withVal("bg1").element;
      } },
    schemeClr_xml: { args: [], fn: (self: any): any => {
        return aSchemeClr().withNsdecls().withVal("accent1").xml();
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_is_used_by_the_parser_for_a_schemeClr_element", args: ["schemeClr"], fn: (self: any, schemeClr: any): any => {
        $assert($isinstance(schemeClr, CT_SchemeColor));
      } },
    { name: "it_knows_the_theme_color_idx", args: ["schemeClr"], fn: (self: any, schemeClr: any): any => {
        $assertEqual(schemeClr.val, MSO_THEME_COLOR.BACKGROUND_1);
      } },
    { name: "it_can_set_the_scheme_color_value", args: ["schemeClr", "schemeClr_xml"], fn: (self: any, schemeClr: any, schemeClr_xml: any): any => {
        schemeClr.val = MSO_THEME_COLOR.ACCENT_1;
        $assertEqual(schemeClr.xml, schemeClr_xml);
      } },
  ],
};
pytestClass("DescribeCT_SchemeColor", DescribeCT_SchemeColor);

const DescribeCT_SRgbColor: any = {
  fixtures: {
    srgbClr: { args: [], fn: (self: any): any => {
        return anSrgbClr().withNsdecls().withVal("123456").element;
      } },
    srgbClr_xml: { args: [], fn: (self: any): any => {
        return anSrgbClr().withNsdecls().withVal("987654").xml();
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_is_used_by_the_parser_for_an_srgbClr_element", args: ["srgbClr"], fn: (self: any, srgbClr: any): any => {
        $assert($isinstance(srgbClr, CT_SRgbColor));
      } },
    { name: "it_knows_the_rgb_str_value", args: ["srgbClr"], fn: (self: any, srgbClr: any): any => {
        $assertEqual(srgbClr.val, "123456");
      } },
    { name: "it_can_set_the_rgb_str_value", args: ["srgbClr", "srgbClr_xml"], fn: (self: any, srgbClr: any, srgbClr_xml: any): any => {
        srgbClr.val = "987654";
        $assertEqual(srgbClr.xml, srgbClr_xml);
      } },
  ],
};
pytestClass("DescribeCT_SRgbColor", DescribeCT_SRgbColor);

