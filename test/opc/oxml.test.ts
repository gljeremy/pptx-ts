// Machine-translated from python-pptx tests/opc/test_oxml.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $call, $cast, $encode, $get, $isinstance, $kw, $len, $list, $str, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { RELATIONSHIP_TARGET_MODE as RTM } from "../../src/opc/constants.js";
import { CT_Default, CT_Override, CT_Relationship, CT_Relationships, CT_Types, nsmap, oxmlTostring, serializePartXml } from "../../src/opc/oxml.js";
import { PackURI } from "../../src/opc/packuri.js";
import { parseXml } from "../../src/oxml/index.js";
import { BaseOxmlElement } from "../../src/oxml/xmlchemy.js";
import { element } from "../unitutil/cxml.js";
void $py;


const DescribeCT_Default: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_read_access_to_xml_values", args: [], fn: (self: any): any => {
        let default$: any;
        default$ = $cast(CT_Default, element("ct:Default{Extension=xml,ContentType=application/xml}"));
        $assertEqual(default$.extension, "xml");
        $assertEqual(default$.contentType, "application/xml");
      } },
  ],
};
pytestClass("DescribeCT_Default", DescribeCT_Default);

const DescribeCT_Override: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_read_access_to_xml_values", args: [], fn: (self: any): any => {
        let override: any;
        override = $cast(CT_Override, element("ct:Override{PartName=/part/name.xml,ContentType=text/plain}"));
        $assertEqual(override.partName, "/part/name.xml");
        $assertEqual(override.contentType, "text/plain");
      } },
  ],
};
pytestClass("DescribeCT_Override", DescribeCT_Override);

const DescribeCT_Relationship: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_read_access_to_xml_values", args: [], fn: (self: any): any => {
        let rel: any;
        rel = $cast(CT_Relationship, element("pr:Relationship{Id=rId9,Type=ReLtYpE,Target=docProps/core.xml}"));
        $assertEqual(rel.rId, "rId9");
        $assertEqual(rel.reltype, "ReLtYpE");
        $assertEqual(rel.targetRef, "docProps/core.xml");
        $assertEqual(rel.targetMode, RTM.INTERNAL);
      } },
    { name: "it_constructs_an_internal_relationship_when_no_target_mode_is_provided", args: [], fn: (self: any): any => {
        let rel: any;
        rel = CT_Relationship.new("rId9", "ReLtYpE", "foo/bar.xml");
        $assertEqual(rel.rId, "rId9");
        $assertEqual(rel.reltype, "ReLtYpE");
        $assertEqual(rel.targetRef, "foo/bar.xml");
        $assertEqual(rel.targetMode, RTM.INTERNAL);
        $assertEqual(rel.xml, `<Relationship xmlns="${$str($get(nsmap, "pr"))}" Id="rId9" Type="ReLtYpE" Target="foo/bar.xml"/>`);
      } },
    { name: "and_it_constructs_an_internal_relationship_when_target_mode_INTERNAL_is_specified", args: [], fn: (self: any): any => {
        let rel: any;
        rel = CT_Relationship.new("rId9", "ReLtYpE", "foo/bar.xml", RTM.INTERNAL);
        $assertEqual(rel.rId, "rId9");
        $assertEqual(rel.reltype, "ReLtYpE");
        $assertEqual(rel.targetRef, "foo/bar.xml");
        $assertEqual(rel.targetMode, RTM.INTERNAL);
        $assertEqual(rel.xml, `<Relationship xmlns="${$str($get(nsmap, "pr"))}" Id="rId9" Type="ReLtYpE" Target="foo/bar.xml"/>`);
      } },
    { name: "and_it_constructs_an_external_relationship_when_target_mode_EXTERNAL_is_specified", args: [], fn: (self: any): any => {
        let rel: any;
        rel = CT_Relationship.new("rId9", "ReLtYpE", "http://some/link", RTM.EXTERNAL);
        $assertEqual(rel.rId, "rId9");
        $assertEqual(rel.reltype, "ReLtYpE");
        $assertEqual(rel.targetRef, "http://some/link");
        $assertEqual(rel.targetMode, RTM.EXTERNAL);
        $assertEqual(rel.xml, `<Relationship xmlns="${$str($get(nsmap, "pr"))}" Id="rId9" Type="ReLtYpE" Target="http://some/link" TargetMode="External"/>`);
      } },
  ],
};
pytestClass("DescribeCT_Relationship", DescribeCT_Relationship);

const DescribeCT_Relationships: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_can_construct_a_new_relationships_element", args: [], fn: (self: any): any => {
        let rels: any;
        rels = CT_Relationships.new();
        $assertEqual(rels.xml, "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>");
      } },
    { name: "it_can_build_rels_element_incrementally", args: [], fn: (self: any): any => {
        let rels: any;
        rels = CT_Relationships.new();
        rels.addRel("rId1", "http://reltype1", "docProps/core.xml");
        rels.addRel("rId2", "http://linktype", "http://some/link", true);
        rels.addRel("rId3", "http://reltype2", "../slides/slide1.xml");
        $assertEqual($call(oxmlTostring, rels, $kw({encoding: "unicode", pretty_print: true})), "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n  <Relationship Id=\"rId1\" Type=\"http://reltype1\" Target=\"docProps/core.xml\"/>\n  <Relationship Id=\"rId2\" Type=\"http://linktype\" Target=\"http://some/link\" TargetMode=\"External\"/>\n  <Relationship Id=\"rId3\" Type=\"http://reltype2\" Target=\"../slides/slide1.xml\"/>\n</Relationships>\n");
      } },
    { name: "it_can_generate_rels_file_xml", args: [], fn: (self: any): any => {
        $assertEqual(CT_Relationships.new().xmlFileBytes, $encode("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>", "utf-8"));
      } },
  ],
};
pytestClass("DescribeCT_Relationships", DescribeCT_Relationships);

const DescribeCT_Types: any = {
  fixtures: {
    types: { args: [], fn: (self: any): any => {
        return $cast(CT_Types, element("ct:Types/(ct:Default{Extension=xml,ContentType=application/xml},ct:Default{Extension=jpeg,ContentType=image/jpeg},ct:Override{PartName=/docProps/core.xml,ContentType=app/vnd.type1},ct:Override{PartName=/ppt/presentation.xml,ContentType=app/vnd.type2},ct:Override{PartName=/docProps/thunbnail.jpeg,ContentType=image/jpeg})"));
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_access_to_default_child_elements", args: ["types"], fn: (self: any, types: any): any => {
        $assertEqual($len(types.defaultLst), 2);
        for (const default$ of $list(types.defaultLst)) {
          $assert($isinstance(default$, CT_Default));
        }
      } },
    { name: "it_provides_access_to_override_child_elements", args: ["types"], fn: (self: any, types: any): any => {
        $assertEqual($len(types.overrideLst), 3);
        for (const override of $list(types.overrideLst)) {
          $assert($isinstance(override, CT_Override));
        }
      } },
    { name: "it_should_have_empty_list_on_no_matching_elements", args: [], fn: (self: any): any => {
        let types: any;
        types = $cast(CT_Types, element("ct:Types"));
        $assertEqual(types.defaultLst, []);
        $assertEqual(types.overrideLst, []);
      } },
    { name: "it_can_construct_a_new_types_element", args: [], fn: (self: any): any => {
        let types: any;
        types = CT_Types.new();
        $assertEqual(types.xml, "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>\n");
      } },
    { name: "it_can_build_types_element_incrementally", args: [], fn: (self: any): any => {
        let types: any;
        types = CT_Types.new();
        types.addDefault("xml", "application/xml");
        types.addDefault("jpeg", "image/jpeg");
        types.addOverride($call(PackURI, "/docProps/core.xml"), "app/vnd.type1");
        types.addOverride($call(PackURI, "/ppt/presentation.xml"), "app/vnd.type2");
        types.addOverride($call(PackURI, "/docProps/thumbnail.jpeg"), "image/jpeg");
        $assertEqual(types.xml, "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n  <Default Extension=\"jpeg\" ContentType=\"image/jpeg\"/>\n  <Override PartName=\"/docProps/core.xml\" ContentType=\"app/vnd.type1\"/>\n  <Override PartName=\"/ppt/presentation.xml\" ContentType=\"app/vnd.type2\"/>\n  <Override PartName=\"/docProps/thumbnail.jpeg\" ContentType=\"image/jpeg\"/>\n</Types>\n");
      } },
  ],
};
pytestClass("DescribeCT_Types", DescribeCT_Types);

const Describe_serialize_part_xml: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_produces_properly_formatted_xml_for_an_opc_part", args: [], fn: (self: any): any => {
        let part_elm: any, xml: any;
        part_elm = $cast(BaseOxmlElement, parseXml("<f:foo xmlns:f=\"http://foo\" xmlns:b=\"http://bar\">\n  <f:bar>fØØbÅr</f:bar>\n</f:foo>\n"));
        xml = serializePartXml(part_elm);
        $assertEqual($len(xml), 137);
        $assertEqual(xml, $encode("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n<f:foo xmlns:f=\"http://foo\" xmlns:b=\"http://bar\"><f:bar>fØØbÅr</f:bar></f:foo>", "utf-8"));
      } },
  ],
};
pytestClass("Describe_serialize_part_xml", Describe_serialize_part_xml);

