// Machine-translated from python-pptx tests/parts/test_coreprops.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $b, $cmp, $datetime, $getattr, $is, $isinstance, $kw, $m, $mod, $replace, $setattr, $truthy, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { CONTENT_TYPE as CT$imported } from "../../src/opc/constants.js";
import { CT_CoreProperties as CT_CoreProperties$imported } from "../../src/oxml/coreprops.js";
import { CorePropertiesPart as CorePropertiesPart$imported } from "../../src/parts/coreprops.js";
// python binds imported names at import time; patches made later don't change them
const CT = CT$imported;
const CT_CoreProperties = CT_CoreProperties$imported;
const CorePropertiesPart = CorePropertiesPart$imported;
void $py;


const DescribeCorePropertiesPart: any = {
  fixtures: {
    core_properties: { args: [], fn: (self: any): any => {
        let xml: any;
        xml = $b("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\n  <cp:contentStatus>DRAFT</cp:contentStatus>\n  <dc:creator>python-pptx</dc:creator>\n  <dcterms:created xsi:type=\"dcterms:W3CDTF\">2012-11-17T11:07:40-05:30</dcterms:created>\n  <dc:description/>\n  <dc:identifier>GXS 10.2.1ab</dc:identifier>\n  <dc:language>US-EN</dc:language>\n  <cp:lastPrinted>2014-06-04T04:28:00Z</cp:lastPrinted>\n  <cp:keywords>foo bar baz</cp:keywords>\n  <cp:lastModifiedBy>Steve Canny</cp:lastModifiedBy>\n  <cp:revision>4</cp:revision>\n  <dc:subject>Spam</dc:subject>\n  <dc:title>Presentation</dc:title>\n  <cp:version>1.2.88</cp:version>\n</cp:coreProperties>\n");
        return CorePropertiesPart.load(null, null, null, xml);
      } },
  },
  helpers: {
    corePropertiesXml: (self: any, tagname: any, str_val: any, attrs: any = ""): any => {
        let child_element: any, tmpl: any;
        tmpl = "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">%s</cp:coreProperties>\n";
        if ($truthy(!$truthy(tagname))) {
          child_element = "";
        } else {
          if ($truthy(!$truthy(str_val))) {
            child_element = $mod("\n  <%s%s/>\n", [tagname, attrs]);
          } else {
            child_element = $mod("\n  <%s%s>%s</%s>\n", [tagname, attrs, str_val, tagname]);
          }
        }
        return $mod(tmpl, child_element);
      },
  },
  tests: [
    { name: "it_knows_the_string_property_values", args: ["core_properties", "prop_name", "expected_value"], parametrize: [{ names: ["prop_name", "expected_value"], values: [["author", "python-pptx"], ["category", ""], ["comments", ""], ["content_status", "DRAFT"], ["identifier", "GXS 10.2.1ab"], ["keywords", "foo bar baz"], ["language", "US-EN"], ["last_modified_by", "Steve Canny"], ["subject", "Spam"], ["title", "Presentation"], ["version", "1.2.88"]] }], fn: (self: any, core_properties: any, prop_name: any, expected_value: any): any => {
        $assertEqual($getattr(core_properties, prop_name), expected_value);
      } },
    { name: "it_can_change_the_string_property_values", args: ["prop_name", "tagname", "value"], parametrize: [{ names: ["prop_name", "tagname", "value"], values: [["author", "dc:creator", "scanny"], ["category", "cp:category", "silly stories"], ["comments", "dc:description", "Bar foo to you"], ["content_status", "cp:contentStatus", "FINAL"], ["identifier", "dc:identifier", "GT 5.2.xab"], ["keywords", "cp:keywords", "dog cat moo"], ["language", "dc:language", "GB-EN"], ["last_modified_by", "cp:lastModifiedBy", "Billy Bob"], ["subject", "dc:subject", "Eggs"], ["title", "dc:title", "Dissertation"], ["version", "cp:version", "81.2.8"]] }], fn: (self: any, prop_name: any, tagname: any, value: any): any => {
        let coreProperties: any, core_properties: any;
        coreProperties = self.corePropertiesXml(null, null);
        core_properties = CorePropertiesPart.load(null, null, null, coreProperties);
        $setattr(core_properties, prop_name, value);
        $assertEqual(core_properties._element.xml, self.corePropertiesXml(tagname, value));
      } },
    { name: "it_knows_the_date_property_values", args: ["core_properties", "prop_name", "expected_value"], parametrize: [{ names: ["prop_name", "expected_value"], values: [["created", $datetime.datetime(2012, 11, 17, 16, 37, 40)], ["last_printed", $datetime.datetime(2014, 6, 4, 4, 28)], ["modified", null]] }], fn: (self: any, core_properties: any, prop_name: any, expected_value: any): any => {
        let actual_datetime: any;
        actual_datetime = $getattr(core_properties, prop_name);
        $assertEqual(actual_datetime, expected_value);
      } },
    { name: "it_can_change_the_date_property_values", args: ["prop_name", "tagname", "value", "str_val", "attrs"], parametrize: [{ names: ["prop_name", "tagname", "value", "str_val", "attrs"], values: [["created", "dcterms:created", $datetime.datetime(2001, 2, 3, 4, 5), "2001-02-03T04:05:00Z", " xsi:type=\"dcterms:W3CDTF\""], ["last_printed", "cp:lastPrinted", $datetime.datetime(2014, 6, 4, 4), "2014-06-04T04:00:00Z", ""], ["modified", "dcterms:modified", $datetime.datetime(2005, 4, 3, 2, 1), "2005-04-03T02:01:00Z", " xsi:type=\"dcterms:W3CDTF\""]] }], fn: (self: any, prop_name: any, tagname: any, value: any, str_val: any, attrs: any): any => {
        let coreProperties: any, core_properties: any;
        coreProperties = self.corePropertiesXml(null, null);
        core_properties = CorePropertiesPart.load(null, null, null, coreProperties);
        $setattr(core_properties, prop_name, value);
        $assertEqual(core_properties._element.xml, self.corePropertiesXml(tagname, str_val, attrs));
      } },
    { name: "it_knows_the_revision_number", args: ["str_val", "expected_value"], parametrize: [{ names: ["str_val", "expected_value"], values: [["42", 42], [null, 0], ["foobar", 0], ["-17", 0], ["32.7", 0]] }], fn: (self: any, str_val: any, expected_value: any): any => {
        let coreProperties: any, core_properties: any, tagname: any;
        tagname = ($truthy($is(str_val, null)) ? "" : "cp:revision");
        coreProperties = self.corePropertiesXml(tagname, str_val);
        core_properties = CorePropertiesPart.load(null, null, null, coreProperties);
        $assertEqual(core_properties.revision, expected_value);
      } },
    { name: "it_can_change_the_revision_number", args: [], fn: (self: any): any => {
        let coreProperties: any, core_properties: any;
        coreProperties = self.corePropertiesXml(null, null);
        core_properties = CorePropertiesPart.load(null, null, null, coreProperties);
        core_properties.revision = 42;
        $assertEqual(core_properties._element.xml, self.corePropertiesXml("cp:revision", "42"));
      } },
    { name: "it_can_construct_a_default_core_props", args: [], fn: (self: any): any => {
        let core_props: any, max_expected_timedelta: any, modified_timedelta: any;
        core_props = CorePropertiesPart.default(null);
        $assert($isinstance(core_props, CorePropertiesPart));
        $assert($is(core_props.contentType, CT.OPC_CORE_PROPERTIES));
        $assertEqual(core_props.partname, "/docProps/core.xml");
        $assert($isinstance(core_props._element, CT_CoreProperties));
        $assertEqual(core_props.title, "PowerPoint Presentation");
        $assertEqual(core_props.lastModifiedBy, "python-pptx");
        $assertEqual(core_props.revision, 1);
        $assert(!$is(core_props.modified, null));
        modified_timedelta = ($replace($datetime.datetime.now($datetime.timezone.utc), $kw({tzinfo: null})) - core_props.modified);
        max_expected_timedelta = $m($datetime, "timedelta", $kw({seconds: 2}));
        $assert($cmp(modified_timedelta, "<", max_expected_timedelta));
      } },
  ],
};
pytestClass("DescribeCorePropertiesPart", DescribeCorePropertiesPart);

