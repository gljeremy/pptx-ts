// Machine-translated from python-pptx tests/oxml/test_ns.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assertEqual, $call, $d, $mod, pytestClass, pytestModule } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { NamespacePrefixedTag as NamespacePrefixedTag$imported, namespaces as namespaces$imported, nsdecls as nsdecls$imported, nsuri as nsuri$imported, qn as qn$imported } from "../../src/oxml/ns.js";
// python binds imported names at import time; patches made later don't change them
const NamespacePrefixedTag = NamespacePrefixedTag$imported;
const namespaces = namespaces$imported;
const nsdecls = nsdecls$imported;
const nsuri = nsuri$imported;
const qn = qn$imported;
void $py;

pytestModule({
  fixtures: {
  clark_name: { args: ["namespace_uri_a", "local_part"], fn: (self: any, namespace_uri_a: any, local_part: any): any => {
      return $mod("{%s}%s", [namespace_uri_a, local_part]);
    } },
  local_part: { args: [], fn: (self: any): any => {
      return "foobar";
    } },
  namespace_uri_a: { args: [], fn: (self: any): any => {
      return "http://schemas.openxmlformats.org/drawingml/2006/main";
    } },
  namespace_uri_p: { args: [], fn: (self: any): any => {
      return "http://schemas.openxmlformats.org/presentationml/2006/main";
    } },
  nsdecls_str: { args: ["namespace_uri_a", "namespace_uri_p"], fn: (self: any, namespace_uri_a: any, namespace_uri_p: any): any => {
      return $mod("xmlns:a=\"%s\" xmlns:p=\"%s\"", [namespace_uri_a, namespace_uri_p]);
    } },
  nsmap: { args: ["namespace_uri_a", "namespace_uri_p"], fn: (self: any, namespace_uri_a: any, namespace_uri_p: any): any => {
      return $d({a: namespace_uri_a, p: namespace_uri_p});
    } },
  nsptag: { args: ["nsptag_str"], fn: (self: any, nsptag_str: any): any => {
      return $call(NamespacePrefixedTag, nsptag_str);
    } },
  nsptag_str: { args: ["local_part"], fn: (self: any, local_part: any): any => {
      return $mod("a:%s", local_part);
    } },
  },
  tests: [
  ],
});


const DescribeNamespacePrefixedTag: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_behaves_like_a_string_when_you_want_it_to", args: ["nsptag"], fn: (self: any, nsptag: any): any => {
        let s: any;
        s = $mod("- %s -", nsptag);
        $assertEqual(s, "- a:foobar -");
      } },
    { name: "it_knows_its_clark_name", args: ["nsptag", "clark_name"], fn: (self: any, nsptag: any, clark_name: any): any => {
        $assertEqual(nsptag.clarkName, clark_name);
      } },
    { name: "it_knows_its_local_part", args: ["nsptag", "local_part"], fn: (self: any, nsptag: any, local_part: any): any => {
        $assertEqual(nsptag.localPart, local_part);
      } },
    { name: "it_can_compose_a_single_entry_nsmap_for_itself", args: ["nsptag", "namespace_uri_a"], fn: (self: any, nsptag: any, namespace_uri_a: any): any => {
        let expected_nsmap: any;
        expected_nsmap = $d({a: namespace_uri_a});
        $assertEqual(nsptag.nsmap, expected_nsmap);
      } },
    { name: "it_knows_its_namespace_prefix", args: ["nsptag"], fn: (self: any, nsptag: any): any => {
        $assertEqual(nsptag.nspfx, "a");
      } },
    { name: "it_knows_its_namespace_uri", args: ["nsptag", "namespace_uri_a"], fn: (self: any, nsptag: any, namespace_uri_a: any): any => {
        $assertEqual(nsptag.nsuri, namespace_uri_a);
      } },
  ],
};
pytestClass("DescribeNamespacePrefixedTag", DescribeNamespacePrefixedTag);

const DescribeNamespaces: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_composes_a_dict_of_ns_uris_keyed_by_ns_pfx", args: ["nsmap"], fn: (self: any, nsmap: any): any => {
        $assertEqual(namespaces("a", "p"), nsmap);
      } },
  ],
};
pytestClass("DescribeNamespaces", DescribeNamespaces);

const DescribeNsdecls: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_formats_namespace_declarations_from_a_list_of_prefixes", args: ["nsdecls_str"], fn: (self: any, nsdecls_str: any): any => {
        $assertEqual(nsdecls("a", "p"), nsdecls_str);
      } },
  ],
};
pytestClass("DescribeNsdecls", DescribeNsdecls);

const DescribeNsuri: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_finds_the_namespace_uri_corresponding_to_a_namespace_prefix", args: ["namespace_uri_a"], fn: (self: any, namespace_uri_a: any): any => {
        $assertEqual(nsuri("a"), namespace_uri_a);
      } },
  ],
};
pytestClass("DescribeNsuri", DescribeNsuri);

const DescribeQn: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_calculates_the_clark_name_for_an_ns_prefixed_tag_string", args: ["nsptag_str", "clark_name"], fn: (self: any, nsptag_str: any, clark_name: any): any => {
        $assertEqual(qn(nsptag_str), clark_name);
      } },
  ],
};
pytestClass("DescribeQn", DescribeQn);









