// Machine-translated from python-pptx tests/test_shared.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "./pyshim.js";
import { $assert, $call, $eq, $is, $list, instanceMock, pytestClass } from "./pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../src/exc.js";
import { defineChildren } from "../src/oxml/xmlchemy.js";
import { XmlPart as XmlPart$imported } from "../src/opc/package.js";
import { ElementProxy as ElementProxy$imported, ParentedElementProxy as ParentedElementProxy$imported } from "../src/shared.js";
import { element } from "./unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const XmlPart = XmlPart$imported;
const ElementProxy = ElementProxy$imported;
const ParentedElementProxy = ParentedElementProxy$imported;
void $py;


const DescribeElementProxy: any = {
  fixtures: {
    element_fixture: { args: [], fn: (self: any): any => {
        let p: any, proxy: any;
        p = element("w:p");
        proxy = $call(ElementProxy, p);
        return [proxy, p];
      } },
    eq_fixture: { args: [], fn: (self: any): any => {
        let not_a_proxy: any, p: any, proxy: any, proxy_2: any, proxy_3: any, q: any;
        [p, q] = $list([element("w:p"), element("w:p")]);
        proxy = $call(ElementProxy, p);
        proxy_2 = $call(ElementProxy, p);
        proxy_3 = $call(ElementProxy, q);
        not_a_proxy = "Foobar";
        return [proxy, proxy_2, proxy_3, not_a_proxy];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_when_its_equal_to_another_proxy_object", args: ["eq_fixture"], fn: (self: any, eq_fixture: any): any => {
        let not_a_proxy: any, proxy: any, proxy_2: any, proxy_3: any;
        [proxy, proxy_2, proxy_3, not_a_proxy] = $list(eq_fixture);
        $assert($is($eq(proxy, proxy_2), true));
        $assert($is($eq(proxy, proxy_3), false));
        $assert($is($eq(proxy, not_a_proxy), false));
        $assert($is(!$eq(proxy, proxy_2), false));
        $assert($is(!$eq(proxy, proxy_3), true));
        $assert($is(!$eq(proxy, not_a_proxy), true));
      } },
    { name: "it_knows_its_element", args: ["element_fixture"], fn: (self: any, element_fixture: any): any => {
        let element: any, proxy: any;
        [proxy, element] = $list(element_fixture);
        $assert($is(proxy.element, element));
      } },
  ],
};
pytestClass("DescribeElementProxy", DescribeElementProxy);

const DescribeParentedElementProxy: any = {
  fixtures: {
    parent_fixture: { args: [], fn: (self: any): any => {
        let parent: any, proxy: any;
        parent = 42;
        proxy = $call(ParentedElementProxy, element("w:p"), parent);
        return [proxy, parent];
      } },
    part_fixture: { args: ["other_proxy_", "part_"], fn: (self: any, other_proxy_: any, part_: any): any => {
        let proxy: any;
        other_proxy_.part = part_;
        proxy = $call(ParentedElementProxy, null, other_proxy_);
        return [proxy, part_];
      } },
    other_proxy_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, ParentedElementProxy);
      } },
    part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, XmlPart);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_parent", args: ["parent_fixture"], fn: (self: any, parent_fixture: any): any => {
        let parent: any, proxy: any;
        [proxy, parent] = $list(parent_fixture);
        $assert($is(proxy.parent, parent));
      } },
    { name: "it_knows_its_part", args: ["part_fixture"], fn: (self: any, part_fixture: any): any => {
        let part_: any, proxy: any;
        [proxy, part_] = $list(part_fixture);
        $assert($is(proxy.part, part_));
      } },
  ],
};
pytestClass("DescribeParentedElementProxy", DescribeParentedElementProxy);

