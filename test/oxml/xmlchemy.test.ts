// Machine-translated from python-pptx tests/oxml/test_xmlchemy.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $call, $cmp, $eq, $get, $hasattr, $is, $isinstance, $kw, $list, $mod, $newm, $orv, $raises, $truthy, $typename, call, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { InvalidXmlError as InvalidXmlError$imported } from "../../src/exc.js";
import { registerElementCls as registerElementCls$imported } from "../../src/oxml/index.js";
import { qn as qn$imported } from "../../src/oxml/ns.js";
import { BaseIntType as BaseIntType$imported } from "../../src/oxml/simpletypes.js";
import { BaseOxmlElement as BaseOxmlElement$imported, Choice as Choice$imported, OneAndOnlyOne as OneAndOnlyOne$imported, OneOrMore as OneOrMore$imported, OptionalAttribute as OptionalAttribute$imported, RequiredAttribute as RequiredAttribute$imported, ZeroOrMore as ZeroOrMore$imported, ZeroOrOne as ZeroOrOne$imported, ZeroOrOneChoice as ZeroOrOneChoice$imported } from "../../src/oxml/xmlchemy.js";
import { BaseBuilder } from "../unitdata.js";
// python binds imported names at import time; patches made later don't change them
const InvalidXmlError = InvalidXmlError$imported;
const registerElementCls = registerElementCls$imported;
const qn = qn$imported;
const BaseIntType = BaseIntType$imported;
const BaseOxmlElement = BaseOxmlElement$imported;
const Choice = Choice$imported;
const OneAndOnlyOne = OneAndOnlyOne$imported;
const OneOrMore = OneOrMore$imported;
const OptionalAttribute = OptionalAttribute$imported;
const RequiredAttribute = RequiredAttribute$imported;
const ZeroOrMore = ZeroOrMore$imported;
const ZeroOrOne = ZeroOrOne$imported;
const ZeroOrOneChoice = ZeroOrOneChoice$imported;
void $py;


const DescribeCustomElementClass: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_has_the_MetaOxmlElement_metaclass", skip: "python metaclass; TS classes call defineChildren() instead", args: [], fn: (self: any): any => {
        $assertEqual($typename(CT_Parent), "MetaOxmlElement");
      } },
  ],
};
pytestClass("DescribeCustomElementClass", DescribeCustomElementClass);

const DescribeChoice: any = {
  fixtures: {
    add_fixture: { args: [], fn: (self: any): any => {
        let expected_xml: any, parent: any;
        parent = self.parentBldr().element;
        expected_xml = self.parentBldr("choice").xml();
        return [parent, expected_xml];
      } },
    get_or_change_to_fixture: { args: ["request"], params: [["choice2", "choice"], [null, "choice"], ["choice", "choice"]], fn: (self: any, request: any): any => {
        let after_member_tag: any, before_member_tag: any, expected_xml: any, parent: any;
        [before_member_tag, after_member_tag] = $list(request.param);
        parent = self.parentBldr(before_member_tag).element;
        expected_xml = self.parentBldr(after_member_tag).xml();
        return [parent, expected_xml];
      } },
    getter_fixture: { args: ["request"], params: ["choice", null], fn: (self: any, request: any): any => {
        let choice_tag: any, expected_choice: any, parent: any;
        choice_tag = request.param;
        parent = self.parentBldr(choice_tag).element;
        expected_choice = parent.find(qn("p:choice"));
        return [parent, expected_choice];
      } },
    insert_fixture: { args: [], fn: (self: any): any => {
        let choice: any, expected_xml: any, parent: any;
        parent = a_parent().withNsdecls().withChild(an_oomChild()).withChild(an_oooChild()).element;
        choice = a_choice().withNsdecls().element;
        expected_xml = a_parent().withNsdecls().withChild(a_choice()).withChild(an_oomChild()).withChild(an_oooChild()).xml();
        return [parent, choice, expected_xml];
      } },
    new_fixture: { args: [], fn: (self: any): any => {
        let expected_xml: any, parent: any;
        parent = self.parentBldr().element;
        expected_xml = a_choice().withNsdecls().xml();
        return [parent, expected_xml];
      } },
  },
  helpers: {
    parentBldr: (self: any, choice_tag: any = null): any => {
        let parent_bldr: any;
        parent_bldr = a_parent().withNsdecls();
        if ($truthy($eq(choice_tag, "choice"))) {
          parent_bldr.withChild(a_choice());
        }
        if ($truthy($eq(choice_tag, "choice2"))) {
          parent_bldr.withChild(a_choice2());
        }
        return parent_bldr;
      },
  },
  tests: [
    { name: "it_adds_a_getter_property_for_the_choice_element", args: ["getter_fixture"], fn: (self: any, getter_fixture: any): any => {
        let expected_choice: any, parent: any;
        [parent, expected_choice] = $list(getter_fixture);
        $assert($is(parent.choice, expected_choice));
      } },
    { name: "it_adds_a_creator_method_for_the_child_element", args: ["new_fixture"], fn: (self: any, new_fixture: any): any => {
        let choice: any, expected_xml: any, parent: any;
        [parent, expected_xml] = $list(new_fixture);
        choice = parent._newChoice();
        $assertEqual(choice.xml, expected_xml);
      } },
    { name: "it_adds_an_insert_method_for_the_child_element", args: ["insert_fixture"], fn: (self: any, insert_fixture: any): any => {
        let choice: any, expected_xml: any, parent: any;
        [parent, choice, expected_xml] = $list(insert_fixture);
        parent._insertChoice(choice);
        $assertEqual(parent.xml, expected_xml);
        // python docstring assertion omitted (no __doc__ in JS)
      } },
    { name: "it_adds_an_add_method_for_the_child_element", args: ["add_fixture"], fn: (self: any, add_fixture: any): any => {
        let choice: any, expected_xml: any, parent: any;
        [parent, expected_xml] = $list(add_fixture);
        choice = parent._addChoice();
        $assertEqual(parent.xml, expected_xml);
        $assert($isinstance(choice, CT_Choice));
        // python docstring assertion omitted (no __doc__ in JS)
      } },
    { name: "it_adds_a_get_or_change_to_method_for_the_child_element", args: ["get_or_change_to_fixture"], fn: (self: any, get_or_change_to_fixture: any): any => {
        let choice: any, expected_xml: any, parent: any;
        [parent, expected_xml] = $list(get_or_change_to_fixture);
        choice = parent.getOrChangeToChoice();
        $assert($isinstance(choice, CT_Choice));
        $assertEqual(parent.xml, expected_xml);
      } },
  ],
};
pytestClass("DescribeChoice", DescribeChoice);

const DescribeOneAndOnlyOne: any = {
  fixtures: {
    getter_fixture: { args: [], fn: (self: any): any => {
        let oooChild: any, parent: any;
        parent = a_parent().withNsdecls().withChild(an_oooChild()).element;
        oooChild = parent.find(qn("p:oooChild"));
        return [parent, oooChild];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_adds_a_getter_property_for_the_child_element", args: ["getter_fixture"], fn: (self: any, getter_fixture: any): any => {
        let oooChild: any, parent: any;
        [parent, oooChild] = $list(getter_fixture);
        $assert($is(parent.oooChild, oooChild));
      } },
  ],
};
pytestClass("DescribeOneAndOnlyOne", DescribeOneAndOnlyOne);

const DescribeOneOrMore: any = {
  fixtures: {
    add_fixture: { args: [], fn: (self: any): any => {
        let expected_xml: any, parent: any;
        parent = self.parentBldr(false).element;
        expected_xml = self.parentBldr(true).xml();
        return [parent, expected_xml];
      } },
    getter_fixture: { args: [], fn: (self: any): any => {
        let oomChild: any, parent: any;
        parent = self.parentBldr(true).element;
        oomChild = parent.find(qn("p:oomChild"));
        return [parent, oomChild];
      } },
    insert_fixture: { args: [], fn: (self: any): any => {
        let expected_xml: any, oomChild: any, parent: any;
        parent = a_parent().withNsdecls().withChild(an_oooChild()).withChild(a_zomChild()).withChild(a_zooChild()).element;
        oomChild = an_oomChild().withNsdecls().element;
        expected_xml = a_parent().withNsdecls().withChild(an_oomChild()).withChild(an_oooChild()).withChild(a_zomChild()).withChild(a_zooChild()).xml();
        return [parent, oomChild, expected_xml];
      } },
    new_fixture: { args: [], fn: (self: any): any => {
        let expected_xml: any, parent: any;
        parent = self.parentBldr(false).element;
        expected_xml = an_oomChild().withNsdecls().xml();
        return [parent, expected_xml];
      } },
  },
  helpers: {
    parentBldr: (self: any, oomChild_is_present: any): any => {
        let parent_bldr: any;
        parent_bldr = a_parent().withNsdecls();
        if ($truthy(oomChild_is_present)) {
          parent_bldr.withChild(an_oomChild());
        }
        return parent_bldr;
      },
  },
  tests: [
    { name: "it_adds_a_getter_property_for_the_child_element_list", args: ["getter_fixture"], fn: (self: any, getter_fixture: any): any => {
        let oomChild: any, parent: any;
        [parent, oomChild] = $list(getter_fixture);
        $assert($is($get(parent.oomChildLst, 0), oomChild));
      } },
    { name: "it_adds_a_creator_method_for_the_child_element", args: ["new_fixture"], fn: (self: any, new_fixture: any): any => {
        let expected_xml: any, oomChild: any, parent: any;
        [parent, expected_xml] = $list(new_fixture);
        oomChild = parent._newOomChild();
        $assertEqual(oomChild.xml, expected_xml);
      } },
    { name: "it_adds_an_insert_method_for_the_child_element", args: ["insert_fixture"], fn: (self: any, insert_fixture: any): any => {
        let expected_xml: any, oomChild: any, parent: any;
        [parent, oomChild, expected_xml] = $list(insert_fixture);
        parent._insertOomChild(oomChild);
        $assertEqual(parent.xml, expected_xml);
        // python docstring assertion omitted (no __doc__ in JS)
      } },
    { name: "it_adds_a_private_add_method_for_the_child_element", args: ["add_fixture"], fn: (self: any, add_fixture: any): any => {
        let expected_xml: any, oomChild: any, parent: any;
        [parent, expected_xml] = $list(add_fixture);
        oomChild = parent._addOomChild();
        $assertEqual(parent.xml, expected_xml);
        $assert($isinstance(oomChild, CT_OomChild));
        // python docstring assertion omitted (no __doc__ in JS)
      } },
    { name: "it_adds_a_public_add_method_for_the_child_element", args: ["add_fixture"], fn: (self: any, add_fixture: any): any => {
        let expected_xml: any, oomChild: any, parent: any;
        [parent, expected_xml] = $list(add_fixture);
        oomChild = parent.addOomChild();
        $assertEqual(parent.xml, expected_xml);
        $assert($isinstance(oomChild, CT_OomChild));
        // python docstring assertion omitted (no __doc__ in JS)
      } },
  ],
};
pytestClass("DescribeOneOrMore", DescribeOneOrMore);

const DescribeOptionalAttribute: any = {
  fixtures: {
    getter_fixture: { args: [], fn: (self: any): any => {
        let parent: any;
        parent = a_parent().withNsdecls().withOptAttr("24").element;
        return [parent, 24];
      } },
    setter_fixture: { args: ["request"], params: [36, null], fn: (self: any, request: any): any => {
        let expected_xml: any, parent: any, value: any;
        value = request.param;
        parent = a_parent().withNsdecls().withOptAttr("42").element;
        if ($truthy($is(value, null))) {
          expected_xml = a_parent().withNsdecls().xml();
        } else {
          expected_xml = a_parent().withNsdecls().withOptAttr(value).xml();
        }
        return [parent, value, expected_xml];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_adds_a_getter_property_for_the_attr_value", args: ["getter_fixture"], fn: (self: any, getter_fixture: any): any => {
        let optAttr_python_value: any, parent: any;
        [parent, optAttr_python_value] = $list(getter_fixture);
        $assertEqual(parent.optAttr, optAttr_python_value);
      } },
    { name: "it_adds_a_setter_property_for_the_attr", args: ["setter_fixture"], fn: (self: any, setter_fixture: any): any => {
        let expected_xml: any, parent: any, value: any;
        [parent, value, expected_xml] = $list(setter_fixture);
        parent.optAttr = value;
        $assertEqual(parent.xml, expected_xml);
      } },
    { name: "it_adds_a_docstring_for_the_property", skip: "python __doc__ strings", args: [], fn: (self: any): any => {
        // python docstring assertion omitted (no __doc__ in JS)
      } },
  ],
};
pytestClass("DescribeOptionalAttribute", DescribeOptionalAttribute);

const DescribeRequiredAttribute: any = {
  fixtures: {
    getter_fixture: { args: [], fn: (self: any): any => {
        let parent: any;
        parent = a_parent().withNsdecls().withReqAttr("42").element;
        return [parent, 42];
      } },
    invalid_assign_fixture: { args: ["request"], params: [[null, TypeError], [-4, ValueError], ["2", TypeError]], fn: (self: any, request: any): any => {
        let expected_exception: any, invalid_value: any, parent: any;
        [invalid_value, expected_exception] = $list(request.param);
        parent = a_parent().withNsdecls().withReqAttr(1).element;
        return [parent, invalid_value, expected_exception];
      } },
    setter_fixture: { args: [], fn: (self: any): any => {
        let expected_xml: any, parent: any, value: any;
        parent = a_parent().withNsdecls().withReqAttr("42").element;
        value = 24;
        expected_xml = a_parent().withNsdecls().withReqAttr(value).xml();
        return [parent, value, expected_xml];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_adds_a_getter_property_for_the_attr_value", args: ["getter_fixture"], fn: (self: any, getter_fixture: any): any => {
        let parent: any, reqAttr_python_value: any;
        [parent, reqAttr_python_value] = $list(getter_fixture);
        $assertEqual(parent.reqAttr, reqAttr_python_value);
      } },
    { name: "it_adds_a_setter_property_for_the_attr", args: ["setter_fixture"], fn: (self: any, setter_fixture: any): any => {
        let expected_xml: any, parent: any, value: any;
        [parent, value, expected_xml] = $list(setter_fixture);
        parent.reqAttr = value;
        $assertEqual(parent.xml, expected_xml);
      } },
    { name: "it_adds_a_docstring_for_the_property", skip: "python __doc__ strings", args: [], fn: (self: any): any => {
        // python docstring assertion omitted (no __doc__ in JS)
      } },
    { name: "it_raises_on_get_when_attribute_not_present", args: [], fn: (self: any): any => {
        let parent: any;
        parent = a_parent().withNsdecls().element;
        $raises(InvalidXmlError, () => {
          parent.reqAttr;
        });
      } },
    { name: "it_raises_on_assign_invalid_value", args: ["invalid_assign_fixture"], fn: (self: any, invalid_assign_fixture: any): any => {
        let expected_exception: any, parent: any, value: any;
        [parent, value, expected_exception] = $list(invalid_assign_fixture);
        $raises(expected_exception, () => {
          parent.reqAttr = value;
        });
      } },
  ],
};
pytestClass("DescribeRequiredAttribute", DescribeRequiredAttribute);

const DescribeZeroOrMore: any = {
  fixtures: {
    add_fixture: { args: [], fn: (self: any): any => {
        let expected_xml: any, parent: any;
        parent = self.parentBldr(false).element;
        expected_xml = self.parentBldr(true).xml();
        return [parent, expected_xml];
      } },
    getter_fixture: { args: [], fn: (self: any): any => {
        let parent: any, zomChild: any;
        parent = self.parentBldr(true).element;
        zomChild = parent.find(qn("p:zomChild"));
        return [parent, zomChild];
      } },
    insert_fixture: { args: [], fn: (self: any): any => {
        let expected_xml: any, parent: any, zomChild: any;
        parent = a_parent().withNsdecls().withChild(an_oomChild()).withChild(an_oooChild()).withChild(a_zooChild()).element;
        zomChild = a_zomChild().withNsdecls().element;
        expected_xml = a_parent().withNsdecls().withChild(an_oomChild()).withChild(an_oooChild()).withChild(a_zomChild()).withChild(a_zooChild()).xml();
        return [parent, zomChild, expected_xml];
      } },
    new_fixture: { args: [], fn: (self: any): any => {
        let expected_xml: any, parent: any;
        parent = self.parentBldr(false).element;
        expected_xml = a_zomChild().withNsdecls().xml();
        return [parent, expected_xml];
      } },
  },
  helpers: {
    parentBldr: (self: any, zomChild_is_present: any): any => {
        let parent_bldr: any;
        parent_bldr = a_parent().withNsdecls();
        if ($truthy(zomChild_is_present)) {
          parent_bldr.withChild(a_zomChild());
        }
        return parent_bldr;
      },
  },
  tests: [
    { name: "it_adds_a_getter_property_for_the_child_element_list", args: ["getter_fixture"], fn: (self: any, getter_fixture: any): any => {
        let parent: any, zomChild: any;
        [parent, zomChild] = $list(getter_fixture);
        $assert($is($get(parent.zomChildLst, 0), zomChild));
      } },
    { name: "it_adds_a_creator_method_for_the_child_element", args: ["new_fixture"], fn: (self: any, new_fixture: any): any => {
        let expected_xml: any, parent: any, zomChild: any;
        [parent, expected_xml] = $list(new_fixture);
        zomChild = parent._newZomChild();
        $assertEqual(zomChild.xml, expected_xml);
      } },
    { name: "it_adds_an_insert_method_for_the_child_element", args: ["insert_fixture"], fn: (self: any, insert_fixture: any): any => {
        let expected_xml: any, parent: any, zomChild: any;
        [parent, zomChild, expected_xml] = $list(insert_fixture);
        parent._insertZomChild(zomChild);
        $assertEqual(parent.xml, expected_xml);
        // python docstring assertion omitted (no __doc__ in JS)
      } },
    { name: "it_adds_an_add_method_for_the_child_element", args: ["add_fixture"], fn: (self: any, add_fixture: any): any => {
        let expected_xml: any, parent: any, zomChild: any;
        [parent, expected_xml] = $list(add_fixture);
        zomChild = parent._addZomChild();
        $assertEqual(parent.xml, expected_xml);
        $assert($isinstance(zomChild, CT_ZomChild));
        // python docstring assertion omitted (no __doc__ in JS)
      } },
    { name: "it_removes_the_property_root_name_used_for_declaration", args: [], fn: (self: any): any => {
        $assert(!$truthy($hasattr(CT_Parent, "zomChild")));
      } },
  ],
};
pytestClass("DescribeZeroOrMore", DescribeZeroOrMore);

const DescribeZeroOrOne: any = {
  fixtures: {
    add_fixture: { args: [], fn: (self: any): any => {
        let expected_xml: any, parent: any;
        parent = self.parentBldr(false).element;
        expected_xml = self.parentBldr(true).xml();
        return [parent, expected_xml];
      } },
    getter_fixture: { args: ["request"], params: [true, false], fn: (self: any, request: any): any => {
        let parent: any, zooChild: any, zooChild_is_present: any;
        zooChild_is_present = request.param;
        parent = self.parentBldr(zooChild_is_present).element;
        zooChild = parent.find(qn("p:zooChild"));
        return [parent, zooChild];
      } },
    get_or_add_fixture: { args: ["request"], params: [true, false], fn: (self: any, request: any): any => {
        let expected_xml: any, parent: any, zooChild_is_present: any;
        zooChild_is_present = request.param;
        parent = self.parentBldr(zooChild_is_present).element;
        expected_xml = self.parentBldr(true).xml();
        return [parent, expected_xml];
      } },
    insert_fixture: { args: [], fn: (self: any): any => {
        let expected_xml: any, parent: any, zooChild: any;
        parent = a_parent().withNsdecls().withChild(an_oomChild()).withChild(an_oooChild()).withChild(a_zomChild()).element;
        zooChild = a_zooChild().withNsdecls().element;
        expected_xml = a_parent().withNsdecls().withChild(an_oomChild()).withChild(an_oooChild()).withChild(a_zomChild()).withChild(a_zooChild()).xml();
        return [parent, zooChild, expected_xml];
      } },
    remove_fixture: { args: ["request"], params: [true, false], fn: (self: any, request: any): any => {
        let expected_xml: any, parent: any, zooChild_is_present: any;
        zooChild_is_present = request.param;
        parent = self.parentBldr(zooChild_is_present).element;
        expected_xml = self.parentBldr(false).xml();
        return [parent, expected_xml];
      } },
  },
  helpers: {
    parentBldr: (self: any, zooChild_is_present: any): any => {
        let parent_bldr: any;
        parent_bldr = a_parent().withNsdecls();
        if ($truthy(zooChild_is_present)) {
          parent_bldr.withChild(a_zooChild());
        }
        return parent_bldr;
      },
  },
  tests: [
    { name: "it_adds_a_getter_property_for_the_child_element", args: ["getter_fixture"], fn: (self: any, getter_fixture: any): any => {
        let parent: any, zooChild: any;
        [parent, zooChild] = $list(getter_fixture);
        $assert($is(parent.zooChild, zooChild));
      } },
    { name: "it_adds_an_add_method_for_the_child_element", args: ["add_fixture"], fn: (self: any, add_fixture: any): any => {
        let expected_xml: any, parent: any, zooChild: any;
        [parent, expected_xml] = $list(add_fixture);
        zooChild = parent._addZooChild();
        $assertEqual(parent.xml, expected_xml);
        $assert($isinstance(zooChild, CT_ZooChild));
        // python docstring assertion omitted (no __doc__ in JS)
      } },
    { name: "it_adds_an_insert_method_for_the_child_element", args: ["insert_fixture"], fn: (self: any, insert_fixture: any): any => {
        let expected_xml: any, parent: any, zooChild: any;
        [parent, zooChild, expected_xml] = $list(insert_fixture);
        parent._insertZooChild(zooChild);
        $assertEqual(parent.xml, expected_xml);
        // python docstring assertion omitted (no __doc__ in JS)
      } },
    { name: "it_adds_a_get_or_add_method_for_the_child_element", args: ["get_or_add_fixture"], fn: (self: any, get_or_add_fixture: any): any => {
        let expected_xml: any, parent: any, zooChild: any;
        [parent, expected_xml] = $list(get_or_add_fixture);
        zooChild = parent.getOrAddZooChild();
        $assert($isinstance(zooChild, CT_ZooChild));
        $assertEqual(parent.xml, expected_xml);
      } },
    { name: "it_adds_a_remover_method_for_the_child_element", args: ["remove_fixture"], fn: (self: any, remove_fixture: any): any => {
        let expected_xml: any, parent: any;
        [parent, expected_xml] = $list(remove_fixture);
        parent._removeZooChild();
        $assertEqual(parent.xml, expected_xml);
      } },
  ],
};
pytestClass("DescribeZeroOrOne", DescribeZeroOrOne);

const DescribeZeroOrOneChoice: any = {
  fixtures: {
    getter_fixture: { args: ["request"], params: [null, "choice", "choice2"], fn: (self: any, request: any): any => {
        let choice_tag: any, expected_choice: any, parent: any, tagname: any;
        choice_tag = request.param;
        parent = self.parentBldr(choice_tag).element;
        tagname = $mod("p:%s", choice_tag);
        expected_choice = parent.find(qn(tagname));
        return [parent, expected_choice];
      } },
  },
  helpers: {
    parentBldr: (self: any, choice_tag: any = null): any => {
        let parent_bldr: any;
        parent_bldr = a_parent().withNsdecls();
        if ($truthy($eq(choice_tag, "choice"))) {
          parent_bldr.withChild(a_choice());
        }
        if ($truthy($eq(choice_tag, "choice2"))) {
          parent_bldr.withChild(a_choice2());
        }
        return parent_bldr;
      },
  },
  tests: [
    { name: "it_adds_a_getter_for_the_current_choice", args: ["getter_fixture"], fn: (self: any, getter_fixture: any): any => {
        let expected_choice: any, parent: any;
        [parent, expected_choice] = $list(getter_fixture);
        $assert($is(parent.egZooChoice, expected_choice));
      } },
  ],
};
pytestClass("DescribeZeroOrOneChoice", DescribeZeroOrOneChoice);

export class ST_IntegerType extends BaseIntType {
  static validate(value: any): any {
    this.validateInt(value);
    if ($truthy($orv(() => $cmp(value, "<", 1), () => $cmp(value, ">", 42)))) {
      throw new ValueError("value must be in range 1 to 42 inclusive");
    }
  }
}

export class CT_Parent extends BaseOxmlElement {
  static {
    defineChildren(this, { egZooChoice: $newm(ZeroOrOneChoice, [$newm(Choice, "p:choice"), $newm(Choice, "p:choice2")], $kw({successors: ["p:oomChild", "p:oooChild"]})), oomChild: $newm(OneOrMore, "p:oomChild", $kw({successors: ["p:oooChild", "p:zomChild", "p:zooChild"]})), oooChild: $newm(OneAndOnlyOne, "p:oooChild"), zomChild: $newm(ZeroOrMore, "p:zomChild", $kw({successors: ["p:zooChild"]})), zooChild: $newm(ZeroOrOne, "p:zooChild", $kw({successors: []})), optAttr: $newm(OptionalAttribute, "p:optAttr", ST_IntegerType), reqAttr: $newm(RequiredAttribute, "reqAttr", ST_IntegerType) });
  }
}

export class CT_Choice extends BaseOxmlElement {
}

export class CT_OomChild extends BaseOxmlElement {
}

export class CT_ZomChild extends BaseOxmlElement {
}

export class CT_ZooChild extends BaseOxmlElement {
}

registerElementCls("p:parent", CT_Parent);

registerElementCls("p:choice", CT_Choice);

registerElementCls("p:oomChild", CT_OomChild);

registerElementCls("p:zomChild", CT_ZomChild);

registerElementCls("p:zooChild", CT_ZooChild);

export class CT_ChoiceBuilder extends BaseBuilder {
  static __tag__: any = "p:choice";
  static __nspfxs__: any = ["p"];
  static __attrs__: any = [];
}

export class CT_Choice2Builder extends BaseBuilder {
  static __tag__: any = "p:choice2";
  static __nspfxs__: any = ["p"];
  static __attrs__: any = [];
}

export class CT_ParentBuilder extends BaseBuilder {
  static __tag__: any = "p:parent";
  static __nspfxs__: any = ["p"];
  static __attrs__: any = ["p:optAttr", "reqAttr"];
}

export class CT_OomChildBuilder extends BaseBuilder {
  static __tag__: any = "p:oomChild";
  static __nspfxs__: any = ["p"];
  static __attrs__: any = [];
}

export class CT_OooChildBuilder extends BaseBuilder {
  static __tag__: any = "p:oooChild";
  static __nspfxs__: any = ["p"];
  static __attrs__: any = [];
}

export class CT_ZomChildBuilder extends BaseBuilder {
  static __tag__: any = "p:zomChild";
  static __nspfxs__: any = ["p"];
  static __attrs__: any = [];
}

export class CT_ZooChildBuilder extends BaseBuilder {
  static __tag__: any = "p:zooChild";
  static __nspfxs__: any = ["p"];
  static __attrs__: any = [];
}

export function a_choice(): any {
  return $call(CT_ChoiceBuilder);
}

export function a_choice2(): any {
  return $call(CT_Choice2Builder);
}

export function a_parent(): any {
  return $call(CT_ParentBuilder);
}

export function a_zomChild(): any {
  return $call(CT_ZomChildBuilder);
}

export function a_zooChild(): any {
  return $call(CT_ZooChildBuilder);
}

export function an_oomChild(): any {
  return $call(CT_OomChildBuilder);
}

export function an_oooChild(): any {
  return $call(CT_OooChildBuilder);
}

