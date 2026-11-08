// Machine-translated from python-pptx tests/test_action.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "./pyshim.js";
import { $assert, $assertEqual, $call, $is, $kw, $list, $raises, $slice, $truthy, call, classMock, instanceMock, methodMock, propertyMock, pytestClass } from "./pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../src/exc.js";
import { defineChildren } from "../src/oxml/xmlchemy.js";
import { ActionSetting as ActionSetting$imported, Hyperlink as Hyperlink$imported } from "../src/action.js";
import { PP_ACTION as PP_ACTION$imported } from "../src/enum/action.js";
import { RELATIONSHIP_TYPE as RT$imported } from "../src/opc/constants.js";
import { XmlPart as XmlPart$imported } from "../src/opc/package.js";
import { SlidePart as SlidePart$imported } from "../src/parts/slide.js";
import { Slide as Slide$imported } from "../src/slide.js";
import { element, xml } from "./unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const ActionSetting = ActionSetting$imported;
const Hyperlink = Hyperlink$imported;
const PP_ACTION = PP_ACTION$imported;
const RT = RT$imported;
const XmlPart = XmlPart$imported;
const SlidePart = SlidePart$imported;
const Slide = Slide$imported;
void $py;


const DescribeActionSetting: any = {
  fixtures: {
    action_fixture: { args: ["request"], params: [["p:cNvPr", null, PP_ACTION.NONE], ["p:cNvPr/a:hlinkClick", null, PP_ACTION.HYPERLINK], ["p:cNvPr/a:hlinkClick", "ppaction://hlinkshowjump?jump=firstslide", PP_ACTION.FIRST_SLIDE], ["p:cNvPr/a:hlinkClick", "ppaction://hlinkshowjump?jump=lastslide", PP_ACTION.LAST_SLIDE], ["p:cNvPr/a:hlinkClick", "ppaction://hlinkshowjump?jump=nextslide", PP_ACTION.NEXT_SLIDE], ["p:cNvPr/a:hlinkClick", "ppaction://hlinkshowjump?jump=previousslide", PP_ACTION.PREVIOUS_SLIDE], ["p:cNvPr/a:hlinkClick", "ppaction://hlinkshowjump?jump=endshow", PP_ACTION.END_SHOW], ["p:cNvPr/a:hlinkClick", "ppaction://hlinksldjump", PP_ACTION.NAMED_SLIDE], ["p:cNvPr/a:hlinkClick", "ppaction://hlinkfile", PP_ACTION.OPEN_FILE], ["p:cNvPr/a:hlinkClick", "ppaction://hlinkpres", PP_ACTION.PLAY], ["p:cNvPr/a:hlinkClick", "ppaction://customshow", PP_ACTION.NAMED_SLIDE_SHOW], ["p:cNvPr/a:hlinkClick", "ppaction://ole", PP_ACTION.OLE_VERB], ["p:cNvPr/a:hlinkClick", "ppaction://macro", PP_ACTION.RUN_MACRO], ["p:cNvPr/a:hlinkClick", "ppaction://program", PP_ACTION.RUN_PROGRAM], ["p:cNvPr/a:hlinkClick", "ppaction://hlinkshowjump?jump=lastslideviewed", PP_ACTION.LAST_SLIDE_VIEWED], ["p:cNvPr/a:hlinkClick", "ppaction://media", PP_ACTION.NONE]], fn: (self: any, request: any): any => {
        let action_setting: any, action_text: any, cNvPr: any, cNvPr_cxml: any, expected_action: any;
        [cNvPr_cxml, action_text, expected_action] = $list(request.param);
        cNvPr = element(cNvPr_cxml);
        if ($truthy(!$is(action_text, null))) {
          cNvPr.hlinkClick.action = action_text;
        }
        action_setting = $call(ActionSetting, cNvPr, null);
        return [action_setting, expected_action];
      } },
    clear_fixture: { args: ["request", "part_prop_", "part_"], params: [["p:cNvPr", null, "p:cNvPr"], ["p:cNvPr{a:b=c,r:s=t}/a:hlinkClick{r:id=rId42}", "rId42", "p:cNvPr{a:b=c,r:s=t}"]], fn: (self: any, request: any, part_prop_: any, part_: any): any => {
        let action_setting: any, calls: any, expected_cxml: any, expected_xml: any, rId: any, xPr_cxml: any;
        [xPr_cxml, rId, expected_cxml] = $list(request.param);
        action_setting = $call(ActionSetting, element(xPr_cxml), null);
        part_prop_.returnValue = part_;
        calls = ($truthy(rId) ? [call(rId)] : []);
        expected_xml = xml(expected_cxml);
        return [action_setting, calls, expected_xml];
      } },
    hyperlink_fixture: { args: ["Hyperlink_", "hyperlink_"], fn: (self: any, Hyperlink_: any, hyperlink_: any): any => {
        let action_setting: any, parent: any, xPr: any;
        [xPr, parent] = $list(["xPr", "parent"]);
        action_setting = $call(ActionSetting, xPr, parent);
        return [action_setting, hyperlink_, Hyperlink_, xPr, parent];
      } },
    _slide_index_fixture: { args: ["request", "part_prop_"], fn: (self: any, request: any, part_prop_: any): any => {
        let action_setting: any, expected_value: any, slide: any, slide_part: any, slides: any;
        action_setting = $call(ActionSetting, null, null);
        slide_part = part_prop_.returnValue;
        slides = slide_part.package.presentationPart.presentation.slides;
        slide = slide_part.slide;
        expected_value = 123;
        slides.index.returnValue = expected_value;
        return [action_setting, slides, slide, expected_value];
      } },
    target_get_fixture: { args: ["request", "action_prop_", "_slide_index_prop_", "part_prop_"], params: [[PP_ACTION.NONE, null], [PP_ACTION.HYPERLINK, null], [PP_ACTION.FIRST_SLIDE, 0], [PP_ACTION.LAST_SLIDE, 5], [PP_ACTION.NEXT_SLIDE, 3], [PP_ACTION.PREVIOUS_SLIDE, 1], [PP_ACTION.END_SHOW, null], [PP_ACTION.NAMED_SLIDE, 4], [PP_ACTION.OPEN_FILE, null], [PP_ACTION.PLAY, null], [PP_ACTION.NAMED_SLIDE_SHOW, null], [PP_ACTION.OLE_VERB, null], [PP_ACTION.RUN_MACRO, null], [PP_ACTION.RUN_PROGRAM, null], [PP_ACTION.LAST_SLIDE_VIEWED, null]], fn: (self: any, request: any, action_prop_: any, _slide_index_prop_: any, part_prop_: any): any => {
        let action_setting: any, action_type: any, cNvPr: any, expected_value: any, prs_part_: any, related_part_: any;
        [action_type, expected_value] = $list(request.param);
        cNvPr = element("p:cNvPr/a:hlinkClick{r:id=rId6}");
        action_setting = $call(ActionSetting, cNvPr, null);
        action_prop_.returnValue = action_type;
        _slide_index_prop_.returnValue = 2;
        prs_part_ = part_prop_.returnValue.package.presentationPart;
        prs_part_.presentation.slides = [0, 1, 2, 3, 4, 5];
        related_part_ = part_prop_.returnValue.relatedPart;
        related_part_.returnValue.slide = 4;
        return [action_setting, expected_value];
      } },
    target_raise_fixture: { args: ["request", "action_prop_", "part_prop_", "_slide_index_prop_"], params: [[PP_ACTION.NEXT_SLIDE, 2], [PP_ACTION.PREVIOUS_SLIDE, 0]], fn: (self: any, request: any, action_prop_: any, part_prop_: any, _slide_index_prop_: any): any => {
        let action_setting: any, action_type: any, slide_idx: any;
        [action_type, slide_idx] = $list(request.param);
        action_setting = $call(ActionSetting, null, null);
        action_prop_.returnValue = action_type;
        part_prop_.returnValue.package.presentation.slides = [0, 1, 2];
        _slide_index_prop_.returnValue = slide_idx;
        return action_setting;
      } },
    action_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, ActionSetting, "action");
      } },
    _clear_click_action_: { args: ["request"], fn: (self: any, request: any): any => {
        return methodMock(request, ActionSetting, "_clear_click_action");
      } },
    Hyperlink_: { args: ["request", "hyperlink_"], fn: (self: any, request: any, hyperlink_: any): any => {
        return $call(classMock, request, "pptx.action.Hyperlink", $kw({return_value: hyperlink_}));
      } },
    hyperlink_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Hyperlink);
      } },
    part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, XmlPart);
      } },
    part_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, ActionSetting, "part");
      } },
    slide_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Slide);
      } },
    _slide_index_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, ActionSetting, "_slide_index");
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_action_type", args: ["action_fixture"], fn: (self: any, action_fixture: any): any => {
        let action: any, action_setting: any, expected_action: any;
        [action_setting, expected_action] = $list(action_fixture);
        action = action_setting.action;
        $assert($is(action, expected_action));
      } },
    { name: "it_provides_access_to_its_hyperlink", args: ["hyperlink_fixture"], fn: (self: any, hyperlink_fixture: any): any => {
        let Hyperlink_: any, action_setting: any, hyperlink: any, hyperlink_: any, parent: any, xPr: any;
        [action_setting, hyperlink_, Hyperlink_] = $list($slice(hyperlink_fixture, null, 3, null));
        [xPr, parent] = $list($slice(hyperlink_fixture, 3, null, null));
        hyperlink = action_setting.hyperlink;
        Hyperlink_.assertCalledOnceWith(xPr, parent, false);
        $assert($is(hyperlink, hyperlink_));
      } },
    { name: "it_can_find_its_slide_jump_target", args: ["target_get_fixture"], fn: (self: any, target_get_fixture: any): any => {
        let action_setting: any, expected_value: any, target_slide: any;
        [action_setting, expected_value] = $list(target_get_fixture);
        target_slide = action_setting.targetSlide;
        $assertEqual(target_slide, expected_value);
      } },
    { name: "it_can_change_its_slide_jump_target", args: ["request", "_clear_click_action_", "slide_", "part_prop_", "part_"], fn: (self: any, request: any, _clear_click_action_: any, slide_: any, part_prop_: any, part_: any): any => {
        let action_setting: any, slide_part_: any;
        part_prop_.returnValue = part_;
        part_.relateTo.returnValue = "rId42";
        slide_part_ = instanceMock(request, SlidePart);
        slide_.part = slide_part_;
        action_setting = $call(ActionSetting, element("p:cNvPr{a:b=c,r:s=t}"), null);
        action_setting.targetSlide = slide_;
        _clear_click_action_.assertCalledOnceWith(action_setting);
        part_.relateTo.assertCalledOnceWith(slide_part_, RT.SLIDE);
        $assertEqual(action_setting._element.xml, xml("p:cNvPr{a:b=c,r:s=t}/a:hlinkClick{action=ppaction://hlinksldjump,r:id=rId42}"));
      } },
    { name: "but_it_clears_the_target_slide_if_None_is_assigned", args: ["_clear_click_action_"], fn: (self: any, _clear_click_action_: any): any => {
        let action_setting: any;
        action_setting = $call(ActionSetting, element("p:cNvPr{a:b=c,r:s=t}"), null);
        action_setting.targetSlide = null;
        _clear_click_action_.assertCalledOnceWith(action_setting);
        $assertEqual(action_setting._element.xml, xml("p:cNvPr{a:b=c,r:s=t}"));
      } },
    { name: "it_raises_on_no_next_prev_slide", args: ["target_raise_fixture"], fn: (self: any, target_raise_fixture: any): any => {
        let action_setting: any;
        action_setting = target_raise_fixture;
        $raises(ValueError, () => {
          action_setting.targetSlide;
        });
      } },
    { name: "it_knows_its_slide_index_to_help", args: ["_slide_index_fixture"], fn: (self: any, _slide_index_fixture: any): any => {
        let action_setting: any, expected_value: any, slide: any, slide_index: any, slides: any;
        [action_setting, slides, slide, expected_value] = $list(_slide_index_fixture);
        slide_index = action_setting._slideIndex;
        slides.index.assertCalledOnceWith(slide);
        $assertEqual(slide_index, expected_value);
      } },
    { name: "it_clears_the_click_action_to_help", args: ["clear_fixture"], fn: (self: any, clear_fixture: any): any => {
        let action_setting: any, calls: any, expected_xml: any;
        [action_setting, calls, expected_xml] = $list(clear_fixture);
        action_setting._clearClickAction();
        $assertEqual(action_setting.part.dropRel.callArgsList, calls);
        $assertEqual(action_setting._element.xml, expected_xml);
      } },
  ],
};
pytestClass("DescribeActionSetting", DescribeActionSetting);

const DescribeHyperlink: any = {
  fixtures: {
    address_fixture: { args: ["part_prop_"], fn: (self: any, part_prop_: any): any => {
        let cNvPr: any, cNvPr_cxml: any, expected_address: any, hyperlink: any, rId: any;
        [cNvPr_cxml, rId] = $list(["p:cNvPr/a:hlinkClick{r:id=rId1}", "rId1"]);
        expected_address = "http://foobar.com";
        cNvPr = element(cNvPr_cxml);
        hyperlink = $call(Hyperlink, cNvPr, null);
        part_prop_.returnValue.targetRef.returnValue = expected_address;
        return [hyperlink, rId, expected_address];
      } },
    no_address_fixture: { args: ["request"], params: ["p:cNvPr", "p:cNvPr/a:hlinkClick"], fn: (self: any, request: any): any => {
        let cNvPr: any, cNvPr_cxml: any, hyperlink: any;
        cNvPr_cxml = request.param;
        cNvPr = element(cNvPr_cxml);
        hyperlink = $call(Hyperlink, cNvPr, null);
        return hyperlink;
      } },
    remove_fixture: { args: ["request", "part_prop_"], params: [["p:cNvPr{a:a=a,r:r=r}", []], ["p:cNvPr{a:a=a,r:r=r}/a:hlinkClick", []], ["p:cNvPr{a:a=a,r:r=r}/a:hlinkClick{r:id=rId3}", [call("rId3")]]], fn: (self: any, request: any, part_prop_: any): any => {
        let cNvPr: any, cNvPr_cxml: any, calls: any, expected_xml: any, hyperlink: any;
        [cNvPr_cxml, calls] = $list(request.param);
        cNvPr = element(cNvPr_cxml);
        expected_xml = xml("p:cNvPr{a:a=a,r:r=r}");
        hyperlink = $call(Hyperlink, cNvPr, null);
        return [hyperlink, calls, expected_xml];
      } },
    update_fixture: { args: ["request", "part_prop_"], params: [["p:cNvPr{a:a=a,r:r=r}", "http://foo.com", false, true, "p:cNvPr{a:a=a,r:r=r}/a:hlinkClick{r:id=rId3}"], ["p:cNvPr{a:a=a,r:r=r}", "http://bar.com", true, true, "p:cNvPr{a:a=a,r:r=r}/a:hlinkHover{r:id=rId3}"], ["p:cNvPr/a:hlinkClick{r:id=rId6}", "http://baz.com", false, true, "p:cNvPr/a:hlinkClick{r:id=rId3}"], ["p:cNvPr/a:hlinkHover{r:id=rId6}", "http://zab.com", true, true, "p:cNvPr/a:hlinkHover{r:id=rId3}"], ["p:cNvPr/a:hlinkHover{r:id=rId6}", "http://boo.com", false, true, "p:cNvPr/(a:hlinkClick{r:id=rId3},a:hlinkHover{r:id=rId6})"], ["p:cNvPr{a:a=a,r:r=r}/a:hlinkClick{r:id=rId6}", null, false, false, "p:cNvPr{a:a=a,r:r=r}"]], fn: (self: any, request: any, part_prop_: any): any => {
        let cNvPr: any, cNvPr_cxml: any, called: any, calls: any, expected_cNvPr_cxml: any, expected_xml: any, hover: any, hyperlink: any, url: any;
        [cNvPr_cxml, url, hover, called, expected_cNvPr_cxml] = $list(request.param);
        cNvPr = element(cNvPr_cxml);
        hyperlink = $call(Hyperlink, cNvPr, null, hover);
        calls = ($truthy(called) ? [$call(call, url, RT.HYPERLINK, true)] : []);
        part_prop_.returnValue.relateTo.returnValue = "rId3";
        expected_xml = xml(expected_cNvPr_cxml);
        return [hyperlink, url, calls, expected_xml];
      } },
    part_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, Hyperlink, "part");
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_the_target_url_of_the_hyperlink", args: ["address_fixture"], fn: (self: any, address_fixture: any): any => {
        let address: any, expected_address: any, hyperlink: any, rId: any;
        [hyperlink, rId, expected_address] = $list(address_fixture);
        address = hyperlink.address;
        hyperlink.part.targetRef.assertCalledOnceWith(rId);
        $assertEqual(address, expected_address);
      } },
    { name: "it_knows_when_theres_no_url", args: ["no_address_fixture"], fn: (self: any, no_address_fixture: any): any => {
        let hyperlink: any;
        hyperlink = no_address_fixture;
        $assert($is(hyperlink.address, null));
      } },
    { name: "it_can_remove_its_url", args: ["remove_fixture"], fn: (self: any, remove_fixture: any): any => {
        let calls: any, expected_xml: any, hyperlink: any;
        [hyperlink, calls, expected_xml] = $list(remove_fixture);
        hyperlink.address = null;
        $assertEqual(hyperlink.part.dropRel.callArgsList, calls);
        $assertEqual(hyperlink._element.xml, expected_xml);
      } },
    { name: "it_can_set_its_target_url", args: ["update_fixture"], fn: (self: any, update_fixture: any): any => {
        let calls: any, expected_xml: any, hyperlink: any, url: any;
        [hyperlink, url, calls, expected_xml] = $list(update_fixture);
        hyperlink.address = url;
        $assertEqual(hyperlink.part.relateTo.callArgsList, calls);
        $assertEqual(hyperlink._element.xml, expected_xml);
      } },
  ],
};
pytestClass("DescribeHyperlink", DescribeHyperlink);

