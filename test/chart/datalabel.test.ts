// Machine-translated from python-pptx tests/chart/test_datalabel.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $call, $get, $getattr, $is, $isinstance, $kw, $list, $truthy, classMock, instanceMock, methodMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { DataLabel as DataLabel$imported, DataLabels as DataLabels$imported } from "../../src/chart/datalabel.js";
import { XL_LABEL_POSITION as XL_LABEL_POSITION$imported } from "../../src/enum/chart.js";
import { Font as Font$imported } from "../../src/text/text.js";
import { element, xml } from "../unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const DataLabel = DataLabel$imported;
const DataLabels = DataLabels$imported;
const XL_LABEL_POSITION = XL_LABEL_POSITION$imported;
const Font = Font$imported;
void $py;


const DescribeDataLabel: any = {
  fixtures: {
    font_fixture: { args: ["request"], params: [["c:ser{a:b=c}", "c:ser{a:b=c}/c:dLbls/(c:dLbl/(c:idx{val=9},c:spPr,c:txPr/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr),c:showLegendKey{val=0},c:showVal{val=1},c:showCatName{val=0},c:showSerName{val=0},c:showPercent{val=0},c:showBubbleSize{val=0}),c:showLegendKey{val=0},c:showVal{val=0},c:showCatName{val=0},c:showSerName{val=0},c:showPercent{val=0},c:showBubbleSize{val=0},c:showLeaderLines{val=1})"], ["c:ser{a:b=c}/c:dLbls/c:dLbl/(c:idx{val=9},c:txPr/(a:bodyPr,a:p))", "c:ser{a:b=c}/c:dLbls/c:dLbl/(c:idx{val=9},c:txPr/(a:bodyPr,a:p/a:pPr/a:defRPr))"]], fn: (self: any, request: any): any => {
        let data_label: any, expected_cxml: any, expected_xml: any, ser_cxml: any;
        [ser_cxml, expected_cxml] = $list(request.param);
        data_label = $call(DataLabel, element(ser_cxml), 9);
        expected_xml = xml(expected_cxml);
        return [data_label, expected_xml];
      } },
    has_tf_get_fixture: { args: ["request"], params: [["c:ser", false], ["c:ser/c:dLbls", false], ["c:ser/c:dLbls/c:dLbl/c:idx{val=42}", false], ["c:ser/c:dLbls/c:dLbl/(c:idx{val=42},c:tx/c:strRef)", false], ["c:ser/c:dLbls/c:dLbl/(c:idx{val=24},c:tx/c:rich)", false], ["c:ser/c:dLbls/c:dLbl/(c:idx{val=42},c:tx/c:rich)", true]], fn: (self: any, request: any): any => {
        let data_label: any, expected_value: any, ser_cxml: any;
        [ser_cxml, expected_value] = $list(request.param);
        data_label = $call(DataLabel, element(ser_cxml), 42);
        return [data_label, expected_value];
      } },
    has_tf_set_fixture: { args: ["request"], params: [["c:ser", false, "c:ser"], ["c:ser/c:dLbls", false, "c:ser/c:dLbls"], ["c:ser/c:dLbls/c:dLbl/c:idx{val=42}", false, "c:ser/c:dLbls/c:dLbl/c:idx{val=42}"], ["c:ser/c:dLbls/c:dLbl/(c:idx{val=42},c:tx/c:strRef)", false, "c:ser/c:dLbls/c:dLbl/(c:idx{val=42},c:tx/c:strRef)"], ["c:ser/c:dLbls/c:dLbl/(c:idx{val=24},c:tx/c:rich)", false, "c:ser/c:dLbls/c:dLbl/(c:idx{val=24},c:tx/c:rich)"], ["c:ser/c:dLbls/c:dLbl/(c:idx{val=42},c:tx/c:rich)", false, "c:ser/c:dLbls/c:dLbl/c:idx{val=42}"], ["c:ser{a:b=c}", true, "c:ser{a:b=c}/c:dLbls/(c:dLbl/(c:idx{val=42},c:tx/c:rich/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr),c:showLegendKey{val=0},c:showVal{val=1},c:showCatName{val=0},c:showSerName{val=0},c:showPercent{val=0},c:showBubbleSize{val=0}),c:showLegendKey{val=0},c:showVal{val=0},c:showCatName{val=0},c:showSerName{val=0},c:showPercent{val=0},c:showBubbleSize{val=0},c:showLeaderLines{val=1})"], ["c:ser{a:b=c}/c:dLbls", true, "c:ser{a:b=c}/c:dLbls/c:dLbl/(c:idx{val=42},c:tx/c:rich/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr),c:showLegendKey{val=0},c:showVal{val=1},c:showCatName{val=0},c:showSerName{val=0},c:showPercent{val=0},c:showBubbleSize{val=0})"], ["c:ser{a:b=c}/c:dLbls/c:dLbl/c:idx{val=42}", true, "c:ser{a:b=c}/c:dLbls/c:dLbl/(c:idx{val=42},c:tx/c:rich/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr))"], ["c:ser{a:b=c}/c:dLbls/c:dLbl/(c:idx{val=42},c:tx/c:strRef)", true, "c:ser{a:b=c}/c:dLbls/c:dLbl/(c:idx{val=42},c:tx/c:rich/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr))"], ["c:ser{a:b=c}/c:dLbls/c:dLbl/(c:idx{val=24},c:tx/c:rich)", true, "c:ser{a:b=c}/c:dLbls/(c:dLbl/(c:idx{val=24},c:tx/c:rich),c:dLbl/(c:idx{val=42},c:tx/c:rich/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr),c:showLegendKey{val=0},c:showVal{val=1},c:showCatName{val=0},c:showSerName{val=0},c:showPercent{val=0},c:showBubbleSize{val=0}))"], ["c:ser{a:b=c}/c:dLbls/c:dLbl/c:idx{val=42}", true, "c:ser{a:b=c}/c:dLbls/c:dLbl/(c:idx{val=42},c:tx/c:rich/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr))"]], fn: (self: any, request: any): any => {
        let data_label: any, expected_cxml: any, expected_xml: any, ser_cxml: any, value: any;
        [ser_cxml, value, expected_cxml] = $list(request.param);
        data_label = $call(DataLabel, element(ser_cxml), 42);
        expected_xml = xml(expected_cxml);
        return [data_label, value, expected_xml];
      } },
    position_get_fixture: { args: ["request"], params: [["c:ser", null], ["c:ser/c:dLbls/c:dLbl/c:idx{val=42}", null], ["c:ser/c:dLbls/c:dLbl/(c:idx{val=42},c:dLblPos{val=b})", "BELOW"]], fn: (self: any, request: any): any => {
        let data_label: any, expected_value: any, ser_cxml: any, value: any;
        [ser_cxml, value] = $list(request.param);
        data_label = $call(DataLabel, element(ser_cxml), 42);
        expected_value = ($truthy($is(value, null)) ? null : $getattr(XL_LABEL_POSITION, value));
        return [data_label, expected_value];
      } },
    position_set_fixture: { args: ["request"], params: [["c:ser{a:b=c}", "CENTER", "c:ser{a:b=c}/c:dLbls/(c:dLbl/(c:idx{val=42},c:spPr,c:txPr/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr),c:dLblPos{val=ctr},c:showLegendKey{val=0},c:showVal{val=1},c:showCatName{val=0},c:showSerName{val=0},c:showPercent{val=0},c:showBubbleSize{val=0}),c:showLegendKey{val=0},c:showVal{val=0},c:showCatName{val=0},c:showSerName{val=0},c:showPercent{val=0},c:showBubbleSize{val=0},c:showLeaderLines{val=1})"], ["c:ser/c:dLbls/c:dLbl/(c:idx{val=42},c:dLblPos{val=ctr})", "BELOW", "c:ser/c:dLbls/c:dLbl/(c:idx{val=42},c:dLblPos{val=b})"], ["c:ser/c:dLbls/c:dLbl/(c:idx{val=42},c:dLblPos{val=b})", null, "c:ser/c:dLbls/c:dLbl/c:idx{val=42}"], ["c:ser", null, "c:ser"]], fn: (self: any, request: any): any => {
        let data_label: any, expected_cxml: any, expected_xml: any, new_value: any, ser_cxml: any, value: any;
        [ser_cxml, value, expected_cxml] = $list(request.param);
        data_label = $call(DataLabel, element(ser_cxml), 42);
        new_value = ($truthy($is(value, null)) ? null : $getattr(XL_LABEL_POSITION, value));
        expected_xml = xml(expected_cxml);
        return [data_label, new_value, expected_xml];
      } },
    rich_fixture: { args: ["request"], params: [["c:ser{a:b=c}", "c:ser{a:b=c}/c:dLbls/(c:dLbl/(c:idx{val=42},c:tx/c:rich/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr),c:showLegendKey{val=0},c:showVal{val=1},c:showCatName{val=0},c:showSerName{val=0},c:showPercent{val=0},c:showBubbleSize{val=0}),c:showLegendKey{val=0},c:showVal{val=0},c:showCatName{val=0},c:showSerName{val=0},c:showPercent{val=0},c:showBubbleSize{val=0},c:showLeaderLines{val=1})"], ["c:ser{a:b=c}/c:dLbls/c:dLbl/(c:idx{val=42},c:tx/c:strRef)", "c:ser{a:b=c}/c:dLbls/c:dLbl/(c:idx{val=42},c:tx/c:rich/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr))"]], fn: (self: any, request: any): any => {
        let data_label: any, expected_cxml: any, expected_xml: any, ser_cxml: any;
        [ser_cxml, expected_cxml] = $list(request.param);
        data_label = $call(DataLabel, element(ser_cxml), 42);
        expected_xml = xml(expected_cxml);
        return [data_label, expected_xml];
      } },
    text_frame_fixture: { args: ["request", "_get_or_add_rich_", "TextFrame_"], fn: (self: any, request: any, _get_or_add_rich_: any, TextFrame_: any): any => {
        let data_label: any, rich_: any, text_frame_: any;
        data_label = $call(DataLabel, null, null);
        rich_ = _get_or_add_rich_.returnValue;
        text_frame_ = TextFrame_.returnValue;
        return [data_label, TextFrame_, rich_, text_frame_];
      } },
    _get_or_add_rich_: { args: ["request"], fn: (self: any, request: any): any => {
        return methodMock(request, DataLabel, "_get_or_add_rich");
      } },
    TextFrame_: { args: ["request"], fn: (self: any, request: any): any => {
        return classMock(request, "pptx.chart.datalabel.TextFrame");
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_has_a_font", args: ["font_fixture"], fn: (self: any, font_fixture: any): any => {
        let data_label: any, expected_xml: any, font: any;
        [data_label, expected_xml] = $list(font_fixture);
        font = data_label.font;
        $assertEqual(data_label._ser.xml, expected_xml);
        $assert($isinstance(font, Font));
      } },
    { name: "it_knows_its_position", args: ["position_get_fixture"], fn: (self: any, position_get_fixture: any): any => {
        let data_label: any, expected_value: any, position: any;
        [data_label, expected_value] = $list(position_get_fixture);
        position = data_label.position;
        $assertEqual(position, expected_value);
      } },
    { name: "it_can_change_its_position", args: ["position_set_fixture"], fn: (self: any, position_set_fixture: any): any => {
        let data_label: any, expected_xml: any, value: any;
        [data_label, value, expected_xml] = $list(position_set_fixture);
        data_label.position = value;
        $assertEqual(data_label._element.xml, expected_xml);
      } },
    { name: "it_knows_whether_it_has_a_text_frame", args: ["has_tf_get_fixture"], fn: (self: any, has_tf_get_fixture: any): any => {
        let data_label: any, expected_value: any, value: any;
        [data_label, expected_value] = $list(has_tf_get_fixture);
        value = data_label.hasTextFrame;
        $assert($is(value, expected_value));
      } },
    { name: "it_can_change_whether_it_has_a_text_frame", args: ["has_tf_set_fixture"], fn: (self: any, has_tf_set_fixture: any): any => {
        let data_label: any, expected_xml: any, value: any;
        [data_label, value, expected_xml] = $list(has_tf_set_fixture);
        data_label.hasTextFrame = value;
        $assertEqual(data_label._element.xml, expected_xml);
      } },
    { name: "it_provides_access_to_its_text_frame", args: ["text_frame_fixture"], fn: (self: any, text_frame_fixture: any): any => {
        let TextFrame_: any, data_label: any, rich_: any, text_frame: any, text_frame_: any;
        [data_label, TextFrame_, rich_, text_frame_] = $list(text_frame_fixture);
        text_frame = data_label.textFrame;
        TextFrame_.assertCalledOnceWith(rich_, data_label);
        $assert($is(text_frame, text_frame_));
      } },
    { name: "it_gets_or_adds_rich_element_to_help", args: ["rich_fixture"], fn: (self: any, rich_fixture: any): any => {
        let data_label: any, expected_xml: any, rich: any;
        [data_label, expected_xml] = $list(rich_fixture);
        rich = data_label._getOrAddRich();
        $assertEqual(data_label._ser.xml, expected_xml);
        $assert($is(rich, $get(data_label._ser.xpath("c:dLbls/c:dLbl/c:tx/c:rich"), 0)));
      } },
  ],
};
pytestClass("DescribeDataLabel", DescribeDataLabel);

const DescribeDataLabels: any = {
  fixtures: {
    catname_get_fixture: { args: ["request"], params: [["c:dLbls", false], ["c:dLbls/c:showCatName{val=0}", false], ["c:dLbls/c:showCatName{val=1}", true], ["c:dLbls/c:showCatName{val=true}", true], ["c:dLbls/c:showCatName", true]], fn: (self: any, request: any): any => {
        let dLbls: any, dLbls_cxml: any, expected_value: any;
        [dLbls_cxml, expected_value] = $list(request.param);
        dLbls = element(dLbls_cxml);
        return [dLbls, expected_value];
      } },
    catname_set_fixture: { args: ["request"], params: [["c:dLbls", false, "c:dLbls/c:showCatName{val=0}"], ["c:dLbls", true, "c:dLbls/c:showCatName{val=1}"], ["c:dLbls/c:showCatName", false, "c:dLbls/c:showCatName{val=0}"], ["c:dLbls/c:showCatName{val=0}", false, "c:dLbls/c:showCatName{val=0}"], ["c:dLbls/c:showCatName{val=0}", true, "c:dLbls/c:showCatName{val=1}"], ["c:dLbls/c:showCatName{val=1}", false, "c:dLbls/c:showCatName{val=0}"]], fn: (self: any, request: any): any => {
        let dLbls: any, dLbls_cxml: any, expected_dLbls_cxml: any, expected_dLbls_xml: any, new_value: any;
        [dLbls_cxml, new_value, expected_dLbls_cxml] = $list(request.param);
        dLbls = element(dLbls_cxml);
        expected_dLbls_xml = xml(expected_dLbls_cxml);
        return [dLbls, new_value, expected_dLbls_xml];
      } },
    font_fixture: { args: ["Font_", "font_"], fn: (self: any, Font_: any, font_: any): any => {
        let dLbls: any, data_labels: any, defRPr: any;
        dLbls = element("c:dLbls/c:txPr/a:p/a:pPr/a:defRPr");
        defRPr = $get(dLbls.xpath(".//a:defRPr"), 0);
        data_labels = $call(DataLabels, dLbls);
        return [data_labels, Font_, defRPr, font_];
      } },
    lgndkey_get_fixture: { args: ["request"], params: [["c:dLbls", false], ["c:dLbls/c:showLegendKey{val=0}", false], ["c:dLbls/c:showLegendKey{val=1}", true], ["c:dLbls/c:showLegendKey{val=true}", true], ["c:dLbls/c:showLegendKey", true]], fn: (self: any, request: any): any => {
        let dLbls: any, dLbls_cxml: any, expected_value: any;
        [dLbls_cxml, expected_value] = $list(request.param);
        dLbls = element(dLbls_cxml);
        return [dLbls, expected_value];
      } },
    lgndkey_set_fixture: { args: ["request"], params: [["c:dLbls", false, "c:dLbls/c:showLegendKey{val=0}"], ["c:dLbls", true, "c:dLbls/c:showLegendKey{val=1}"], ["c:dLbls/c:showLegendKey", false, "c:dLbls/c:showLegendKey{val=0}"], ["c:dLbls/c:showLegendKey{val=0}", false, "c:dLbls/c:showLegendKey{val=0}"], ["c:dLbls/c:showLegendKey{val=0}", true, "c:dLbls/c:showLegendKey{val=1}"], ["c:dLbls/c:showLegendKey{val=1}", false, "c:dLbls/c:showLegendKey{val=0}"]], fn: (self: any, request: any): any => {
        let dLbls: any, dLbls_cxml: any, expected_dLbls_cxml: any, expected_dLbls_xml: any, new_value: any;
        [dLbls_cxml, new_value, expected_dLbls_cxml] = $list(request.param);
        dLbls = element(dLbls_cxml);
        expected_dLbls_xml = xml(expected_dLbls_cxml);
        return [dLbls, new_value, expected_dLbls_xml];
      } },
    number_format_get_fixture: { args: ["request"], params: [["c:dLbls", "General"], ["c:dLbls/c:numFmt{formatCode=foobar}", "foobar"]], fn: (self: any, request: any): any => {
        let dLbls_cxml: any, data_labels: any, expected_value: any;
        [dLbls_cxml, expected_value] = $list(request.param);
        data_labels = $call(DataLabels, element(dLbls_cxml));
        return [data_labels, expected_value];
      } },
    number_format_set_fixture: { args: ["request"], params: [["c:dLbls", "General", "c:dLbls/c:numFmt{formatCode=General,sourceLinked=0}"], ["c:dLbls/c:numFmt{formatCode=General}", "00.00", "c:dLbls/c:numFmt{formatCode=00.00,sourceLinked=0}"]], fn: (self: any, request: any): any => {
        let dLbls_cxml: any, data_labels: any, expected_dLbls_cxml: any, expected_xml: any, new_value: any;
        [dLbls_cxml, new_value, expected_dLbls_cxml] = $list(request.param);
        data_labels = $call(DataLabels, element(dLbls_cxml));
        expected_xml = xml(expected_dLbls_cxml);
        return [data_labels, new_value, expected_xml];
      } },
    number_format_is_linked_get_fixture: { args: ["request"], params: [["c:dLbls", true], ["c:dLbls/c:numFmt", true], ["c:dLbls/c:numFmt{sourceLinked=0}", false], ["c:dLbls/c:numFmt{sourceLinked=1}", true]], fn: (self: any, request: any): any => {
        let dLbls_cxml: any, data_labels: any, expected_value: any;
        [dLbls_cxml, expected_value] = $list(request.param);
        data_labels = $call(DataLabels, element(dLbls_cxml));
        return [data_labels, expected_value];
      } },
    number_format_is_linked_set_fixture: { args: ["request"], params: [["c:dLbls", true, "c:dLbls/c:numFmt{sourceLinked=1}"], ["c:dLbls", false, "c:dLbls/c:numFmt{sourceLinked=0}"], ["c:dLbls", null, "c:dLbls/c:numFmt"], ["c:dLbls/c:numFmt", true, "c:dLbls/c:numFmt{sourceLinked=1}"], ["c:dLbls/c:numFmt{sourceLinked=1}", false, "c:dLbls/c:numFmt{sourceLinked=0}"]], fn: (self: any, request: any): any => {
        let dLbls_cxml: any, data_labels: any, expected_dLbls_cxml: any, expected_xml: any, new_value: any;
        [dLbls_cxml, new_value, expected_dLbls_cxml] = $list(request.param);
        data_labels = $call(DataLabels, element(dLbls_cxml));
        expected_xml = xml(expected_dLbls_cxml);
        return [data_labels, new_value, expected_xml];
      } },
    percent_get_fixture: { args: ["request"], params: [["c:dLbls", false], ["c:dLbls/c:showPercent{val=0}", false], ["c:dLbls/c:showPercent{val=1}", true], ["c:dLbls/c:showPercent{val=true}", true], ["c:dLbls/c:showPercent", true]], fn: (self: any, request: any): any => {
        let dLbls: any, dLbls_cxml: any, expected_value: any;
        [dLbls_cxml, expected_value] = $list(request.param);
        dLbls = element(dLbls_cxml);
        return [dLbls, expected_value];
      } },
    percent_set_fixture: { args: ["request"], params: [["c:dLbls", false, "c:dLbls/c:showPercent{val=0}"], ["c:dLbls", true, "c:dLbls/c:showPercent{val=1}"], ["c:dLbls/c:showPercent", false, "c:dLbls/c:showPercent{val=0}"], ["c:dLbls/c:showPercent{val=0}", false, "c:dLbls/c:showPercent{val=0}"], ["c:dLbls/c:showPercent{val=0}", true, "c:dLbls/c:showPercent{val=1}"], ["c:dLbls/c:showPercent{val=1}", false, "c:dLbls/c:showPercent{val=0}"]], fn: (self: any, request: any): any => {
        let dLbls: any, dLbls_cxml: any, expected_dLbls_cxml: any, expected_dLbls_xml: any, new_value: any;
        [dLbls_cxml, new_value, expected_dLbls_cxml] = $list(request.param);
        dLbls = element(dLbls_cxml);
        expected_dLbls_xml = xml(expected_dLbls_cxml);
        return [dLbls, new_value, expected_dLbls_xml];
      } },
    position_get_fixture: { args: ["request"], params: [["c:dLbls", null], ["c:dLbls/c:dLblPos{val=inBase}", XL_LABEL_POSITION.INSIDE_BASE]], fn: (self: any, request: any): any => {
        let dLbls_cxml: any, data_labels: any, expected_value: any;
        [dLbls_cxml, expected_value] = $list(request.param);
        data_labels = $call(DataLabels, element(dLbls_cxml));
        return [data_labels, expected_value];
      } },
    position_set_fixture: { args: ["request"], params: [["c:dLbls", XL_LABEL_POSITION.INSIDE_BASE, "c:dLbls/c:dLblPos{val=inBase}"], ["c:dLbls/c:dLblPos{val=inBase}", XL_LABEL_POSITION.OUTSIDE_END, "c:dLbls/c:dLblPos{val=outEnd}"], ["c:dLbls/c:dLblPos{val=inBase}", null, "c:dLbls"], ["c:dLbls", null, "c:dLbls"]], fn: (self: any, request: any): any => {
        let dLbls_cxml: any, data_labels: any, expected_dLbls_cxml: any, expected_xml: any, new_value: any;
        [dLbls_cxml, new_value, expected_dLbls_cxml] = $list(request.param);
        data_labels = $call(DataLabels, element(dLbls_cxml));
        expected_xml = xml(expected_dLbls_cxml);
        return [data_labels, new_value, expected_xml];
      } },
    sername_get_fixture: { args: ["request"], params: [["c:dLbls", false], ["c:dLbls/c:showSerName{val=0}", false], ["c:dLbls/c:showSerName{val=1}", true], ["c:dLbls/c:showSerName{val=true}", true], ["c:dLbls/c:showSerName", true]], fn: (self: any, request: any): any => {
        let dLbls: any, dLbls_cxml: any, expected_value: any;
        [dLbls_cxml, expected_value] = $list(request.param);
        dLbls = element(dLbls_cxml);
        return [dLbls, expected_value];
      } },
    sername_set_fixture: { args: ["request"], params: [["c:dLbls", false, "c:dLbls/c:showSerName{val=0}"], ["c:dLbls", true, "c:dLbls/c:showSerName{val=1}"], ["c:dLbls/c:showSerName", false, "c:dLbls/c:showSerName{val=0}"], ["c:dLbls/c:showSerName{val=0}", false, "c:dLbls/c:showSerName{val=0}"], ["c:dLbls/c:showSerName{val=0}", true, "c:dLbls/c:showSerName{val=1}"], ["c:dLbls/c:showSerName{val=1}", false, "c:dLbls/c:showSerName{val=0}"]], fn: (self: any, request: any): any => {
        let dLbls: any, dLbls_cxml: any, expected_dLbls_cxml: any, expected_dLbls_xml: any, new_value: any;
        [dLbls_cxml, new_value, expected_dLbls_cxml] = $list(request.param);
        dLbls = element(dLbls_cxml);
        expected_dLbls_xml = xml(expected_dLbls_cxml);
        return [dLbls, new_value, expected_dLbls_xml];
      } },
    txPr_fixture: { args: ["request"], params: [["c:dLbls{a:b=c}", "c:dLbls{a:b=c}/c:txPr/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr)"], ["c:dLbls{a:b=c}/c:txPr/(a:bodyPr,a:p)", "c:dLbls{a:b=c}/c:txPr/(a:bodyPr,a:p/a:pPr/a:defRPr)"], ["c:dLbls{a:b=c}/c:txPr/(a:bodyPr,a:p/a:pPr)", "c:dLbls{a:b=c}/c:txPr/(a:bodyPr,a:p/a:pPr/a:defRPr)"]], fn: (self: any, request: any): any => {
        let dLbls_cxml: any, data_labels: any, expected_cxml: any, expected_xml: any;
        [dLbls_cxml, expected_cxml] = $list(request.param);
        data_labels = $call(DataLabels, element(dLbls_cxml));
        expected_xml = xml(expected_cxml);
        return [data_labels, expected_xml];
      } },
    value_get_fixture: { args: ["request"], params: [["c:dLbls", false], ["c:dLbls/c:showVal{val=0}", false], ["c:dLbls/c:showVal{val=1}", true], ["c:dLbls/c:showVal{val=true}", true], ["c:dLbls/c:showVal", true]], fn: (self: any, request: any): any => {
        let dLbls: any, dLbls_cxml: any, expected_value: any;
        [dLbls_cxml, expected_value] = $list(request.param);
        dLbls = element(dLbls_cxml);
        return [dLbls, expected_value];
      } },
    value_set_fixture: { args: ["request"], params: [["c:dLbls", false, "c:dLbls/c:showVal{val=0}"], ["c:dLbls", true, "c:dLbls/c:showVal{val=1}"], ["c:dLbls/c:showVal", false, "c:dLbls/c:showVal{val=0}"], ["c:dLbls/c:showVal{val=0}", false, "c:dLbls/c:showVal{val=0}"], ["c:dLbls/c:showVal{val=0}", true, "c:dLbls/c:showVal{val=1}"], ["c:dLbls/c:showVal{val=1}", false, "c:dLbls/c:showVal{val=0}"]], fn: (self: any, request: any): any => {
        let dLbls: any, dLbls_cxml: any, expected_dLbls_cxml: any, expected_dLbls_xml: any, new_value: any;
        [dLbls_cxml, new_value, expected_dLbls_cxml] = $list(request.param);
        dLbls = element(dLbls_cxml);
        expected_dLbls_xml = xml(expected_dLbls_cxml);
        return [dLbls, new_value, expected_dLbls_xml];
      } },
    Font_: { args: ["request", "font_"], fn: (self: any, request: any, font_: any): any => {
        return $call(classMock, request, "pptx.chart.datalabel.Font", $kw({return_value: font_}));
      } },
    font_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Font);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_access_to_its_font", args: ["font_fixture"], fn: (self: any, font_fixture: any): any => {
        let Font_: any, data_labels: any, defRPr: any, font: any, font_: any;
        [data_labels, Font_, defRPr, font_] = $list(font_fixture);
        font = data_labels.font;
        Font_.assertCalledOnceWith(defRPr);
        $assert($is(font, font_));
      } },
    { name: "it_adds_a_txPr_to_help_font", args: ["txPr_fixture"], fn: (self: any, txPr_fixture: any): any => {
        let data_labels: any, expected_xml: any;
        [data_labels, expected_xml] = $list(txPr_fixture);
        data_labels.font;
        $assertEqual(data_labels._element.xml, expected_xml);
      } },
    { name: "it_knows_its_number_format", args: ["number_format_get_fixture"], fn: (self: any, number_format_get_fixture: any): any => {
        let data_labels: any, expected_value: any;
        [data_labels, expected_value] = $list(number_format_get_fixture);
        $assertEqual(data_labels.numberFormat, expected_value);
      } },
    { name: "it_can_change_its_number_format", args: ["number_format_set_fixture"], fn: (self: any, number_format_set_fixture: any): any => {
        let data_labels: any, expected_xml: any, new_value: any;
        [data_labels, new_value, expected_xml] = $list(number_format_set_fixture);
        data_labels.numberFormat = new_value;
        $assertEqual(data_labels._element.xml, expected_xml);
      } },
    { name: "it_knows_whether_its_number_format_is_linked", args: ["number_format_is_linked_get_fixture"], fn: (self: any, number_format_is_linked_get_fixture: any): any => {
        let data_labels: any, expected_value: any;
        [data_labels, expected_value] = $list(number_format_is_linked_get_fixture);
        $assert($is(data_labels.numberFormatIsLinked, expected_value));
      } },
    { name: "it_can_change_whether_its_number_format_is_linked", args: ["number_format_is_linked_set_fixture"], fn: (self: any, number_format_is_linked_set_fixture: any): any => {
        let data_labels: any, expected_xml: any, new_value: any;
        [data_labels, new_value, expected_xml] = $list(number_format_is_linked_set_fixture);
        data_labels.numberFormatIsLinked = new_value;
        $assertEqual(data_labels._element.xml, expected_xml);
      } },
    { name: "it_knows_its_position", args: ["position_get_fixture"], fn: (self: any, position_get_fixture: any): any => {
        let data_labels: any, expected_value: any;
        [data_labels, expected_value] = $list(position_get_fixture);
        $assertEqual(data_labels.position, expected_value);
      } },
    { name: "it_can_change_its_position", args: ["position_set_fixture"], fn: (self: any, position_set_fixture: any): any => {
        let data_labels: any, expected_xml: any, new_value: any;
        [data_labels, new_value, expected_xml] = $list(position_set_fixture);
        data_labels.position = new_value;
        $assertEqual(data_labels._element.xml, expected_xml);
      } },
    { name: "it_knows_whether_it_shows_category_name", args: ["catname_get_fixture"], fn: (self: any, catname_get_fixture: any): any => {
        let dLbls: any, data_labels: any, expected_value: any, show_category_name: any;
        [dLbls, expected_value] = $list(catname_get_fixture);
        data_labels = $call(DataLabels, dLbls);
        show_category_name = data_labels.showCategoryName;
        $assertEqual(show_category_name, expected_value);
      } },
    { name: "it_can_set_whether_it_shows_category_name", args: ["catname_set_fixture"], fn: (self: any, catname_set_fixture: any): any => {
        let dLbls: any, data_labels: any, expected_xml: any, new_value: any;
        [dLbls, new_value, expected_xml] = $list(catname_set_fixture);
        data_labels = $call(DataLabels, dLbls);
        data_labels.showCategoryName = new_value;
        $assertEqual(data_labels._element.xml, expected_xml);
      } },
    { name: "it_knows_whether_it_shows_legend_key", args: ["lgndkey_get_fixture"], fn: (self: any, lgndkey_get_fixture: any): any => {
        let dLbls: any, data_labels: any, expected_value: any, show_legend_key: any;
        [dLbls, expected_value] = $list(lgndkey_get_fixture);
        data_labels = $call(DataLabels, dLbls);
        show_legend_key = data_labels.showLegendKey;
        $assertEqual(show_legend_key, expected_value);
      } },
    { name: "it_can_set_whether_it_shows_legend_key", args: ["lgndkey_set_fixture"], fn: (self: any, lgndkey_set_fixture: any): any => {
        let dLbls: any, data_labels: any, expected_xml: any, new_value: any;
        [dLbls, new_value, expected_xml] = $list(lgndkey_set_fixture);
        data_labels = $call(DataLabels, dLbls);
        data_labels.showLegendKey = new_value;
        $assertEqual(data_labels._element.xml, expected_xml);
      } },
    { name: "it_knows_whether_it_shows_percentage", args: ["percent_get_fixture"], fn: (self: any, percent_get_fixture: any): any => {
        let dLbls: any, data_labels: any, expected_value: any, show_percentage: any;
        [dLbls, expected_value] = $list(percent_get_fixture);
        data_labels = $call(DataLabels, dLbls);
        show_percentage = data_labels.showPercentage;
        $assertEqual(show_percentage, expected_value);
      } },
    { name: "it_can_set_whether_it_shows_percentage", args: ["percent_set_fixture"], fn: (self: any, percent_set_fixture: any): any => {
        let dLbls: any, data_labels: any, expected_xml: any, new_value: any;
        [dLbls, new_value, expected_xml] = $list(percent_set_fixture);
        data_labels = $call(DataLabels, dLbls);
        data_labels.showPercentage = new_value;
        $assertEqual(data_labels._element.xml, expected_xml);
      } },
    { name: "it_knows_whether_it_shows_series_name", args: ["sername_get_fixture"], fn: (self: any, sername_get_fixture: any): any => {
        let dLbls: any, data_labels: any, expected_value: any, show_series_name: any;
        [dLbls, expected_value] = $list(sername_get_fixture);
        data_labels = $call(DataLabels, dLbls);
        show_series_name = data_labels.showSeriesName;
        $assertEqual(show_series_name, expected_value);
      } },
    { name: "it_can_set_whether_it_shows_series_name", args: ["sername_set_fixture"], fn: (self: any, sername_set_fixture: any): any => {
        let dLbls: any, data_labels: any, expected_xml: any, new_value: any;
        [dLbls, new_value, expected_xml] = $list(sername_set_fixture);
        data_labels = $call(DataLabels, dLbls);
        data_labels.showSeriesName = new_value;
        $assertEqual(data_labels._element.xml, expected_xml);
      } },
    { name: "it_knows_whether_it_shows_value", args: ["value_get_fixture"], fn: (self: any, value_get_fixture: any): any => {
        let dLbls: any, data_labels: any, expected_value: any, show_value: any;
        [dLbls, expected_value] = $list(value_get_fixture);
        data_labels = $call(DataLabels, dLbls);
        show_value = data_labels.showValue;
        $assertEqual(show_value, expected_value);
      } },
    { name: "it_can_set_whether_it_shows_value", args: ["value_set_fixture"], fn: (self: any, value_set_fixture: any): any => {
        let dLbls: any, data_labels: any, expected_xml: any, new_value: any;
        [dLbls, new_value, expected_xml] = $list(value_set_fixture);
        data_labels = $call(DataLabels, dLbls);
        data_labels.showValue = new_value;
        $assertEqual(data_labels._element.xml, expected_xml);
      } },
  ],
};
pytestClass("DescribeDataLabels", DescribeDataLabels);

