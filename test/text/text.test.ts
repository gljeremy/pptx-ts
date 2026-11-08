// Machine-translated from python-pptx tests/text/test_text.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $call, $cast, $clear, $enumerate, $get, $getattr, $is, $isinstance, $kw, $len, $list, $mod, $raises, $setattr, $slice, classMock, instanceMock, looseMock, methodMock, propertyMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { ColorFormat as ColorFormat$imported } from "../../src/dml/color.js";
import { FillFormat as FillFormat$imported } from "../../src/dml/fill.js";
import { MSO_LANGUAGE_ID as MSO_LANGUAGE_ID$imported } from "../../src/enum/lang.js";
import { MSO_ANCHOR as MSO_ANCHOR$imported, MSO_AUTO_SIZE as MSO_AUTO_SIZE$imported, MSO_UNDERLINE as MSO_UNDERLINE$imported, PP_ALIGN as PP_ALIGN$imported } from "../../src/enum/text.js";
import { RELATIONSHIP_TYPE as RT$imported } from "../../src/opc/constants.js";
import { XmlPart as XmlPart$imported } from "../../src/opc/package.js";
import { Shape as Shape$imported } from "../../src/shapes/autoshape.js";
import { Font as Font$imported, TextFrame as TextFrame$imported, _Hyperlink as _Hyperlink$imported, _Paragraph as _Paragraph$imported, _Run as _Run$imported } from "../../src/text/text.js";
import { Inches as Inches$imported, Pt as Pt$imported } from "../../src/util.js";
import { aP, aT, anHlinkClick, anR, anRPr } from "../oxml/unitdata/text.js";
import { element, xml } from "../unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const ColorFormat = ColorFormat$imported;
const FillFormat = FillFormat$imported;
const MSO_LANGUAGE_ID = MSO_LANGUAGE_ID$imported;
const MSO_ANCHOR = MSO_ANCHOR$imported;
const MSO_AUTO_SIZE = MSO_AUTO_SIZE$imported;
const MSO_UNDERLINE = MSO_UNDERLINE$imported;
const PP_ALIGN = PP_ALIGN$imported;
const RT = RT$imported;
const XmlPart = XmlPart$imported;
const Shape = Shape$imported;
const Font = Font$imported;
const TextFrame = TextFrame$imported;
const _Hyperlink = _Hyperlink$imported;
const _Paragraph = _Paragraph$imported;
const _Run = _Run$imported;
const Inches = Inches$imported;
const Pt = Pt$imported;
void $py;



const DescribeTextFrame: any = {
  fixtures: {
    add_paragraph_fixture: { args: ["request"], params: [["p:txBody/a:bodyPr", "p:txBody/(a:bodyPr,a:p)"], ["p:txBody/(a:bodyPr,a:p)", "p:txBody/(a:bodyPr,a:p,a:p)"]], fn: (self: any, request: any): any => {
        let expected_cxml: any, expected_xml: any, text_frame: any, txBody_cxml: any;
        [txBody_cxml, expected_cxml] = $list(request.param);
        text_frame = $call(TextFrame, element(txBody_cxml), null);
        expected_xml = xml(expected_cxml);
        return [text_frame, expected_xml];
      } },
    autosize_get_fixture: { args: ["request"], params: [["p:txBody/a:bodyPr", null], ["p:txBody/a:bodyPr/a:noAutofit", MSO_AUTO_SIZE.NONE], ["p:txBody/a:bodyPr/a:spAutoFit", MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT], ["p:txBody/a:bodyPr/a:normAutofit", MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE]], fn: (self: any, request: any): any => {
        let expected_value: any, text_frame: any, txBody_cxml: any;
        [txBody_cxml, expected_value] = $list(request.param);
        text_frame = $call(TextFrame, element(txBody_cxml), null);
        return [text_frame, expected_value];
      } },
    margin_get_fixture: { args: ["request"], params: [["p:txBody/a:bodyPr", "left", "emu", $call(Inches, 0.1)], ["p:txBody/a:bodyPr", "top", "emu", $call(Inches, 0.05)], ["p:txBody/a:bodyPr", "right", "emu", $call(Inches, 0.1)], ["p:txBody/a:bodyPr", "bottom", "emu", $call(Inches, 0.05)], ["p:txBody/a:bodyPr{lIns=9144}", "left", "cm", 0.0254], ["p:txBody/a:bodyPr{tIns=18288}", "top", "mm", 0.508], ["p:txBody/a:bodyPr{rIns=76200}", "right", "pt", 6.0], ["p:txBody/a:bodyPr{bIns=36576}", "bottom", "inches", 0.04]], fn: (self: any, request: any): any => {
        let expected_value: any, prop_name: any, side: any, text_frame: any, txBody_cxml: any, unit: any;
        [txBody_cxml, side, unit, expected_value] = $list(request.param);
        text_frame = $call(TextFrame, element(txBody_cxml), null);
        prop_name = $mod("margin_%s", side);
        return [text_frame, prop_name, unit, expected_value];
      } },
    margin_set_fixture: { args: ["request"], params: [["p:txBody/a:bodyPr", "left", $call(Inches, 0.11), "p:txBody/a:bodyPr{lIns=100584}"], ["p:txBody/a:bodyPr{tIns=1234}", "top", $call(Inches, 0.12), "p:txBody/a:bodyPr{tIns=109728}"], ["p:txBody/a:bodyPr{rIns=2345}", "right", $call(Inches, 0.13), "p:txBody/a:bodyPr{rIns=118872}"], ["p:txBody/a:bodyPr{bIns=3456}", "bottom", $call(Inches, 0.14), "p:txBody/a:bodyPr{bIns=128016}"], ["p:txBody/a:bodyPr", "left", $call(Inches, 0.1), "p:txBody/a:bodyPr"], ["p:txBody/a:bodyPr", "top", $call(Inches, 0.05), "p:txBody/a:bodyPr"], ["p:txBody/a:bodyPr", "right", $call(Inches, 0.1), "p:txBody/a:bodyPr"], ["p:txBody/a:bodyPr", "bottom", $call(Inches, 0.05), "p:txBody/a:bodyPr"]], fn: (self: any, request: any): any => {
        let expected_txBody_cxml: any, expected_xml: any, new_value: any, prop_name: any, side: any, text_frame: any, txBody_cxml: any;
        [txBody_cxml, side, new_value, expected_txBody_cxml] = $list(request.param);
        text_frame = $call(TextFrame, element(txBody_cxml), null);
        prop_name = $mod("margin_%s", side);
        expected_xml = xml(expected_txBody_cxml);
        return [text_frame, prop_name, new_value, expected_xml];
      } },
    paragraphs_fixture: { args: ["request"], params: ["p:txBody", "p:txBody/a:p", "p:txBody/(a:p,a:p)"], fn: (self: any, request: any): any => {
        let ps: any, text_frame: any, txBody: any, txBody_cxml: any;
        txBody_cxml = request.param;
        txBody = element(txBody_cxml);
        text_frame = $call(TextFrame, txBody, null);
        ps = txBody.xpath(".//a:p");
        return [text_frame, ps];
      } },
    set_font_fixture: { args: ["request"], params: [["p:txBody/(a:bodyPr,a:p/a:r)", true, false, "p:txBody/(a:bodyPr,a:p/(a:r/a:rPr{sz=600,b=1,i=0}/a:latin{typeface=F},a:endParaRPr{sz=600,b=1,i=0}/a:latin{typeface=F}))"], ["p:txBody/a:p/a:br", true, false, "p:txBody/a:p/(a:br/a:rPr{sz=600,b=1,i=0}/a:latin{typeface=F},a:endParaRPr{sz=600,b=1,i=0}/a:latin{typeface=F})"], ["p:txBody/a:p/a:fld", true, false, "p:txBody/a:p/(a:fld/a:rPr{sz=600,b=1,i=0}/a:latin{typeface=F},a:endParaRPr{sz=600,b=1,i=0}/a:latin{typeface=F})"]], fn: (self: any, request: any): any => {
        let bold: any, expected_cxml: any, expected_xml: any, family: any, italic: any, size: any, text_frame: any, txBody_cxml: any;
        [txBody_cxml, bold, italic, expected_cxml] = $list(request.param);
        [family, size] = $list(["F", 6]);
        text_frame = $call(TextFrame, element(txBody_cxml), null);
        expected_xml = xml(expected_cxml);
        return [text_frame, family, size, bold, italic, expected_xml];
      } },
    size_font_fixture: { args: ["FontFiles_", "TextFitter_", "text_prop_", "_extents_prop_"], fn: (self: any, FontFiles_: any, TextFitter_: any, text_prop_: any, _extents_prop_: any): any => {
        let bold: any, extents: any, family: any, font_file: any, font_size: any, italic: any, max_size: any, text: any, text_frame: any;
        text_frame = $call(TextFrame, null, null);
        [family, max_size, bold, italic] = $list(["Family", 42, true, false]);
        [text, extents, font_size, font_file] = $list(["text", [111, 222], 21, "f.ttf"]);
        text_prop_.returnValue = text;
        _extents_prop_.returnValue = extents;
        FontFiles_.find.returnValue = font_file;
        TextFitter_.bestFitFontSize.returnValue = font_size;
        return [text_frame, family, max_size, bold, italic, FontFiles_, TextFitter_, text, extents, font_file, font_size];
      } },
    text_get_fixture: { args: ["request"], params: [[["foobar"], "foobar"], [["foo", "bar", "baz"], "foo\nbar\nbaz"]], fn: (self: any, request: any): any => {
        let expected_value: any, paragraph_texts: any;
        [paragraph_texts, expected_value] = $list(request.param);
        return [paragraph_texts, expected_value];
      } },
    text_set_fixture: { args: ["request"], params: [["p:txBody/a:p", "foobar", "p:txBody/a:p/a:r/a:t\"foobar\""], ["p:txBody/a:p/a:r/a:t\"foobar\"", "barfoo", "p:txBody/a:p/a:r/a:t\"barfoo\""], ["p:txBody/a:p/a:r/a:t\"barfoo\"", "foo\nbar", "p:txBody/(a:p/a:r/a:t\"foo\",a:p/a:r/a:t\"bar\")"], ["p:txBody/(a:p/a:r/a:t\"foo\",a:p/a:r/a:t\"bar\")", "barfoo", "p:txBody/a:p/a:r/a:t\"barfoo\""], ["p:txBody/a:p/a:r/a:t\"foobar\"", "", "p:txBody/a:p"], ["p:txBody/a:p", "a\u000bb", "p:txBody/a:p/(a:r/a:t\"a\",a:br,a:r/a:t\"b\")"]], fn: (self: any, request: any): any => {
        let expected_cxml: any, expected_xml: any, text: any, txBody: any, txBody_cxml: any;
        [txBody_cxml, text, expected_cxml] = $list(request.param);
        txBody = element(txBody_cxml);
        expected_xml = xml(expected_cxml);
        return [txBody, text, expected_xml];
      } },
    wrap_get_fixture: { args: ["request"], params: [["p:txBody/a:bodyPr", null], ["p:txBody/a:bodyPr{wrap=square}", true], ["p:txBody/a:bodyPr{wrap=none}", false]], fn: (self: any, request: any): any => {
        let expected_value: any, text_frame: any, txBody_cxml: any;
        [txBody_cxml, expected_value] = $list(request.param);
        text_frame = $call(TextFrame, element(txBody_cxml), null);
        return [text_frame, expected_value];
      } },
    wrap_set_fixture: { args: ["request"], params: [["p:txBody/a:bodyPr", true, "p:txBody/a:bodyPr{wrap=square}"], ["p:txBody/a:bodyPr{wrap=square}", false, "p:txBody/a:bodyPr{wrap=none}"], ["p:txBody/a:bodyPr{wrap=none}", null, "p:txBody/a:bodyPr"]], fn: (self: any, request: any): any => {
        let expected_txBody_cxml: any, expected_xml: any, new_value: any, text_frame: any, txBody_cxml: any;
        [txBody_cxml, new_value, expected_txBody_cxml] = $list(request.param);
        text_frame = $call(TextFrame, element(txBody_cxml), null);
        expected_xml = xml(expected_txBody_cxml);
        return [text_frame, new_value, expected_xml];
      } },
    _extents_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, TextFrame, "_extents");
      } },
    FontFiles_: { args: ["request"], fn: (self: any, request: any): any => {
        return classMock(request, "pptx.text.text.FontFiles");
      } },
    paragraphs_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, TextFrame, "paragraphs");
      } },
    TextFitter_: { args: ["request"], fn: (self: any, request: any): any => {
        return classMock(request, "pptx.text.text.TextFitter");
      } },
    text_frame_with_parent_: { args: ["request"], fn: (self: any, request: any): any => {
        let parent_: any, text_frame: any;
        parent_ = $call(looseMock, request, $kw({name: "parent_"}));
        text_frame = $call(TextFrame, null, parent_);
        return [text_frame, parent_];
      } },
    text_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, TextFrame, "text");
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_add_a_paragraph_to_itself", args: ["add_paragraph_fixture"], fn: (self: any, add_paragraph_fixture: any): any => {
        let expected_xml: any, text_frame: any;
        [text_frame, expected_xml] = $list(add_paragraph_fixture);
        text_frame.addParagraph();
        $assertEqual(text_frame._txBody.xml, expected_xml);
      } },
    { name: "it_knows_its_autosize_setting", args: ["autosize_get_fixture"], fn: (self: any, autosize_get_fixture: any): any => {
        let expected_value: any, text_frame: any;
        [text_frame, expected_value] = $list(autosize_get_fixture);
        $assertEqual(text_frame.autoSize, expected_value);
      } },
    { name: "it_can_change_its_autosize_setting", args: ["txBody_cxml", "value", "expected_cxml"], parametrize: [{ names: ["txBody_cxml", "value", "expected_cxml"], values: [["p:txBody/a:bodyPr", MSO_AUTO_SIZE.NONE, "p:txBody/a:bodyPr/a:noAutofit"], ["p:txBody/a:bodyPr/a:noAutofit", MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT, "p:txBody/a:bodyPr/a:spAutoFit"], ["p:txBody/a:bodyPr/a:spAutoFit", MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE, "p:txBody/a:bodyPr/a:normAutofit"], ["p:txBody/a:bodyPr/a:normAutofit", null, "p:txBody/a:bodyPr"]] }], fn: (self: any, txBody_cxml: any, value: any, expected_cxml: any): any => {
        let text_frame: any;
        text_frame = $call(TextFrame, element(txBody_cxml), null);
        text_frame.autoSize = value;
        $assertEqual(text_frame._txBody.xml, xml(expected_cxml));
      } },
    { name: "it_can_clear_itself_of_content", args: ["txBody_cxml"], parametrize: [{ names: ["txBody_cxml"], values: ["p:txBody/(a:p,a:p,a:p)", "p:txBody/a:p/a:r/a:t\"foo\"", "p:txBody/a:p/(a:br,a:r/a:t\"foo\")", "p:txBody/a:p/(a:fld,a:br,a:r/a:t\"foo\")"] }], fn: (self: any, txBody_cxml: any): any => {
        let text_frame: any;
        text_frame = $call(TextFrame, element(txBody_cxml), null);
        $clear(text_frame);
        $assertEqual(text_frame._element.xml, xml("p:txBody/a:p"));
      } },
    { name: "it_knows_its_margin_settings", args: ["margin_get_fixture"], fn: (self: any, margin_get_fixture: any): any => {
        let expected_value: any, margin_value: any, prop_name: any, text_frame: any, unit: any;
        [text_frame, prop_name, unit, expected_value] = $list(margin_get_fixture);
        margin_value = $getattr(text_frame, prop_name);
        $assertEqual($getattr(margin_value, unit), expected_value);
      } },
    { name: "it_can_change_its_margin_settings", args: ["margin_set_fixture"], fn: (self: any, margin_set_fixture: any): any => {
        let expected_xml: any, new_value: any, prop_name: any, text_frame: any;
        [text_frame, prop_name, new_value, expected_xml] = $list(margin_set_fixture);
        $setattr(text_frame, prop_name, new_value);
        $assertEqual(text_frame._txBody.xml, expected_xml);
      } },
    { name: "it_knows_its_vertical_alignment", args: ["txBody_cxml", "expected_value"], parametrize: [{ names: ["txBody_cxml", "expected_value"], values: [["p:txBody/a:bodyPr", null], ["p:txBody/a:bodyPr{anchor=t}", MSO_ANCHOR.TOP], ["p:txBody/a:bodyPr{anchor=b}", MSO_ANCHOR.BOTTOM]] }], fn: (self: any, txBody_cxml: any, expected_value: any): any => {
        let text_frame: any;
        text_frame = $call(TextFrame, $cast("CT_TextBody", element(txBody_cxml)), null);
        $assertEqual(text_frame.verticalAnchor, expected_value);
      } },
    { name: "it_can_change_its_vertical_alignment", args: ["txBody_cxml", "new_value", "expected_cxml"], parametrize: [{ names: ["txBody_cxml", "new_value", "expected_cxml"], values: [["p:txBody/a:bodyPr", MSO_ANCHOR.TOP, "p:txBody/a:bodyPr{anchor=t}"], ["p:txBody/a:bodyPr{anchor=t}", MSO_ANCHOR.MIDDLE, "p:txBody/a:bodyPr{anchor=ctr}"], ["p:txBody/a:bodyPr{anchor=ctr}", MSO_ANCHOR.BOTTOM, "p:txBody/a:bodyPr{anchor=b}"], ["p:txBody/a:bodyPr{anchor=b}", null, "p:txBody/a:bodyPr"]] }], fn: (self: any, txBody_cxml: any, new_value: any, expected_cxml: any): any => {
        let text_frame: any;
        text_frame = $call(TextFrame, $cast("CT_TextBody", element(txBody_cxml)), null);
        text_frame.verticalAnchor = new_value;
        $assertEqual(text_frame._element.xml, xml(expected_cxml));
      } },
    { name: "it_knows_its_word_wrap_setting", args: ["wrap_get_fixture"], fn: (self: any, wrap_get_fixture: any): any => {
        let expected_value: any, text_frame: any;
        [text_frame, expected_value] = $list(wrap_get_fixture);
        $assertEqual(text_frame.wordWrap, expected_value);
      } },
    { name: "it_can_change_its_word_wrap_setting", args: ["wrap_set_fixture"], fn: (self: any, wrap_set_fixture: any): any => {
        let expected_xml: any, new_value: any, text_frame: any;
        [text_frame, new_value, expected_xml] = $list(wrap_set_fixture);
        text_frame.wordWrap = new_value;
        $assertEqual(text_frame._element.xml, expected_xml);
      } },
    { name: "it_provides_access_to_its_paragraphs", args: ["paragraphs_fixture"], fn: (self: any, paragraphs_fixture: any): any => {
        let paragraphs: any, ps: any, text_frame: any;
        [text_frame, ps] = $list(paragraphs_fixture);
        paragraphs = text_frame.paragraphs;
        $assertEqual($len(paragraphs), $len(ps));
        for (const [idx, paragraph] of $list($enumerate(paragraphs))) {
          $assert($isinstance(paragraph, _Paragraph));
          $assert($is(paragraph._element, $get(ps, idx)));
        }
      } },
    { name: "it_raises_on_attempt_to_set_margin_to_non_int", args: [], fn: (self: any): any => {
        let text_frame: any;
        text_frame = $call(TextFrame, element("p:txBody/a:bodyPr"), null);
        $raises(TypeError, () => {
          text_frame.marginBottom = "0.1";
        });
      } },
    { name: "it_knows_the_part_it_belongs_to", args: ["text_frame_with_parent_"], fn: (self: any, text_frame_with_parent_: any): any => {
        let parent_: any, part: any, text_frame: any;
        [text_frame, parent_] = $list(text_frame_with_parent_);
        part = text_frame.part;
        $assert($is(part, parent_.part));
      } },
    { name: "it_knows_what_text_it_contains", args: ["request", "text_get_fixture", "paragraphs_prop_"], fn: (self: any, request: any, text_get_fixture: any, paragraphs_prop_: any): any => {
        let expected_value: any, paragraph_texts: any, text: any, text_frame: any;
        [paragraph_texts, expected_value] = $list(text_get_fixture);
        paragraphs_prop_.returnValue = $list($list(paragraph_texts).map((text: any) => $call(instanceMock, request, _Paragraph, $kw({text: text}))));
        text_frame = $call(TextFrame, null, null);
        text = text_frame.text;
        $assertEqual(text, expected_value);
      } },
    { name: "it_can_replace_the_text_it_contains", args: ["text_set_fixture"], fn: (self: any, text_set_fixture: any): any => {
        let expected_xml: any, text: any, text_frame: any, txBody: any;
        [txBody, text, expected_xml] = $list(text_set_fixture);
        text_frame = $call(TextFrame, txBody, null);
        text_frame.text = text;
        $assertEqual(text_frame._element.xml, expected_xml);
      } },
    { name: "it_can_resize_its_text_to_best_fit", args: ["request", "text_prop_"], fn: (self: any, request: any, text_prop_: any): any => {
        let _apply_fit_: any, _best_fit_font_size_: any, bold: any, family: any, font_file: any, font_size: any, italic: any, max_size: any, text_frame: any;
        [family, max_size, bold, italic, font_file, font_size] = $list(["Family", 42, "bold", "italic", "font_file", 21]);
        text_prop_.returnValue = "some text";
        _best_fit_font_size_ = $call(methodMock, request, TextFrame, "_best_fit_font_size", $kw({return_value: font_size}));
        _apply_fit_ = methodMock(request, TextFrame, "_apply_fit");
        text_frame = $call(TextFrame, null, null);
        text_frame.fitText(family, max_size, bold, italic, font_file);
        _best_fit_font_size_.assertCalledOnceWith(text_frame, family, max_size, bold, italic, font_file);
        _apply_fit_.assertCalledOnceWith(text_frame, family, font_size, bold, italic);
      } },
    { name: "it_calculates_its_best_fit_font_size_to_help_fit_text", args: ["size_font_fixture"], fn: (self: any, size_font_fixture: any): any => {
        let FontFiles_: any, TextFitter_: any, bold: any, extents: any, family: any, font_file_: any, font_size: any, font_size_: any, italic: any, max_size: any, text: any, text_frame: any;
        [text_frame, family, max_size, bold, italic] = $list($slice(size_font_fixture, null, 5, null));
        [FontFiles_, TextFitter_, text, extents] = $list($slice(size_font_fixture, 5, 9, null));
        [font_file_, font_size_] = $list($slice(size_font_fixture, 9, null, null));
        font_size = text_frame._bestFitFontSize(family, max_size, bold, italic, null);
        FontFiles_.find.assertCalledOnceWith(family, bold, italic);
        TextFitter_.bestFitFontSize.assertCalledOnceWith(text, extents, max_size, font_file_);
        $assert($is(font_size, font_size_));
      } },
    { name: "it_calculates_its_effective_size_to_help_fit_text", args: [], fn: (self: any): any => {
        let sp_cxml: any, text_frame: any;
        sp_cxml = "p:sp/(p:spPr/a:xfrm/(a:off{x=914400,y=914400},a:ext{cx=914400,cy=914400}),p:txBody/(a:bodyPr,a:p))";
        text_frame = $call(Shape, element(sp_cxml), null).textFrame;
        $assertEqual(text_frame._extents, [731520, 822960]);
      } },
    { name: "it_applies_fit_to_help_fit_text", args: ["request"], fn: (self: any, request: any): any => {
        let _set_font_: any, bold: any, family: any, font_size: any, italic: any, text_frame: any;
        [family, font_size, bold, italic] = $list(["Family", 42, true, false]);
        _set_font_ = methodMock(request, TextFrame, "_set_font");
        text_frame = $call(TextFrame, element("p:txBody/a:bodyPr"), null);
        text_frame._applyFit(family, font_size, bold, italic);
        $assert($is(text_frame.autoSize, MSO_AUTO_SIZE.NONE));
        $assert($is(text_frame.wordWrap, true));
        _set_font_.assertCalledOnceWith(text_frame, family, font_size, bold, italic);
      } },
    { name: "it_sets_its_font_to_help_fit_text", args: ["set_font_fixture"], fn: (self: any, set_font_fixture: any): any => {
        let bold: any, expected_xml: any, family: any, italic: any, size: any, text_frame: any;
        [text_frame, family, size, bold, italic, expected_xml] = $list(set_font_fixture);
        text_frame._setFont(family, size, bold, italic);
        $assertEqual(text_frame._element.xml, expected_xml);
      } },
  ],
};
pytestClass("DescribeTextFrame", DescribeTextFrame);

const DescribeFont: any = {
  fixtures: {
    bold_get_fixture: { args: ["request"], params: [["a:rPr", null], ["a:rPr{b=0}", false], ["a:rPr{b=1}", true]], fn: (self: any, request: any): any => {
        let expected_value: any, font: any, rPr_cxml: any;
        [rPr_cxml, expected_value] = $list(request.param);
        font = $call(Font, element(rPr_cxml));
        return [font, expected_value];
      } },
    bold_set_fixture: { args: ["request"], params: [["a:rPr", true, "a:rPr{b=1}"], ["a:rPr{b=1}", false, "a:rPr{b=0}"], ["a:rPr{b=0}", null, "a:rPr"]], fn: (self: any, request: any): any => {
        let expected_rPr_cxml: any, expected_xml: any, font: any, new_value: any, rPr_cxml: any;
        [rPr_cxml, new_value, expected_rPr_cxml] = $list(request.param);
        font = $call(Font, element(rPr_cxml));
        expected_xml = xml(expected_rPr_cxml);
        return [font, new_value, expected_xml];
      } },
    italic_get_fixture: { args: ["request"], params: [["a:rPr", null], ["a:rPr{i=0}", false], ["a:rPr{i=1}", true]], fn: (self: any, request: any): any => {
        let expected_value: any, font: any, rPr_cxml: any;
        [rPr_cxml, expected_value] = $list(request.param);
        font = $call(Font, element(rPr_cxml));
        return [font, expected_value];
      } },
    italic_set_fixture: { args: ["request"], params: [["a:rPr", true, "a:rPr{i=1}"], ["a:rPr{i=1}", false, "a:rPr{i=0}"], ["a:rPr{i=0}", null, "a:rPr"]], fn: (self: any, request: any): any => {
        let expected_rPr_cxml: any, expected_xml: any, font: any, new_value: any, rPr_cxml: any;
        [rPr_cxml, new_value, expected_rPr_cxml] = $list(request.param);
        font = $call(Font, element(rPr_cxml));
        expected_xml = xml(expected_rPr_cxml);
        return [font, new_value, expected_xml];
      } },
    language_id_get_fixture: { args: ["request"], params: [["a:rPr", MSO_LANGUAGE_ID.NONE], ["a:rPr{lang=pl-PL}", MSO_LANGUAGE_ID.POLISH], ["a:rPr{lang=de-AT}", MSO_LANGUAGE_ID.GERMAN_AUSTRIA], ["a:rPr{lang=fr-FR}", MSO_LANGUAGE_ID.FRENCH]], fn: (self: any, request: any): any => {
        let expected_value: any, font: any, rPr_cxml: any;
        [rPr_cxml, expected_value] = $list(request.param);
        font = $call(Font, element(rPr_cxml));
        return [font, expected_value];
      } },
    language_id_set_fixture: { args: ["request"], params: [["a:rPr", MSO_LANGUAGE_ID.ZULU, "a:rPr{lang=zu-ZA}"], ["a:rPr{lang=zu-ZA}", MSO_LANGUAGE_ID.URDU, "a:rPr{lang=ur-PK}"], ["a:rPr{lang=ur-PK}", MSO_LANGUAGE_ID.NONE, "a:rPr"], ["a:rPr{lang=ur-PK}", null, "a:rPr"], ["a:rPr", MSO_LANGUAGE_ID.NONE, "a:rPr"], ["a:rPr", null, "a:rPr"]], fn: (self: any, request: any): any => {
        let expected_rPr_cxml: any, expected_xml: any, font: any, new_value: any, rPr_cxml: any;
        [rPr_cxml, new_value, expected_rPr_cxml] = $list(request.param);
        font = $call(Font, element(rPr_cxml));
        expected_xml = xml(expected_rPr_cxml);
        return [font, new_value, expected_xml];
      } },
    name_get_fixture: { args: ["request"], params: [["a:rPr", null], ["a:rPr/a:latin{typeface=Foobar}", "Foobar"]], fn: (self: any, request: any): any => {
        let expected_value: any, font: any, rPr_cxml: any;
        [rPr_cxml, expected_value] = $list(request.param);
        font = $call(Font, element(rPr_cxml));
        return [font, expected_value];
      } },
    name_set_fixture: { args: ["request"], params: [["a:rPr", "Foobar", "a:rPr/a:latin{typeface=Foobar}"], ["a:rPr/a:latin{typeface=Foobar}", "Barfoo", "a:rPr/a:latin{typeface=Barfoo}"], ["a:rPr/a:latin{typeface=Barfoo}", null, "a:rPr"]], fn: (self: any, request: any): any => {
        let expected_rPr_cxml: any, expected_xml: any, font: any, new_value: any, rPr_cxml: any;
        [rPr_cxml, new_value, expected_rPr_cxml] = $list(request.param);
        font = $call(Font, element(rPr_cxml));
        expected_xml = xml(expected_rPr_cxml);
        return [font, new_value, expected_xml];
      } },
    size_get_fixture: { args: ["request"], params: [["a:rPr", null], ["a:rPr{sz=2400}", 304800]], fn: (self: any, request: any): any => {
        let expected_value: any, font: any, rPr_cxml: any;
        [rPr_cxml, expected_value] = $list(request.param);
        font = $call(Font, element(rPr_cxml));
        return [font, expected_value];
      } },
    size_set_fixture: { args: ["request"], params: [["a:rPr", $call(Pt, 24), "a:rPr{sz=2400}"], ["a:rPr{sz=2400}", null, "a:rPr"]], fn: (self: any, request: any): any => {
        let expected_rPr_cxml: any, expected_xml: any, font: any, new_value: any, rPr_cxml: any;
        [rPr_cxml, new_value, expected_rPr_cxml] = $list(request.param);
        font = $call(Font, element(rPr_cxml));
        expected_xml = xml(expected_rPr_cxml);
        return [font, new_value, expected_xml];
      } },
    underline_get_fixture: { args: ["request"], params: [["a:rPr", null], ["a:rPr{u=none}", false], ["a:rPr{u=sng}", true], ["a:rPr{u=dbl}", MSO_UNDERLINE.DOUBLE_LINE]], fn: (self: any, request: any): any => {
        let expected_value: any, font: any, rPr_cxml: any;
        [rPr_cxml, expected_value] = $list(request.param);
        font = $call(Font, element(rPr_cxml));
        return [font, expected_value];
      } },
    underline_set_fixture: { args: ["request"], params: [["a:rPr", true, "a:rPr{u=sng}"], ["a:rPr{u=sng}", false, "a:rPr{u=none}"], ["a:rPr{u=none}", MSO_UNDERLINE.WAVY_LINE, "a:rPr{u=wavy}"], ["a:rPr{u=wavy}", MSO_UNDERLINE.NONE, "a:rPr{u=none}"], ["a:rPr{u=wavy}", null, "a:rPr"]], fn: (self: any, request: any): any => {
        let expected_rPr_cxml: any, expected_xml: any, font: any, new_value: any, rPr_cxml: any;
        [rPr_cxml, new_value, expected_rPr_cxml] = $list(request.param);
        font = $call(Font, element(rPr_cxml));
        expected_xml = xml(expected_rPr_cxml);
        return [font, new_value, expected_xml];
      } },
    font: { args: [], fn: (self: any): any => {
        return $call(Font, element("a:rPr"));
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_bold_setting", args: ["bold_get_fixture"], fn: (self: any, bold_get_fixture: any): any => {
        let expected_value: any, font: any;
        [font, expected_value] = $list(bold_get_fixture);
        $assertEqual(font.bold, expected_value);
      } },
    { name: "it_can_change_its_bold_setting", args: ["bold_set_fixture"], fn: (self: any, bold_set_fixture: any): any => {
        let expected_xml: any, font: any, new_value: any;
        [font, new_value, expected_xml] = $list(bold_set_fixture);
        font.bold = new_value;
        $assertEqual(font._element.xml, expected_xml);
      } },
    { name: "it_knows_its_italic_setting", args: ["italic_get_fixture"], fn: (self: any, italic_get_fixture: any): any => {
        let expected_value: any, font: any;
        [font, expected_value] = $list(italic_get_fixture);
        $assertEqual(font.italic, expected_value);
      } },
    { name: "it_can_change_its_italic_setting", args: ["italic_set_fixture"], fn: (self: any, italic_set_fixture: any): any => {
        let expected_xml: any, font: any, new_value: any;
        [font, new_value, expected_xml] = $list(italic_set_fixture);
        font.italic = new_value;
        $assertEqual(font._element.xml, expected_xml);
      } },
    { name: "it_knows_its_language_id", args: ["language_id_get_fixture"], fn: (self: any, language_id_get_fixture: any): any => {
        let expected_value: any, font: any;
        [font, expected_value] = $list(language_id_get_fixture);
        $assertEqual(font.languageId, expected_value);
      } },
    { name: "it_can_change_its_language_id_setting", args: ["language_id_set_fixture"], fn: (self: any, language_id_set_fixture: any): any => {
        let expected_xml: any, font: any, new_value: any;
        [font, new_value, expected_xml] = $list(language_id_set_fixture);
        font.languageId = new_value;
        $assertEqual(font._element.xml, expected_xml);
      } },
    { name: "it_knows_its_underline_setting", args: ["underline_get_fixture"], fn: (self: any, underline_get_fixture: any): any => {
        let expected_value: any, font: any;
        [font, expected_value] = $list(underline_get_fixture);
        $assert($is(font.underline, expected_value), $mod("got %s", font.underline));
      } },
    { name: "it_can_change_its_underline_setting", args: ["underline_set_fixture"], fn: (self: any, underline_set_fixture: any): any => {
        let expected_xml: any, font: any, new_value: any;
        [font, new_value, expected_xml] = $list(underline_set_fixture);
        font.underline = new_value;
        $assertEqual(font._element.xml, expected_xml);
      } },
    { name: "it_knows_its_size", args: ["size_get_fixture"], fn: (self: any, size_get_fixture: any): any => {
        let expected_value: any, font: any;
        [font, expected_value] = $list(size_get_fixture);
        $assertEqual(font.size, expected_value);
      } },
    { name: "it_can_change_its_size", args: ["size_set_fixture"], fn: (self: any, size_set_fixture: any): any => {
        let expected_xml: any, font: any, new_value: any;
        [font, new_value, expected_xml] = $list(size_set_fixture);
        font.size = new_value;
        $assertEqual(font._element.xml, expected_xml);
      } },
    { name: "it_knows_its_latin_typeface", args: ["name_get_fixture"], fn: (self: any, name_get_fixture: any): any => {
        let expected_value: any, font: any;
        [font, expected_value] = $list(name_get_fixture);
        $assertEqual(font.name, expected_value);
      } },
    { name: "it_can_change_its_latin_typeface", args: ["name_set_fixture"], fn: (self: any, name_set_fixture: any): any => {
        let expected_xml: any, font: any, new_value: any;
        [font, new_value, expected_xml] = $list(name_set_fixture);
        font.name = new_value;
        $assertEqual(font._element.xml, expected_xml);
      } },
    { name: "it_provides_access_to_its_color", args: ["font"], fn: (self: any, font: any): any => {
        $assert($isinstance(font.color, ColorFormat));
      } },
    { name: "it_provides_access_to_its_fill", args: ["font"], fn: (self: any, font: any): any => {
        $assert($isinstance(font.fill, FillFormat));
      } },
  ],
};
pytestClass("DescribeFont", DescribeFont);

const Describe_Hyperlink: any = {
  fixtures: {
    change_hlink_fixture_: { args: ["request", "hlink_with_hlinkClick", "rId", "rId_2", "part_", "url_2"], fn: (self: any, request: any, hlink_with_hlinkClick: any, rId: any, rId_2: any, part_: any, url_2: any): any => {
        let hlinkClick_bldr: any, new_rPr_xml: any;
        hlinkClick_bldr = anHlinkClick().withRId(rId_2);
        new_rPr_xml = anRPr().withNsdecls("a", "r").withChild(hlinkClick_bldr).xml();
        part_.relateTo.returnValue = rId_2;
        $call(propertyMock, request, _Hyperlink, "part", $kw({return_value: part_}));
        return [hlink_with_hlinkClick, rId, url_2, new_rPr_xml];
      } },
    hlink: { args: ["request", "part_"], fn: (self: any, request: any, part_: any): any => {
        let hlink: any, rPr: any;
        rPr = anRPr().withNsdecls("a", "r").element;
        hlink = $call(_Hyperlink, rPr, null);
        $call(propertyMock, request, _Hyperlink, "part", $kw({return_value: part_}));
        return hlink;
      } },
    hlink_with_hlinkClick: { args: ["request", "rPr_with_hlinkClick_bldr"], fn: (self: any, request: any, rPr_with_hlinkClick_bldr: any): any => {
        let rPr: any;
        rPr = rPr_with_hlinkClick_bldr.element;
        return $call(_Hyperlink, rPr, null);
      } },
    hlink_with_url_: { args: ["request", "part_", "hlink_with_hlinkClick", "rId", "url"], fn: (self: any, request: any, part_: any, hlink_with_hlinkClick: any, rId: any, url: any): any => {
        $call(propertyMock, request, _Hyperlink, "part", $kw({return_value: part_}));
        return [hlink_with_hlinkClick, rId, url];
      } },
    part_: { args: ["request", "url", "rId"], fn: (self: any, request: any, url: any, rId: any): any => {
        let part_: any;
        part_ = instanceMock(request, XmlPart);
        part_.targetRef.returnValue = url;
        part_.relateTo.returnValue = rId;
        return part_;
      } },
    rId: { args: [], fn: (self: any): any => {
        return "rId2";
      } },
    rId_2: { args: [], fn: (self: any): any => {
        return "rId6";
      } },
    remove_hlink_fixture_: { args: ["request", "hlink_with_hlinkClick", "rPr_xml", "rId"], fn: (self: any, request: any, hlink_with_hlinkClick: any, rPr_xml: any, rId: any): any => {
        propertyMock(request, _Hyperlink, "part");
        return [hlink_with_hlinkClick, rPr_xml, rId];
      } },
    rPr_with_hlinkClick_bldr: { args: ["rId"], fn: (self: any, rId: any): any => {
        let hlinkClick_bldr: any, rPr_bldr: any;
        hlinkClick_bldr = anHlinkClick().withRId(rId);
        rPr_bldr = anRPr().withNsdecls("a", "r").withChild(hlinkClick_bldr);
        return rPr_bldr;
      } },
    rPr_with_hlinkClick_xml: { args: ["rPr_with_hlinkClick_bldr"], fn: (self: any, rPr_with_hlinkClick_bldr: any): any => {
        return rPr_with_hlinkClick_bldr.xml();
      } },
    rPr_xml: { args: [], fn: (self: any): any => {
        return anRPr().withNsdecls("a", "r").xml();
      } },
    url: { args: [], fn: (self: any): any => {
        return "https://github.com/scanny/python-pptx";
      } },
    url_2: { args: [], fn: (self: any): any => {
        return "https://pypi.python.org/pypi/python-pptx";
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_the_target_url_of_the_hyperlink", args: ["hlink_with_url_"], fn: (self: any, hlink_with_url_: any): any => {
        let hlink: any, rId: any, url: any;
        [hlink, rId, url] = $list(hlink_with_url_);
        $assertEqual(hlink.address, url);
        hlink.part.targetRef.assertCalledOnceWith(rId);
      } },
    { name: "it_has_None_for_address_when_no_hyperlink_is_present", args: ["hlink"], fn: (self: any, hlink: any): any => {
        $assert($is(hlink.address, null));
      } },
    { name: "it_can_set_the_target_url", args: ["hlink", "rPr_with_hlinkClick_xml", "url"], fn: (self: any, hlink: any, rPr_with_hlinkClick_xml: any, url: any): any => {
        hlink.address = url;
        hlink.part.relateTo.assertCalledOnceWith(url, RT.HYPERLINK, $kw({is_external: true}));
        $assertEqual(hlink._rPr.xml, rPr_with_hlinkClick_xml);
        $assertEqual(hlink.address, url);
      } },
    { name: "it_can_remove_the_hyperlink", args: ["remove_hlink_fixture_"], fn: (self: any, remove_hlink_fixture_: any): any => {
        let hlink: any, rId: any, rPr_xml: any;
        [hlink, rPr_xml, rId] = $list(remove_hlink_fixture_);
        hlink.address = null;
        $assertEqual(hlink._rPr.xml, rPr_xml);
        hlink.part.dropRel.assertCalledOnceWith(rId);
      } },
    { name: "it_should_remove_the_hyperlink_when_url_set_to_empty_string", args: ["remove_hlink_fixture_"], fn: (self: any, remove_hlink_fixture_: any): any => {
        let hlink: any, rId: any, rPr_xml: any;
        [hlink, rPr_xml, rId] = $list(remove_hlink_fixture_);
        hlink.address = "";
        $assertEqual(hlink._rPr.xml, rPr_xml);
        hlink.part.dropRel.assertCalledOnceWith(rId);
      } },
    { name: "it_can_change_the_target_url", args: ["change_hlink_fixture_"], fn: (self: any, change_hlink_fixture_: any): any => {
        let hlink: any, new_rPr_xml: any, new_url: any, rId_existing: any;
        [hlink, rId_existing, new_url, new_rPr_xml] = $list(change_hlink_fixture_);
        hlink.address = new_url;
        $assertEqual(hlink._rPr.xml, new_rPr_xml);
        hlink.part.dropRel.assertCalledOnceWith(rId_existing);
        hlink.part.relateTo.assertCalledOnceWith(new_url, RT.HYPERLINK, $kw({is_external: true}));
      } },
  ],
};
pytestClass("Describe_Hyperlink", Describe_Hyperlink);

const Describe_Paragraph: any = {
  fixtures: {
    after_get_fixture: { args: ["request"], params: [["a:p", null], ["a:p/a:pPr", null], ["a:p/a:pPr/a:spcAft/a:spcPct{val=150000}", null], ["a:p/a:pPr/a:spcAft/a:spcPts{val=600}", 76200]], fn: (self: any, request: any): any => {
        let expected_value: any, p_cxml: any, paragraph: any;
        [p_cxml, expected_value] = $list(request.param);
        paragraph = $call(_Paragraph, element(p_cxml), null);
        return [paragraph, expected_value];
      } },
    after_set_fixture: { args: ["request"], params: [["a:p", $call(Pt, 8.333), "a:p/a:pPr/a:spcAft/a:spcPts{val=833}"], ["a:p/a:pPr/a:spcAft/a:spcPts{val=600}", $call(Pt, 42), "a:p/a:pPr/a:spcAft/a:spcPts{val=4200}"], ["a:p/a:pPr/a:spcAft/a:spcPct{val=150000}", $call(Pt, 24), "a:p/a:pPr/a:spcAft/a:spcPts{val=2400}"], ["a:p/a:pPr/a:spcAft/a:spcPts{val=600}", null, "a:p/a:pPr"]], fn: (self: any, request: any): any => {
        let expected_p_cxml: any, expected_xml: any, new_value: any, p_cxml: any, paragraph: any;
        [p_cxml, new_value, expected_p_cxml] = $list(request.param);
        paragraph = $call(_Paragraph, element(p_cxml), null);
        expected_xml = xml(expected_p_cxml);
        return [paragraph, new_value, expected_xml];
      } },
    alignment_get_fixture: { args: ["request"], params: [["a:p", null], ["a:p/a:pPr{algn=ctr}", PP_ALIGN.CENTER], ["a:p/a:pPr{algn=r}", PP_ALIGN.RIGHT]], fn: (self: any, request: any): any => {
        let expected_value: any, p_cxml: any, paragraph: any;
        [p_cxml, expected_value] = $list(request.param);
        paragraph = $call(_Paragraph, element(p_cxml), null);
        return [paragraph, expected_value];
      } },
    alignment_set_fixture: { args: ["request"], params: [["a:p", PP_ALIGN.LEFT, "a:p/a:pPr{algn=l}"], ["a:p/a:pPr{algn=l}", PP_ALIGN.JUSTIFY, "a:p/a:pPr{algn=just}"], ["a:p/a:pPr{algn=just}", null, "a:p/a:pPr"]], fn: (self: any, request: any): any => {
        let expected_p_cxml: any, expected_xml: any, new_value: any, p_cxml: any, paragraph: any;
        [p_cxml, new_value, expected_p_cxml] = $list(request.param);
        paragraph = $call(_Paragraph, element(p_cxml), null);
        expected_xml = xml(expected_p_cxml);
        return [paragraph, new_value, expected_xml];
      } },
    before_get_fixture: { args: ["request"], params: [["a:p", null], ["a:p/a:pPr", null], ["a:p/a:pPr/a:spcBef/a:spcPct{val=150000}", null], ["a:p/a:pPr/a:spcBef/a:spcPts{val=600}", 76200]], fn: (self: any, request: any): any => {
        let expected_value: any, p_cxml: any, paragraph: any;
        [p_cxml, expected_value] = $list(request.param);
        paragraph = $call(_Paragraph, element(p_cxml), null);
        return [paragraph, expected_value];
      } },
    before_set_fixture: { args: ["request"], params: [["a:p", $call(Pt, 8.333), "a:p/a:pPr/a:spcBef/a:spcPts{val=833}"], ["a:p/a:pPr/a:spcBef/a:spcPts{val=600}", $call(Pt, 42), "a:p/a:pPr/a:spcBef/a:spcPts{val=4200}"], ["a:p/a:pPr/a:spcBef/a:spcPct{val=150000}", $call(Pt, 24), "a:p/a:pPr/a:spcBef/a:spcPts{val=2400}"], ["a:p/a:pPr/a:spcBef/a:spcPts{val=600}", null, "a:p/a:pPr"]], fn: (self: any, request: any): any => {
        let expected_p_cxml: any, expected_xml: any, new_value: any, p_cxml: any, paragraph: any;
        [p_cxml, new_value, expected_p_cxml] = $list(request.param);
        paragraph = $call(_Paragraph, element(p_cxml), null);
        expected_xml = xml(expected_p_cxml);
        return [paragraph, new_value, expected_xml];
      } },
    clear_fixture: { args: ["request"], params: [["a:p/a:r/a:t\"foo\"", "a:p"], ["a:p/(a:br,a:r/a:t\"foo\")", "a:p"], ["a:p/(a:fld,a:br,a:r/a:t\"foo\")", "a:p"]], fn: (self: any, request: any): any => {
        let expected_p_cxml: any, expected_xml: any, p_cxml: any, paragraph: any;
        [p_cxml, expected_p_cxml] = $list(request.param);
        paragraph = $call(_Paragraph, element(p_cxml), null);
        expected_xml = xml(expected_p_cxml);
        return [paragraph, expected_xml];
      } },
    level_get_fixture: { args: ["request"], params: [["a:p", 0], ["a:p/a:pPr{lvl=2}", 2]], fn: (self: any, request: any): any => {
        let expected_value: any, p_cxml: any, paragraph: any;
        [p_cxml, expected_value] = $list(request.param);
        paragraph = $call(_Paragraph, element(p_cxml), null);
        return [paragraph, expected_value];
      } },
    level_set_fixture: { args: ["request"], params: [["a:p", 1, "a:p/a:pPr{lvl=1}"], ["a:p/a:pPr{lvl=1}", 2, "a:p/a:pPr{lvl=2}"], ["a:p/a:pPr{lvl=2}", 0, "a:p/a:pPr"]], fn: (self: any, request: any): any => {
        let expected_p_cxml: any, expected_xml: any, new_value: any, p_cxml: any, paragraph: any;
        [p_cxml, new_value, expected_p_cxml] = $list(request.param);
        paragraph = $call(_Paragraph, element(p_cxml), null);
        expected_xml = xml(expected_p_cxml);
        return [paragraph, new_value, expected_xml];
      } },
    line_break_fixture: { args: ["request"], params: [["a:p", "a:p/a:br"], ["a:p/a:r", "a:p/(a:r,a:br)"], ["a:p/a:br", "a:p/(a:br,a:br)"]], fn: (self: any, request: any): any => {
        let cxml: any, expected_cxml: any, expected_xml: any, paragraph: any;
        [cxml, expected_cxml] = $list(request.param);
        paragraph = $call(_Paragraph, element(cxml), null);
        expected_xml = xml(expected_cxml);
        return [paragraph, expected_xml];
      } },
    runs_fixture: { args: [], fn: (self: any): any => {
        let expected_text: any, p_cxml: any, paragraph: any;
        p_cxml = "a:p/(a:r/a:t\"Foo\",a:r/a:t\"Bar\",a:r/a:t\"Baz\")";
        paragraph = $call(_Paragraph, element(p_cxml), null);
        expected_text = ["Foo", "Bar", "Baz"];
        return [paragraph, expected_text];
      } },
    spacing_get_fixture: { args: ["request"], params: [["a:p", null], ["a:p/a:pPr", null], ["a:p/a:pPr/a:lnSpc/a:spcPts{val=1800}", 228600], ["a:p/a:pPr/a:lnSpc/a:spcPct{val=142000}", 1.42], ["a:p/a:pPr/a:lnSpc/a:spcPct{val=124.64%}", 1.2464]], fn: (self: any, request: any): any => {
        let expected_value: any, p_cxml: any, paragraph: any;
        [p_cxml, expected_value] = $list(request.param);
        paragraph = $call(_Paragraph, element(p_cxml), null);
        return [paragraph, expected_value];
      } },
    spacing_set_fixture: { args: ["request"], params: [["a:p", 1.42, "a:p/a:pPr/a:lnSpc/a:spcPct{val=142000}"], ["a:p", $call(Pt, 42), "a:p/a:pPr/a:lnSpc/a:spcPts{val=4200}"], ["a:p/a:pPr/a:lnSpc/a:spcPct{val=110000}", 0.875, "a:p/a:pPr/a:lnSpc/a:spcPct{val=87500}"], ["a:p/a:pPr/a:lnSpc/a:spcPts{val=600}", $call(Pt, 42), "a:p/a:pPr/a:lnSpc/a:spcPts{val=4200}"], ["a:p/a:pPr/a:lnSpc/a:spcPts{val=1900}", 0.925, "a:p/a:pPr/a:lnSpc/a:spcPct{val=92500}"], ["a:p/a:pPr/a:lnSpc/a:spcPct{val=150000}", $call(Pt, 24), "a:p/a:pPr/a:lnSpc/a:spcPts{val=2400}"], ["a:p/a:pPr/a:lnSpc/a:spcPts{val=600}", null, "a:p/a:pPr"], ["a:p/a:pPr/a:lnSpc/a:spcPct{val=150000}", null, "a:p/a:pPr"]], fn: (self: any, request: any): any => {
        let expected_p_cxml: any, expected_xml: any, new_value: any, p_cxml: any, paragraph: any;
        [p_cxml, new_value, expected_p_cxml] = $list(request.param);
        paragraph = $call(_Paragraph, element(p_cxml), null);
        expected_xml = xml(expected_p_cxml);
        return [paragraph, new_value, expected_xml];
      } },
    text_get_fixture: { args: ["request"], params: [["a:p/a:r/a:t\"foobar\"", "foobar"], ["a:p/(a:r/a:t\"foo\",a:r/a:t\"bar\")", "foobar"], ["a:p/(a:r/a:t\"foo\",a:br,a:r/a:t\"bar\")", "foo\u000bbar"], ["a:p/(a:r/a:t\"foo \",a:fld/a:t\"42\",a:r/a:t\" bar\")", "foo 42 bar"], ["a:p/(a:r/a:t\" foo\",a:br,a:fld/a:t\"42\")", " foo\u000b42"], ["a:p/(a:pPr,a:r/a:t\"foobar\",a:endParaRPr)", "foobar"], ["a:p/a:fld/a:t\"42\"", "42"], ["a:p/a:br", "\u000b"]], fn: (self: any, request: any): any => {
        let expected_value: any, p: any, p_cxml: any;
        [p_cxml, expected_value] = $list(request.param);
        p = element(p_cxml);
        return [p, expected_value];
      } },
    Font_: { args: ["request"], fn: (self: any, request: any): any => {
        return classMock(request, "pptx.text.text.Font");
      } },
    p_bldr: { args: [], fn: (self: any): any => {
        return aP().withNsdecls();
      } },
    p_with_r_xml: { args: [], fn: (self: any): any => {
        let run_bldr: any;
        run_bldr = anR().withChild(aT());
        return aP().withNsdecls().withChild(run_bldr).xml();
      } },
    paragraph: { args: ["p_bldr"], fn: (self: any, p_bldr: any): any => {
        return $call(_Paragraph, p_bldr.element, null);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_add_a_line_break", args: ["line_break_fixture"], fn: (self: any, line_break_fixture: any): any => {
        let expected_xml: any, paragraph: any;
        [paragraph, expected_xml] = $list(line_break_fixture);
        paragraph.addLineBreak();
        $assertEqual(paragraph._p.xml, expected_xml);
      } },
    { name: "it_can_add_a_run", args: ["paragraph", "p_with_r_xml"], fn: (self: any, paragraph: any, p_with_r_xml: any): any => {
        let run: any;
        run = paragraph.addRun();
        $assertEqual(paragraph._p.xml, p_with_r_xml);
        $assert($isinstance(run, _Run));
      } },
    { name: "it_knows_its_horizontal_alignment", args: ["alignment_get_fixture"], fn: (self: any, alignment_get_fixture: any): any => {
        let expected_value: any, paragraph: any;
        [paragraph, expected_value] = $list(alignment_get_fixture);
        $assertEqual(paragraph.alignment, expected_value);
      } },
    { name: "it_can_change_its_horizontal_alignment", args: ["alignment_set_fixture"], fn: (self: any, alignment_set_fixture: any): any => {
        let expected_xml: any, new_value: any, paragraph: any;
        [paragraph, new_value, expected_xml] = $list(alignment_set_fixture);
        paragraph.alignment = new_value;
        $assertEqual(paragraph._element.xml, expected_xml);
      } },
    { name: "it_can_clear_itself_of_content", args: ["clear_fixture"], fn: (self: any, clear_fixture: any): any => {
        let expected_xml: any, paragraph: any;
        [paragraph, expected_xml] = $list(clear_fixture);
        $clear(paragraph);
        $assertEqual(paragraph._element.xml, expected_xml);
      } },
    { name: "it_provides_access_to_the_default_paragraph_font", args: ["paragraph", "Font_"], fn: (self: any, paragraph: any, Font_: any): any => {
        let font: any;
        font = paragraph.font;
        Font_.assertCalledOnceWith(paragraph._defRPr);
        $assertEqual(font, Font_.returnValue);
      } },
    { name: "it_knows_its_indentation_level", args: ["level_get_fixture"], fn: (self: any, level_get_fixture: any): any => {
        let expected_value: any, paragraph: any;
        [paragraph, expected_value] = $list(level_get_fixture);
        $assertEqual(paragraph.level, expected_value);
      } },
    { name: "it_can_change_its_indentation_level", args: ["level_set_fixture"], fn: (self: any, level_set_fixture: any): any => {
        let expected_xml: any, new_value: any, paragraph: any;
        [paragraph, new_value, expected_xml] = $list(level_set_fixture);
        paragraph.level = new_value;
        $assertEqual(paragraph._element.xml, expected_xml);
      } },
    { name: "it_knows_its_line_spacing", args: ["spacing_get_fixture"], fn: (self: any, spacing_get_fixture: any): any => {
        let expected_value: any, paragraph: any;
        [paragraph, expected_value] = $list(spacing_get_fixture);
        $assertEqual(paragraph.lineSpacing, expected_value);
      } },
    { name: "it_can_change_its_line_spacing", args: ["spacing_set_fixture"], fn: (self: any, spacing_set_fixture: any): any => {
        let expected_xml: any, new_value: any, paragraph: any;
        [paragraph, new_value, expected_xml] = $list(spacing_set_fixture);
        paragraph.lineSpacing = new_value;
        $assertEqual(paragraph._element.xml, expected_xml);
      } },
    { name: "it_provides_access_to_its_runs", args: ["runs_fixture"], fn: (self: any, runs_fixture: any): any => {
        let expected_text: any, paragraph: any, runs: any;
        [paragraph, expected_text] = $list(runs_fixture);
        runs = paragraph.runs;
        $assertEqual($list($list(runs).map((r: any) => r.text)), expected_text);
        for (const r of $list(runs)) {
          $assert($isinstance(r, _Run));
          $assertEqual(r._parent, paragraph);
        }
      } },
    { name: "it_knows_its_space_after", args: ["after_get_fixture"], fn: (self: any, after_get_fixture: any): any => {
        let expected_value: any, paragraph: any;
        [paragraph, expected_value] = $list(after_get_fixture);
        $assertEqual(paragraph.spaceAfter, expected_value);
      } },
    { name: "it_can_change_its_space_after", args: ["after_set_fixture"], fn: (self: any, after_set_fixture: any): any => {
        let expected_xml: any, new_value: any, paragraph: any;
        [paragraph, new_value, expected_xml] = $list(after_set_fixture);
        paragraph.spaceAfter = new_value;
        $assertEqual(paragraph._element.xml, expected_xml);
      } },
    { name: "it_knows_its_space_before", args: ["before_get_fixture"], fn: (self: any, before_get_fixture: any): any => {
        let expected_value: any, paragraph: any;
        [paragraph, expected_value] = $list(before_get_fixture);
        $assertEqual(paragraph.spaceBefore, expected_value);
      } },
    { name: "it_can_change_its_space_before", args: ["before_set_fixture"], fn: (self: any, before_set_fixture: any): any => {
        let expected_xml: any, new_value: any, paragraph: any;
        [paragraph, new_value, expected_xml] = $list(before_set_fixture);
        paragraph.spaceBefore = new_value;
        $assertEqual(paragraph._element.xml, expected_xml);
      } },
    { name: "it_knows_what_text_it_contains", args: ["text_get_fixture"], fn: (self: any, text_get_fixture: any): any => {
        let expected_value: any, p: any, paragraph: any, text: any;
        [p, expected_value] = $list(text_get_fixture);
        paragraph = $call(_Paragraph, p, null);
        text = paragraph.text;
        $assertEqual(text, expected_value);
        $assert($isinstance(text, String));
      } },
    { name: "it_can_change_its_text", args: ["p_cxml", "value", "expected_cxml"], parametrize: [{ names: ["p_cxml", "value", "expected_cxml"], values: [["a:p/(a:r/a:t\"foo\",a:r/a:t\"bar\")", "foobar", "a:p/a:r/a:t\"foobar\""], ["a:p", "", "a:p"], ["a:p", "foobar", "a:p/a:r/a:t\"foobar\""], ["a:p", "foo\nbar", "a:p/(a:r/a:t\"foo\",a:br,a:r/a:t\"bar\")"], ["a:p", "\u000bfoo\n", "a:p/(a:br,a:r/a:t\"foo\",a:br)"], ["a:p", "\n\nfoo", "a:p/(a:br,a:br,a:r/a:t\"foo\")"], ["a:p", "foo\n", "a:p/(a:r/a:t\"foo\",a:br)"], ["a:p", "foo\u0007\n", "a:p/(a:r/a:t\"foo_x0007_\",a:br)"], ["a:p", "ŮŦƑ-8\u001bliteral", "a:p/a:r/a:t\"ŮŦƑ-8_x001B_literal\""], ["a:p", "utf-8 unicode: Hér er texti", "a:p/a:r/a:t\"utf-8 unicode: Hér er texti\""]] }], fn: (self: any, p_cxml: any, value: any, expected_cxml: any): any => {
        let p: any, paragraph: any;
        p = $cast("CT_TextParagraph", element(p_cxml));
        paragraph = $call(_Paragraph, p, null);
        paragraph.text = value;
        $assertEqual(paragraph._element.xml, xml(expected_cxml));
      } },
  ],
};
pytestClass("Describe_Paragraph", Describe_Paragraph);

const Describe_Run: any = {
  fixtures: {
    font_fixture: { args: ["Font_", "font_"], fn: (self: any, Font_: any, font_: any): any => {
        let r: any, rPr: any, run: any;
        r = element("a:r/a:rPr");
        rPr = r.rPr;
        run = $call(_Run, r, null);
        return [run, rPr, Font_, font_];
      } },
    hyperlink_fixture: { args: ["_Hyperlink_", "hlink_"], fn: (self: any, _Hyperlink_: any, hlink_: any): any => {
        let r: any, rPr: any, run: any;
        r = element("a:r/a:rPr");
        rPr = r.rPr;
        run = $call(_Run, r, null);
        return [run, rPr, _Hyperlink_, hlink_];
      } },
    text_get_fixture: { args: [], fn: (self: any): any => {
        let r: any, run: any;
        r = element("a:r/a:t\"foobar\"");
        run = $call(_Run, r, null);
        return [run, "foobar"];
      } },
    Font_: { args: ["request", "font_"], fn: (self: any, request: any, font_: any): any => {
        return $call(classMock, request, "pptx.text.text.Font", $kw({return_value: font_}));
      } },
    font_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Font);
      } },
    _Hyperlink_: { args: ["request", "hlink_"], fn: (self: any, request: any, hlink_: any): any => {
        return $call(classMock, request, "pptx.text.text._Hyperlink", $kw({return_value: hlink_}));
      } },
    hlink_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _Hyperlink);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_access_to_its_font", args: ["font_fixture"], fn: (self: any, font_fixture: any): any => {
        let Font_: any, font: any, font_: any, rPr: any, run: any;
        [run, rPr, Font_, font_] = $list(font_fixture);
        font = run.font;
        Font_.assertCalledOnceWith(rPr);
        $assertEqual(font, font_);
      } },
    { name: "it_provides_access_to_a_hyperlink_proxy", args: ["hyperlink_fixture"], fn: (self: any, hyperlink_fixture: any): any => {
        let _Hyperlink_: any, hlink: any, hlink_: any, rPr: any, run: any;
        [run, rPr, _Hyperlink_, hlink_] = $list(hyperlink_fixture);
        hlink = run.hyperlink;
        _Hyperlink_.assertCalledOnceWith(rPr, run);
        $assert($is(hlink, hlink_));
      } },
    { name: "it_can_get_the_text_of_the_run", args: ["text_get_fixture"], fn: (self: any, text_get_fixture: any): any => {
        let expected_value: any, run: any, text: any;
        [run, expected_value] = $list(text_get_fixture);
        text = run.text;
        $assertEqual(text, expected_value);
        $assert($isinstance(text, String));
      } },
    { name: "it_can_change_its_text", args: ["r_cxml", "new_value", "expected_r_cxml"], parametrize: [{ names: ["r_cxml", "new_value", "expected_r_cxml"], values: [["a:r/a:t", "barfoo", "a:r/a:t\"barfoo\""], ["a:r/a:t", "bar\u001bfoo", "a:r/a:t\"bar_x001B_foo\""], ["a:r/a:t", "bar\tfoo", "a:r/a:t\"bar\tfoo\""]] }], fn: (self: any, r_cxml: any, new_value: any, expected_r_cxml: any): any => {
        let run: any;
        run = $call(_Run, element(r_cxml), null);
        run.text = new_value;
        $assertEqual(run._r.xml, xml(expected_r_cxml));
      } },
  ],
};
pytestClass("Describe_Run", Describe_Run);

