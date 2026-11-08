// Machine-translated from python-pptx tests/oxml/shapes/test_autoshape.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $assert, $assertEqual, $in, $is, $list, $mod, $mul, $slice, $truthy, $zip, pytestClass } from "../../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../../src/exc.js";
import { defineChildren } from "../../../src/oxml/xmlchemy.js";
import { PP_PLACEHOLDER as PP_PLACEHOLDER$imported } from "../../../src/enum/shapes.js";
import { nsdecls as nsdecls$imported } from "../../../src/oxml/ns.js";
import { CT_Shape as CT_Shape$imported } from "../../../src/oxml/shapes/autoshape.js";
import { ST_Direction as ST_Direction$imported, ST_PlaceholderSize as ST_PlaceholderSize$imported } from "../../../src/oxml/shapes/shared.js";
import { aGd, aPrstGeom, anAvLst } from "../../oxml/unitdata/shape.js";
import { element } from "../../unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const PP_PLACEHOLDER = PP_PLACEHOLDER$imported;
const nsdecls = nsdecls$imported;
const CT_Shape = CT_Shape$imported;
const ST_Direction = ST_Direction$imported;
const ST_PlaceholderSize = ST_PlaceholderSize$imported;
void $py;


const DescribeCT_PresetGeometry2D: any = {
  fixtures: {
    gd_lst_fixture: { args: ["request"], params: [[], [["adj1", "val 111"]], [["adj2", "val 222"], ["adj3", "val 333"], ["adj4", "val 444"], ["adj5", "val 555"]]], fn: (self: any, request: any): any => {
        let expected_vals: any, prstGeom: any;
        expected_vals = request.param;
        prstGeom = self.prstGeomBldr("foobar", expected_vals).element;
        return [prstGeom, expected_vals];
      } },
    rewrite_guides_fixture_: { args: ["request"], params: [["circularArrow", 5], ["chevron", 0], ["chevron", 1]], fn: (self: any, request: any): any => {
        let after_vals: any, before_vals: any, expected_xml: any, fmlas: any, gd_count: any, guides: any, names: any, prst: any, prstGeom: any, vals: any;
        [prst, gd_count] = $list(request.param);
        names = ["adj1", "adj2", "adj3", "adj4", "adj5"];
        vals = [111, 222, 333, 444, 555];
        fmlas = $list(vals).map((v: any) => $mod("val %d", v));
        before_vals = $mul([["adj6", "val 666"]], 3);
        after_vals = $zip($slice(names, null, gd_count, null), $slice(fmlas, null, gd_count, null));
        prstGeom = self.prstGeomBldr(prst, before_vals).element;
        guides = $zip($slice(names, null, gd_count, null), $slice(vals, null, gd_count, null));
        expected_xml = self.prstGeomBldr(prst, after_vals).xml();
        return [prstGeom, guides, expected_xml];
      } },
  },
  helpers: {
    prstGeomBldr: (self: any, prst: any, gd_vals: any): any => {
        let avLst_bldr: any, gd_bldr: any, prstGeom_bldr: any;
        avLst_bldr = anAvLst();
        for (const [name, fmla] of $list(gd_vals)) {
          gd_bldr = aGd().withName(name).withFmla(fmla);
          avLst_bldr.withChild(gd_bldr);
        }
        prstGeom_bldr = aPrstGeom().withNsdecls().withPrst(prst).withChild(avLst_bldr);
        return prstGeom_bldr;
      },
  },
  tests: [
    { name: "it_can_get_the_gd_elms_as_a_sequence", args: ["gd_lst_fixture"], fn: (self: any, gd_lst_fixture: any): any => {
        let actual_vals: any, expected_vals: any, prstGeom: any;
        [prstGeom, expected_vals] = $list(gd_lst_fixture);
        actual_vals = $list(prstGeom.gdLst).map((gd: any) => [gd.name, gd.fmla]);
        $assertEqual(actual_vals, expected_vals);
      } },
    { name: "it_can_rewrite_the_gd_elms", args: ["rewrite_guides_fixture_"], fn: (self: any, rewrite_guides_fixture_: any): any => {
        let expected_xml: any, guides: any, prstGeom: any;
        [prstGeom, guides, expected_xml] = $list(rewrite_guides_fixture_);
        console.log(prstGeom.xml);
        prstGeom.rewriteGuides(guides);
        $assertEqual(prstGeom.xml, expected_xml);
      } },
  ],
};
pytestClass("DescribeCT_PresetGeometry2D", DescribeCT_PresetGeometry2D);

const DescribeCT_Shape: any = {
  fixtures: {
    is_autoshape_fixture: { args: ["request"], params: [["p:sp/(p:nvSpPr/p:cNvSpPr, p:spPr/a:prstGeom)", true], ["p:sp/(p:nvSpPr/p:nvPr/p:ph, p:spPr)", false], ["p:sp/(p:nvSpPr/p:cNvSpPr{txBox=1}, p:spPr/a:prstGeom)", false]], fn: (self: any, request: any): any => {
        let expected_value: any, sp: any, sp_cxml: any;
        [sp_cxml, expected_value] = $list(request.param);
        sp = element(sp_cxml);
        return [sp, expected_value];
      } },
    is_textbox_fixture: { args: ["request"], params: [["p:sp/(p:nvSpPr/p:cNvSpPr, p:spPr/a:prstGeom)", false], ["p:sp/(p:nvSpPr/(p:nvPr/p:ph, p:cNvSpPr), p:spPr)", false], ["p:sp/(p:nvSpPr/p:cNvSpPr{txBox=1}, p:spPr/a:prstGeom)", true]], fn: (self: any, request: any): any => {
        let expected_value: any, sp: any, sp_cxml: any;
        [sp_cxml, expected_value] = $list(request.param);
        sp = element(sp_cxml);
        return [sp, expected_value];
      } },
    new_ph_sp_fixture: { args: ["request"], params: [[2, "Title 1", PP_PLACEHOLDER.CENTER_TITLE, ST_Direction.HORZ, ST_PlaceholderSize.FULL, 0, " type=\"ctrTitle\""], [3, "Date Placeholder 2", PP_PLACEHOLDER.DATE, ST_Direction.HORZ, ST_PlaceholderSize.HALF, 10, " type=\"dt\" sz=\"half\" idx=\"10\""], [4, "Vertical Subtitle 3", PP_PLACEHOLDER.SUBTITLE, ST_Direction.VERT, ST_PlaceholderSize.FULL, 1, " type=\"subTitle\" orient=\"vert\" idx=\"1\""], [5, "Table Placeholder 4", PP_PLACEHOLDER.TABLE, ST_Direction.HORZ, ST_PlaceholderSize.QUARTER, 14, " type=\"tbl\" sz=\"quarter\" idx=\"14\""], [6, "Slide Number Placeholder 5", PP_PLACEHOLDER.SLIDE_NUMBER, ST_Direction.HORZ, ST_PlaceholderSize.QUARTER, 12, " type=\"sldNum\" sz=\"quarter\" idx=\"12\""], [7, "Footer Placeholder 6", PP_PLACEHOLDER.FOOTER, ST_Direction.HORZ, ST_PlaceholderSize.QUARTER, 11, " type=\"ftr\" sz=\"quarter\" idx=\"11\""], [8, "Content Placeholder 7", PP_PLACEHOLDER.OBJECT, ST_Direction.HORZ, ST_PlaceholderSize.FULL, 15, " idx=\"15\""]], fn: (self: any, request: any): any => {
        let expected_attrs: any, expected_values: any, expected_xml: any, expected_xml_tmpl: any, id_: any, idx: any, name: any, orient: any, ph_type: any, sz: any, txBody_snippet: any, txBody_str: any;
        [id_, name, ph_type, orient, sz, idx, expected_attrs] = $list(request.param);
        expected_xml_tmpl = $mod("<p:sp %s>\n  <p:nvSpPr>\n    <p:cNvPr id=\"%s\" name=\"%s\"/>\n    <p:cNvSpPr>\n      <a:spLocks noGrp=\"1\"/>\n    </p:cNvSpPr>\n    <p:nvPr>\n      <p:ph%s/>\n    </p:nvPr>\n  </p:nvSpPr>\n  <p:spPr/>\n%s</p:sp>\n", [nsdecls("a", "p"), "%d", "%s", "%s", "%s"]);
        txBody_snippet = "  <p:txBody>\n    <a:bodyPr/>\n    <a:lstStyle/>\n    <a:p/>\n  </p:txBody>\n";
        txBody_str = ($truthy($in(id_, [2, 4, 8])) ? txBody_snippet : "");
        expected_values = [id_, name, expected_attrs, txBody_str];
        expected_xml = $mod(expected_xml_tmpl, expected_values);
        return [id_, name, ph_type, orient, sz, idx, expected_xml];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_how_to_create_a_new_autoshape_sp", args: [], fn: (self: any): any => {
        let height: any, id_: any, left: any, name: any, prst: any, sp: any, top: any, width: any, xml: any;
        id_ = 9;
        name = "Rounded Rectangle 8";
        prst = "roundRect";
        [left, top, width, height] = $list([111, 222, 333, 444]);
        xml = $mod("<p:sp %s>\n  <p:nvSpPr>\n    <p:cNvPr id=\"%d\" name=\"%s\"/>\n    <p:cNvSpPr/>\n    <p:nvPr/>\n  </p:nvSpPr>\n  <p:spPr>\n    <a:xfrm>\n      <a:off x=\"%d\" y=\"%d\"/>\n      <a:ext cx=\"%d\" cy=\"%d\"/>\n    </a:xfrm>\n    <a:prstGeom prst=\"%s\">\n      <a:avLst/>\n    </a:prstGeom>\n  </p:spPr>\n  <p:style>\n    <a:lnRef idx=\"1\">\n      <a:schemeClr val=\"accent1\"/>\n    </a:lnRef>\n    <a:fillRef idx=\"3\">\n      <a:schemeClr val=\"accent1\"/>\n    </a:fillRef>\n    <a:effectRef idx=\"2\">\n      <a:schemeClr val=\"accent1\"/>\n    </a:effectRef>\n    <a:fontRef idx=\"minor\">\n      <a:schemeClr val=\"lt1\"/>\n    </a:fontRef>\n  </p:style>\n  <p:txBody>\n    <a:bodyPr rtlCol=\"0\" anchor=\"ctr\"/>\n    <a:lstStyle/>\n    <a:p>\n      <a:pPr algn=\"ctr\"/>\n    </a:p>\n  </p:txBody>\n</p:sp>\n", [nsdecls("a", "p"), id_, name, left, top, width, height, prst]);
        sp = CT_Shape.newAutoshapeSp(id_, name, prst, left, top, width, height);
        $assertEqual(sp.xml, xml);
      } },
    { name: "it_knows_how_to_create_a_new_placeholder_sp", args: ["new_ph_sp_fixture"], fn: (self: any, new_ph_sp_fixture: any): any => {
        let expected_xml: any, id_: any, idx: any, name: any, orient: any, ph_type: any, sp: any, sz: any;
        [id_, name, ph_type, orient, sz, idx, expected_xml] = $list(new_ph_sp_fixture);
        sp = CT_Shape.newPlaceholderSp(id_, name, ph_type, orient, sz, idx);
        $assertEqual(sp.xml, expected_xml);
      } },
    { name: "it_knows_how_to_create_a_new_textbox_sp", args: [], fn: (self: any): any => {
        let height: any, id_: any, left: any, name: any, sp: any, top: any, width: any, xml: any;
        id_ = 9;
        name = "TextBox 8";
        [left, top, width, height] = $list([111, 222, 333, 444]);
        xml = $mod("<p:sp %s>\n  <p:nvSpPr>\n    <p:cNvPr id=\"%d\" name=\"%s\"/>\n    <p:cNvSpPr txBox=\"1\"/>\n    <p:nvPr/>\n  </p:nvSpPr>\n  <p:spPr>\n    <a:xfrm>\n      <a:off x=\"%d\" y=\"%d\"/>\n      <a:ext cx=\"%d\" cy=\"%d\"/>\n    </a:xfrm>\n    <a:prstGeom prst=\"rect\">\n      <a:avLst/>\n    </a:prstGeom>\n    <a:noFill/>\n  </p:spPr>\n  <p:txBody>\n    <a:bodyPr wrap=\"none\">\n      <a:spAutoFit/>\n    </a:bodyPr>\n    <a:lstStyle/>\n    <a:p/>\n  </p:txBody>\n</p:sp>\n", [nsdecls("a", "p"), id_, name, left, top, width, height]);
        sp = CT_Shape.newTextboxSp(id_, name, left, top, width, height);
        $assertEqual(sp.xml, xml);
      } },
    { name: "it_knows_whether_it_is_an_autoshape", args: ["is_autoshape_fixture"], fn: (self: any, is_autoshape_fixture: any): any => {
        let expected_value: any, sp: any;
        [sp, expected_value] = $list(is_autoshape_fixture);
        $assert($is(sp.isAutoshape, expected_value));
      } },
    { name: "it_knows_whether_it_is_an_textbox", args: ["is_textbox_fixture"], fn: (self: any, is_textbox_fixture: any): any => {
        let expected_value: any, sp: any;
        [sp, expected_value] = $list(is_textbox_fixture);
        $assert($is(sp.isTextbox, expected_value));
      } },
  ],
};
pytestClass("DescribeCT_Shape", DescribeCT_Shape);

