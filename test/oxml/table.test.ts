// Machine-translated from python-pptx tests/oxml/test_table.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $call, $get, $is, $list, $mod, $truthy, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { nsdecls as nsdecls$imported } from "../../src/oxml/ns.js";
import { CT_Table as CT_Table$imported, TcRange as TcRange$imported } from "../../src/oxml/table.js";
import { element } from "../unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const nsdecls = nsdecls$imported;
const CT_Table = CT_Table$imported;
const TcRange = TcRange$imported;
void $py;


const DescribeCT_Table: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_can_create_a_new_tbl_element_tree", args: [], fn: (self: any): any => {
        let expected_xml: any, tbl: any;
        expected_xml = $mod("<a:tbl %s>\n  <a:tblPr firstRow=\"1\" bandRow=\"1\">\n    <a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>\n  </a:tblPr>\n  <a:tblGrid>\n    <a:gridCol w=\"111\"/>\n    <a:gridCol w=\"111\"/>\n    <a:gridCol w=\"112\"/>\n  </a:tblGrid>\n  <a:tr h=\"222\">\n    <a:tc>\n      <a:txBody>\n        <a:bodyPr/>\n        <a:lstStyle/>\n        <a:p/>\n      </a:txBody>\n      <a:tcPr/>\n    </a:tc>\n    <a:tc>\n      <a:txBody>\n        <a:bodyPr/>\n        <a:lstStyle/>\n        <a:p/>\n      </a:txBody>\n      <a:tcPr/>\n    </a:tc>\n    <a:tc>\n      <a:txBody>\n        <a:bodyPr/>\n        <a:lstStyle/>\n        <a:p/>\n      </a:txBody>\n      <a:tcPr/>\n    </a:tc>\n  </a:tr>\n  <a:tr h=\"223\">\n    <a:tc>\n      <a:txBody>\n        <a:bodyPr/>\n        <a:lstStyle/>\n        <a:p/>\n      </a:txBody>\n      <a:tcPr/>\n    </a:tc>\n    <a:tc>\n      <a:txBody>\n        <a:bodyPr/>\n        <a:lstStyle/>\n        <a:p/>\n      </a:txBody>\n      <a:tcPr/>\n    </a:tc>\n    <a:tc>\n      <a:txBody>\n        <a:bodyPr/>\n        <a:lstStyle/>\n        <a:p/>\n      </a:txBody>\n      <a:tcPr/>\n    </a:tc>\n  </a:tr>\n</a:tbl>\n", nsdecls("a"));
        tbl = CT_Table.newTbl(2, 3, 334, 445);
        $assertEqual(tbl.xml, expected_xml);
      } },
    { name: "it_provides_access_to_its_tc_elements", args: [], fn: (self: any): any => {
        let tbl: any, tbl_cxml: any, tcs: any;
        tbl_cxml = "a:tbl/(a:tr/(a:tc,a:tc),a:tr/(a:tc,a:tc))";
        tbl = element(tbl_cxml);
        tcs = tbl.xpath(".//a:tc");
        $assert($is(tbl.tc(0, 0), $get(tcs, 0)));
        $assert($is(tbl.tc(0, 1), $get(tcs, 1)));
        $assert($is(tbl.tc(1, 0), $get(tcs, 2)));
        $assert($is(tbl.tc(1, 1), $get(tcs, 3)));
      } },
  ],
};
pytestClass("DescribeCT_Table", DescribeCT_Table);

const DescribeTcRange: any = {
  fixtures: {
    contains_merge_fixture: { args: ["request"], params: [["a:tbl/a:tr/(a:tc,a:tc)", false], ["a:tbl/a:tr/(a:tc{gridSpan=1},a:tc{hMerge=false})", false], ["a:tbl/a:tr/(a:tc{gridSpan=2},a:tc{hMerge=1})", true], ["a:tbl/(a:tr/a:tc,a:tr/a:tc)", false], ["a:tbl/(a:tr/a:tc{rowSpan=1},a:tr/a:tc{vMerge=false})", false], ["a:tbl/(a:tr/a:tc{rowSpan=2},a:tr/a:tc{vMerge=true})", true]], fn: (self: any, request: any): any => {
        let expected_value: any, tbl_cxml: any, tcs: any;
        [tbl_cxml, expected_value] = $list(request.param);
        tcs = element(tbl_cxml).xpath("//a:tc");
        return [$get(tcs, 0), $get(tcs, 1), expected_value];
      } },
    dimensions_fixture: { args: ["request"], params: [["a:tbl/a:tr/(a:tc,a:tc)", [1, 2]], ["a:tbl/(a:tr/a:tc,a:tr/a:tc)", [2, 1]], ["a:tbl/(a:tr/(a:tc,a:tc),a:tr/(a:tc,a:tc))", [2, 2]]], fn: (self: any, request: any): any => {
        let expected_value: any, tbl_cxml: any, tcs: any;
        [tbl_cxml, expected_value] = $list(request.param);
        tcs = element(tbl_cxml).xpath("//a:tc");
        return [$get(tcs, 0), $get(tcs, -1), expected_value];
      } },
    except_left_fixture: { args: ["request"], params: [["a:tbl/a:tr/(a:tc,a:tc)", [0, 1], [1]], ["a:tbl/(a:tr/a:tc,a:tr/a:tc)", [0, 1], []], ["a:tbl/(a:tr/(a:tc,a:tc),a:tr/(a:tc,a:tc))", [2, 1], [1, 3]], ["a:tbl/(a:tr/(a:tc,a:tc,a:tc),a:tr/(a:tc,a:tc,a:tc),a:tr/(a:tc,a:tc,a:tc))", [0, 8], [1, 2, 4, 5, 7, 8]]], fn: (self: any, request: any): any => {
        let expected_tc_idxs: any, expected_value: any, other_tc: any, tbl_cxml: any, tc: any, tc_idxs: any, tcs: any;
        [tbl_cxml, tc_idxs, expected_tc_idxs] = $list(request.param);
        tcs = element(tbl_cxml).xpath("//a:tc");
        [tc, other_tc] = $list([$get(tcs, $get(tc_idxs, 0)), $get(tcs, $get(tc_idxs, 1))]);
        expected_value = $list(expected_tc_idxs).map((idx: any) => $get(tcs, idx));
        return [tc, other_tc, expected_value];
      } },
    except_top_fixture: { args: ["request"], params: [["a:tbl/a:tr/(a:tc,a:tc)", [0, 1], []], ["a:tbl/(a:tr/a:tc,a:tr/a:tc)", [0, 1], [1]], ["a:tbl/(a:tr/(a:tc,a:tc),a:tr/(a:tc,a:tc))", [2, 1], [2, 3]], ["a:tbl/(a:tr/(a:tc,a:tc,a:tc),a:tr/(a:tc,a:tc,a:tc),a:tr/(a:tc,a:tc,a:tc))", [0, 8], [3, 4, 5, 6, 7, 8]]], fn: (self: any, request: any): any => {
        let expected_tc_idxs: any, expected_value: any, other_tc: any, tbl_cxml: any, tc: any, tc_idxs: any, tcs: any;
        [tbl_cxml, tc_idxs, expected_tc_idxs] = $list(request.param);
        tcs = element(tbl_cxml).xpath("//a:tc");
        [tc, other_tc] = $list([$get(tcs, $get(tc_idxs, 0)), $get(tcs, $get(tc_idxs, 1))]);
        expected_value = $list(expected_tc_idxs).map((idx: any) => $get(tcs, idx));
        return [tc, other_tc, expected_value];
      } },
    in_same_table_fixture: { args: ["request"], params: [true, false], fn: (self: any, request: any): any => {
        let expected_value: any, other_tbl: any, other_tc: any, tbl: any, tc: any;
        expected_value = request.param;
        tbl = element("a:tbl/a:tr/(a:tc,a:tc)");
        other_tbl = element("a:tbl/a:tr/(a:tc,a:tc)");
        tc = $get(tbl.xpath("//a:tc"), 0);
        other_tc = ($truthy(expected_value) ? $get(tbl.xpath("//a:tc"), 1) : $get(other_tbl.xpath("//a:tc"), 1));
        return [tc, other_tc, expected_value];
      } },
    left_col_fixture: { args: ["request"], params: [["a:tbl/a:tr/(a:tc,a:tc)", [0, 1], [0]], ["a:tbl/(a:tr/a:tc,a:tr/a:tc)", [0, 1], [0, 1]], ["a:tbl/(a:tr/(a:tc,a:tc),a:tr/(a:tc,a:tc))", [2, 1], [0, 2]], ["a:tbl/(a:tr/(a:tc,a:tc,a:tc),a:tr/(a:tc,a:tc,a:tc),a:tr/(a:tc,a:tc,a:tc))", [4, 8], [4, 7]]], fn: (self: any, request: any): any => {
        let expected_tc_idxs: any, expected_value: any, other_tc: any, tbl_cxml: any, tc: any, tc_idxs: any, tcs: any;
        [tbl_cxml, tc_idxs, expected_tc_idxs] = $list(request.param);
        tcs = element(tbl_cxml).xpath("//a:tc");
        [tc, other_tc] = $list([$get(tcs, $get(tc_idxs, 0)), $get(tcs, $get(tc_idxs, 1))]);
        expected_value = $list(expected_tc_idxs).map((idx: any) => $get(tcs, idx));
        return [tc, other_tc, expected_value];
      } },
    move_fixture: { args: ["request"], params: [["a:tbl/a:tr/(a:tc/a:txBody/a:p,a:tc/a:txBody/a:p)", ""], ["a:tbl/a:tr/(a:tc/a:txBody/a:p,a:tc/a:txBody/a:p/a:r/a:t\"b\")", "b"], ["a:tbl/a:tr/(a:tc/a:txBody/a:p/a:r/a:t\"a\",a:tc/a:txBody/a:p)", "a"], ["a:tbl/a:tr/(a:tc/a:txBody/a:p/a:r/a:t\"a\",a:tc/a:txBody/a:p/a:r/a:t\"b\")", "a\nb"], ["a:tbl/a:tr/(a:tc/a:txBody/a:p/a:r/a:t\"a\",a:tc/a:txBody/(a:p,a:p))", "a\n\n"], ["a:tbl/a:tr/(a:tc/a:txBody/(a:p,a:p),a:tc/a:txBody/a:p/a:r/a:t\"b\")", "\n\nb"]], fn: (self: any, request: any): any => {
        let expected_text: any, tbl_cxml: any, tcs: any;
        [tbl_cxml, expected_text] = $list(request.param);
        tcs = element(tbl_cxml).xpath("//a:tc");
        return [$get(tcs, 0), $get(tcs, 1), expected_text];
      } },
    top_row_fixture: { args: ["request"], params: [["a:tbl/a:tr/(a:tc,a:tc)", [0, 1], [0, 1]], ["a:tbl/(a:tr/a:tc,a:tr/a:tc)", [0, 1], [0]], ["a:tbl/(a:tr/(a:tc,a:tc),a:tr/(a:tc,a:tc))", [2, 1], [0, 1]], ["a:tbl/(a:tr/(a:tc,a:tc,a:tc),a:tr/(a:tc,a:tc,a:tc),a:tr/(a:tc,a:tc,a:tc))", [4, 8], [4, 5]]], fn: (self: any, request: any): any => {
        let expected_tc_idxs: any, expected_value: any, other_tc: any, tbl_cxml: any, tc: any, tc_idxs: any, tcs: any;
        [tbl_cxml, tc_idxs, expected_tc_idxs] = $list(request.param);
        tcs = element(tbl_cxml).xpath("//a:tc");
        [tc, other_tc] = $list([$get(tcs, $get(tc_idxs, 0)), $get(tcs, $get(tc_idxs, 1))]);
        expected_value = $list(expected_tc_idxs).map((idx: any) => $get(tcs, idx));
        return [tc, other_tc, expected_value];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_when_the_range_contains_a_merged_cell", args: ["contains_merge_fixture"], fn: (self: any, contains_merge_fixture: any): any => {
        let contains_merged_cell: any, expected_value: any, other_tc: any, tc: any, tc_range: any;
        [tc, other_tc, expected_value] = $list(contains_merge_fixture);
        tc_range = $call(TcRange, tc, other_tc);
        contains_merged_cell = tc_range.containsMergedCell;
        $assert($is(contains_merged_cell, expected_value));
      } },
    { name: "it_knows_how_big_the_merge_range_is", args: ["dimensions_fixture"], fn: (self: any, dimensions_fixture: any): any => {
        let dimensions: any, expected_value: any, other_tc: any, tc: any, tc_range: any;
        [tc, other_tc, expected_value] = $list(dimensions_fixture);
        tc_range = $call(TcRange, tc, other_tc);
        dimensions = tc_range.dimensions;
        $assertEqual(dimensions, expected_value);
      } },
    { name: "it_knows_when_tcs_are_in_the_same_tbl", args: ["in_same_table_fixture"], fn: (self: any, in_same_table_fixture: any): any => {
        let expected_value: any, in_same_table: any, other_tc: any, tc: any, tc_range: any;
        [tc, other_tc, expected_value] = $list(in_same_table_fixture);
        tc_range = $call(TcRange, tc, other_tc);
        in_same_table = tc_range.inSameTable;
        $assert($is(in_same_table, expected_value));
      } },
    { name: "it_can_iterate_tcs_not_in_left_col_of_range", args: ["except_left_fixture"], fn: (self: any, except_left_fixture: any): any => {
        let expected_value: any, other_tc: any, tc: any, tc_range: any, tcs: any;
        [tc, other_tc, expected_value] = $list(except_left_fixture);
        tc_range = $call(TcRange, tc, other_tc);
        tcs = $list(tc_range.iterExceptLeftColTcs());
        $assertEqual(tcs, expected_value);
      } },
    { name: "it_can_iterate_tcs_not_in_top_row_of_range", args: ["except_top_fixture"], fn: (self: any, except_top_fixture: any): any => {
        let expected_value: any, other_tc: any, tc: any, tc_range: any, tcs: any;
        [tc, other_tc, expected_value] = $list(except_top_fixture);
        tc_range = $call(TcRange, tc, other_tc);
        tcs = $list(tc_range.iterExceptTopRowTcs());
        $assertEqual(tcs, expected_value);
      } },
    { name: "it_can_iterate_left_col_of_range_tcs", args: ["left_col_fixture"], fn: (self: any, left_col_fixture: any): any => {
        let expected_value: any, other_tc: any, tc: any, tc_range: any, tcs: any;
        [tc, other_tc, expected_value] = $list(left_col_fixture);
        tc_range = $call(TcRange, tc, other_tc);
        tcs = $list(tc_range.iterLeftColTcs());
        $assertEqual(tcs, expected_value);
      } },
    { name: "it_can_iterate_top_row_of_range_tcs", args: ["top_row_fixture"], fn: (self: any, top_row_fixture: any): any => {
        let expected_value: any, other_tc: any, tc: any, tc_range: any, tcs: any;
        [tc, other_tc, expected_value] = $list(top_row_fixture);
        tc_range = $call(TcRange, tc, other_tc);
        tcs = $list(tc_range.iterTopRowTcs());
        $assertEqual(tcs, expected_value);
      } },
    { name: "it_can_migrate_range_content_to_origin_cell", args: ["move_fixture"], fn: (self: any, move_fixture: any): any => {
        let expected_text: any, other_tc: any, tc: any, tc_range: any;
        [tc, other_tc, expected_text] = $list(move_fixture);
        tc_range = $call(TcRange, tc, other_tc);
        tc_range.moveContentToOrigin();
        $assertEqual(tc.text, expected_text);
        $assertEqual(other_tc.text, "");
      } },
  ],
};
pytestClass("DescribeTcRange", DescribeTcRange);

