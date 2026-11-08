// Machine-translated from python-pptx tests/test_table.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "./pyshim.js";
import { $add, $all, $assert, $assertEqual, $call, $enumerate, $eq, $get, $getattr, $in, $is, $isinstance, $kw, $len, $list, $m, $mod, $raises, $range, $setattr, $split, $str, $t1, $truthy, call, classMock, instanceMock, propertyMock, pytestClass } from "./pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../src/exc.js";
import { defineChildren } from "../src/oxml/xmlchemy.js";
import { FillFormat as FillFormat$imported } from "../src/dml/fill.js";
import { MSO_ANCHOR as MSO_ANCHOR$imported } from "../src/enum/text.js";
import { qn as qn$imported } from "../src/oxml/ns.js";
import { CT_Table as CT_Table$imported, CT_TableCell as CT_TableCell$imported, TcRange as TcRange$imported } from "../src/oxml/table.js";
import { GraphicFrame as GraphicFrame$imported } from "../src/shapes/graphfrm.js";
import { Table as Table$imported, _Cell as _Cell$imported, _CellCollection as _CellCollection$imported, _Column as _Column$imported, _ColumnCollection as _ColumnCollection$imported, _Row as _Row$imported, _RowCollection as _RowCollection$imported } from "../src/table.js";
import { TextFrame as TextFrame$imported } from "../src/text/text.js";
import { Inches as Inches$imported, Length as Length$imported, Pt as Pt$imported } from "../src/util.js";
import { element, xml } from "./unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const FillFormat = FillFormat$imported;
const MSO_ANCHOR = MSO_ANCHOR$imported;
const qn = qn$imported;
const CT_Table = CT_Table$imported;
const CT_TableCell = CT_TableCell$imported;
const TcRange = TcRange$imported;
const GraphicFrame = GraphicFrame$imported;
const Table = Table$imported;
const _Cell = _Cell$imported;
const _CellCollection = _CellCollection$imported;
const _Column = _Column$imported;
const _ColumnCollection = _ColumnCollection$imported;
const _Row = _Row$imported;
const _RowCollection = _RowCollection$imported;
const TextFrame = TextFrame$imported;
const Inches = Inches$imported;
const Length = Length$imported;
const Pt = Pt$imported;
void $py;


const DescribeTable: any = {
  fixtures: {
    dx_fixture: { args: ["graphic_frame_"], fn: (self: any, graphic_frame_: any): any => {
        let expected_width: any, table: any, tbl_cxml: any;
        tbl_cxml = "a:tbl/a:tblGrid/(a:gridCol{w=111},a:gridCol{w=222})";
        table = $call(Table, element(tbl_cxml), graphic_frame_);
        expected_width = 333;
        return [table, expected_width];
      } },
    dy_fixture: { args: ["graphic_frame_"], fn: (self: any, graphic_frame_: any): any => {
        let expected_height: any, table: any, tbl_cxml: any;
        tbl_cxml = "a:tbl/(a:tr{h=100},a:tr{h=200})";
        table = $call(Table, element(tbl_cxml), graphic_frame_);
        expected_height = 300;
        return [table, expected_height];
      } },
    _Cell_: { args: ["request"], fn: (self: any, request: any): any => {
        return classMock(request, "pptx.table._Cell");
      } },
    cell_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _Cell);
      } },
    graphic_frame_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, GraphicFrame);
      } },
    tbl_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, CT_Table);
      } },
    tc_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, CT_TableCell);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_access_to_its_cells", args: ["tbl_", "tc_", "_Cell_", "cell_"], fn: (self: any, tbl_: any, tc_: any, _Cell_: any, cell_: any): any => {
        let cell: any, col_idx: any, row_idx: any, table: any;
        [row_idx, col_idx] = $list([4, 2]);
        tbl_.tc.returnValue = tc_;
        _Cell_.returnValue = cell_;
        table = $call(Table, tbl_, null);
        cell = table.cell(row_idx, col_idx);
        tbl_.tc.assertCalledOnceWith(row_idx, col_idx);
        _Cell_.assertCalledOnceWith(tc_, table);
        $assert($is(cell, cell_));
      } },
    { name: "it_provides_access_to_its_columns", args: ["request"], fn: (self: any, request: any): any => {
        let _ColumnCollection_: any, columns: any, columns_: any, table: any, tbl: any;
        columns_ = instanceMock(request, _ColumnCollection);
        _ColumnCollection_ = $call(classMock, request, "pptx.table._ColumnCollection", $kw({return_value: columns_}));
        tbl = element("a:tbl");
        table = $call(Table, tbl, null);
        columns = table.columns;
        _ColumnCollection_.assertCalledOnceWith(tbl, table);
        $assert($is(columns, columns_));
      } },
    { name: "it_can_iterate_its_grid_cells", args: ["request", "_Cell_"], fn: (self: any, request: any, _Cell_: any): any => {
        let cells: any, expected_cells: any, expected_tcs: any, table: any, tbl: any;
        tbl = element("a:tbl/(a:tr/(a:tc,a:tc),a:tr/(a:tc,a:tc))");
        expected_tcs = tbl.xpath(".//a:tc");
        const $t1 = $list($range(4)).map((idx: any) => $call(instanceMock, request, _Cell, $kw({name: $mod("cell%d", idx)})));
        expected_cells = $t1;
        _Cell_.sideEffect = $t1;
        table = $call(Table, tbl, null);
        cells = $list(table.iterCells());
        $assertEqual(cells, expected_cells);
        $assertEqual(_Cell_.callArgsList, $list(expected_tcs).map((tc: any) => call(tc, table)));
      } },
    { name: "it_provides_access_to_its_rows", args: ["request"], fn: (self: any, request: any): any => {
        let _RowCollection_: any, rows: any, rows_: any, table: any, tbl: any;
        rows_ = instanceMock(request, _RowCollection);
        _RowCollection_ = $call(classMock, request, "pptx.table._RowCollection", $kw({return_value: rows_}));
        tbl = element("a:tbl");
        table = $call(Table, tbl, null);
        rows = table.rows;
        _RowCollection_.assertCalledOnceWith(tbl, table);
        $assert($is(rows, rows_));
      } },
    { name: "it_updates_graphic_frame_width_on_width_change", args: ["dx_fixture"], fn: (self: any, dx_fixture: any): any => {
        let expected_width: any, table: any;
        [table, expected_width] = $list(dx_fixture);
        table.notifyWidthChanged();
        $assertEqual(table._graphicFrame.width, expected_width);
      } },
    { name: "it_updates_graphic_frame_height_on_height_change", args: ["dy_fixture"], fn: (self: any, dy_fixture: any): any => {
        let expected_height: any, table: any;
        [table, expected_height] = $list(dy_fixture);
        table.notifyHeightChanged();
        $assertEqual(table._graphicFrame.height, expected_height);
      } },
  ],
};
pytestClass("DescribeTable", DescribeTable);

const DescribeTableBooleanProperties: any = {
  fixtures: {
    boolprop_get_fixture: { args: ["request"], params: [["a:tbl", "first_row", false], ["a:tbl/a:tblPr", "first_row", false], ["a:tbl/a:tblPr{firstRow=1}", "first_row", true], ["a:tbl/a:tblPr{firstRow=0}", "first_row", false], ["a:tbl/a:tblPr{firstRow=true}", "first_row", true], ["a:tbl/a:tblPr{firstRow=false}", "first_row", false], ["a:tbl/a:tblPr{firstCol=1}", "first_col", true], ["a:tbl/a:tblPr{lastRow=0}", "last_row", false], ["a:tbl/a:tblPr{lastCol=true}", "last_col", true], ["a:tbl/a:tblPr{bandRow=false}", "horz_banding", false], ["a:tbl/a:tblPr", "vert_banding", false]], fn: (self: any, request: any): any => {
        let boolprop_name: any, expected_value: any, table: any, tbl_cxml: any;
        [tbl_cxml, boolprop_name, expected_value] = $list(request.param);
        table = $call(Table, element(tbl_cxml), null);
        return [table, boolprop_name, expected_value];
      } },
    boolprop_set_fixture: { args: ["request"], params: [["a:tbl", "first_row", true, "a:tbl/a:tblPr{firstRow=1}"], ["a:tbl", "first_row", false, "a:tbl/a:tblPr"], ["a:tbl/a:tblPr", "first_row", true, "a:tbl/a:tblPr{firstRow=1}"], ["a:tbl/a:tblPr", "first_row", false, "a:tbl/a:tblPr"], ["a:tbl/a:tblPr{firstRow=true}", "first_row", true, "a:tbl/a:tblPr{firstRow=1}"], ["a:tbl/a:tblPr{firstRow=false}", "first_row", false, "a:tbl/a:tblPr"], ["a:tbl/a:tblPr{bandRow=1}", "first_row", true, "a:tbl/a:tblPr{bandRow=1,firstRow=1}"], ["a:tbl", "first_col", true, "a:tbl/a:tblPr{firstCol=1}"], ["a:tbl", "last_row", true, "a:tbl/a:tblPr{lastRow=1}"], ["a:tbl", "last_col", true, "a:tbl/a:tblPr{lastCol=1}"], ["a:tbl", "horz_banding", true, "a:tbl/a:tblPr{bandRow=1}"], ["a:tbl", "vert_banding", true, "a:tbl/a:tblPr{bandCol=1}"]], fn: (self: any, request: any): any => {
        let boolprop_name: any, expected_tbl_cxml: any, expected_xml: any, new_value: any, table: any, tbl_cxml: any;
        [tbl_cxml, boolprop_name, new_value, expected_tbl_cxml] = $list(request.param);
        table = $call(Table, element(tbl_cxml), null);
        expected_xml = xml(expected_tbl_cxml);
        return [table, boolprop_name, new_value, expected_xml];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_boolean_property_settings", args: ["boolprop_get_fixture"], fn: (self: any, boolprop_get_fixture: any): any => {
        let boolprop_name: any, boolprop_value: any, expected_value: any, table: any;
        [table, boolprop_name, expected_value] = $list(boolprop_get_fixture);
        boolprop_value = $getattr(table, boolprop_name);
        $assert($is(boolprop_value, expected_value));
      } },
    { name: "it_can_change_its_boolean_property_settings", args: ["boolprop_set_fixture"], fn: (self: any, boolprop_set_fixture: any): any => {
        let boolprop_name: any, expected_xml: any, new_value: any, table: any;
        [table, boolprop_name, new_value, expected_xml] = $list(boolprop_set_fixture);
        $setattr(table, boolprop_name, new_value);
        $assertEqual(table._tbl.xml, expected_xml);
      } },
  ],
};
pytestClass("DescribeTableBooleanProperties", DescribeTableBooleanProperties);

const Describe_Cell: any = {
  fixtures: {
    anchor_get_fixture: { args: ["request"], params: [["a:tc", null], ["a:tc/a:tcPr", null], ["a:tc/a:tcPr{anchor=t}", MSO_ANCHOR.TOP], ["a:tc/a:tcPr{anchor=ctr}", MSO_ANCHOR.MIDDLE], ["a:tc/a:tcPr{anchor=b}", MSO_ANCHOR.BOTTOM]], fn: (self: any, request: any): any => {
        let cell: any, expected_value: any, tc_cxml: any;
        [tc_cxml, expected_value] = $list(request.param);
        cell = $call(_Cell, element(tc_cxml), null);
        return [cell, expected_value];
      } },
    anchor_set_fixture: { args: ["request"], params: [["a:tc", null, "a:tc"], ["a:tc", MSO_ANCHOR.TOP, "a:tc/a:tcPr{anchor=t}"], ["a:tc", MSO_ANCHOR.MIDDLE, "a:tc/a:tcPr{anchor=ctr}"], ["a:tc", MSO_ANCHOR.BOTTOM, "a:tc/a:tcPr{anchor=b}"], ["a:tc/a:tcPr{anchor=t}", MSO_ANCHOR.MIDDLE, "a:tc/a:tcPr{anchor=ctr}"], ["a:tc/a:tcPr{anchor=ctr}", null, "a:tc/a:tcPr"]], fn: (self: any, request: any): any => {
        let cell: any, expected_tc_cxml: any, expected_xml: any, new_value: any, tc_cxml: any;
        [tc_cxml, new_value, expected_tc_cxml] = $list(request.param);
        cell = $call(_Cell, element(tc_cxml), null);
        expected_xml = xml(expected_tc_cxml);
        return [cell, new_value, expected_xml];
      } },
    fill_fixture: { args: ["cell"], fn: (self: any, cell: any): any => {
        return cell;
      } },
    height_fixture: { args: ["request"], params: [["a:tc", 1], ["a:tc{gridSpan=2}", 1], ["a:tc{rowSpan=42}", 42]], fn: (self: any, request: any): any => {
        let expected_value: any, tc: any, tc_cxml: any;
        [tc_cxml, expected_value] = $list(request.param);
        tc = element(tc_cxml);
        return [tc, expected_value];
      } },
    margin_get_fixture: { args: ["request"], params: [["a:tc/a:tcPr{marL=82296}", "margin_left", $call(Inches, 0.09)], ["a:tc/a:tcPr{marR=73152}", "margin_right", $call(Inches, 0.08)], ["a:tc/a:tcPr{marT=64008}", "margin_top", $call(Inches, 0.07)], ["a:tc/a:tcPr{marB=54864}", "margin_bottom", $call(Inches, 0.06)], ["a:tc", "margin_left", $call(Inches, 0.1)], ["a:tc/a:tcPr", "margin_right", $call(Inches, 0.1)], ["a:tc", "margin_top", $call(Inches, 0.05)], ["a:tc/a:tcPr", "margin_bottom", $call(Inches, 0.05)]], fn: (self: any, request: any): any => {
        let cell: any, expected_value: any, margin_prop_name: any, tc_cxml: any;
        [tc_cxml, margin_prop_name, expected_value] = $list(request.param);
        cell = $call(_Cell, element(tc_cxml), null);
        return [cell, margin_prop_name, expected_value];
      } },
    margin_set_fixture: { args: ["request"], params: [["a:tc", "margin_left", $call(Inches, 0.08), "a:tc/a:tcPr{marL=73152}"], ["a:tc", "margin_right", $call(Inches, 0.08), "a:tc/a:tcPr{marR=73152}"], ["a:tc", "margin_top", $call(Inches, 0.08), "a:tc/a:tcPr{marT=73152}"], ["a:tc", "margin_bottom", $call(Inches, 0.08), "a:tc/a:tcPr{marB=73152}"], ["a:tc", "margin_left", null, "a:tc"], ["a:tc/a:tcPr{marL=42}", "margin_left", null, "a:tc/a:tcPr"]], fn: (self: any, request: any): any => {
        let cell: any, expected_tc_cxml: any, expected_xml: any, margin_prop_name: any, new_value: any, tc_cxml: any;
        [tc_cxml, margin_prop_name, new_value, expected_tc_cxml] = $list(request.param);
        cell = $call(_Cell, element(tc_cxml), null);
        expected_xml = xml(expected_tc_cxml);
        return [cell, margin_prop_name, new_value, expected_xml];
      } },
    margin_raises_fixture: { args: ["request"], params: ["margin_left", "margin_right", "margin_top", "margin_bottom"], fn: (self: any, request: any): any => {
        let cell: any, margin_prop_name: any, val_of_invalid_type: any;
        margin_prop_name = request.param;
        cell = $call(_Cell, element("a:tc"), null);
        val_of_invalid_type = "foobar";
        return [cell, margin_prop_name, val_of_invalid_type];
      } },
    origin_fixture: { args: ["request"], params: [["a:tc", false], ["a:tc{gridSpan=1}", false], ["a:tc{hMerge=1}", false], ["a:tc{gridSpan=2,vMerge=1}", false], ["a:tc{gridSpan=2}", true], ["a:tc{rowSpan=2}", true], ["a:tc{gridSpan=2,rowSpan=3}", true]], fn: (self: any, request: any): any => {
        let expected_value: any, tc: any, tc_cxml: any;
        [tc_cxml, expected_value] = $list(request.param);
        tc = element(tc_cxml);
        return [tc, expected_value];
      } },
    spanned_fixture: { args: ["request"], params: [["a:tc", false], ["a:tc{gridSpan=2}", false], ["a:tc{hMerge=1}", true], ["a:tc{gridSpan=2,vMerge=1}", true], ["a:tc{rowSpan=2,hMerge=true}", true], ["a:tc{gridSpan=2,rowSpan=3}", false]], fn: (self: any, request: any): any => {
        let expected_value: any, tc: any, tc_cxml: any;
        [tc_cxml, expected_value] = $list(request.param);
        tc = element(tc_cxml);
        return [tc, expected_value];
      } },
    split_fixture: { args: ["request"], params: [["a:tbl/(a:tr/(a:tc{gridSpan=2},a:tc{hMerge=1}),a:tr/(a:tc,a:tc))", 0, [0, 1]], ["a:tbl/(a:tr/(a:tc{rowSpan=2},a:tc),a:tr/(a:tc{vMerge=1},a:tc))", 0, [0, 2]], ["a:tbl/(a:tr/(a:tc{gridSpan=2,rowSpan=2},a:tc{hMerge=1,rowSpan=2}),a:tr/(a:tc{gridSpan=2,vMerge=1},a:tc{hMerge=1,vMerge=1}))", 0, [0, 1, 2, 3]]], fn: (self: any, request: any): any => {
        let origin_tc: any, origin_tc_idx: any, range_tc_idxs: any, range_tcs: any, tbl_cxml: any, tcs: any;
        [tbl_cxml, origin_tc_idx, range_tc_idxs] = $list(request.param);
        tcs = element(tbl_cxml).xpath("//a:tc");
        origin_tc = $get(tcs, origin_tc_idx);
        range_tcs = $list($list(range_tc_idxs).map((idx: any) => $get(tcs, idx)));
        return [origin_tc, range_tcs];
      } },
    width_fixture: { args: ["request"], params: [["a:tc", 1], ["a:tc{rowSpan=2}", 1], ["a:tc{gridSpan=24}", 24]], fn: (self: any, request: any): any => {
        let expected_value: any, tc: any, tc_cxml: any;
        [tc_cxml, expected_value] = $list(request.param);
        tc = element(tc_cxml);
        return [tc, expected_value];
      } },
    cell: { args: [], fn: (self: any): any => {
        return $call(_Cell, element("a:tc"), null);
      } },
    TcRange_: { args: ["request"], fn: (self: any, request: any): any => {
        return classMock(request, "pptx.table.TcRange");
      } },
    tc_range_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, TcRange);
      } },
    text_frame_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, TextFrame);
      } },
    text_frame_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, _Cell, "text_frame");
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_is_equal_to_other_instance_having_same_tc", args: [], fn: (self: any): any => {
        let cell: any, cell_with_other_tc: any, cell_with_same_tc: any, other_tc: any, tc: any;
        tc = element("a:tc");
        other_tc = element("a:tc");
        cell = $call(_Cell, tc, null);
        cell_with_same_tc = $call(_Cell, tc, null);
        cell_with_other_tc = $call(_Cell, other_tc, null);
        $assertEqual(cell, cell_with_same_tc);
        $assert(!$eq(cell, cell_with_other_tc));
      } },
    { name: "it_has_a_fill", args: ["fill_fixture"], fn: (self: any, fill_fixture: any): any => {
        let cell: any;
        cell = fill_fixture;
        $assert($isinstance(cell.fill, FillFormat));
      } },
    { name: "it_knows_whether_it_is_merge_origin_cell", args: ["origin_fixture"], fn: (self: any, origin_fixture: any): any => {
        let cell: any, expected_value: any, is_merge_origin: any, tc: any;
        [tc, expected_value] = $list(origin_fixture);
        cell = $call(_Cell, tc, null);
        is_merge_origin = cell.isMergeOrigin;
        $assert($is(is_merge_origin, expected_value));
      } },
    { name: "it_knows_whether_it_is_spanned", args: ["spanned_fixture"], fn: (self: any, spanned_fixture: any): any => {
        let cell: any, expected_value: any, is_spanned: any, tc: any;
        [tc, expected_value] = $list(spanned_fixture);
        cell = $call(_Cell, tc, null);
        is_spanned = cell.isSpanned;
        $assert($is(is_spanned, expected_value));
      } },
    { name: "it_knows_its_margin_settings", args: ["margin_get_fixture"], fn: (self: any, margin_get_fixture: any): any => {
        let cell: any, expected_value: any, margin_prop_name: any, margin_value: any;
        [cell, margin_prop_name, expected_value] = $list(margin_get_fixture);
        margin_value = $getattr(cell, margin_prop_name);
        $assertEqual(margin_value, expected_value);
      } },
    { name: "it_can_change_its_margin_settings", args: ["margin_set_fixture"], fn: (self: any, margin_set_fixture: any): any => {
        let cell: any, expected_xml: any, margin_prop_name: any, new_value: any;
        [cell, margin_prop_name, new_value, expected_xml] = $list(margin_set_fixture);
        $setattr(cell, margin_prop_name, new_value);
        $assertEqual(cell._tc.xml, expected_xml);
      } },
    { name: "it_raises_on_margin_assigned_other_than_int_or_None", args: ["margin_raises_fixture"], fn: (self: any, margin_raises_fixture: any): any => {
        let cell: any, margin_attr_name: any, val_of_invalid_type: any;
        [cell, margin_attr_name, val_of_invalid_type] = $list(margin_raises_fixture);
        $raises(TypeError, () => {
          $setattr(cell, margin_attr_name, val_of_invalid_type);
        });
      } },
    { name: "it_can_merge_a_range_of_cells", args: ["TcRange_", "tc_range_"], fn: (self: any, TcRange_: any, tc_range_: any): any => {
        let cell: any, expected_xml: any, other_cell: any, other_tc: any, tbl: any, tc: any;
        tbl = element("a:tbl/(a:tr/(a:tc,a:tc),a:tr/(a:tc,a:tc))");
        [tc, other_tc] = $list([tbl.tc(0, 0), tbl.tc(1, 1)]);
        TcRange_.returnValue = tc_range_;
        tc_range_.containsMergedCell = false;
        tc_range_.dimensions = [2, 2];
        const tcs = (...rowcols: any[]): any => {
          return $list(rowcols).map((rowcol: any) => $m(tbl, "tc", ...$list(rowcol)));
        };
        tc_range_.iterTopRowTcs.returnValue = tcs([0, 0], [0, 1]);
        tc_range_.iterLeftColTcs.returnValue = tcs([0, 0], [1, 0]);
        tc_range_.iterExceptLeftColTcs.returnValue = tcs([0, 1], [1, 1]);
        tc_range_.iterExceptTopRowTcs.returnValue = tcs([1, 0], [1, 1]);
        expected_xml = xml("a:tbl/(a:tr/(a:tc{gridSpan=2,rowSpan=2},a:tc{rowSpan=2,hMerge=1}),a:tr/(a:tc{gridSpan=2,vMerge=1},a:tc{hMerge=1,vMerge=1}))");
        [cell, other_cell] = $list([$call(_Cell, tc, null), $call(_Cell, other_tc, null)]);
        cell.merge(other_cell);
        TcRange_.assertCalledOnceWith(tc, other_tc);
        tc_range_.moveContentToOrigin.assertCalledOnceWith();
        $assertEqual(tbl.xml, expected_xml);
      } },
    { name: "but_it_raises_when_cells_are_from_different_tables", args: ["TcRange_", "tc_range_"], fn: (self: any, TcRange_: any, tc_range_: any): any => {
        let cell: any, e: any, other_cell: any;
        TcRange_.returnValue = tc_range_;
        tc_range_.inSameTable = false;
        [cell, other_cell] = $list([$call(_Cell, null, null), $call(_Cell, null, null)]);
        e = $raises(ValueError, () => {
          cell.merge(other_cell);
        });
        $assert($in("different table", $str(e.value)));
      } },
    { name: "and_it_raises_when_range_contains_merged_cell", args: ["TcRange_", "tc_range_"], fn: (self: any, TcRange_: any, tc_range_: any): any => {
        let cell: any, e: any, other_cell: any;
        TcRange_.returnValue = tc_range_;
        tc_range_.containsMergedCell = true;
        [cell, other_cell] = $list([$call(_Cell, null, null), $call(_Cell, null, null)]);
        e = $raises(ValueError, () => {
          cell.merge(other_cell);
        });
        $assert($in("contains one or more merged cells", $str(e.value)));
      } },
    { name: "it_knows_how_many_rows_the_merge_spans", args: ["height_fixture"], fn: (self: any, height_fixture: any): any => {
        let cell: any, expected_value: any, span_height: any, tc: any;
        [tc, expected_value] = $list(height_fixture);
        cell = $call(_Cell, tc, null);
        span_height = cell.spanHeight;
        $assertEqual(span_height, expected_value);
      } },
    { name: "it_knows_how_many_columns_the_merge_spans", args: ["width_fixture"], fn: (self: any, width_fixture: any): any => {
        let cell: any, expected_value: any, span_width: any, tc: any;
        [tc, expected_value] = $list(width_fixture);
        cell = $call(_Cell, tc, null);
        span_width = cell.spanWidth;
        $assertEqual(span_width, expected_value);
      } },
    { name: "it_can_split_a_merged_cell", args: ["split_fixture"], fn: (self: any, split_fixture: any): any => {
        let cell: any, origin_tc: any, range_tcs: any;
        [origin_tc, range_tcs] = $list(split_fixture);
        cell = $call(_Cell, origin_tc, null);
        $split(cell);
        $assert($all($list(range_tcs).map((tc: any) => $eq(tc.gridSpan, 1))));
        $assert($all($list(range_tcs).map((tc: any) => $eq(tc.rowSpan, 1))));
        $assert($all($list(range_tcs).map((tc: any) => !$truthy(tc.hMerge))));
        $assert($all($list(range_tcs).map((tc: any) => !$truthy(tc.vMerge))));
      } },
    { name: "but_it_raises_when_cell_to_be_split_is_not_merge_origin", args: [], fn: (self: any): any => {
        let cell: any, e: any, tc: any;
        tc = $get(element("a:tbl/a:tr/a:tc").xpath("//a:tc"), 0);
        cell = $call(_Cell, tc, null);
        e = $raises(ValueError, () => {
          $split(cell);
        });
        $assert($in("not a merge-origin cell", $str(e.value)));
      } },
    { name: "it_knows_what_text_it_contains", args: ["text_frame_prop_", "text_frame_"], fn: (self: any, text_frame_prop_: any, text_frame_: any): any => {
        let cell: any, text: any;
        text_frame_prop_.returnValue = text_frame_;
        text_frame_.text = "foobar";
        cell = $call(_Cell, null, null);
        text = cell.text;
        $assertEqual(text, "foobar");
      } },
    { name: "it_can_change_its_text", args: ["text_frame_prop_", "text_frame_"], fn: (self: any, text_frame_prop_: any, text_frame_: any): any => {
        let cell: any;
        text_frame_prop_.returnValue = text_frame_;
        cell = $call(_Cell, null, null);
        cell.text = "føøbår";
        $assertEqual(text_frame_.text, "føøbår");
      } },
    { name: "it_knows_its_vertical_anchor_setting", args: ["anchor_get_fixture"], fn: (self: any, anchor_get_fixture: any): any => {
        let cell: any, expected_value: any;
        [cell, expected_value] = $list(anchor_get_fixture);
        $assertEqual(cell.verticalAnchor, expected_value);
      } },
    { name: "it_can_change_its_vertical_anchor", args: ["anchor_set_fixture"], fn: (self: any, anchor_set_fixture: any): any => {
        let cell: any, expected_xml: any, new_value: any;
        [cell, new_value, expected_xml] = $list(anchor_set_fixture);
        cell.verticalAnchor = new_value;
        $assertEqual(cell._tc.xml, expected_xml);
      } },
  ],
};
pytestClass("Describe_Cell", Describe_Cell);

const Describe_CellCollection: any = {
  fixtures: {
    iter_fixture: { args: ["request", "_Cell_"], params: ["a:tr", "a:tr/a:tc", "a:tr/(a:tc, a:tc, a:tc)"], fn: (self: any, request: any, _Cell_: any): any => {
        let calls: any, cell_collection: any, expected_cells: any, tcs: any, tr: any, tr_cxml: any;
        tr_cxml = request.param;
        tr = element(tr_cxml);
        tcs = tr.xpath("//a:tc");
        cell_collection = $call(_CellCollection, tr, null);
        expected_cells = $list($range($len(tcs))).map((idx: any) => $call(instanceMock, request, _Cell, $kw({name: $mod("cell%d", idx)})));
        _Cell_.sideEffect = expected_cells;
        calls = $list(tcs).map((tc: any) => call(tc, cell_collection));
        return [cell_collection, _Cell_, calls, expected_cells];
      } },
    len_fixture: { args: ["request"], params: [["a:tr", 0], ["a:tr/a:tc", 1], ["a:tr/(a:tc, a:tc)", 2]], fn: (self: any, request: any): any => {
        let cells: any, expected_len: any, tr_cxml: any;
        [tr_cxml, expected_len] = $list(request.param);
        cells = $call(_CellCollection, element(tr_cxml), null);
        return [cells, expected_len];
      } },
    _Cell_: { args: ["request"], fn: (self: any, request: any): any => {
        return classMock(request, "pptx.table._Cell");
      } },
    cell_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _Cell);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_how_many_cells_it_contains", args: ["len_fixture"], fn: (self: any, len_fixture: any): any => {
        let cells: any, expected_count: any;
        [cells, expected_count] = $list(len_fixture);
        $assertEqual($len(cells), expected_count);
      } },
    { name: "it_can_iterate_over_the_cells_it_contains", args: ["iter_fixture"], fn: (self: any, iter_fixture: any): any => {
        let _Cell_: any, calls: any, cell_collection: any, cells: any, expected_cells: any;
        [cell_collection, _Cell_, calls, expected_cells] = $list(iter_fixture);
        cells = $list(cell_collection);
        $assertEqual(_Cell_.callArgsList, calls);
        $assertEqual(cells, expected_cells);
      } },
    { name: "it_supports_indexed_access", args: ["_Cell_", "cell_"], fn: (self: any, _Cell_: any, cell_: any): any => {
        let cell: any, cell_collection: any, tcs: any, tr: any;
        tr = element("a:tr/(a:tc, a:tc, a:tc)");
        tcs = tr.xpath("//a:tc");
        _Cell_.returnValue = cell_;
        cell_collection = $call(_CellCollection, tr, null);
        cell = $get(cell_collection, 1);
        _Cell_.assertCalledOnceWith($get(tcs, 1), cell_collection);
        $assert($is(cell, cell_));
      } },
    { name: "it_raises_on_indexed_access_out_of_range", args: [], fn: (self: any): any => {
        let cells: any;
        cells = $call(_CellCollection, element("a:tr/a:tc"), null);
        $raises(IndexError, () => {
          $get(cells, -1);
        });
        $raises(IndexError, () => {
          $get(cells, 9);
        });
      } },
  ],
};
pytestClass("Describe_CellCollection", Describe_CellCollection);

const Describe_Column: any = {
  fixtures: {
    width_get_fixture: { args: ["request"], params: [["a:gridCol{w=914400}", $call(Inches, 1)], ["a:gridCol{w=10pt}", $call(Pt, 10)]], fn: (self: any, request: any): any => {
        let column: any, expected_value: any, gridCol_cxml: any;
        [gridCol_cxml, expected_value] = $list(request.param);
        column = $call(_Column, element(gridCol_cxml), null);
        return [column, expected_value];
      } },
    width_set_fixture: { args: ["request", "parent_"], params: [["a:gridCol{w=12pt}", $call(Inches, 1), "a:gridCol{w=914400}"], ["a:gridCol{w=1234}", $call(Inches, 1), "a:gridCol{w=914400}"]], fn: (self: any, request: any, parent_: any): any => {
        let column: any, expected_gridCol_cxml: any, expected_xml: any, gridCol_cxml: any, new_width: any;
        [gridCol_cxml, new_width, expected_gridCol_cxml] = $list(request.param);
        column = $call(_Column, element(gridCol_cxml), parent_);
        expected_xml = xml(expected_gridCol_cxml);
        return [column, new_width, expected_xml, parent_];
      } },
    parent_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _ColumnCollection);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_width", args: ["width_get_fixture"], fn: (self: any, width_get_fixture: any): any => {
        let column: any, expected_value: any, width: any;
        [column, expected_value] = $list(width_get_fixture);
        width = column.width;
        $assertEqual(width, expected_value);
        $assert($isinstance(width, Length));
      } },
    { name: "it_can_change_its_width", args: ["width_set_fixture"], fn: (self: any, width_set_fixture: any): any => {
        let column: any, expected_xml: any, new_width: any, parent_: any;
        [column, new_width, expected_xml, parent_] = $list(width_set_fixture);
        column.width = new_width;
        $assertEqual(column._gridCol.xml, expected_xml);
        parent_.notifyWidthChanged.assertCalledOnceWith();
      } },
  ],
};
pytestClass("Describe_Column", Describe_Column);

const Describe_ColumnCollection: any = {
  fixtures: {
    getitem_fixture: { args: ["request"], params: ["a:tbl/a:tblGrid", "a:tbl/a:tblGrid/a:gridCol", "a:tbl/a:tblGrid/(a:gridCol, a:gridCol, a:gridCol)"], fn: (self: any, request: any): any => {
        let columns: any, expected_column_lst: any, tbl: any, tbl_cxml: any;
        tbl_cxml = request.param;
        tbl = element(tbl_cxml);
        columns = $call(_ColumnCollection, tbl, null);
        expected_column_lst = tbl.xpath("//a:gridCol");
        return [columns, expected_column_lst];
      } },
    iter_fixture: { args: ["request"], params: ["a:tbl/a:tblGrid", "a:tbl/a:tblGrid/a:gridCol", "a:tbl/a:tblGrid/(a:gridCol, a:gridCol, a:gridCol)"], fn: (self: any, request: any): any => {
        let columns: any, expected_column_lst: any, tbl: any, tbl_cxml: any;
        tbl_cxml = request.param;
        tbl = element(tbl_cxml);
        columns = $call(_ColumnCollection, tbl, null);
        expected_column_lst = tbl.xpath("//a:gridCol");
        return [columns, expected_column_lst];
      } },
    len_fixture: { args: ["request"], params: [["a:tbl/a:tblGrid", 0], ["a:tbl/a:tblGrid/a:gridCol", 1], ["a:tbl/a:tblGrid/(a:gridCol,a:gridCol)", 2]], fn: (self: any, request: any): any => {
        let columns: any, expected_len: any, tbl_cxml: any;
        [tbl_cxml, expected_len] = $list(request.param);
        columns = $call(_ColumnCollection, element(tbl_cxml), null);
        return [columns, expected_len];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_how_many_columns_it_contains", args: ["len_fixture"], fn: (self: any, len_fixture: any): any => {
        let columns: any, expected_count: any;
        [columns, expected_count] = $list(len_fixture);
        $assertEqual($len(columns), expected_count);
      } },
    { name: "it_can_iterate_over_the_columns_it_contains", args: ["iter_fixture"], fn: (self: any, iter_fixture: any): any => {
        let columns: any, count: any, expected_gridCol_lst: any;
        [columns, expected_gridCol_lst] = $list(iter_fixture);
        count = 0;
        for (const [idx, column] of $list($enumerate(columns))) {
          $assert($isinstance(column, _Column));
          $assert($is(column._gridCol, $get(expected_gridCol_lst, idx)));
          count = $add(count, 1);
        }
        $assertEqual(count, $len(expected_gridCol_lst));
      } },
    { name: "it_supports_indexed_access", args: ["getitem_fixture"], fn: (self: any, getitem_fixture: any): any => {
        let column: any, columns: any, expected_gridCol_lst: any;
        [columns, expected_gridCol_lst] = $list(getitem_fixture);
        for (const [idx, gridCol] of $list($enumerate(expected_gridCol_lst))) {
          column = $get(columns, idx);
          $assert($isinstance(column, _Column));
          $assert($is(column._gridCol, gridCol));
        }
      } },
    { name: "it_raises_on_indexed_access_out_of_range", args: [], fn: (self: any): any => {
        let columns: any;
        columns = $call(_ColumnCollection, element("a:tbl/a:tblGrid/a:gridCol"), null);
        $raises(IndexError, () => {
          $get(columns, -1);
        });
        $raises(IndexError, () => {
          $get(columns, 9);
        });
      } },
  ],
};
pytestClass("Describe_ColumnCollection", Describe_ColumnCollection);

const Describe_Row: any = {
  fixtures: {
    cells_fixture: { args: ["_CellCollection_", "cells_"], fn: (self: any, _CellCollection_: any, cells_: any): any => {
        let row: any;
        row = $call(_Row, element("a:tr"), null);
        return [row, _CellCollection_, cells_];
      } },
    height_get_fixture: { args: ["request"], params: [["a:tr{h=914400}", $call(Inches, 1)], ["a:tr{h=10pt}", $call(Pt, 10)]], fn: (self: any, request: any): any => {
        let expected_value: any, row: any, tr_cxml: any;
        [tr_cxml, expected_value] = $list(request.param);
        row = $call(_Row, element(tr_cxml), null);
        return [row, expected_value];
      } },
    height_set_fixture: { args: ["request", "parent_"], params: [["a:tr{h=12pt}", $call(Inches, 1), "a:tr{h=914400}"], ["a:tr{h=1234}", $call(Inches, 1), "a:tr{h=914400}"]], fn: (self: any, request: any, parent_: any): any => {
        let expected_tr_cxml: any, expected_xml: any, new_height: any, row: any, tr_cxml: any;
        [tr_cxml, new_height, expected_tr_cxml] = $list(request.param);
        row = $call(_Row, element(tr_cxml), parent_);
        expected_xml = xml(expected_tr_cxml);
        return [row, new_height, expected_xml, parent_];
      } },
    _CellCollection_: { args: ["request", "cells_"], fn: (self: any, request: any, cells_: any): any => {
        return $call(classMock, request, "pptx.table._CellCollection", $kw({return_value: cells_}));
      } },
    cells_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _CellCollection);
      } },
    parent_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _RowCollection);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_height", args: ["height_get_fixture"], fn: (self: any, height_get_fixture: any): any => {
        let expected_value: any, height: any, row: any;
        [row, expected_value] = $list(height_get_fixture);
        height = row.height;
        $assertEqual(height, expected_value);
        $assert($isinstance(height, Length));
      } },
    { name: "it_can_change_its_height", args: ["height_set_fixture"], fn: (self: any, height_set_fixture: any): any => {
        let expected_xml: any, new_height: any, parent_: any, row: any;
        [row, new_height, expected_xml, parent_] = $list(height_set_fixture);
        row.height = new_height;
        $assertEqual(row._tr.xml, expected_xml);
        parent_.notifyHeightChanged.assertCalledOnceWith();
      } },
    { name: "it_provides_access_to_its_cells", args: ["cells_fixture"], fn: (self: any, cells_fixture: any): any => {
        let _CellCollection_: any, cells: any, cells_: any, row: any;
        [row, _CellCollection_, cells_] = $list(cells_fixture);
        cells = row.cells;
        _CellCollection_.assertCalledOnceWith(row._tr, row);
        $assert($is(cells, cells_));
      } },
  ],
};
pytestClass("Describe_Row", Describe_Row);

const Describe_RowCollection: any = {
  fixtures: {
    getitem_fixture: { args: ["request"], params: ["a:tbl", "a:tbl/a:tr", "a:tbl/(a:tr, a:tr, a:tr)"], fn: (self: any, request: any): any => {
        let expected_row_lst: any, rows: any, tbl: any, tbl_cxml: any;
        tbl_cxml = request.param;
        tbl = element(tbl_cxml);
        rows = $call(_RowCollection, tbl, null);
        expected_row_lst = tbl.findall(qn("a:tr"));
        return [rows, expected_row_lst];
      } },
    iter_fixture: { args: ["request"], params: ["a:tbl", "a:tbl/a:tr", "a:tbl/(a:tr, a:tr, a:tr)"], fn: (self: any, request: any): any => {
        let expected_row_lst: any, rows: any, tbl: any, tbl_cxml: any;
        tbl_cxml = request.param;
        tbl = element(tbl_cxml);
        rows = $call(_RowCollection, tbl, null);
        expected_row_lst = tbl.findall(qn("a:tr"));
        return [rows, expected_row_lst];
      } },
    len_fixture: { args: ["request"], params: [["a:tbl", 0], ["a:tbl/a:tr", 1], ["a:tbl/(a:tr, a:tr)", 2]], fn: (self: any, request: any): any => {
        let expected_len: any, rows: any, tbl_cxml: any;
        [tbl_cxml, expected_len] = $list(request.param);
        rows = $call(_RowCollection, element(tbl_cxml), null);
        return [rows, expected_len];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_how_many_rows_it_contains", args: ["len_fixture"], fn: (self: any, len_fixture: any): any => {
        let expected_count: any, rows: any;
        [rows, expected_count] = $list(len_fixture);
        $assertEqual($len(rows), expected_count);
      } },
    { name: "it_can_iterate_over_the_rows_it_contains", args: ["iter_fixture"], fn: (self: any, iter_fixture: any): any => {
        let count: any, expected_tr_lst: any, rows: any;
        [rows, expected_tr_lst] = $list(iter_fixture);
        count = 0;
        for (const [idx, row] of $list($enumerate(rows))) {
          $assert($isinstance(row, _Row));
          $assert($is(row._tr, $get(expected_tr_lst, idx)));
          count = $add(count, 1);
        }
        $assertEqual(count, $len(expected_tr_lst));
      } },
    { name: "it_supports_indexed_access", args: ["getitem_fixture"], fn: (self: any, getitem_fixture: any): any => {
        let expected_tr_lst: any, row: any, rows: any;
        [rows, expected_tr_lst] = $list(getitem_fixture);
        for (const [idx, tr] of $list($enumerate(expected_tr_lst))) {
          row = $get(rows, idx);
          $assert($isinstance(row, _Row));
          $assert($is(row._tr, tr));
        }
      } },
    { name: "it_raises_on_indexed_access_out_of_range", args: [], fn: (self: any): any => {
        let rows: any;
        rows = $call(_RowCollection, element("a:tbl/a:tr"), null);
        $raises(IndexError, () => {
          $get(rows, -1);
        });
        $raises(IndexError, () => {
          $get(rows, 9);
        });
      } },
  ],
};
pytestClass("Describe_RowCollection", Describe_RowCollection);

