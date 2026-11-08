// Machine-translated from python-pptx tests/chart/test_xlsx.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $b, $call, $cm1, $d, $enter, $enumerate, $exit, $is, $iter, $kw, $len, $list, $m, $raises, $slice, $t2, ANY, BytesIO, call, classMock, instanceMock, methodMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { BubbleChartData as BubbleChartData$imported, Categories as Categories$imported, CategoryChartData as CategoryChartData$imported, CategorySeriesData as CategorySeriesData$imported, XyChartData as XyChartData$imported } from "../../src/chart/data.js";
import { BubbleWorkbookWriter as BubbleWorkbookWriter$imported, CategoryWorkbookWriter as CategoryWorkbookWriter$imported, XyWorkbookWriter as XyWorkbookWriter$imported, _BaseWorkbookWriter as _BaseWorkbookWriter$imported } from "../../src/chart/xlsx.js";
import { Workbook as Workbook$imported, Worksheet as Worksheet$imported } from "../../src/xlsx/workbook.js";
// python binds imported names at import time; patches made later don't change them
const BubbleChartData = BubbleChartData$imported;
const Categories = Categories$imported;
const CategoryChartData = CategoryChartData$imported;
const CategorySeriesData = CategorySeriesData$imported;
const XyChartData = XyChartData$imported;
const BubbleWorkbookWriter = BubbleWorkbookWriter$imported;
const CategoryWorkbookWriter = CategoryWorkbookWriter$imported;
const XyWorkbookWriter = XyWorkbookWriter$imported;
const _BaseWorkbookWriter = _BaseWorkbookWriter$imported;
const Workbook = Workbook$imported;
const Worksheet = Worksheet$imported;
void $py;


const Describe_BaseWorkbookWriter: any = {
  fixtures: {
    open_fixture: { args: ["xlsx_file_", "workbook_", "worksheet_", "Workbook_"], fn: (self: any, xlsx_file_: any, workbook_: any, worksheet_: any, Workbook_: any): any => {
        let workbook_writer: any;
        workbook_writer = $call(_BaseWorkbookWriter, null);
        workbook_.addWorksheet.returnValue = worksheet_;
        return [workbook_writer, xlsx_file_, workbook_, worksheet_, Workbook_];
      } },
    populate_fixture: { args: [], fn: (self: any): any => {
        let workbook_writer: any;
        workbook_writer = $call(_BaseWorkbookWriter, null);
        return workbook_writer;
      } },
    BytesIO_: { args: ["request"], fn: (self: any, request: any): any => {
        return classMock(request, "pptx.fileio.BytesIO");
      } },
    Workbook_: { args: ["request", "workbook_"], fn: (self: any, request: any, workbook_: any): any => {
        return $call(classMock, request, "pptx.chart.xlsx.Workbook", $kw({return_value: workbook_}));
      } },
    workbook_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Workbook);
      } },
    worksheet_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Worksheet);
      } },
    xlsx_file_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, BytesIO);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_generate_a_chart_data_Excel_blob", args: ["request", "xlsx_file_", "workbook_", "worksheet_", "BytesIO_"], fn: (self: any, request: any, xlsx_file_: any, workbook_: any, worksheet_: any, BytesIO_: any): any => {
        let _open_worksheet_: any, _populate_worksheet_: any, workbook_writer: any, xlsx_blob: any;
        _populate_worksheet_ = methodMock(request, _BaseWorkbookWriter, "_populate_worksheet");
        _open_worksheet_ = methodMock(request, _BaseWorkbookWriter, "_open_worksheet");
        _open_worksheet_.returnValue.__enter__.returnValue = [workbook_, worksheet_];
        BytesIO_.returnValue = xlsx_file_;
        xlsx_file_.getvalue.returnValue = $b("xlsx-blob");
        workbook_writer = $call(_BaseWorkbookWriter, null);
        xlsx_blob = workbook_writer.xlsxBlob;
        _open_worksheet_.assertCalledOnceWith(workbook_writer, xlsx_file_);
        _populate_worksheet_.assertCalledOnceWith(workbook_writer, workbook_, worksheet_);
        $assertEqual(xlsx_blob, $b("xlsx-blob"));
      } },
    { name: "it_can_open_a_worksheet_in_a_context", args: ["open_fixture"], fn: (self: any, open_fixture: any): any => {
        let Workbook_: any, wb_writer: any, workbook: any, workbook_: any, worksheet: any, worksheet_: any, xlsx_file_: any;
        [wb_writer, xlsx_file_, workbook_, worksheet_, Workbook_] = $list(open_fixture);
        const $cm1 = wb_writer._openWorksheet(xlsx_file_);
        [workbook, worksheet] = $list($enter($cm1));
        try {
          Workbook_.assertCalledOnceWith(xlsx_file_, $d({in_memory: true}));
          workbook_.addWorksheet.assertCalledOnceWith();
          $assert($is(workbook, workbook_));
          $assert($is(worksheet, worksheet_));
        } finally {
          $exit($cm1);
        }
        workbook_.close.assertCalledOnceWith();
      } },
    { name: "it_raises_on_no_override_of_populate", args: ["populate_fixture"], fn: (self: any, populate_fixture: any): any => {
        let workbook_writer: any;
        workbook_writer = populate_fixture;
        $raises(NotImplementedError, () => {
          workbook_writer._populateWorksheet(null, null);
        });
      } },
  ],
};
pytestClass("Describe_BaseWorkbookWriter", Describe_BaseWorkbookWriter);

const DescribeCategoryWorkbookWriter: any = {
  fixtures: {
    categories_ref_fixture: { args: ["request", "chart_data_", "categories_"], params: [[1, 1, "Sheet1!$A$2:$A$2"], [1, 3, "Sheet1!$A$2:$A$4"], [2, 4, "Sheet1!$A$2:$B$5"], [3, 8, "Sheet1!$A$2:$C$9"]], fn: (self: any, request: any, chart_data_: any, categories_: any): any => {
        let depth: any, expected_value: any, leaf_count: any, workbook_writer: any;
        [depth, leaf_count, expected_value] = $list(request.param);
        workbook_writer = $call(CategoryWorkbookWriter, chart_data_);
        chart_data_.categories = categories_;
        [categories_.depth, categories_.leafCount] = $list([depth, leaf_count]);
        return [workbook_writer, expected_value];
      } },
    cat_ref_raises_fixture: { args: ["request", "chart_data_", "categories_"], fn: (self: any, request: any, chart_data_: any, categories_: any): any => {
        let workbook_writer: any;
        workbook_writer = $call(CategoryWorkbookWriter, chart_data_);
        chart_data_.categories = categories_;
        categories_.depth = 0;
        return workbook_writer;
      } },
    col_ref_fixture: { args: ["request"], params: [[2, "B"], [26, "Z"], [27, "AA"], [52, "AZ"], [53, "BA"], [676, "YZ"], [677, "ZA"], [702, "ZZ"], [703, "AAA"], [728, "AAZ"], [729, "ABA"], [1378, "AZZ"], [1379, "BAA"], [16384, "XFD"]], fn: (self: any, request: any): any => {
        let column_number: any, expected_value: any;
        [column_number, expected_value] = $list(request.param);
        return [column_number, expected_value];
      } },
    populate_fixture: { args: ["workbook_", "worksheet_", "_write_categories_", "_write_series_"], fn: (self: any, workbook_: any, worksheet_: any, _write_categories_: any, _write_series_: any): any => {
        let workbook_writer: any;
        workbook_writer = $call(CategoryWorkbookWriter, null);
        return [workbook_writer, workbook_, worksheet_];
      } },
    raises_fixture: { args: ["request"], params: [0, -1, 16385, 30433], fn: (self: any, request: any): any => {
        let column_number: any;
        column_number = request.param;
        return column_number;
      } },
    ser_name_ref_fixture: { args: ["request", "series_data_", "categories_"], params: [[1, 0, "Sheet1!$B$1"], [1, 3, "Sheet1!$E$1"], [3, 0, "Sheet1!$D$1"], [3, 3, "Sheet1!$G$1"]], fn: (self: any, request: any, series_data_: any, categories_: any): any => {
        let cat_depth: any, expected_value: any, series_index: any, workbook_writer: any;
        [cat_depth, series_index, expected_value] = $list(request.param);
        workbook_writer = $call(CategoryWorkbookWriter, null);
        series_data_.categories = categories_;
        categories_.depth = cat_depth;
        series_data_.index = series_index;
        return [workbook_writer, series_data_, expected_value];
      } },
    values_ref_fixture: { args: ["request", "series_data_", "categories_"], params: [[1, 0, 3, "Sheet1!$B$2:$B$4"], [1, 1, 3, "Sheet1!$C$2:$C$4"], [2, 0, 5, "Sheet1!$C$2:$C$6"], [3, 2, 7, "Sheet1!$F$2:$F$8"]], fn: (self: any, request: any, series_data_: any, categories_: any): any => {
        let cat_depth: any, expected_value: any, ser_idx: any, val_count: any, workbook_writer: any;
        [cat_depth, ser_idx, val_count, expected_value] = $list(request.param);
        workbook_writer = $call(CategoryWorkbookWriter, null);
        series_data_.categories = categories_;
        categories_.depth = cat_depth;
        series_data_.index = ser_idx;
        series_data_.__len__.returnValue = val_count;
        return [workbook_writer, series_data_, expected_value];
      } },
    write_cat_fixture: { args: ["worksheet_"], fn: (self: any, worksheet_: any): any => {
        let calls: any, col: any, level: any, num_format: any, workbook_writer: any;
        workbook_writer = $call(CategoryWorkbookWriter, null);
        [col, level] = $list([6, [[0, "Foo"], [1, "Bar"]]]);
        num_format = Object();
        calls = [call.setColumn(col, col, 10), call.write(1, col, "Foo", num_format), call.write(2, col, "Bar", num_format)];
        return [workbook_writer, worksheet_, col, level, num_format, calls];
      } },
    write_cats_fixture: { args: ["request", "chart_data_", "workbook_", "worksheet_", "categories_", "_write_cat_column_"], params: [[[[0, "Foo"], [1, "Bar"], [2, "Baz"]]], [[[0, "CA"], [1, "NV"], [2, "NY"], [3, "NJ"]], [[0, "WEST"], [2, "EAST"]]]], fn: (self: any, request: any, chart_data_: any, workbook_: any, worksheet_: any, categories_: any, _write_cat_column_: any): any => {
        let _depth: any, calls: any, levels: any, num_format: any, number_format: any, workbook_writer: any;
        levels = request.param;
        workbook_writer = $call(CategoryWorkbookWriter, chart_data_);
        const $t2 = "$#0.#";
        number_format = $t2;
        categories_.numberFormat = $t2;
        num_format = workbook_.addFormat.returnValue;
        _depth = $len(levels);
        calls = $list($enumerate(levels)).map(([idx, level]: any) => call(workbook_writer, worksheet_, ((_depth - idx) - 1), level, num_format));
        chart_data_.categories = categories_;
        categories_.depth = $len(levels);
        categories_.levels = levels;
        return [workbook_writer, workbook_, worksheet_, number_format, calls];
      } },
    write_sers_fixture: { args: ["request", "chart_data_", "workbook_", "worksheet_", "categories_"], fn: (self: any, request: any, chart_data_: any, workbook_: any, worksheet_: any, categories_: any): any => {
        let calls: any, num_format: any, ser: any, workbook_writer: any;
        workbook_writer = $call(CategoryWorkbookWriter, chart_data_);
        num_format = workbook_.addFormat.returnValue;
        calls = [call.write(0, 1, "S1"), call.writeColumn(1, 1, [42, 24], num_format)];
        ser = $call(instanceMock, request, CategorySeriesData, $kw({values: [42, 24]}));
        ser.name = "S1";
        chart_data_.categories = categories_;
        categories_.depth = 1;
        chart_data_.__iter__.returnValue = $iter([ser]);
        return [workbook_writer, workbook_, worksheet_, calls];
      } },
    categories_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Categories);
      } },
    chart_data_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, CategoryChartData);
      } },
    series_data_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, CategorySeriesData);
      } },
    workbook_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Workbook);
      } },
    worksheet_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Worksheet);
      } },
    _write_cat_column_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, CategoryWorkbookWriter, "_write_cat_column", $kw({autospec: true}));
      } },
    _write_categories_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, CategoryWorkbookWriter, "_write_categories", $kw({autospec: true}));
      } },
    _write_series_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, CategoryWorkbookWriter, "_write_series", $kw({autospec: true}));
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_the_categories_range_ref", args: ["categories_ref_fixture"], fn: (self: any, categories_ref_fixture: any): any => {
        let expected_value: any, workbook_writer: any;
        [workbook_writer, expected_value] = $list(categories_ref_fixture);
        $assertEqual(workbook_writer.categoriesRef, expected_value);
      } },
    { name: "it_raises_on_cat_ref_on_no_categories", args: ["cat_ref_raises_fixture"], fn: (self: any, cat_ref_raises_fixture: any): any => {
        let workbook_writer: any;
        workbook_writer = cat_ref_raises_fixture;
        $raises(ValueError, () => {
          workbook_writer.categoriesRef;
        });
      } },
    { name: "it_knows_the_ref_for_a_series_name", args: ["ser_name_ref_fixture"], fn: (self: any, ser_name_ref_fixture: any): any => {
        let expected_value: any, series_: any, workbook_writer: any;
        [workbook_writer, series_, expected_value] = $list(ser_name_ref_fixture);
        $assertEqual(workbook_writer.seriesNameRef(series_), expected_value);
      } },
    { name: "it_knows_the_values_range_ref", args: ["values_ref_fixture"], fn: (self: any, values_ref_fixture: any): any => {
        let expected_value: any, series_: any, workbook_writer: any;
        [workbook_writer, series_, expected_value] = $list(values_ref_fixture);
        $assertEqual(workbook_writer.valuesRef(series_), expected_value);
      } },
    { name: "it_can_populate_a_worksheet_with_chart_data", args: ["populate_fixture"], fn: (self: any, populate_fixture: any): any => {
        let workbook_: any, workbook_writer: any, worksheet_: any;
        [workbook_writer, workbook_, worksheet_] = $list(populate_fixture);
        workbook_writer._populateWorksheet(workbook_, worksheet_);
        workbook_writer._writeCategories.assertCalledOnceWith(workbook_writer, workbook_, worksheet_);
        workbook_writer._writeSeries.assertCalledOnceWith(workbook_writer, workbook_, worksheet_);
      } },
    { name: "it_calculates_a_column_reference_to_help", args: ["col_ref_fixture"], fn: (self: any, col_ref_fixture: any): any => {
        let col_num: any, column_reference: any, expected_value: any;
        [col_num, expected_value] = $list(col_ref_fixture);
        column_reference = CategoryWorkbookWriter._columnReference(col_num);
        $assertEqual(column_reference, expected_value);
      } },
    { name: "it_raises_on_col_number_out_of_range", args: ["raises_fixture"], fn: (self: any, raises_fixture: any): any => {
        let col_num: any;
        col_num = raises_fixture;
        $raises(ValueError, () => {
          CategoryWorkbookWriter._columnReference(col_num);
        });
      } },
    { name: "it_writes_categories_to_help", args: ["write_cats_fixture"], fn: (self: any, write_cats_fixture: any): any => {
        let calls: any, number_format: any, workbook_: any, workbook_writer: any, worksheet_: any;
        [workbook_writer, workbook_, worksheet_] = $list($slice(write_cats_fixture, null, 3, null));
        [number_format, calls] = $list($slice(write_cats_fixture, 3, null, null));
        workbook_writer._writeCategories(workbook_, worksheet_);
        workbook_.addFormat.assertCalledOnceWith($d({num_format: number_format}));
        $assertEqual(workbook_writer._writeCatColumn.callArgsList, calls);
      } },
    { name: "it_writes_a_category_column_to_help", args: ["write_cat_fixture"], fn: (self: any, write_cat_fixture: any): any => {
        let calls: any, col: any, level: any, num_format_: any, workbook_writer: any, worksheet_: any;
        [workbook_writer, worksheet_, col] = $list($slice(write_cat_fixture, null, 3, null));
        [level, num_format_, calls] = $list($slice(write_cat_fixture, 3, null, null));
        workbook_writer._writeCatColumn(worksheet_, col, level, num_format_);
        $assertEqual(worksheet_.mockCalls, calls);
      } },
    { name: "it_writes_series_to_help", args: ["write_sers_fixture"], fn: (self: any, write_sers_fixture: any): any => {
        let calls: any, workbook_: any, workbook_writer: any, worksheet_: any;
        [workbook_writer, workbook_, worksheet_, calls] = $list(write_sers_fixture);
        workbook_writer._writeSeries(workbook_, worksheet_);
        $assertEqual(worksheet_.mockCalls, calls);
      } },
  ],
};
pytestClass("DescribeCategoryWorkbookWriter", DescribeCategoryWorkbookWriter);

const DescribeBubbleWorkbookWriter: any = {
  fixtures: {
    populate_fixture: { args: ["workbook_", "worksheet_"], fn: (self: any, workbook_: any, worksheet_: any): any => {
        let chart_data: any, expected_calls: any, series_1: any, series_2: any, workbook_writer: any;
        chart_data = $call(BubbleChartData);
        series_1 = chart_data.addSeries("Series 1");
        for (const pt of $list([[1, 1.1, 10], [2, 2.2, 20]])) {
          $m(series_1, "addDataPoint", ...$list(pt));
        }
        series_2 = chart_data.addSeries("Series 2");
        for (const pt of $list([[3, 3.3, 30], [4, 4.4, 40]])) {
          $m(series_2, "addDataPoint", ...$list(pt));
        }
        workbook_writer = $call(BubbleWorkbookWriter, chart_data);
        expected_calls = [call.writeColumn(1, 0, [1, 2], ANY), call.write(0, 1, "Series 1"), call.writeColumn(1, 1, [1.1, 2.2], ANY), call.write(0, 2, "Size"), call.writeColumn(1, 2, [10, 20], ANY), call.writeColumn(5, 0, [3, 4], ANY), call.write(4, 1, "Series 2"), call.writeColumn(5, 1, [3.3, 4.4], ANY), call.write(4, 2, "Size"), call.writeColumn(5, 2, [30, 40], ANY)];
        return [workbook_writer, workbook_, worksheet_, expected_calls];
      } },
    workbook_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Workbook);
      } },
    worksheet_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Worksheet);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_populate_a_worksheet_with_chart_data", args: ["populate_fixture"], fn: (self: any, populate_fixture: any): any => {
        let expected_calls: any, workbook_: any, workbook_writer: any, worksheet_: any;
        [workbook_writer, workbook_, worksheet_, expected_calls] = $list(populate_fixture);
        workbook_writer._populateWorksheet(workbook_, worksheet_);
        $assertEqual(worksheet_.mockCalls, expected_calls);
      } },
  ],
};
pytestClass("DescribeBubbleWorkbookWriter", DescribeBubbleWorkbookWriter);

const DescribeXyWorkbookWriter: any = {
  fixtures: {
    populate_fixture: { args: ["workbook_", "worksheet_"], fn: (self: any, workbook_: any, worksheet_: any): any => {
        let chart_data: any, expected_calls: any, series_1: any, series_2: any, workbook_writer: any;
        chart_data = $call(XyChartData);
        series_1 = chart_data.addSeries("Series 1");
        for (const pt of $list([[1, 1.1], [2, 2.2]])) {
          $m(series_1, "addDataPoint", ...$list(pt));
        }
        series_2 = chart_data.addSeries("Series 2");
        for (const pt of $list([[3, 3.3], [4, 4.4]])) {
          $m(series_2, "addDataPoint", ...$list(pt));
        }
        workbook_writer = $call(XyWorkbookWriter, chart_data);
        expected_calls = [call.writeColumn(1, 0, [1, 2], ANY), call.write(0, 1, "Series 1"), call.writeColumn(1, 1, [1.1, 2.2], ANY), call.writeColumn(5, 0, [3, 4], ANY), call.write(4, 1, "Series 2"), call.writeColumn(5, 1, [3.3, 4.4], ANY)];
        return [workbook_writer, workbook_, worksheet_, expected_calls];
      } },
    workbook_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Workbook);
      } },
    worksheet_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Worksheet);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_populate_a_worksheet_with_chart_data", args: ["populate_fixture"], fn: (self: any, populate_fixture: any): any => {
        let expected_calls: any, workbook_: any, workbook_writer: any, worksheet_: any;
        [workbook_writer, workbook_, worksheet_, expected_calls] = $list(populate_fixture);
        workbook_writer._populateWorksheet(workbook_, worksheet_);
        $assertEqual(worksheet_.mockCalls, expected_calls);
      } },
  ],
};
pytestClass("DescribeXyWorkbookWriter", DescribeXyWorkbookWriter);

