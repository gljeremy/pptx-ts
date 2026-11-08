// Machine-translated from python-pptx tests/chart/test_data.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $append, $assert, $assertEqual, $call, $encode, $get, $index, $is, $isinstance, $iter, $kw, $list, $raises, $slice, $t1, $t2, $t3, $truthy, call, classMock, instanceMock, propertyMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { BubbleChartData as BubbleChartData$imported, BubbleDataPoint as BubbleDataPoint$imported, BubbleSeriesData as BubbleSeriesData$imported, Categories as Categories$imported, Category as Category$imported, CategoryChartData as CategoryChartData$imported, CategoryDataPoint as CategoryDataPoint$imported, CategorySeriesData as CategorySeriesData$imported, ChartData as ChartData$imported, XyChartData as XyChartData$imported, XyDataPoint as XyDataPoint$imported, XySeriesData as XySeriesData$imported, _BaseChartData as _BaseChartData$imported, _BaseDataPoint as _BaseDataPoint$imported, _BaseSeriesData as _BaseSeriesData$imported } from "../../src/chart/data.js";
import { CategoryWorkbookWriter as CategoryWorkbookWriter$imported } from "../../src/chart/xlsx.js";
import { XL_CHART_TYPE as XL_CHART_TYPE$imported } from "../../src/enum/chart.js";
// python binds imported names at import time; patches made later don't change them
const BubbleChartData = BubbleChartData$imported;
const BubbleDataPoint = BubbleDataPoint$imported;
const BubbleSeriesData = BubbleSeriesData$imported;
const Categories = Categories$imported;
const Category = Category$imported;
const CategoryChartData = CategoryChartData$imported;
const CategoryDataPoint = CategoryDataPoint$imported;
const CategorySeriesData = CategorySeriesData$imported;
const ChartData = ChartData$imported;
const XyChartData = XyChartData$imported;
const XyDataPoint = XyDataPoint$imported;
const XySeriesData = XySeriesData$imported;
const _BaseChartData = _BaseChartData$imported;
const _BaseDataPoint = _BaseDataPoint$imported;
const _BaseSeriesData = _BaseSeriesData$imported;
const CategoryWorkbookWriter = CategoryWorkbookWriter$imported;
const XL_CHART_TYPE = XL_CHART_TYPE$imported;
const date = $py.$datetime.date;
const datetime = $py.$datetime.datetime;
void $py;


const DescribeChartData: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_is_a_CategoryChartData_object", args: [], fn: (self: any): any => {
        $assert($isinstance($call(ChartData), CategoryChartData));
      } },
  ],
};
pytestClass("DescribeChartData", DescribeChartData);

const Describe_BaseChartData: any = {
  fixtures: {
    number_format_fixture: { args: ["request"], params: [[null, "General"], [42, 42]], fn: (self: any, request: any): any => {
        let argv: any, chart_data: any, expected_value: any, number_format: any;
        [number_format, expected_value] = $list(request.param);
        argv = ($truthy($is(number_format, null)) ? [] : [number_format]);
        chart_data = $call(_BaseChartData, ...$list(argv));
        return [chart_data, expected_value];
      } },
    ChartXmlWriter_: { args: ["request"], fn: (self: any, request: any): any => {
        let ChartXmlWriter_: any;
        ChartXmlWriter_ = classMock(request, "pptx.chart.data.ChartXmlWriter");
        ChartXmlWriter_.returnValue.xml = "ƒøØßår";
        return ChartXmlWriter_;
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_generate_chart_part_XML_for_its_data", args: ["ChartXmlWriter_"], fn: (self: any, ChartXmlWriter_: any): any => {
        let chart_data: any, xml_bytes: any;
        ChartXmlWriter_.returnValue.xml = "ƒøØßår";
        chart_data = $call(_BaseChartData);
        xml_bytes = chart_data.xmlBytes(XL_CHART_TYPE.PIE);
        ChartXmlWriter_.assertCalledOnceWith(XL_CHART_TYPE.PIE, chart_data);
        $assertEqual(xml_bytes, $encode("ƒøØßår", "utf-8"));
      } },
    { name: "it_knows_its_number_format", args: ["number_format_fixture"], fn: (self: any, number_format_fixture: any): any => {
        let chart_data: any, expected_value: any;
        [chart_data, expected_value] = $list(number_format_fixture);
        $assertEqual(chart_data.numberFormat, expected_value);
      } },
  ],
};
pytestClass("Describe_BaseChartData", Describe_BaseChartData);

const Describe_BaseSeriesData: any = {
  fixtures: {
    name_fixture: { args: ["request"], params: [["Tincture of Foo", "Tincture of Foo"], [null, ""]], fn: (self: any, request: any): any => {
        let expected_value: any, name: any;
        [name, expected_value] = $list(request.param);
        return [name, expected_value];
      } },
    number_format_fixture: { args: ["request", "chart_data_"], params: [["General", 42, 42], ["General", null, "General"]], fn: (self: any, request: any, chart_data_: any): any => {
        let expected_value: any, number_format: any, parent_number_format: any, series_data: any;
        [parent_number_format, number_format, expected_value] = $list(request.param);
        chart_data_.numberFormat = parent_number_format;
        series_data = $call(_BaseSeriesData, chart_data_, "Foobar", number_format);
        return [series_data, expected_value];
      } },
    chart_data_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _BaseChartData);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_name", args: ["name_fixture"], fn: (self: any, name_fixture: any): any => {
        let expected_value: any, name: any, name_arg: any, series_data: any;
        [name_arg, expected_value] = $list(name_fixture);
        series_data = $call(_BaseSeriesData, null, name_arg, null);
        name = series_data.name;
        $assertEqual(name, expected_value);
      } },
    { name: "it_knows_its_number_format", args: ["number_format_fixture"], fn: (self: any, number_format_fixture: any): any => {
        let expected_value: any, series_data: any;
        [series_data, expected_value] = $list(number_format_fixture);
        $assertEqual(series_data.numberFormat, expected_value);
      } },
  ],
};
pytestClass("Describe_BaseSeriesData", Describe_BaseSeriesData);

const Describe_BaseDataPoint: any = {
  fixtures: {
    number_format_fixture: { args: ["request", "series_data_"], params: [[42, 24, 24], [42, null, 42]], fn: (self: any, request: any, series_data_: any): any => {
        let data_point: any, expected_value: any, number_format: any, parent_number_format: any;
        [parent_number_format, number_format, expected_value] = $list(request.param);
        series_data_.numberFormat = parent_number_format;
        data_point = $call(_BaseDataPoint, series_data_, number_format);
        return [data_point, expected_value];
      } },
    series_data_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _BaseSeriesData);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_number_format", args: ["number_format_fixture"], fn: (self: any, number_format_fixture: any): any => {
        let data_point: any, expected_value: any;
        [data_point, expected_value] = $list(number_format_fixture);
        $assertEqual(data_point.numberFormat, expected_value);
      } },
  ],
};
pytestClass("Describe_BaseDataPoint", Describe_BaseDataPoint);

const DescribeCategoryChartData: any = {
  fixtures: {
    add_cat_fixture: { args: ["categories_prop_", "categories_", "category_"], fn: (self: any, categories_prop_: any, categories_: any, category_: any): any => {
        let chart_data: any, name: any;
        chart_data = $call(CategoryChartData);
        name = "foobar";
        categories_.addCategory.returnValue = category_;
        return [chart_data, name, categories_, category_];
      } },
    add_ser_fixture: { args: ["CategorySeriesData_", "series_"], fn: (self: any, CategorySeriesData_: any, series_: any): any => {
        let calls: any, chart_data: any, name: any, number_format: any, values: any;
        chart_data = $call(CategoryChartData);
        [name, values, number_format] = $list(["foobar", $iter([1, 2, 3]), "0.0"]);
        calls = [call(1), call(2), call(3)];
        return [chart_data, name, values, number_format, CategorySeriesData_, calls, series_];
      } },
    categories_fixture: { args: ["Categories_", "categories_"], fn: (self: any, Categories_: any, categories_: any): any => {
        let chart_data: any;
        chart_data = $call(CategoryChartData);
        return [chart_data, Categories_, categories_];
      } },
    categories_ref_fixture: { args: ["_workbook_writer_prop_", "workbook_writer_"], fn: (self: any, _workbook_writer_prop_: any, workbook_writer_: any): any => {
        let categories_ref: any, chart_data: any, expected_value: any;
        chart_data = $call(CategoryChartData);
        const $t1 = "Sheet42!$G$24";
        expected_value = $t1;
        categories_ref = $t1;
        workbook_writer_.categoriesRef = categories_ref;
        return [chart_data, expected_value];
      } },
    categories_set_fixture: { args: ["Categories_", "categories_"], fn: (self: any, Categories_: any, categories_: any): any => {
        let calls: any, chart_data: any, names: any;
        chart_data = $call(CategoryChartData);
        names = $iter(["a", "b", "c"]);
        calls = [call("a"), call("b"), call("c")];
        return [chart_data, names, Categories_, categories_, calls];
      } },
    values_ref_fixture: { args: ["_workbook_writer_prop_", "workbook_writer_", "series_"], fn: (self: any, _workbook_writer_prop_: any, workbook_writer_: any, series_: any): any => {
        let chart_data: any, values_ref_: any;
        chart_data = $call(CategoryChartData);
        values_ref_ = "Sheet1!$V$24";
        workbook_writer_.valuesRef.returnValue = values_ref_;
        return [chart_data, workbook_writer_, series_, values_ref_];
      } },
    Categories_: { args: ["request", "categories_"], fn: (self: any, request: any, categories_: any): any => {
        return $call(classMock, request, "pptx.chart.data.Categories", $kw({return_value: categories_}));
      } },
    CategorySeriesData_: { args: ["request", "series_"], fn: (self: any, request: any, series_: any): any => {
        return $call(classMock, request, "pptx.chart.data.CategorySeriesData", $kw({return_value: series_}));
      } },
    categories_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Categories);
      } },
    categories_prop_: { args: ["request", "categories_"], fn: (self: any, request: any, categories_: any): any => {
        return $call(propertyMock, request, CategoryChartData, "categories", $kw({return_value: categories_}));
      } },
    category_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Category);
      } },
    series_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, CategorySeriesData);
      } },
    workbook_writer_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, CategoryWorkbookWriter);
      } },
    _workbook_writer_prop_: { args: ["request", "workbook_writer_"], fn: (self: any, request: any, workbook_writer_: any): any => {
        return $call(propertyMock, request, CategoryChartData, "_workbook_writer", $kw({return_value: workbook_writer_}));
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_is_a__BaseChartData_object", args: [], fn: (self: any): any => {
        $assert($isinstance($call(CategoryChartData), _BaseChartData));
      } },
    { name: "it_knows_the_categories_range_ref", args: ["categories_ref_fixture"], fn: (self: any, categories_ref_fixture: any): any => {
        let chart_data: any, expected_value: any;
        [chart_data, expected_value] = $list(categories_ref_fixture);
        $assertEqual(chart_data.categoriesRef, expected_value);
      } },
    { name: "it_knows_the_values_range_ref_for_a_series", args: ["values_ref_fixture"], fn: (self: any, values_ref_fixture: any): any => {
        let chart_data: any, series_: any, values_ref: any, values_ref_: any, workbook_writer_: any;
        [chart_data, workbook_writer_, series_, values_ref_] = $list(values_ref_fixture);
        values_ref = chart_data.valuesRef(series_);
        workbook_writer_.valuesRef.assertCalledOnceWith(series_);
        $assert($is(values_ref, values_ref_));
      } },
    { name: "it_provides_access_to_its_categories", args: ["categories_fixture"], fn: (self: any, categories_fixture: any): any => {
        let Categories_: any, categories: any, categories_: any, chart_data: any;
        [chart_data, Categories_, categories_] = $list(categories_fixture);
        categories = chart_data.categories;
        Categories_.assertCalledOnceWith();
        $assert($is(categories, categories_));
      } },
    { name: "it_can_add_a_category", args: ["add_cat_fixture"], fn: (self: any, add_cat_fixture: any): any => {
        let categories_: any, category: any, category_: any, chart_data: any, name: any;
        [chart_data, name, categories_, category_] = $list(add_cat_fixture);
        category = chart_data.addCategory(name);
        categories_.addCategory.assertCalledOnceWith(name);
        $assert($is(category, category_));
      } },
    { name: "it_can_add_a_series", args: ["add_ser_fixture"], fn: (self: any, add_ser_fixture: any): any => {
        let CategorySeriesData_: any, calls: any, chart_data: any, name: any, number_format: any, series: any, series_: any, values: any;
        [chart_data, name, values, number_format] = $list($slice(add_ser_fixture, null, 4, null));
        [CategorySeriesData_, calls, series_] = $list($slice(add_ser_fixture, 4, null, null));
        series = chart_data.addSeries(name, values, number_format);
        CategorySeriesData_.assertCalledOnceWith(chart_data, name, number_format);
        $assert($is($get(chart_data, -1), series));
        $assertEqual(series.addDataPoint.callArgsList, calls);
        $assert($is(series, series_));
      } },
    { name: "it_can_set_its_categories", args: ["categories_set_fixture"], fn: (self: any, categories_set_fixture: any): any => {
        let Categories_: any, calls: any, categories_: any, chart_data: any, names: any;
        [chart_data, names, Categories_, categories_, calls] = $list(categories_set_fixture);
        chart_data.categories = names;
        Categories_.assertCalledOnceWith();
        $assertEqual(categories_.addCategory.callArgsList, calls);
        $assert($is(chart_data._categories, categories_));
      } },
  ],
};
pytestClass("DescribeCategoryChartData", DescribeCategoryChartData);

const DescribeCategories: any = {
  fixtures: {
    add_fixture: { args: ["Category_", "category_"], fn: (self: any, Category_: any, category_: any): any => {
        let categories: any, name: any;
        categories = $call(Categories);
        name = "foobar";
        return [categories, name, Category_, category_];
      } },
    are_dates_fixture: { args: ["request"], params: [[[], false], [["foo", "bar"], false], [[1, 2], false], [[1.2, 2.3], false], [[date(2016, 12, 21)], true]], fn: (self: any, request: any): any => {
        let categories: any, expected_value: any, labels: any;
        [labels, expected_value] = $list(request.param);
        categories = $call(Categories);
        for (const label of $list(labels)) {
          categories.addCategory(label);
        }
        return [categories, expected_value];
      } },
    are_numeric_fixture: { args: ["request"], params: [[[], false], [["foo", "bar"], false], [[1, 2], true], [[1.2, 2.3], true], [[date(2016, 12, 21)], true]], fn: (self: any, request: any): any => {
        let categories: any, expected_value: any, labels: any;
        [labels, expected_value] = $list(request.param);
        categories = $call(Categories);
        for (const label of $list(labels)) {
          categories.addCategory(label);
        }
        return [categories, expected_value];
      } },
    depth_fixture: { args: ["request"], params: [[[], 0], [[1], 1], [[3], 3], [[1, 1, 1], 1], [[3, 3, 3], 3]], fn: (self: any, request: any): any => {
        let categories: any, depths: any, expected_value: any;
        [depths, expected_value] = $list(request.param);
        categories = $call(Categories);
        for (const depth of $list(depths)) {
          $append(categories._categories, $call(instanceMock, request, Category, $kw({depth: depth})));
        }
        return [categories, expected_value];
      } },
    depth_raises_fixture: { args: ["request"], fn: (self: any, request: any): any => {
        let categories: any;
        categories = $call(Categories);
        for (const depth of $list([1, 2, 1])) {
          $append(categories._categories, $call(instanceMock, request, Category, $kw({depth: depth})));
        }
        return categories;
      } },
    index_fixture: { args: ["request"], fn: (self: any, request: any): any => {
        let categories: any, categories_: any, category_: any, expected_value: any;
        categories = $call(Categories);
        categories_ = [$call(instanceMock, request, Category, $kw({leaf_count: 3})), $call(instanceMock, request, Category, $kw({leaf_count: 6})), $call(instanceMock, request, Category, $kw({leaf_count: 9}))];
        category_ = $get(categories_, 1);
        expected_value = 3;
        categories._categories = categories_;
        return [categories, category_, expected_value];
      } },
    leaf_fixture: { args: ["request"], params: [[[], 0], [[1], 1], [[1, 2, 3], 6]], fn: (self: any, request: any): any => {
        let categories: any, expected_value: any, leaf_counts: any;
        [leaf_counts, expected_value] = $list(request.param);
        categories = $call(Categories);
        for (const leaf_count of $list(leaf_counts)) {
          $append(categories._categories, $call(instanceMock, request, Category, $kw({leaf_count: leaf_count})));
        }
        return [categories, expected_value];
      } },
    levels_fixture: { args: ["request"], params: [[[[0, "a", []], [1, "b", []]], [[[0, "a"], [1, "b"]]]], [[[0, "WEST", [[0, "CA", []], [1, "LA", []]]], [2, "EAST", [[2, "NY", []], [3, "NJ", []]]]], [[[0, "CA"], [1, "LA"], [2, "NY"], [3, "NJ"]], [[0, "WEST"], [2, "EAST"]]]]], fn: function* (self: any, request: any): any {
        let cat_data: any, categories: any, expected_value: any;
        [cat_data, expected_value] = $list(request.param);
        categories = $call(Categories);
        const iter_cats = function* (cat_tree: any): any {
          let category_: any;
          for (const [idx, cat_label, sub_cats] of $list(cat_tree)) {
            category_ = $call(instanceMock, request, Category, $kw({idx: idx}));
            category_.label = cat_label;
            category_.subCategories = $list(iter_cats(sub_cats));
            yield category_;
          }
        };
        categories._categories = $list(iter_cats(cat_data));
        return [categories, expected_value];
      } },
    number_format_get_fixture: { args: ["request"], params: [[null, null, null, "General"], [null, "foo", null, "General"], [null, "foo", "bar", "General"], [null, date(2016, 12, 22), null, "yyyy\\-mm\\-dd"], [null, date(2016, 12, 22), "foo", "General"], ["#0", 42.24, null, "#0"]], fn: (self: any, request: any): any => {
        let cat: any, categories: any, category: any, expected_value: any, number_format: any, subcat: any;
        [number_format, cat, subcat, expected_value] = $list(request.param);
        categories = $call(Categories);
        if ($truthy(!$is(cat, null))) {
          category = categories.addCategory(cat);
          if ($truthy(!$is(subcat, null))) {
            category.addSubCategory(subcat);
          }
        }
        if ($truthy(!$is(number_format, null))) {
          categories._numberFormat = number_format;
        }
        return [categories, expected_value];
      } },
    number_format_set_fixture: { args: ["request"], params: [["0.0", "0.0"], [null, "General"]], fn: (self: any, request: any): any => {
        let categories: any, expected_value: any, new_value: any;
        [new_value, expected_value] = $list(request.param);
        categories = $call(Categories);
        return [categories, new_value, expected_value];
      } },
    Category_: { args: ["request", "category_"], fn: (self: any, request: any, category_: any): any => {
        return $call(classMock, request, "pptx.chart.data.Category", $kw({return_value: category_}));
      } },
    category_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Category);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_when_its_categories_are_numeric", args: ["are_numeric_fixture"], fn: (self: any, are_numeric_fixture: any): any => {
        let categories: any, expected_value: any;
        [categories, expected_value] = $list(are_numeric_fixture);
        $assertEqual(categories.areNumeric, expected_value);
      } },
    { name: "it_knows_when_its_categories_are_dates", args: ["are_dates_fixture"], fn: (self: any, are_dates_fixture: any): any => {
        let categories: any, expected_value: any;
        [categories, expected_value] = $list(are_dates_fixture);
        $assertEqual(categories.areDates, expected_value);
      } },
    { name: "it_knows_the_category_hierarchy_depth", args: ["depth_fixture"], fn: (self: any, depth_fixture: any): any => {
        let categories: any, expected_value: any;
        [categories, expected_value] = $list(depth_fixture);
        $assertEqual(categories.depth, expected_value);
      } },
    { name: "it_knows_the_idx_of_a_category", args: ["index_fixture"], fn: (self: any, index_fixture: any): any => {
        let categories: any, category_: any, expected_value: any;
        [categories, category_, expected_value] = $list(index_fixture);
        $assertEqual($index(categories, category_), expected_value);
      } },
    { name: "it_knows_its_leaf_category_count", args: ["leaf_fixture"], fn: (self: any, leaf_fixture: any): any => {
        let categories: any, expected_value: any;
        [categories, expected_value] = $list(leaf_fixture);
        $assertEqual(categories.leafCount, expected_value);
      } },
    { name: "it_knows_its_levels", args: ["levels_fixture"], fn: (self: any, levels_fixture: any): any => {
        let categories: any, expected_value: any;
        [categories, expected_value] = $list(levels_fixture);
        $assertEqual($list(categories.levels), expected_value);
      } },
    { name: "it_knows_its_number_format", args: ["number_format_get_fixture"], fn: (self: any, number_format_get_fixture: any): any => {
        let categories: any, expected_value: any;
        [categories, expected_value] = $list(number_format_get_fixture);
        $assertEqual(categories.numberFormat, expected_value);
      } },
    { name: "it_can_change_its_number_format", args: ["number_format_set_fixture"], fn: (self: any, number_format_set_fixture: any): any => {
        let categories: any, expected_value: any, new_value: any;
        [categories, new_value, expected_value] = $list(number_format_set_fixture);
        categories.numberFormat = new_value;
        $assertEqual(categories.numberFormat, expected_value);
      } },
    { name: "it_raises_on_category_depth_not_uniform", args: ["depth_raises_fixture"], fn: (self: any, depth_raises_fixture: any): any => {
        let categories: any;
        categories = depth_raises_fixture;
        $raises(ValueError, () => {
          categories.depth;
        });
      } },
    { name: "it_can_add_a_category", args: ["add_fixture"], fn: (self: any, add_fixture: any): any => {
        let Category_: any, categories: any, category: any, category_: any, name: any;
        [categories, name, Category_, category_] = $list(add_fixture);
        category = categories.addCategory(name);
        Category_.assertCalledOnceWith(name, categories);
        $assert($is($get(categories._categories, -1), category));
        $assert($is(category, category_));
      } },
  ],
};
pytestClass("DescribeCategories", DescribeCategories);

const DescribeCategory: any = {
  fixtures: {
    add_sub_fixture: { args: ["request", "category_"], fn: (self: any, request: any, category_: any): any => {
        let Category_: any, category: any, name: any;
        category = $call(Category, null, null);
        name = "foobar";
        Category_ = $call(classMock, request, "pptx.chart.data.Category", $kw({return_value: category_}));
        return [category, name, Category_, category_];
      } },
    depth_fixture: { args: ["request"], params: [[[], 1], [[1], 2], [[1, 1, 1], 2], [[2, 2, 2], 3]], fn: (self: any, request: any): any => {
        let category: any, depths: any, expected_value: any;
        [depths, expected_value] = $list(request.param);
        category = $call(Category, null, null);
        for (const depth of $list(depths)) {
          $append(category._subCategories, $call(instanceMock, request, Category, $kw({depth: depth})));
        }
        return [category, expected_value];
      } },
    depth_raises_fixture: { args: ["request"], fn: (self: any, request: any): any => {
        let category: any;
        category = $call(Category, null, null);
        for (const depth of $list([1, 2, 1])) {
          $append(category._subCategories, $call(instanceMock, request, Category, $kw({depth: depth})));
        }
        return category;
      } },
    excel_date_fixture: { args: ["request"], params: [[date(2016, 12, 22), false, 42726], [date(1999, 12, 31), true, 35063], [datetime(1900, 2, 28), false, 59], [datetime(1990, 9, 1), true, 31655]], fn: (self: any, request: any): any => {
        let category: any, date_1904: any, expected_value: any, label: any;
        [label, date_1904, expected_value] = $list(request.param);
        category = $call(Category, label, null);
        return [category, date_1904, expected_value];
      } },
    idx_fixture: { args: ["request"], params: [Categories, Category], fn: (self: any, request: any): any => {
        let category: any, idx_: any, parent_: any, parent_cls: any;
        parent_cls = request.param;
        parent_ = instanceMock(request, parent_cls);
        category = $call(Category, null, parent_);
        const $t2 = 42;
        parent_.index.returnValue = $t2;
        idx_ = $t2;
        return [category, parent_, idx_];
      } },
    index_fixture: { args: ["request", "categories_"], fn: (self: any, request: any, categories_: any): any => {
        let category: any, expected_value: any, sub_categories_: any, sub_category_: any;
        category = $call(Category, null, categories_);
        sub_categories_ = [$call(instanceMock, request, Category, $kw({leaf_count: 2})), $call(instanceMock, request, Category, $kw({leaf_count: 4})), $call(instanceMock, request, Category, $kw({leaf_count: 6}))];
        sub_category_ = $get(sub_categories_, 1);
        expected_value = 6;
        categories_.index.returnValue = 4;
        category._subCategories = sub_categories_;
        return [category, sub_category_, expected_value];
      } },
    label_fixture: { args: ["request"], params: [["Able the Label", "Able the Label"], [null, ""]], fn: (self: any, request: any): any => {
        let expected_value: any, label: any;
        [label, expected_value] = $list(request.param);
        return [label, expected_value];
      } },
    leaf_fixture: { args: ["request"], params: [[[], 1], [[1], 1], [[1, 2, 3], 6]], fn: (self: any, request: any): any => {
        let category: any, expected_value: any, leaf_counts: any;
        [leaf_counts, expected_value] = $list(request.param);
        category = $call(Category, null, null);
        for (const leaf_count of $list(leaf_counts)) {
          $append(category._subCategories, $call(instanceMock, request, Category, $kw({leaf_count: leaf_count})));
        }
        return [category, expected_value];
      } },
    num_str_fixture: { args: ["request"], params: [[date(2016, 12, 23), false, "42727.0"], [42.24, false, "42.24"], [42, false, "42"], ["foobar", false, "foobar"]], fn: (self: any, request: any): any => {
        let category: any, date_1904: any, expected_value: any, label: any;
        [label, date_1904, expected_value] = $list(request.param);
        category = $call(Category, label, null);
        return [category, date_1904, expected_value];
      } },
    subs_fixture: { args: [], fn: (self: any): any => {
        let category: any, sub_categories_: any;
        sub_categories_ = [42, 24];
        category = $call(Category, null, null);
        category._subCategories = sub_categories_;
        return [category, sub_categories_];
      } },
    categories_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Categories);
      } },
    category_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Category);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_depth", args: ["depth_fixture"], fn: (self: any, depth_fixture: any): any => {
        let category: any, expected_value: any;
        [category, expected_value] = $list(depth_fixture);
        $assertEqual(category.depth, expected_value);
      } },
    { name: "it_knows_its_idx", args: ["idx_fixture"], fn: (self: any, idx_fixture: any): any => {
        let category: any, idx: any, idx_: any, parent_: any;
        [category, parent_, idx_] = $list(idx_fixture);
        idx = category.idx;
        parent_.index.assertCalledOnceWith(category);
        $assertEqual(idx, idx_);
      } },
    { name: "it_knows_the_index_of_a_sub_category", args: ["index_fixture"], fn: (self: any, index_fixture: any): any => {
        let category: any, expected_value: any, index: any, sub_category_: any;
        [category, sub_category_, expected_value] = $list(index_fixture);
        index = $index(category, sub_category_);
        category._parent.index.assertCalledOnceWith(category);
        $assertEqual(index, expected_value);
      } },
    { name: "it_knows_its_leaf_category_count", args: ["leaf_fixture"], fn: (self: any, leaf_fixture: any): any => {
        let category: any, expected_value: any;
        [category, expected_value] = $list(leaf_fixture);
        $assertEqual(category.leafCount, expected_value);
      } },
    { name: "it_raises_on_depth_not_uniform", args: ["depth_raises_fixture"], fn: (self: any, depth_raises_fixture: any): any => {
        let category: any;
        category = depth_raises_fixture;
        $raises(ValueError, () => {
          category.depth;
        });
      } },
    { name: "it_knows_its_label", args: ["label_fixture"], fn: (self: any, label_fixture: any): any => {
        let category: any, expected_value: any, label: any, label_arg: any;
        [label_arg, expected_value] = $list(label_fixture);
        category = $call(Category, label_arg, null);
        label = category.label;
        $assertEqual(label, expected_value);
      } },
    { name: "it_knows_its_numeric_string_value", args: ["num_str_fixture"], fn: (self: any, num_str_fixture: any): any => {
        let category: any, date_1904: any, expected_value: any;
        [category, date_1904, expected_value] = $list(num_str_fixture);
        $assertEqual(category.numericStrVal(date_1904), expected_value);
      } },
    { name: "it_provides_access_to_its_sub_categories", args: ["subs_fixture"], fn: (self: any, subs_fixture: any): any => {
        let category: any, sub_categories_: any;
        [category, sub_categories_] = $list(subs_fixture);
        $assert($is(category.subCategories, sub_categories_));
      } },
    { name: "it_can_add_a_sub_category", args: ["add_sub_fixture"], fn: (self: any, add_sub_fixture: any): any => {
        let Category_: any, category: any, name: any, sub_category: any, sub_category_: any;
        [category, name, Category_, sub_category_] = $list(add_sub_fixture);
        sub_category = category.addSubCategory(name);
        Category_.assertCalledOnceWith(name, category);
        $assert($is($get(category._subCategories, -1), sub_category));
        $assert($is(sub_category, sub_category_));
      } },
    { name: "it_calculates_an_excel_date_number_to_help", args: ["excel_date_fixture"], fn: (self: any, excel_date_fixture: any): any => {
        let category: any, date_1904: any, date_number: any, expected_value: any;
        [category, date_1904, expected_value] = $list(excel_date_fixture);
        date_number = category._excelDateNumber(date_1904);
        $assertEqual(date_number, expected_value);
      } },
  ],
};
pytestClass("DescribeCategory", DescribeCategory);

const DescribeCategorySeriesData: any = {
  fixtures: {
    add_fixture: { args: ["request", "CategoryDataPoint_", "data_point_"], fn: (self: any, request: any, CategoryDataPoint_: any, data_point_: any): any => {
        let number_format: any, series_data: any, value: any;
        series_data = $call(CategorySeriesData, null, null, null);
        [value, number_format] = $list([42, "0.0"]);
        return [series_data, value, number_format, CategoryDataPoint_, data_point_];
      } },
    categories_fixture: { args: ["chart_data_", "categories_"], fn: (self: any, chart_data_: any, categories_: any): any => {
        let series_data: any;
        series_data = $call(CategorySeriesData, chart_data_, null, null);
        chart_data_.categories = categories_;
        return [series_data, categories_];
      } },
    categories_ref_fixture: { args: ["chart_data_"], fn: (self: any, chart_data_: any): any => {
        let categories_ref: any, expected_value: any, series_data: any;
        series_data = $call(CategorySeriesData, chart_data_, null, null);
        const $t3 = "Sheet1!$F$42";
        expected_value = $t3;
        categories_ref = $t3;
        chart_data_.categoriesRef = categories_ref;
        return [series_data, expected_value];
      } },
    values_fixture: { args: ["request"], fn: (self: any, request: any): any => {
        let expected_values: any, series_data: any;
        series_data = $call(CategorySeriesData, null, null, null);
        expected_values = [1, 2, 3];
        for (const value of $list(expected_values)) {
          $append(series_data._dataPoints, $call(instanceMock, request, CategoryDataPoint, $kw({value: value})));
        }
        return [series_data, expected_values];
      } },
    values_ref_fixture: { args: ["chart_data_"], fn: (self: any, chart_data_: any): any => {
        let series_data: any, values_ref_: any;
        series_data = $call(CategorySeriesData, chart_data_, null, null);
        values_ref_ = "Sheet1!$V$42";
        chart_data_.valuesRef.returnValue = values_ref_;
        return [series_data, chart_data_, values_ref_];
      } },
    CategoryDataPoint_: { args: ["request", "data_point_"], fn: (self: any, request: any, data_point_: any): any => {
        return $call(classMock, request, "pptx.chart.data.CategoryDataPoint", $kw({return_value: data_point_}));
      } },
    categories_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Categories);
      } },
    chart_data_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, CategoryChartData);
      } },
    data_point_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, CategoryDataPoint);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_the_categories_range_ref", args: ["categories_ref_fixture"], fn: (self: any, categories_ref_fixture: any): any => {
        let expected_value: any, series_data: any;
        [series_data, expected_value] = $list(categories_ref_fixture);
        $assertEqual(series_data.categoriesRef, expected_value);
      } },
    { name: "it_knows_its_values", args: ["values_fixture"], fn: (self: any, values_fixture: any): any => {
        let expected_values: any, series_data: any;
        [series_data, expected_values] = $list(values_fixture);
        $assertEqual(series_data.values, expected_values);
      } },
    { name: "it_knows_its_values_range_ref", args: ["values_ref_fixture"], fn: (self: any, values_ref_fixture: any): any => {
        let chart_data_: any, series_data: any, values_ref: any, values_ref_: any;
        [series_data, chart_data_, values_ref_] = $list(values_ref_fixture);
        values_ref = series_data.valuesRef;
        chart_data_.valuesRef.assertCalledOnceWith(series_data);
        $assert($is(values_ref, values_ref_));
      } },
    { name: "it_provides_access_to_the_chart_categories", args: ["categories_fixture"], fn: (self: any, categories_fixture: any): any => {
        let categories_: any, series_data: any;
        [series_data, categories_] = $list(categories_fixture);
        $assert($is(series_data.categories, categories_));
      } },
    { name: "it_can_add_a_data_point", args: ["add_fixture"], fn: (self: any, add_fixture: any): any => {
        let CategoryDataPoint_: any, data_point: any, data_point_: any, number_format: any, series_data: any, value: any;
        [series_data, value, number_format] = $list($slice(add_fixture, null, 3, null));
        [CategoryDataPoint_, data_point_] = $list($slice(add_fixture, 3, null, null));
        data_point = series_data.addDataPoint(value, number_format);
        CategoryDataPoint_.assertCalledOnceWith(series_data, value, number_format);
        $assert($is($get(series_data, -1), data_point_));
        $assert($is(data_point, data_point_));
      } },
  ],
};
pytestClass("DescribeCategorySeriesData", DescribeCategorySeriesData);

const DescribeBubbleChartData: any = {
  fixtures: {
    add_series_fixture: { args: ["request", "BubbleSeriesData_", "series_data_"], fn: (self: any, request: any, BubbleSeriesData_: any, series_data_: any): any => {
        let chart_data: any, name: any;
        chart_data = $call(BubbleChartData);
        name = "Series Name";
        return [chart_data, name, BubbleSeriesData_, series_data_];
      } },
    BubbleSeriesData_: { args: ["request", "series_data_"], fn: (self: any, request: any, series_data_: any): any => {
        return $call(classMock, request, "pptx.chart.data.BubbleSeriesData", $kw({return_value: series_data_}));
      } },
    series_data_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, BubbleSeriesData);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_add_a_series", args: ["add_series_fixture"], fn: (self: any, add_series_fixture: any): any => {
        let BubbleSeriesData_: any, chart_data: any, name: any, series_data: any, series_data_: any;
        [chart_data, name, BubbleSeriesData_, series_data_] = $list(add_series_fixture);
        series_data = chart_data.addSeries(name);
        BubbleSeriesData_.assertCalledOnceWith(chart_data, name, null);
        $assert($is($get(chart_data, -1), series_data_));
        $assert($is(series_data, series_data_));
      } },
  ],
};
pytestClass("DescribeBubbleChartData", DescribeBubbleChartData);

const DescribeXyChartData: any = {
  fixtures: {
    add_series_fixture: { args: ["request", "XySeriesData_", "series_data_"], fn: (self: any, request: any, XySeriesData_: any, series_data_: any): any => {
        let chart_data: any, label: any;
        chart_data = $call(XyChartData);
        label = "Series Label";
        return [chart_data, label, XySeriesData_, series_data_];
      } },
    XySeriesData_: { args: ["request", "series_data_"], fn: (self: any, request: any, series_data_: any): any => {
        return $call(classMock, request, "pptx.chart.data.XySeriesData", $kw({return_value: series_data_}));
      } },
    series_data_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, XySeriesData);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_is_a__BaseChartData_object", args: [], fn: (self: any): any => {
        $assert($isinstance($call(XyChartData), _BaseChartData));
      } },
    { name: "it_can_add_a_series", args: ["add_series_fixture"], fn: (self: any, add_series_fixture: any): any => {
        let XySeriesData_: any, chart_data: any, label: any, series_data: any, series_data_: any;
        [chart_data, label, XySeriesData_, series_data_] = $list(add_series_fixture);
        series_data = chart_data.addSeries(label);
        XySeriesData_.assertCalledOnceWith(chart_data, label, null);
        $assert($is($get(chart_data, -1), series_data_));
        $assert($is(series_data, series_data_));
      } },
  ],
};
pytestClass("DescribeXyChartData", DescribeXyChartData);

const DescribeBubbleSeriesData: any = {
  fixtures: {
    add_data_point_fixture: { args: ["request", "BubbleDataPoint_", "data_point_"], fn: (self: any, request: any, BubbleDataPoint_: any, data_point_: any): any => {
        let series_data: any, size: any, x: any, y: any;
        series_data = $call(BubbleSeriesData, null, null, null);
        [x, y, size] = $list([42, 24, 17]);
        return [series_data, x, y, size, BubbleDataPoint_, data_point_];
      } },
    BubbleDataPoint_: { args: ["request", "data_point_"], fn: (self: any, request: any, data_point_: any): any => {
        return $call(classMock, request, "pptx.chart.data.BubbleDataPoint", $kw({return_value: data_point_}));
      } },
    data_point_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, BubbleDataPoint);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_add_a_data_point", args: ["add_data_point_fixture"], fn: (self: any, add_data_point_fixture: any): any => {
        let BubbleDataPoint_: any, data_point: any, data_point_: any, series_data: any, size: any, x: any, y: any;
        [series_data, x, y, size, BubbleDataPoint_, data_point_] = $list(add_data_point_fixture);
        data_point = series_data.addDataPoint(x, y, size);
        BubbleDataPoint_.assertCalledOnceWith(series_data, x, y, size, null);
        $assert($is($get(series_data, -1), data_point_));
        $assert($is(data_point, data_point_));
      } },
  ],
};
pytestClass("DescribeBubbleSeriesData", DescribeBubbleSeriesData);

const DescribeXySeriesData: any = {
  fixtures: {
    add_data_point_fixture: { args: ["request", "XyDataPoint_", "data_point_"], fn: (self: any, request: any, XyDataPoint_: any, data_point_: any): any => {
        let series_data: any, x: any, y: any;
        series_data = $call(XySeriesData, null, null, null);
        [x, y] = $list([42, 24]);
        return [series_data, x, y, XyDataPoint_, data_point_];
      } },
    chart_data_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _BaseChartData);
      } },
    data_point_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, XyDataPoint);
      } },
    XyDataPoint_: { args: ["request", "data_point_"], fn: (self: any, request: any, data_point_: any): any => {
        return $call(classMock, request, "pptx.chart.data.XyDataPoint", $kw({return_value: data_point_}));
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_is_a__BaseSeriesData_object", args: ["chart_data_"], fn: (self: any, chart_data_: any): any => {
        let name: any, number_format: any, series_data: any;
        [name, number_format] = $list(["Series 42", "#0.0"]);
        series_data = $call(XySeriesData, chart_data_, name, number_format);
        $assert($isinstance(series_data, _BaseSeriesData));
        $assert($is(series_data._chartData, chart_data_));
        $assertEqual(series_data.name, name);
        $assertEqual(series_data.numberFormat, number_format);
      } },
    { name: "it_can_add_a_data_point", args: ["add_data_point_fixture"], fn: (self: any, add_data_point_fixture: any): any => {
        let XyDataPoint_: any, data_point: any, data_point_: any, series_data: any, x: any, y: any;
        [series_data, x, y, XyDataPoint_, data_point_] = $list(add_data_point_fixture);
        data_point = series_data.addDataPoint(x, y);
        XyDataPoint_.assertCalledOnceWith(series_data, x, y, null);
        $assert($is($get(series_data, -1), data_point_));
        $assert($is(data_point, data_point_));
      } },
  ],
};
pytestClass("DescribeXySeriesData", DescribeXySeriesData);

const DescribeCategoryDataPoint: any = {
  fixtures: {
    value_fixture: { args: [], fn: (self: any): any => {
        let data_point: any, value: any;
        value = 42;
        data_point = $call(CategoryDataPoint, null, value, null);
        return [data_point, value];
      } },
    series_data_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, CategorySeriesData);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_is_a__BaseDataPoint_object", args: ["series_data_"], fn: (self: any, series_data_: any): any => {
        let data_point: any;
        data_point = $call(CategoryDataPoint, series_data_, 42, "#,##0.0");
        $assert($isinstance(data_point, _BaseDataPoint));
        $assertEqual(data_point.numberFormat, "#,##0.0");
        $assert($is(data_point._seriesData, series_data_));
      } },
    { name: "it_knows_its_value", args: ["value_fixture"], fn: (self: any, value_fixture: any): any => {
        let data_point: any, expected_value: any;
        [data_point, expected_value] = $list(value_fixture);
        $assertEqual(data_point.value, expected_value);
      } },
  ],
};
pytestClass("DescribeCategoryDataPoint", DescribeCategoryDataPoint);

const DescribeXyDataPoint: any = {
  fixtures: {
    value_fixture: { args: [], fn: (self: any): any => {
        let data_point: any, x: any, y: any;
        [x, y] = $list([42, 24]);
        data_point = $call(XyDataPoint, null, x, y, null);
        return [data_point, x, y];
      } },
    series_data_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, XySeriesData);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_is_a__BaseDataPoint_object", args: ["series_data_"], fn: (self: any, series_data_: any): any => {
        let data_point: any;
        data_point = $call(XyDataPoint, series_data_, 42, 24, "00.0");
        $assert($isinstance(data_point, _BaseDataPoint));
        $assert($is(data_point._seriesData, series_data_));
        $assertEqual(data_point.numberFormat, "00.0");
      } },
    { name: "it_knows_its_x_y_values", args: ["value_fixture"], fn: (self: any, value_fixture: any): any => {
        let data_point: any, x: any, y: any;
        [data_point, x, y] = $list(value_fixture);
        $assertEqual(data_point.x, x);
        $assertEqual(data_point.y, y);
      } },
  ],
};
pytestClass("DescribeXyDataPoint", DescribeXyDataPoint);

const DescribeBubbleDataPoint: any = {
  fixtures: {
    value_fixture: { args: [], fn: (self: any): any => {
        let data_point: any, size: any, x: any, y: any;
        [x, y, size] = $list([42, 24, 100]);
        data_point = $call(BubbleDataPoint, null, x, y, size, null);
        return [data_point, x, y, size];
      } },
    series_data_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, BubbleSeriesData);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_is_an_XyDataPoint_subclass", args: ["series_data_"], fn: (self: any, series_data_: any): any => {
        let data_point: any, number_format: any, size: any, x: any, y: any;
        [x, y, size, number_format] = $list([1, 2, 10, "#00.0"]);
        data_point = $call(BubbleDataPoint, series_data_, x, y, size, number_format);
        $assert($isinstance(data_point, XyDataPoint));
        $assert($is(data_point._seriesData, series_data_));
        $assertEqual(data_point.x, x);
        $assertEqual(data_point.y, y);
        $assertEqual(data_point.numberFormat, number_format);
      } },
    { name: "it_knows_its_x_y_size_values", args: ["value_fixture"], fn: (self: any, value_fixture: any): any => {
        let data_point: any, size: any, x: any, y: any;
        [data_point, x, y, size] = $list(value_fixture);
        $assertEqual(data_point.x, x);
        $assertEqual(data_point.y, y);
        $assertEqual(data_point.bubbleSize, size);
      } },
  ],
};
pytestClass("DescribeBubbleDataPoint", DescribeBubbleDataPoint);

