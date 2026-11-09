// Machine-translated from python-pptx tests/chart/test_category.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $call, $get, $is, $isinstance, $kw, $len, $list, $range, $truthy, call, classMock, instanceMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { Categories as Categories$imported, Category as Category$imported, CategoryLevel as CategoryLevel$imported } from "../../src/chart/category.js";
import { parseXml as parseXml$imported } from "../../src/oxml/index.js";
import { element } from "../unitutil/cxml.js";
import { snippetSeq } from "../unitutil/file.js";
// python binds imported names at import time; patches made later don't change them
const Categories = Categories$imported;
const Category = Category$imported;
const CategoryLevel = CategoryLevel$imported;
const parseXml = parseXml$imported;
void $py;


const DescribeCategories: any = {
  fixtures: {
    depth_fixture: { args: ["request"], params: [["c:barChart", 0], ["c:barChart/c:ser/c:cat", 1], ["c:barChart/c:ser/c:cat/c:multiLvlStrRef/c:lvl", 1], ["c:barChart/c:ser/c:cat/c:multiLvlStrRef/(c:lvl,c:lvl)", 2], ["c:barChart/c:ser/c:cat/c:multiLvlStrRef/(c:lvl,c:lvl,c:lvl)", 3]], fn: (self: any, request: any): any => {
        let categories: any, expected_value: any, xChart_cxml: any;
        [xChart_cxml, expected_value] = $list(request.param);
        categories = $call(Categories, element(xChart_cxml));
        return [categories, expected_value];
      } },
    flat_fixture: { args: ["request"], params: [[0, [["Foo"], [""], ["Baz"]]], [1, [["CA", "SF"], ["CA", "LA"], ["NY", "NY"], ["NY", "Albany"]]], [2, [["USA", "CA", "SF"], ["USA", "CA", "LA"], ["USA", "NY", "NY"], ["USA", "NY", "Albany"], ["CAN", "AL", "Calgary"], ["CAN", "AL", "Edmonton"], ["CAN", "ON", "Toronto"], ["CAN", "ON", "Ottawa"]]], [3, [["SF"], ["LA"], ["NY"]]]], fn: (self: any, request: any): any => {
        let categories: any, expected_values: any, snippet_idx: any, xChart_xml: any;
        [snippet_idx, expected_values] = $list(request.param);
        xChart_xml = $get(snippetSeq("cat-labels"), snippet_idx);
        categories = $call(Categories, parseXml(xChart_xml));
        return [categories, expected_values];
      } },
    getitem_fixture: { args: ["request", "Category_", "category_"], params: [["c:barChart/c:ser/c:cat/(c:ptCount{val=2})", 0, null], ["c:barChart/c:ser/c:cat/(c:ptCount{val=2},c:pt{idx=1})", 0, null], ["c:barChart/c:ser/c:cat/(c:ptCount{val=2},c:pt{idx=1})", 1, 0], ["c:barChart/c:ser/c:cat/c:lvl/(c:ptCount{val=2},c:pt{idx=1})", 1, 0]], fn: (self: any, request: any, Category_: any, category_: any): any => {
        let categories: any, idx: any, pt: any, pt_offset: any, xChart: any, xChart_cxml: any;
        [xChart_cxml, idx, pt_offset] = $list(request.param);
        xChart = element(xChart_cxml);
        pt = ($truthy($is(pt_offset, null)) ? null : $get(xChart.xpath(".//c:pt"), pt_offset));
        categories = $call(Categories, xChart);
        return [categories, idx, Category_, pt, category_];
      } },
    iter_fixture: { args: ["Category_", "category_"], fn: (self: any, Category_: any, category_: any): any => {
        let calls: any, categories: any, expected_categories: any, pt: any, xChart: any;
        xChart = element("c:barChart/c:ser/c:cat/(c:ptCount{val=2},c:pt{idx=1})");
        pt = $get(xChart.xpath(".//c:pt"), 0);
        categories = $call(Categories, xChart);
        expected_categories = [category_, category_];
        calls = [call(null, 0), call(pt, 1)];
        return [categories, expected_categories, Category_, calls];
      } },
    len_fixture: { args: ["request"], params: [["c:barChart", 0], ["c:barChart/c:ser/c:cat/c:ptCount{val=4}", 4]], fn: (self: any, request: any): any => {
        let categories: any, expected_len: any, xChart_cxml: any;
        [xChart_cxml, expected_len] = $list(request.param);
        categories = $call(Categories, element(xChart_cxml));
        return [categories, expected_len];
      } },
    levels_fixture: { args: ["request", "CategoryLevel_", "category_level_"], params: [["c:barChart", 0], ["c:barChart/c:ser/c:cat", 0], ["c:barChart/c:ser/c:cat/c:multiLvlStrRef/c:lvl", 1], ["c:barChart/c:ser/c:cat/c:multiLvlStrRef/(c:lvl,c:lvl)", 2], ["c:barChart/c:ser/c:cat/c:multiLvlStrRef/(c:lvl,c:lvl,c:lvl)", 3]], fn: (self: any, request: any, CategoryLevel_: any, category_level_: any): any => {
        let calls: any, categories: any, expected_levels: any, level_count: any, lvls: any, xChart: any, xChart_cxml: any;
        [xChart_cxml, level_count] = $list(request.param);
        xChart = element(xChart_cxml);
        lvls = xChart.xpath(".//c:lvl");
        categories = $call(Categories, xChart);
        calls = $list(lvls).map((lvl: any) => call(lvl));
        expected_levels = $list($range(level_count)).map((_: any) => category_level_);
        return [categories, CategoryLevel_, calls, expected_levels];
      } },
    Category_: { args: ["request", "category_"], fn: (self: any, request: any, category_: any): any => {
        return $call(classMock, request, "pptx.chart.category.Category", $kw({return_value: category_}));
      } },
    category_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Category);
      } },
    CategoryLevel_: { args: ["request", "category_level_"], fn: (self: any, request: any, category_level_: any): any => {
        return $call(classMock, request, "pptx.chart.category.CategoryLevel", $kw({return_value: category_level_}));
      } },
    category_level_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, CategoryLevel);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_length", args: ["len_fixture"], fn: (self: any, len_fixture: any): any => {
        let categories: any, expected_len: any;
        [categories, expected_len] = $list(len_fixture);
        $assertEqual($len(categories), expected_len);
      } },
    { name: "it_supports_indexed_access", args: ["getitem_fixture"], fn: (self: any, getitem_fixture: any): any => {
        let Category_: any, categories: any, category: any, category_: any, idx: any, pt: any;
        [categories, idx, Category_, pt, category_] = $list(getitem_fixture);
        category = $get(categories, idx);
        Category_.assertCalledOnceWith(pt, idx);
        $assert($is(category, category_));
      } },
    { name: "it_can_iterate_over_the_categories_it_contains", args: ["iter_fixture"], fn: (self: any, iter_fixture: any): any => {
        let Category_: any, calls: any, categories: any, expected_categories: any;
        [categories, expected_categories, Category_, calls] = $list(iter_fixture);
        $assertEqual($list(categories).map((c: any) => c), expected_categories);
        $assertEqual(Category_.callArgsList, calls);
      } },
    { name: "it_knows_its_depth", args: ["depth_fixture"], fn: (self: any, depth_fixture: any): any => {
        let categories: any, expected_value: any;
        [categories, expected_value] = $list(depth_fixture);
        $assertEqual(categories.depth, expected_value);
      } },
    { name: "it_provides_a_flattened_representation", args: ["flat_fixture"], fn: (self: any, flat_fixture: any): any => {
        let categories: any, expected_values: any, flattened_labels: any;
        [categories, expected_values] = $list(flat_fixture);
        flattened_labels = categories.flattenedLabels;
        $assertEqual(flattened_labels, expected_values);
      } },
    { name: "it_provides_access_to_its_levels", args: ["levels_fixture"], fn: (self: any, levels_fixture: any): any => {
        let CategoryLevel_: any, calls: any, categories: any, expected_levels: any, levels: any;
        [categories, CategoryLevel_, calls, expected_levels] = $list(levels_fixture);
        levels = categories.levels;
        $assertEqual(CategoryLevel_.callArgsList, calls);
        $assertEqual(levels, expected_levels);
      } },
  ],
};
pytestClass("DescribeCategories", DescribeCategories);

const DescribeCategory: any = {
  fixtures: {
    base_class_fixture: { args: ["request"], params: [[null, ""], ["c:pt/c:v\"Foo\"", "Foo"]], fn: (self: any, request: any): any => {
        let category: any, pt: any, pt_cxml: any, str_value: any;
        [pt_cxml, str_value] = $list(request.param);
        pt = ($truthy($is(pt_cxml, null)) ? null : element(pt_cxml));
        category = $call(Category, pt);
        return [category, str_value];
      } },
    idx_fixture: { args: ["request"], params: [[null, 42, 42], ["c:pt{idx=9}/c:v\"Bar\"", null, 9]], fn: (self: any, request: any): any => {
        let category: any, expected_value: any, idx: any, pt: any, pt_cxml: any;
        [pt_cxml, idx, expected_value] = $list(request.param);
        pt = ($truthy($is(pt_cxml, null)) ? null : element(pt_cxml));
        category = $call(Category, pt, idx);
        return [category, expected_value];
      } },
    label_fixture: { args: ["request"], params: [[null, ""], ["c:pt{idx=9}/c:v\"Bar\"", "Bar"]], fn: (self: any, request: any): any => {
        let category: any, expected_value: any, pt: any, pt_cxml: any;
        [pt_cxml, expected_value] = $list(request.param);
        pt = ($truthy($is(pt_cxml, null)) ? null : element(pt_cxml));
        category = $call(Category, pt);
        return [category, expected_value];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_extends_str", args: ["base_class_fixture"], fn: (self: any, base_class_fixture: any): any => {
        let category: any, str_value: any;
        [category, str_value] = $list(base_class_fixture);
        $assert($isinstance(category, String));
        $assertEqual(category, str_value);
      } },
    { name: "it_knows_its_idx", args: ["idx_fixture"], fn: (self: any, idx_fixture: any): any => {
        let category: any, expected_value: any, idx: any;
        [category, expected_value] = $list(idx_fixture);
        idx = category.idx;
        $assertEqual(idx, expected_value);
      } },
    { name: "it_knows_its_label", args: ["label_fixture"], fn: (self: any, label_fixture: any): any => {
        let category: any, expected_value: any, label: any;
        [category, expected_value] = $list(label_fixture);
        label = category.label;
        $assertEqual(label, expected_value);
      } },
  ],
};
pytestClass("DescribeCategory", DescribeCategory);

const DescribeCategoryLevel: any = {
  fixtures: {
    getitem_fixture: { args: ["request", "Category_", "category_"], params: [0, 1, 2], fn: (self: any, request: any, Category_: any, category_: any): any => {
        let category_level: any, idx: any, lvl: any, lvl_cxml: any, pt: any;
        idx = request.param;
        lvl_cxml = "c:lvl/(c:pt,c:pt,c:pt)";
        lvl = element(lvl_cxml);
        pt = $get(lvl.xpath("./c:pt"), idx);
        category_level = $call(CategoryLevel, lvl);
        return [category_level, idx, Category_, pt, category_];
      } },
    len_fixture: { args: ["request"], params: [["c:lvl", 0], ["c:lvl/c:pt", 1], ["c:lvl/(c:pt,c:pt)", 2], ["c:lvl/(c:pt,c:pt,c:pt)", 3]], fn: (self: any, request: any): any => {
        let category_level: any, expected_len: any, lvl_cxml: any;
        [lvl_cxml, expected_len] = $list(request.param);
        category_level = $call(CategoryLevel, element(lvl_cxml));
        return [category_level, expected_len];
      } },
    Category_: { args: ["request", "category_"], fn: (self: any, request: any, category_: any): any => {
        return $call(classMock, request, "pptx.chart.category.Category", $kw({return_value: category_}));
      } },
    category_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Category);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_length", args: ["len_fixture"], fn: (self: any, len_fixture: any): any => {
        let category_level: any, expected_len: any;
        [category_level, expected_len] = $list(len_fixture);
        $assertEqual($len(category_level), expected_len);
      } },
    { name: "it_supports_indexed_access", args: ["getitem_fixture"], fn: (self: any, getitem_fixture: any): any => {
        let Category_: any, category: any, category_: any, category_level: any, idx: any, pt: any;
        [category_level, idx, Category_, pt, category_] = $list(getitem_fixture);
        category = $get(category_level, idx);
        Category_.assertCalledOnceWith(pt);
        $assert($is(category, category_));
      } },
  ],
};
pytestClass("DescribeCategoryLevel", DescribeCategoryLevel);

