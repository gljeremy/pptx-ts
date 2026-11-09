// Machine-translated from python-pptx features/steps/category.py by
// scripts/py2ts_steps.py. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $assertEqual, $call, $d, $enumerate, $eq, $get, $int, $len, $list, $mod, $truthy, $typename } from "../../pyshim.js";
import { AttributeError, IndexError, KeyError, NotImplementedError, ValueError } from "../../../src/errors.js";
import { step } from "../runner.js";
import { test_pptx } from "./helpers.js";
import { Presentation } from "../../../src/index.js";
void $py;


function given_a_Categories_object_containing_3_categories(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("cht-category-access"));
  context.categories = $get($get($get(prs.slides, 0).shapes, 0).chart.plots, 0).categories;
}
step("given", "a Categories object containing 3 categories", [], given_a_Categories_object_containing_3_categories);

function given_a_Categories_object_having_count_category_levels(context: any, count: any): any {
  let slide: any, slide_idx: any;
  slide_idx = $get([2, 0, 3, 1], $int(count));
  slide = $get($call(Presentation, test_pptx("cht-category-access")).slides, slide_idx);
  context.categories = $get($get(slide.shapes, 0).chart.plots, 0).categories;
}
step("given", "a Categories object having {count} category levels", ["count"], given_a_Categories_object_having_count_category_levels);

function given_a_Categories_obj_having_leafs_and_levels(context: any, leafs: any, levels: any): any {
  let slide: any, slide_idx: any;
  slide_idx = $get($d({[[3, 1]]: 0, [[8, 3]]: 1, [[0, 0]]: 2, [[4, 2]]: 3}), [$int(leafs), $int(levels)]);
  slide = $get($call(Presentation, test_pptx("cht-category-access")).slides, slide_idx);
  context.categories = $get($get(slide.shapes, 0).chart.plots, 0).categories;
}
step("given", "a Categories object having {leafs} categories and {levels} levels", ["leafs", "levels"], given_a_Categories_obj_having_leafs_and_levels);

function given_a_Category_object_having_idx_value_idx(context: any, idx: any): any {
  let cat_offset: any, slide: any;
  cat_offset = $int(idx);
  slide = $get($call(Presentation, test_pptx("cht-category-access")).slides, 0);
  context.category = $get($get($get(slide.shapes, 0).chart.plots, 0).categories, cat_offset);
}
step("given", "a Category object having idx value {idx}", ["idx"], given_a_Category_object_having_idx_value_idx);

function given_a_Category_object_having_label(context: any, label: any): any {
  let cat_offset: any, slide: any;
  cat_offset = $get($d({"label 'Foo'": 0, "no label": 1}), label);
  slide = $get($call(Presentation, test_pptx("cht-category-access")).slides, 0);
  context.category = $get($get($get(slide.shapes, 0).chart.plots, 0).categories, cat_offset);
}
step("given", "a Category object having {label}", ["label"], given_a_Category_object_having_label);

function given_a_CategoryLevel_object_containing_4_categories(context: any): any {
  let chart: any, slide: any;
  slide = $get($call(Presentation, test_pptx("cht-category-access")).slides, 1);
  chart = $get(slide.shapes, 0).chart;
  context.categoryLevel = $get($get(chart.plots, 0).categories.levels, 1);
}
step("given", "a CategoryLevel object containing 4 categories", [], given_a_CategoryLevel_object_containing_4_categories);

function then_categories_2_is_a_Category_object(context: any): any {
  let type_name: any;
  type_name = $typename($get(context.categories, 2));
  $assertEqual(type_name, "Category", $mod("got %s", type_name));
}
step("then", "categories[2] is a Category object", [], then_categories_2_is_a_Category_object);

function then_categories_depth_is_value(context: any, value: any): any {
  let depth: any, expected_value: any;
  expected_value = $int(value);
  depth = context.categories.depth;
  $assertEqual(depth, expected_value, $mod("got %s", expected_value));
}
step("then", "categories.depth is {value}", ["value"], then_categories_depth_is_value);

function then_categories_flattened_labels_is_tuple_of_tuples(context: any, leafs: any): any {
  let flattened_labels: any, length: any, type_name: any;
  flattened_labels = context.categories.flattenedLabels;
  type_name = $typename(flattened_labels);
  length = $len(flattened_labels);
  $assertEqual(type_name, "Array", $mod("got %s", type_name));
  $assertEqual(length, $int(leafs), $mod("got %s", length));
  for (const labels of $list(flattened_labels)) {
    type_name = $typename(labels);
    $assertEqual(type_name, "Array", $mod("got %s", type_name));
  }
}
step("then", "categories.flattened_labels is a tuple of {leafs} tuples", ["leafs"], then_categories_flattened_labels_is_tuple_of_tuples);

function then_categories_levels_contains_count_CategoryLevel_objs(context: any, count: any): any {
  let expected_idx: any, idx: any, type_name: any;
  expected_idx = ($int(count) - 1);
  idx = -1;
  let category_level: any;
  for ([idx, category_level] of $list($enumerate(context.categories.levels))) {
    type_name = $typename(category_level);
    $assertEqual(type_name, "CategoryLevel", $mod("got %s", type_name));
  }
  $assertEqual(idx, expected_idx, $mod("got %s", idx));
}
step("then", "categories.levels contains {count} CategoryLevel objects", ["count"], then_categories_levels_contains_count_CategoryLevel_objs);

function then_category_idx_is_value(context: any, value: any): any {
  let expected_value: any, idx: any;
  expected_value = ($truthy($eq(value, "None")) ? null : $int(value));
  idx = context.category.idx;
  $assertEqual(idx, expected_value, $mod("got %s", idx));
}
step("then", "category.idx is {value}", ["value"], then_category_idx_is_value);

function then_category_label_is_value(context: any, value: any): any {
  let expected_value: any, label: any;
  expected_value = $get($d({"'Foo'": "Foo", "''": ""}), value);
  label = context.category.label;
  $assertEqual(label, expected_value, $mod("got %s", label));
}
step("then", "category.label is {value}", ["value"], then_category_label_is_value);

function then_category_level_2_is_a_Category_object(context: any): any {
  let type_name: any;
  type_name = $typename($get(context.categoryLevel, 2));
  $assertEqual(type_name, "Category", $mod("got %s", type_name));
}
step("then", "category_level[2] is a Category object", [], then_category_level_2_is_a_Category_object);

function then_each_label_tuple_contains_levels_labels(context: any, levels: any): any {
  let flattened_labels: any, length: any, type_name: any;
  flattened_labels = context.categories.flattenedLabels;
  for (const labels of $list(flattened_labels)) {
    length = $len(labels);
    $assertEqual(length, $int(levels), $mod("got %s", levels));
    for (const label of $list(labels)) {
      type_name = $typename(label);
      $assertEqual(type_name, "str", $mod("got %s", type_name));
    }
  }
}
step("then", "each label tuple contains {levels} labels", ["levels"], then_each_label_tuple_contains_levels_labels);

function then_iterating_categories_produces_3_category_objects(context: any): any {
  let categories: any, idx: any;
  categories = context.categories;
  idx = -1;
  let category: any;
  for ([idx, category] of $list($enumerate(categories))) {
    $assertEqual($typename(category), "Category");
  }
  $assertEqual(idx, 2, $mod("got %s", idx));
}
step("then", "iterating categories produces 3 Category objects", [], then_iterating_categories_produces_3_category_objects);

function then_iterating_category_level_produces_4_Category_objects(context: any): any {
  let idx: any, type_name: any;
  idx = -1;
  let category: any;
  for ([idx, category] of $list($enumerate(context.categoryLevel))) {
    type_name = $typename(category);
    $assertEqual(type_name, "Category", $mod("got %s", type_name));
  }
  $assertEqual(idx, 3, $mod("got %s", idx));
}
step("then", "iterating category_level produces 4 Category objects", [], then_iterating_category_level_produces_4_Category_objects);

function then_len_categories_is_count(context: any, count: any): any {
  let expected_count: any;
  expected_count = $int(count);
  $assertEqual($len(context.categories), expected_count);
}
step("then", "len(categories) is {count}", ["count"], then_len_categories_is_count);

function then_len_category_level_is_count(context: any, count: any): any {
  let actual_count: any, expected_count: any;
  expected_count = $int(count);
  actual_count = $len(context.categoryLevel);
  $assertEqual(actual_count, expected_count, $mod("got %s", actual_count));
}
step("then", "len(category_level) is {count}", ["count"], then_len_category_level_is_count);

function then_list_categories_is_Foo_empty_Baz(context: any): any {
  let cats_list: any;
  cats_list = $list(context.categories);
  $assertEqual(cats_list, ["Foo", "", "Baz"], $mod("got %s", cats_list));
}
step("then", "list(categories) == ['Foo', '', 'Baz']", [], then_list_categories_is_Foo_empty_Baz);

// the names python's eval() sees in this module
const $scope = { Presentation };
