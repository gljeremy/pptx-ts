// Machine-translated from python-pptx features/steps/chartdata.py by
// scripts/py2ts_steps.py. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $assert, $assertEqual, $call, $d, $datetime, $eq, $get, $getattr, $int, $is, $kw, $list, $m, $mod, $set, $t1, $t2, $t3, $truthy, $typename } from "../../pyshim.js";
import { AttributeError, IndexError, KeyError, NotImplementedError, ValueError } from "../../../src/errors.js";
import { step } from "../runner.js";
import { BubbleChartData, Category, CategoryChartData, XyChartData } from "../../../src/chart/data.js";
import { XL_CHART_TYPE } from "../../../src/enum/chart.js";
import { Inches } from "../../../src/util.js";
void $py;


function given_a_BubbleChartData_object_with_number_format(context: any, strval: any): any {
  let params: any;
  params = $d({});
  if ($truthy(!$eq(strval, "None"))) {
    $set(params, "number_format", $int(strval));
  }
  context.chartData = $call(BubbleChartData, $kw({...params}));
}
step("given", "a BubbleChartData object with number format {strval}", ["strval"], given_a_BubbleChartData_object_with_number_format);

function given_a_Categories_object_with_number_format_init_nf(context: any, init_nf: any): any {
  let categories: any;
  categories = $call(CategoryChartData).categories;
  if ($truthy(!$eq(init_nf, "left as default"))) {
    categories.numberFormat = init_nf;
  }
  context.categories = categories;
}
step("given", "a Categories object with number format {init_nf}", ["init_nf"], given_a_Categories_object_with_number_format_init_nf);

function given_a_Category_object(context: any): any {
  context.category = $call(Category, null, null);
}
step("given", "a Category object", [], given_a_Category_object);

function given_a_CategoryChartData_object(context: any): any {
  context.chartData = $call(CategoryChartData);
}
step("given", "a CategoryChartData object", [], given_a_CategoryChartData_object);

function given_a_CategoryChartData_object_having_date_categories(context: any): any {
  let chart_data: any;
  chart_data = $call(CategoryChartData);
  chart_data.categories = [$datetime.date(2016, 12, 27), $datetime.date(2016, 12, 28), $datetime.date(2016, 12, 29)];
  context.chartData = chart_data;
}
step("given", "a CategoryChartData object having date categories", [], given_a_CategoryChartData_object_having_date_categories);

function given_a_CategoryChartData_object_with_number_format(context: any, strval: any): any {
  let params: any;
  params = $d({});
  if ($truthy(!$eq(strval, "None"))) {
    $set(params, "number_format", $int(strval));
  }
  context.chartData = $call(CategoryChartData, $kw({...params}));
}
step("given", "a CategoryChartData object with number format {strval}", ["strval"], given_a_CategoryChartData_object_with_number_format);

function given_a_XyChartData_object_with_number_format(context: any, strval: any): any {
  let params: any;
  params = $d({});
  if ($truthy(!$eq(strval, "None"))) {
    $set(params, "number_format", $int(strval));
  }
  context.chartData = $call(XyChartData, $kw({...params}));
}
step("given", "a XyChartData object with number format {strval}", ["strval"], given_a_XyChartData_object_with_number_format);

function given_the_categories_are_of_type(context: any, type_: any): any {
  let label: any;
  label = $get($d({date: $datetime.date(2016, 12, 22), float: 42.24, int: 42, str: "foobar"}), type_);
  context.categories.addCategory(label);
}
step("given", "the categories are of type {type_}", ["type_"], given_the_categories_are_of_type);

function when_I_add_a_bubble_data_point_with_number_format(context: any, strval: any): any {
  let params: any, series_data: any;
  series_data = context.seriesData;
  params = $d({x: 1, y: 2, size: 10});
  if ($truthy(!$eq(strval, "None"))) {
    $set(params, "number_format", $int(strval));
  }
  context.dataPoint = $m(series_data, "addDataPoint", $kw({...params}));
}
step("when", "I add a bubble data point with number format {strval}", ["strval"], when_I_add_a_bubble_data_point_with_number_format);

function when_I_add_a_data_point_with_number_format(context: any, strval: any): any {
  let params: any, series_data: any;
  series_data = context.seriesData;
  params = $d({value: 42});
  if ($truthy(!$eq(strval, "None"))) {
    $set(params, "number_format", $int(strval));
  }
  context.dataPoint = $m(series_data, "addDataPoint", $kw({...params}));
}
step("when", "I add a data point with number format {strval}", ["strval"], when_I_add_a_data_point_with_number_format);

function when_I_add_an_XY_data_point_with_number_format(context: any, strval: any): any {
  let params: any, series_data: any;
  series_data = context.seriesData;
  params = $d({x: 1, y: 2});
  if ($truthy(!$eq(strval, "None"))) {
    $set(params, "number_format", $int(strval));
  }
  context.dataPoint = $m(series_data, "addDataPoint", $kw({...params}));
}
step("when", "I add an XY data point with number format {strval}", ["strval"], when_I_add_an_XY_data_point_with_number_format);

function when_I_add_an_xy_chart_having_2_series_of_3_points(context: any, xy_type: any): any {
  let chart_data: any, chart_type: any, data: any, points: any, series: any, series_label: any, x: any, y: any;
  chart_type = $getattr(XL_CHART_TYPE, xy_type);
  data = [["Series 1", [[-0.1, 0.5], [16.2, 0.0], [8.0, 0.2]]], ["Series 2", [[12.4, 0.8], [-7.5, -0.5], [-5.1, -0.2]]]];
  chart_data = $call(XyChartData);
  for (const series_data of $list(data)) {
    [series_label, points] = $list(series_data);
    series = chart_data.addSeries(series_label);
    for (const point of $list(points)) {
      [x, y] = $list(point);
      series.addDataPoint(x, y);
    }
  }
  context.chart = context.slide.shapes.addChart(chart_type, $call(Inches, 1), $call(Inches, 1), $call(Inches, 8), $call(Inches, 5), chart_data).chart;
}
step("when", "I add an {xy_type} chart having 2 series of 3 points each", ["xy_type"], when_I_add_an_xy_chart_having_2_series_of_3_points);

function when_I_assign_a_b_c_to_chart_data_categories(context: any): any {
  let chart_data: any;
  chart_data = context.chartData;
  chart_data.categories = ["a", "b", "c"];
}
step("when", "I assign ['a', 'b', 'c'] to chart_data.categories", [], when_I_assign_a_b_c_to_chart_data_categories);

function then_c_label_for_c_in_chart_data_categories_is_a_b_c(context: any): any {
  let chart_data: any;
  chart_data = context.chartData;
  $assertEqual($list(chart_data.categories).map((c: any) => c.label), ["a", "b", "c"]);
}
step("then", "[c.label for c in chart_data.categories] is ['a', 'b', 'c']", [], then_c_label_for_c_in_chart_data_categories_is_a_b_c);

function then_categories_number_format_is_value(context: any, value: any): any {
  let expected_value: any, number_format: any;
  expected_value = value;
  number_format = context.categories.numberFormat;
  $assertEqual(number_format, expected_value, $mod("got %s", number_format));
}
step("then", "categories.number_format is {value}", ["value"], then_categories_number_format_is_value);

function then_category_add_sub_category_is_a_Category_object(context: any): any {
  let category: any, sub_category: any;
  category = context.category;
  const $t1 = category.addSubCategory("foobar");
  context.subCategory = $t1;
  sub_category = $t1;
  $assertEqual($typename(sub_category), "Category");
}
step("then", "category.add_sub_category(name) is a Category object", [], then_category_add_sub_category_is_a_Category_object);

function then_category_sub_categories_minus_1_is_the_new_category(context: any): any {
  let category: any, sub_category: any;
  [category, sub_category] = $list([context.category, context.subCategory]);
  $assert($is($get(category.subCategories, -1), sub_category));
}
step("then", "category.sub_categories[-1] is the new category", [], then_category_sub_categories_minus_1_is_the_new_category);

function then_chart_data_add_category_name_is_a_Category_object(context: any): any {
  let category: any, chart_data: any;
  chart_data = context.chartData;
  const $t2 = chart_data.addCategory("foobar");
  context.category = $t2;
  category = $t2;
  $assertEqual($typename(category), "Category");
}
step("then", "chart_data.add_category(name) is a Category object", [], then_chart_data_add_category_name_is_a_Category_object);

function then_chart_data_add_series_is_a_CategorySeriesData_object(context: any): any {
  let chart_data: any, series: any;
  chart_data = context.chartData;
  const $t3 = chart_data.addSeries("Series X", [1, 2, 3]);
  context.series = $t3;
  series = $t3;
  $assertEqual($typename(series), "CategorySeriesData");
}
step("then", "chart_data.add_series(name, values) is a CategorySeriesData object", [], then_chart_data_add_series_is_a_CategorySeriesData_object);

function then_chart_data_categories_is_a_Categories_object(context: any): any {
  let chart_data: any;
  chart_data = context.chartData;
  $assertEqual($typename(chart_data.categories), "Categories");
}
step("then", "chart_data.categories is a Categories object", [], then_chart_data_categories_is_a_Categories_object);

function then_chart_data_categories_minus_1_is_the_category(context: any): any {
  let category: any, chart_data: any;
  [chart_data, category] = $list([context.chartData, context.category]);
  $assert($is($get(chart_data.categories, -1), category));
}
step("then", "chart_data.categories[-1] is the category", [], then_chart_data_categories_minus_1_is_the_category);

function then_chart_data_number_format_is(context: any, value_str: any): any {
  let chart_data: any, number_format: any;
  chart_data = context.chartData;
  number_format = ($truthy($eq(value_str, "General")) ? value_str : $int(value_str));
  $assertEqual(chart_data.numberFormat, number_format);
}
step("then", "chart_data.number_format is {value_str}", ["value_str"], then_chart_data_number_format_is);

function then_chart_data_minus_1_is_the_new_series(context: any): any {
  let chart_data: any, series: any;
  [chart_data, series] = $list([context.chartData, context.series]);
  $assert($is($get(chart_data, -1), series));
}
step("then", "chart_data[-1] is the new series", [], then_chart_data_minus_1_is_the_new_series);

function then_series_data_number_format_is(context: any, value_str: any): any {
  let number_format: any, series_data: any;
  series_data = context.seriesData;
  number_format = ($truthy($eq(value_str, "General")) ? value_str : $int(value_str));
  $assertEqual(series_data.numberFormat, number_format);
}
step("then", "series_data.number_format is {value_str}", ["value_str"], then_series_data_number_format_is);

// the names python's eval() sees in this module
const $scope = { BubbleChartData, Category, CategoryChartData, Inches, XL_CHART_TYPE, XyChartData };
