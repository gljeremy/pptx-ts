// Machine-translated from python-pptx features/steps/plot.py by
// scripts/py2ts_steps.py. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $assert, $assertEqual, $call, $d, $eq, $get, $int, $is, $len, $mod, $truthy, $typename } from "../../pyshim.js";
import { AttributeError, IndexError, KeyError, NotImplementedError, ValueError } from "../../../src/errors.js";
import { step } from "../runner.js";
import { test_pptx } from "./helpers.js";
import { Presentation } from "../../../src/index.js";
void $py;


function given_a_bar_plot_having_or_not_data_labels(context: any, having_or_not: any): any {
  let prs: any, slide_idx: any;
  slide_idx = $get($d({having: 0, "not having": 1}), having_or_not);
  prs = $call(Presentation, test_pptx("cht-plot-props"));
  context.plot = $get($get($get(prs.slides, slide_idx).shapes, 0).chart.plots, 0);
}
step("given", "a bar plot {having_or_not} data labels", ["having_or_not"], given_a_bar_plot_having_or_not_data_labels);

function given_a_bar_plot_having_gap_width_of_width(context: any, width: any): any {
  let prs: any, slide_idx: any;
  slide_idx = $get($d({"no explicit value": 0, "300": 1}), width);
  prs = $call(Presentation, test_pptx("cht-plot-props"));
  context.plot = $get($get($get(prs.slides, slide_idx).shapes, 0).chart.plots, 0);
}
step("given", "a bar plot having gap width of {width}", ["width"], given_a_bar_plot_having_gap_width_of_width);

function given_a_bar_plot_having_overlap_of_overlap(context: any, overlap: any): any {
  let prs: any, slide_idx: any;
  slide_idx = $get($d({"no explicit value": 0, "42": 1, "-42": 2}), overlap);
  prs = $call(Presentation, test_pptx("cht-plot-props"));
  context.plot = $get($get($get(prs.slides, slide_idx).shapes, 0).chart.plots, 0);
}
step("given", "a bar plot having overlap of {overlap}", ["overlap"], given_a_bar_plot_having_overlap_of_overlap);

function given_a_bar_plot_having_vary_color_by_category_setting(context: any, setting: any): any {
  let prs: any, slide_idx: any;
  slide_idx = $get($d({"no explicit setting": 0, True: 1, False: 2}), setting);
  prs = $call(Presentation, test_pptx("cht-plot-props"));
  context.plot = $get($get($get(prs.slides, slide_idx).shapes, 0).chart.plots, 0);
}
step("given", "a bar plot having vary color by category set to {setting}", ["setting"], given_a_bar_plot_having_vary_color_by_category_setting);

function given_a_bubble_plot_having_bubble_scale_of_percent(context: any, percent: any): any {
  let prs: any, slide_idx: any;
  slide_idx = $get($d({"no explicit value": 3, "70%": 4}), percent);
  prs = $call(Presentation, test_pptx("cht-plot-props"));
  context.bubblePlot = $get($get($get(prs.slides, slide_idx).shapes, 0).chart.plots, 0);
}
step("given", "a bubble plot having bubble scale of {percent}", ["percent"], given_a_bubble_plot_having_bubble_scale_of_percent);

function given_a_category_plot(context: any): any {
  let prs: any;
  prs = $call(Presentation, test_pptx("cht-plot-props"));
  context.plot = $get($get($get(prs.slides, 2).shapes, 0).chart.plots, 0);
}
step("given", "a category plot", [], given_a_category_plot);

function when_I_assign_value_to_bubble_plot_bubble_scale(context: any, value: any): any {
  let new_value: any;
  new_value = ($truthy($eq(value, "None")) ? null : $int(value));
  context.bubblePlot.bubbleScale = new_value;
}
step("when", "I assign {value} to bubble_plot.bubble_scale", ["value"], when_I_assign_value_to_bubble_plot_bubble_scale);

function when_I_assign_value_to_plot_gap_width(context: any, value: any): any {
  let new_value: any;
  new_value = $int(value);
  context.plot.gapWidth = new_value;
}
step("when", "I assign {value} to plot.gap_width", ["value"], when_I_assign_value_to_plot_gap_width);

function when_I_assign_value_to_plot_has_data_labels(context: any, value: any): any {
  let new_value: any;
  new_value = $get($d({True: true, False: false}), value);
  context.plot.hasDataLabels = new_value;
}
step("when", "I assign {value} to plot.has_data_labels", ["value"], when_I_assign_value_to_plot_has_data_labels);

function when_I_assign_value_to_plot_overlap(context: any, value: any): any {
  let new_value: any;
  new_value = $int(value);
  context.plot.overlap = new_value;
}
step("when", "I assign {value} to plot.overlap", ["value"], when_I_assign_value_to_plot_overlap);

function when_I_assign_value_to_plot_vary_by_categories(context: any, value: any): any {
  let new_value: any;
  new_value = $get($d({True: true, False: false}), value);
  context.plot.varyByCategories = new_value;
}
step("when", "I assign {value} to plot.vary_by_categories", ["value"], when_I_assign_value_to_plot_vary_by_categories);

function then_bubble_plot_bubble_scale_is_value(context: any, value: any): any {
  let bubble_plot: any, expected_value: any;
  expected_value = $int(value);
  bubble_plot = context.bubblePlot;
  $assertEqual(bubble_plot.bubbleScale, expected_value, $mod("got %s", bubble_plot.bubbleScale));
}
step("then", "bubble_plot.bubble_scale is {value}", ["value"], then_bubble_plot_bubble_scale_is_value);

function then_len_plot_categories_is_count(context: any, count: any): any {
  let expected_count: any, plot: any;
  plot = $get(context.chart.plots, 0);
  expected_count = $int(count);
  $assertEqual($len(plot.categories), expected_count);
}
step("then", "len(plot.categories) is {count}", ["count"], then_len_plot_categories_is_count);

function then_plot_categories_is_a_Categories_object(context: any): any {
  let plot: any, type_name: any;
  plot = context.plot;
  type_name = $typename(plot.categories);
  $assertEqual(type_name, "Categories", $mod("got %s", type_name));
}
step("then", "plot.categories is a Categories object", [], then_plot_categories_is_a_Categories_object);

function then_plot_gap_width_is_value(context: any, value: any): any {
  let expected_value: any, plot: any;
  expected_value = $int(value);
  plot = context.plot;
  $assertEqual(plot.gapWidth, expected_value, $mod("got %s", plot.gapWidth));
}
step("then", "plot.gap_width is {value}", ["value"], then_plot_gap_width_is_value);

function then_plot_has_data_labels_is_value(context: any, value: any): any {
  let expected_value: any;
  expected_value = $get($d({True: true, False: false}), value);
  $assert($is(context.plot.hasDataLabels, expected_value));
}
step("then", "plot.has_data_labels is {value}", ["value"], then_plot_has_data_labels_is_value);

function then_plot_overlap_is_expected_value(context: any, value: any): any {
  let expected_value: any, plot: any;
  expected_value = $int(value);
  plot = context.plot;
  $assertEqual(plot.overlap, expected_value, $mod("got %s", plot.overlap));
}
step("then", "plot.overlap is {value}", ["value"], then_plot_overlap_is_expected_value);

function then_plot_vary_by_categories_is_value(context: any, value: any): any {
  let expected_value: any, plot: any;
  expected_value = $get($d({True: true, False: false}), value);
  plot = context.plot;
  $assert($is(plot.varyByCategories, expected_value));
}
step("then", "plot.vary_by_categories is {value}", ["value"], then_plot_vary_by_categories_is_value);

// the names python's eval() sees in this module
const $scope = { Presentation };
