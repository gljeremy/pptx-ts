// Machine-translated from python-pptx tests/chart/test_plot.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $add, $assert, $assertEqual, $call, $endswith, $get, $is, $kw, $list, $mod, $slice, $truthy, classMock, instanceMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { Categories as Categories$imported } from "../../src/chart/category.js";
import { Chart as Chart$imported } from "../../src/chart/chart.js";
import { Area3DPlot as Area3DPlot$imported, AreaPlot as AreaPlot$imported, BarPlot as BarPlot$imported, BubblePlot as BubblePlot$imported, DataLabels as DataLabels$imported, DoughnutPlot as DoughnutPlot$imported, LinePlot as LinePlot$imported, PiePlot as PiePlot$imported, PlotFactory as PlotFactory$imported, PlotTypeInspector as PlotTypeInspector$imported, RadarPlot as RadarPlot$imported, XyPlot as XyPlot$imported, _BasePlot as _BasePlot$imported } from "../../src/chart/plot.js";
import { SeriesCollection as SeriesCollection$imported } from "../../src/chart/series.js";
import { XL_CHART_TYPE as XL$imported } from "../../src/enum/chart.js";
import { element, xml } from "../unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const Categories = Categories$imported;
const Chart = Chart$imported;
const Area3DPlot = Area3DPlot$imported;
const AreaPlot = AreaPlot$imported;
const BarPlot = BarPlot$imported;
const BubblePlot = BubblePlot$imported;
const DataLabels = DataLabels$imported;
const DoughnutPlot = DoughnutPlot$imported;
const LinePlot = LinePlot$imported;
const PiePlot = PiePlot$imported;
const PlotFactory = PlotFactory$imported;
const PlotTypeInspector = PlotTypeInspector$imported;
const RadarPlot = RadarPlot$imported;
const XyPlot = XyPlot$imported;
const _BasePlot = _BasePlot$imported;
const SeriesCollection = SeriesCollection$imported;
const XL = XL$imported;
void $py;


const Describe_BasePlot: any = {
  fixtures: {
    categories_fixture: { args: ["categories_", "Categories_"], fn: (self: any, categories_: any, Categories_: any): any => {
        let plot: any, xChart: any;
        xChart = element("c:barChart");
        plot = $call(_BasePlot, xChart, null);
        return [plot, categories_, Categories_, xChart];
      } },
    chart_fixture: { args: ["chart_"], fn: (self: any, chart_: any): any => {
        let expected_value: any, plot: any;
        plot = $call(_BasePlot, null, chart_);
        expected_value = chart_;
        return [plot, expected_value];
      } },
    data_labels_fixture: { args: ["DataLabels_", "data_labels_"], fn: (self: any, DataLabels_: any, data_labels_: any): any => {
        let barChart: any, dLbls: any, plot: any;
        barChart = element("c:barChart/c:dLbls");
        dLbls = $get(barChart, 0);
        plot = $call(_BasePlot, barChart, null);
        return [plot, data_labels_, DataLabels_, dLbls];
      } },
    has_data_labels_get_fixture: { args: ["request"], params: [["c:barChart", false], ["c:barChart/c:dLbls", true], ["c:lineChart", false], ["c:lineChart/c:dLbls", true], ["c:pieChart", false], ["c:pieChart/c:dLbls", true]], fn: (self: any, request: any): any => {
        let expected_value: any, plot: any, xChart_cxml: any;
        [xChart_cxml, expected_value] = $list(request.param);
        plot = $call(_BasePlot, element(xChart_cxml), null);
        return [plot, expected_value];
      } },
    has_data_labels_set_fixture: { args: ["request"], params: [["c:barChart", true, "c:barChart/c:dLbls/+"], ["c:barChart/c:dLbls", true, "c:barChart/c:dLbls"], ["c:barChart", false, "c:barChart"], ["c:barChart/c:dLbls", false, "c:barChart"], ["c:bubbleChart", true, "c:bubbleChart/c:dLbls/+"], ["c:bubbleChart/c:dLbls", false, "c:bubbleChart"], ["c:lineChart", true, "c:lineChart/c:dLbls/+"], ["c:lineChart/c:dLbls", false, "c:lineChart"], ["c:pieChart", true, "c:pieChart/c:dLbls/+"], ["c:pieChart/c:dLbls", false, "c:pieChart"]], fn: (self: any, request: any): any => {
        let expected_cxml: any, expected_xml: any, new_value: any, plot: any, xChart_cxml: any;
        [xChart_cxml, new_value, expected_cxml] = $list(request.param);
        if ($truthy($endswith(expected_cxml, "+"))) {
          expected_cxml = $add($slice(expected_cxml, null, -1, null), "(c:showLegendKey{val=0},c:showVal{val=1},c:showCatName{val=0},c:showSerName{val=0},c:showPercent{val=0},c:showBubbleSize{val=0},c:showLeaderLines{val=1})");
        }
        plot = $call(PlotFactory, element(xChart_cxml), null);
        expected_xml = xml(expected_cxml);
        return [plot, new_value, expected_xml];
      } },
    series_fixture: { args: ["SeriesCollection_", "series_collection_"], fn: (self: any, SeriesCollection_: any, series_collection_: any): any => {
        let plot: any, xChart: any;
        xChart = element("c:barChart");
        plot = $call(_BasePlot, xChart, null);
        return [plot, series_collection_, SeriesCollection_, xChart];
      } },
    vary_by_categories_get_fixture: { args: ["request"], params: [["c:barChart", true], ["c:lineChart/c:varyColors", true], ["c:pieChart/c:varyColors{val=0}", false]], fn: (self: any, request: any): any => {
        let expected_value: any, plot: any, xChart_cxml: any;
        [xChart_cxml, expected_value] = $list(request.param);
        plot = $call(_BasePlot, element(xChart_cxml), null);
        return [plot, expected_value];
      } },
    vary_by_categories_set_fixture: { args: ["request"], params: [["c:barChart", false, "c:barChart/c:varyColors{val=0}"], ["c:barChart", true, "c:barChart/c:varyColors"], ["c:lineChart/c:varyColors{val=0}", true, "c:lineChart/c:varyColors"], ["c:pieChart/c:varyColors{val=1}", false, "c:pieChart/c:varyColors{val=0}"]], fn: (self: any, request: any): any => {
        let expected_xChart_cxml: any, expected_xml: any, new_value: any, plot: any, xChart_cxml: any;
        [xChart_cxml, new_value, expected_xChart_cxml] = $list(request.param);
        plot = $call(_BasePlot, element(xChart_cxml), null);
        expected_xml = xml(expected_xChart_cxml);
        return [plot, new_value, expected_xml];
      } },
    Categories_: { args: ["request", "categories_"], fn: (self: any, request: any, categories_: any): any => {
        return $call(classMock, request, "pptx.chart.plot.Categories", $kw({return_value: categories_}));
      } },
    categories_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Categories);
      } },
    chart_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Chart);
      } },
    DataLabels_: { args: ["request", "data_labels_"], fn: (self: any, request: any, data_labels_: any): any => {
        return $call(classMock, request, "pptx.chart.plot.DataLabels", $kw({return_value: data_labels_}));
      } },
    data_labels_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, DataLabels);
      } },
    SeriesCollection_: { args: ["request", "series_collection_"], fn: (self: any, request: any, series_collection_: any): any => {
        return $call(classMock, request, "pptx.chart.plot.SeriesCollection", $kw({return_value: series_collection_}));
      } },
    series_collection_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SeriesCollection);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_which_chart_it_belongs_to", args: ["chart_fixture"], fn: (self: any, chart_fixture: any): any => {
        let expected_value: any, plot: any;
        [plot, expected_value] = $list(chart_fixture);
        $assertEqual(plot.chart, expected_value);
      } },
    { name: "it_knows_whether_it_has_data_labels", args: ["has_data_labels_get_fixture"], fn: (self: any, has_data_labels_get_fixture: any): any => {
        let expected_value: any, plot: any;
        [plot, expected_value] = $list(has_data_labels_get_fixture);
        $assertEqual(plot.hasDataLabels, expected_value);
      } },
    { name: "it_can_change_whether_it_has_data_labels", args: ["has_data_labels_set_fixture"], fn: (self: any, has_data_labels_set_fixture: any): any => {
        let expected_xml: any, new_value: any, plot: any;
        [plot, new_value, expected_xml] = $list(has_data_labels_set_fixture);
        plot.hasDataLabels = new_value;
        $assertEqual(plot._element.xml, expected_xml);
      } },
    { name: "it_knows_whether_it_varies_color_by_category", args: ["vary_by_categories_get_fixture"], fn: (self: any, vary_by_categories_get_fixture: any): any => {
        let expected_value: any, plot: any;
        [plot, expected_value] = $list(vary_by_categories_get_fixture);
        $assertEqual(plot.varyByCategories, expected_value);
      } },
    { name: "it_can_change_whether_it_varies_color_by_category", args: ["vary_by_categories_set_fixture"], fn: (self: any, vary_by_categories_set_fixture: any): any => {
        let expected_xml: any, new_value: any, plot: any;
        [plot, new_value, expected_xml] = $list(vary_by_categories_set_fixture);
        plot.varyByCategories = new_value;
        $assertEqual(plot._element.xml, expected_xml);
      } },
    { name: "it_provides_access_to_its_categories", args: ["categories_fixture"], fn: (self: any, categories_fixture: any): any => {
        let Categories_: any, categories: any, categories_: any, plot: any, xChart: any;
        [plot, categories_, Categories_, xChart] = $list(categories_fixture);
        categories = plot.categories;
        Categories_.assertCalledOnceWith(xChart);
        $assert($is(categories, categories_));
      } },
    { name: "it_provides_access_to_the_data_labels", args: ["data_labels_fixture"], fn: (self: any, data_labels_fixture: any): any => {
        let DataLabels_: any, dLbls: any, data_labels: any, data_labels_: any, plot: any;
        [plot, data_labels_, DataLabels_, dLbls] = $list(data_labels_fixture);
        data_labels = plot.dataLabels;
        DataLabels_.assertCalledOnceWith(dLbls);
        $assert($is(data_labels, data_labels_));
      } },
    { name: "it_provides_access_to_its_series", args: ["series_fixture"], fn: (self: any, series_fixture: any): any => {
        let SeriesCollection_: any, plot: any, series: any, series_: any, xChart: any;
        [plot, series_, SeriesCollection_, xChart] = $list(series_fixture);
        series = plot.series;
        SeriesCollection_.assertCalledOnceWith(xChart);
        $assert($is(series, series_));
      } },
  ],
};
pytestClass("Describe_BasePlot", Describe_BasePlot);

const DescribeBarPlot: any = {
  fixtures: {
    gap_width_get_fixture: { args: ["request"], params: [["c:barChart", 150], ["c:barChart/c:gapWidth", 150], ["c:barChart/c:gapWidth{val=175}", 175], ["c:barChart/c:gapWidth{val=042%}", 42]], fn: (self: any, request: any): any => {
        let barChart_cxml: any, bar_plot: any, expected_value: any;
        [barChart_cxml, expected_value] = $list(request.param);
        bar_plot = $call(BarPlot, element(barChart_cxml), null);
        return [bar_plot, expected_value];
      } },
    gap_width_set_fixture: { args: ["request"], params: [["c:barChart", 42, "c:barChart/c:gapWidth{val=42}"], ["c:barChart", 150, "c:barChart/c:gapWidth"], ["c:barChart/c:gapWidth{val=200%}", 300, "c:barChart/c:gapWidth{val=300}"]], fn: (self: any, request: any): any => {
        let barChart_cxml: any, bar_plot: any, expected_barChart_cxml: any, expected_xml: any, new_value: any;
        [barChart_cxml, new_value, expected_barChart_cxml] = $list(request.param);
        bar_plot = $call(BarPlot, element(barChart_cxml), null);
        expected_xml = xml(expected_barChart_cxml);
        return [bar_plot, new_value, expected_xml];
      } },
    overlap_get_fixture: { args: ["request"], params: [["c:barChart", 0], ["c:barChart/c:overlap", 0], ["c:barChart/c:overlap{val=42}", 42], ["c:barChart/c:overlap{val=-42}", -42], ["c:barChart/c:overlap{val=-042%}", -42]], fn: (self: any, request: any): any => {
        let barChart_cxml: any, bar_plot: any, expected_value: any;
        [barChart_cxml, expected_value] = $list(request.param);
        bar_plot = $call(BarPlot, element(barChart_cxml), null);
        return [bar_plot, expected_value];
      } },
    overlap_set_fixture: { args: ["request"], params: [["c:barChart", 42, "c:barChart/c:overlap{val=42}"], ["c:barChart/c:overlap{val=42}", 24, "c:barChart/c:overlap{val=24}"], ["c:barChart/c:overlap{val=42}", -5, "c:barChart/c:overlap{val=-5}"], ["c:barChart/c:overlap{val=42}", 0, "c:barChart"]], fn: (self: any, request: any): any => {
        let barChart_cxml: any, bar_plot: any, expected_barChart_cxml: any, expected_xml: any, new_value: any;
        [barChart_cxml, new_value, expected_barChart_cxml] = $list(request.param);
        bar_plot = $call(BarPlot, element(barChart_cxml), null);
        expected_xml = xml(expected_barChart_cxml);
        return [bar_plot, new_value, expected_xml];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_gap_width", args: ["gap_width_get_fixture"], fn: (self: any, gap_width_get_fixture: any): any => {
        let bar_plot: any, expected_value: any;
        [bar_plot, expected_value] = $list(gap_width_get_fixture);
        $assertEqual(bar_plot.gapWidth, expected_value);
      } },
    { name: "it_can_change_its_gap_width", args: ["gap_width_set_fixture"], fn: (self: any, gap_width_set_fixture: any): any => {
        let bar_plot: any, expected_xml: any, new_value: any;
        [bar_plot, new_value, expected_xml] = $list(gap_width_set_fixture);
        bar_plot.gapWidth = new_value;
        $assertEqual(bar_plot._element.xml, expected_xml);
      } },
    { name: "it_knows_how_much_it_overlaps_the_adjacent_bar", args: ["overlap_get_fixture"], fn: (self: any, overlap_get_fixture: any): any => {
        let bar_plot: any, expected_value: any;
        [bar_plot, expected_value] = $list(overlap_get_fixture);
        $assertEqual(bar_plot.overlap, expected_value);
      } },
    { name: "it_can_change_how_much_it_overlaps_the_adjacent_bar", args: ["overlap_set_fixture"], fn: (self: any, overlap_set_fixture: any): any => {
        let bar_plot: any, expected_xml: any, new_value: any;
        [bar_plot, new_value, expected_xml] = $list(overlap_set_fixture);
        bar_plot.overlap = new_value;
        $assertEqual(bar_plot._element.xml, expected_xml);
      } },
  ],
};
pytestClass("DescribeBarPlot", DescribeBarPlot);

const DescribeBubblePlot: any = {
  fixtures: {
    bubble_scale_get_fixture: { args: ["request"], params: [["c:bubbleChart", 100], ["c:bubbleChart/c:bubbleScale", 100], ["c:bubbleChart/c:bubbleScale{val=175}", 175], ["c:bubbleChart/c:bubbleScale{val=070%}", 70]], fn: (self: any, request: any): any => {
        let bubbleChart_cxml: any, bubble_plot: any, expected_value: any;
        [bubbleChart_cxml, expected_value] = $list(request.param);
        bubble_plot = $call(BubblePlot, element(bubbleChart_cxml), null);
        return [bubble_plot, expected_value];
      } },
    bubble_scale_set_fixture: { args: ["request"], params: [["c:bubbleChart", 42, "c:bubbleChart/c:bubbleScale{val=42}"], ["c:bubbleChart/c:bubbleScale{val=042%}", 150, "c:bubbleChart/c:bubbleScale{val=150}"], ["c:bubbleChart/c:bubbleScale{val=150}", null, "c:bubbleChart"], ["c:bubbleChart/c:bubbleScale{val=150}", 100, "c:bubbleChart/c:bubbleScale"], ["c:bubbleChart", null, "c:bubbleChart"], ["c:bubbleChart", 100, "c:bubbleChart/c:bubbleScale"]], fn: (self: any, request: any): any => {
        let bubbleChart_cxml: any, bubble_plot: any, expected_cxml: any, expected_xml: any, new_value: any;
        [bubbleChart_cxml, new_value, expected_cxml] = $list(request.param);
        bubble_plot = $call(BubblePlot, element(bubbleChart_cxml), null);
        expected_xml = xml(expected_cxml);
        return [bubble_plot, new_value, expected_xml];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_bubble_scale", args: ["bubble_scale_get_fixture"], fn: (self: any, bubble_scale_get_fixture: any): any => {
        let bubble_plot: any, expected_value: any;
        [bubble_plot, expected_value] = $list(bubble_scale_get_fixture);
        $assertEqual(bubble_plot.bubbleScale, expected_value);
      } },
    { name: "it_can_change_its_bubble_scale", args: ["bubble_scale_set_fixture"], fn: (self: any, bubble_scale_set_fixture: any): any => {
        let bubble_plot: any, expected_xml: any, new_value: any;
        [bubble_plot, new_value, expected_xml] = $list(bubble_scale_set_fixture);
        bubble_plot.bubbleScale = new_value;
        $assertEqual(bubble_plot._element.xml, expected_xml);
      } },
  ],
};
pytestClass("DescribeBubblePlot", DescribeBubblePlot);

const DescribePlotFactory: any = {
  fixtures: {
    call_fixture: { args: ["request", "chart_"], params: [["c:areaChart", AreaPlot], ["c:area3DChart", Area3DPlot], ["c:barChart", BarPlot], ["c:bubbleChart", BubblePlot], ["c:doughnutChart", DoughnutPlot], ["c:lineChart", LinePlot], ["c:pieChart", PiePlot], ["c:radarChart", RadarPlot], ["c:scatterChart", XyPlot]], fn: (self: any, request: any, chart_: any): any => {
        let PlotClass_: any, PlotCls: any, class_spec: any, plot_: any, xChart: any, xChart_cxml: any;
        [xChart_cxml, PlotCls] = $list(request.param);
        plot_ = $call(instanceMock, request, PlotCls, $kw({name: "plot_"}));
        class_spec = $mod("pptx.chart.plot.%s", PlotCls.name);
        PlotClass_ = $call(classMock, request, class_spec, $kw({return_value: plot_}));
        xChart = element(xChart_cxml);
        return [xChart, chart_, PlotClass_, plot_];
      } },
    chart_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Chart);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_contructs_a_plot_object_from_a_plot_element", args: ["call_fixture"], fn: (self: any, call_fixture: any): any => {
        let PlotClass_: any, chart_: any, plot: any, plot_: any, xChart: any;
        [xChart, chart_, PlotClass_, plot_] = $list(call_fixture);
        plot = $call(PlotFactory, xChart, chart_);
        PlotClass_.assertCalledOnceWith(xChart, chart_);
        $assert($is(plot, plot_));
      } },
  ],
};
pytestClass("DescribePlotFactory", DescribePlotFactory);

const DescribePlotTypeInspector: any = {
  fixtures: {
    chart_type_fixture: { args: ["request"], params: [["c:areaChart", XL.AREA], ["c:areaChart/c:grouping", XL.AREA], ["c:areaChart/c:grouping{val=standard}", XL.AREA], ["c:areaChart/c:grouping{val=stacked}", XL.AREA_STACKED], ["c:areaChart/c:grouping{val=percentStacked}", XL.AREA_STACKED_100], ["c:area3DChart", XL.THREE_D_AREA], ["c:area3DChart/c:grouping", XL.THREE_D_AREA], ["c:area3DChart/c:grouping{val=standard}", XL.THREE_D_AREA], ["c:area3DChart/c:grouping{val=stacked}", XL.THREE_D_AREA_STACKED], ["c:area3DChart/c:grouping{val=percentStacked}", XL.THREE_D_AREA_STACKED_100], ["c:barChart/c:barDir", XL.COLUMN_CLUSTERED], ["c:barChart/c:barDir{val=col}", XL.COLUMN_CLUSTERED], ["c:barChart/(c:barDir{val=col},c:grouping)", XL.COLUMN_CLUSTERED], ["c:barChart/(c:barDir{val=col},c:grouping{val=clustered})", XL.COLUMN_CLUSTERED], ["c:barChart/(c:barDir{val=col},c:grouping{val=stacked})", XL.COLUMN_STACKED], ["c:barChart/(c:barDir{val=col},c:grouping{val=percentStacked})", XL.COLUMN_STACKED_100], ["c:barChart/c:barDir{val=bar}", XL.BAR_CLUSTERED], ["c:barChart/(c:barDir{val=bar},c:grouping)", XL.BAR_CLUSTERED], ["c:barChart/(c:barDir{val=bar},c:grouping{val=clustered})", XL.BAR_CLUSTERED], ["c:barChart/(c:barDir{val=bar},c:grouping{val=stacked})", XL.BAR_STACKED], ["c:barChart/(c:barDir{val=bar},c:grouping{val=percentStacked})", XL.BAR_STACKED_100], ["c:doughnutChart", XL.DOUGHNUT], ["c:doughnutChart/c:ser/c:explosion{val=25}", XL.DOUGHNUT_EXPLODED], ["c:lineChart", XL.LINE_MARKERS], ["c:lineChart/c:grouping", XL.LINE_MARKERS], ["c:lineChart/c:grouping{val=standard}", XL.LINE_MARKERS], ["c:lineChart/c:grouping{val=stacked}", XL.LINE_MARKERS_STACKED], ["c:lineChart/c:grouping{val=percentStacked}", XL.LINE_MARKERS_STACKED_100], ["c:lineChart/c:ser/c:marker/c:symbol{val=none}", XL.LINE], ["c:lineChart/(c:grouping{val=stacked},c:ser/c:marker/c:symbol{val=none})", XL.LINE_STACKED], ["c:lineChart/(c:grouping{val=percentStacked},c:ser/c:marker/c:symbol{val=none})", XL.LINE_STACKED_100], ["c:pieChart", XL.PIE], ["c:pieChart/c:ser/c:explosion{val=25}", XL.PIE_EXPLODED], ["c:scatterChart/c:scatterStyle", XL.XY_SCATTER], ["c:scatterChart/(c:scatterStyle{val=lineMarker},c:ser/c:spPr/a:ln/a:noFill)", XL.XY_SCATTER], ["c:scatterChart/c:scatterStyle{val=lineMarker}", XL.XY_SCATTER_LINES], ["c:scatterChart/(c:scatterStyle{val=lineMarker},c:ser/c:marker/c:symbol{val=none})", XL.XY_SCATTER_LINES_NO_MARKERS], ["c:scatterChart/(c:scatterStyle{val=lineMarker},c:ser/c:marker/c:symbol{val=diamond})", XL.XY_SCATTER_LINES], ["c:scatterChart/c:scatterStyle{val=smoothMarker}", XL.XY_SCATTER_SMOOTH], ["c:scatterChart/(c:scatterStyle{val=smoothMarker},c:ser/c:marker/c:symbol{val=none})", XL.XY_SCATTER_SMOOTH_NO_MARKERS], ["c:bubbleChart", XL.BUBBLE], ["c:bubbleChart/c:ser", XL.BUBBLE], ["c:bubbleChart/c:ser/c:bubble3D", XL.BUBBLE_THREE_D_EFFECT], ["c:bubbleChart/c:ser/c:bubble3D{val=0}", XL.BUBBLE], ["c:bubbleChart/c:ser/c:bubble3D{val=1}", XL.BUBBLE_THREE_D_EFFECT], ["c:radarChart/c:radarStyle", XL.RADAR], ["c:radarChart/c:radarStyle{val=marker}", XL.RADAR_MARKERS], ["c:radarChart/c:radarStyle{val=filled}", XL.RADAR_FILLED], ["c:radarChart/(c:radarStyle{val=marker},c:ser/c:marker/c:symbol{val=none})", XL.RADAR]], fn: (self: any, request: any): any => {
        let expected_chart_type: any, plot: any, xChart_cxml: any;
        [xChart_cxml, expected_chart_type] = $list(request.param);
        plot = $call(PlotFactory, element(xChart_cxml), null);
        return [plot, expected_chart_type];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_determine_the_chart_type_of_a_plot", args: ["chart_type_fixture"], fn: (self: any, chart_type_fixture: any): any => {
        let chart_type: any, expected_chart_type: any, plot: any;
        [plot, expected_chart_type] = $list(chart_type_fixture);
        chart_type = PlotTypeInspector.chartType(plot);
        $assert($is(chart_type, expected_chart_type));
      } },
  ],
};
pytestClass("DescribePlotTypeInspector", DescribePlotTypeInspector);

