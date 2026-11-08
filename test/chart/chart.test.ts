// Machine-translated from python-pptx tests/chart/test_chart.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $append, $assert, $assertEqual, $call, $d, $get, $is, $kw, $len, $list, $mod, $raises, $truthy, call, classMock, functionMock, instanceMock, propertyMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { CategoryAxis as CategoryAxis$imported, DateAxis as DateAxis$imported, ValueAxis as ValueAxis$imported } from "../../src/chart/axis.js";
import { Chart as Chart$imported, ChartTitle as ChartTitle$imported, Legend as Legend$imported, _Plots as _Plots$imported } from "../../src/chart/chart.js";
import { ChartData as ChartData$imported } from "../../src/chart/data.js";
import { _BasePlot as _BasePlot$imported } from "../../src/chart/plot.js";
import { SeriesCollection as SeriesCollection$imported } from "../../src/chart/series.js";
import { _BaseSeriesXmlRewriter as _BaseSeriesXmlRewriter$imported } from "../../src/chart/xmlwriter.js";
import { ChartFormat as ChartFormat$imported } from "../../src/dml/chtfmt.js";
import { XL_CHART_TYPE as XL_CHART_TYPE$imported } from "../../src/enum/chart.js";
import { ChartWorkbook as ChartWorkbook$imported } from "../../src/parts/chart.js";
import { Font as Font$imported } from "../../src/text/text.js";
import { element, xml } from "../unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const CategoryAxis = CategoryAxis$imported;
const DateAxis = DateAxis$imported;
const ValueAxis = ValueAxis$imported;
const Chart = Chart$imported;
const ChartTitle = ChartTitle$imported;
const Legend = Legend$imported;
const _Plots = _Plots$imported;
const ChartData = ChartData$imported;
const _BasePlot = _BasePlot$imported;
const SeriesCollection = SeriesCollection$imported;
const _BaseSeriesXmlRewriter = _BaseSeriesXmlRewriter$imported;
const ChartFormat = ChartFormat$imported;
const XL_CHART_TYPE = XL_CHART_TYPE$imported;
const ChartWorkbook = ChartWorkbook$imported;
const Font = Font$imported;
void $py;


const DescribeChart: any = {
  fixtures: {
    category_axis_fixture: { args: ["request", "CategoryAxis_", "DateAxis_", "ValueAxis_"], params: ["c:catAx", "c:dateAx", "c:valAx"], fn: (self: any, request: any, CategoryAxis_: any, DateAxis_: any, ValueAxis_: any): any => {
        let AxisCls_: any, ax_tag: any, axis_: any, chart: any, chartSpace: any, chartSpace_cxml: any, xAx: any;
        ax_tag = request.param;
        chartSpace_cxml = $mod("c:chartSpace/c:chart/c:plotArea/%s", ax_tag);
        chartSpace = element(chartSpace_cxml);
        chart = $call(Chart, chartSpace, null);
        AxisCls_ = $get($d({"c:catAx": CategoryAxis_, "c:dateAx": DateAxis_, "c:valAx": ValueAxis_}), ax_tag);
        axis_ = AxisCls_.returnValue;
        xAx = $get(chartSpace.xpath($mod(".//%s", ax_tag)), 0);
        return [chart, axis_, AxisCls_, xAx];
      } },
    cat_ax_raise_fixture: { args: [], fn: (self: any): any => {
        let chart: any;
        chart = $call(Chart, element("c:chartSpace/c:chart/c:plotArea"), null);
        return chart;
      } },
    font_fixture: { args: ["request"], params: [["c:chartSpace{a:b=c}", "c:chartSpace{a:b=c}/c:txPr/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr)"], ["c:chartSpace/c:txPr/a:p", "c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr"], ["c:chartSpace/c:txPr/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr)", "c:chartSpace/c:txPr/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr)"]], fn: (self: any, request: any): any => {
        let chartSpace: any, chartSpace_cxml: any, expected_cxml: any, expected_xml: any;
        [chartSpace_cxml, expected_cxml] = $list(request.param);
        chartSpace = element(chartSpace_cxml);
        expected_xml = xml(expected_cxml);
        return [chartSpace, expected_xml];
      } },
    has_legend_get_fixture: { args: ["request"], params: [["c:chartSpace/c:chart", false], ["c:chartSpace/c:chart/c:legend", true]], fn: (self: any, request: any): any => {
        let chart: any, chartSpace_cxml: any, expected_value: any;
        [chartSpace_cxml, expected_value] = $list(request.param);
        chart = $call(Chart, element(chartSpace_cxml), null);
        return [chart, expected_value];
      } },
    has_legend_set_fixture: { args: ["request"], params: [["c:chartSpace/c:chart", true, "c:chartSpace/c:chart/c:legend"]], fn: (self: any, request: any): any => {
        let chart: any, chartSpace_cxml: any, expected_chartSpace_cxml: any, expected_xml: any, new_value: any;
        [chartSpace_cxml, new_value, expected_chartSpace_cxml] = $list(request.param);
        chart = $call(Chart, element(chartSpace_cxml), null);
        expected_xml = xml(expected_chartSpace_cxml);
        return [chart, new_value, expected_xml];
      } },
    has_title_get_fixture: { args: ["request"], params: [["c:chartSpace/c:chart", false], ["c:chartSpace/c:chart/c:title", true]], fn: (self: any, request: any): any => {
        let chart: any, chartSpace_cxml: any, expected_value: any;
        [chartSpace_cxml, expected_value] = $list(request.param);
        chart = $call(Chart, element(chartSpace_cxml), null);
        return [chart, expected_value];
      } },
    has_title_set_fixture: { args: ["request"], params: [["c:chart", true, "c:chart/c:title/(c:layout,c:overlay{val=0})"], ["c:chart/c:title", true, "c:chart/c:title"], ["c:chart/c:title", false, "c:chart/c:autoTitleDeleted{val=1}"], ["c:chart", false, "c:chart/c:autoTitleDeleted{val=1}"]], fn: (self: any, request: any): any => {
        let chart: any, chart_cxml: any, expected_cxml: any, expected_xml: any, new_value: any;
        [chart_cxml, new_value, expected_cxml] = $list(request.param);
        chart = $call(Chart, element($mod("c:chartSpace/%s", chart_cxml)), null);
        expected_xml = xml(expected_cxml);
        return [chart, new_value, expected_xml];
      } },
    legend_fixture: { args: ["request", "Legend_", "legend_"], params: [["c:chartSpace/c:chart", false], ["c:chartSpace/c:chart/c:legend", true]], fn: (self: any, request: any, Legend_: any, legend_: any): any => {
        let chart: any, chartSpace: any, chartSpace_cxml: any, expected_calls: any, expected_value: any, has_legend: any, legend_elm: any;
        [chartSpace_cxml, has_legend] = $list(request.param);
        chartSpace = element(chartSpace_cxml);
        chart = $call(Chart, chartSpace, null);
        [expected_value, expected_calls] = $list([null, []]);
        if ($truthy(has_legend)) {
          expected_value = legend_;
          legend_elm = chartSpace.chart.legend;
          $append(expected_calls, call(legend_elm));
        }
        return [chart, Legend_, expected_calls, expected_value];
      } },
    plots_fixture: { args: ["_Plots_", "plots_"], fn: (self: any, _Plots_: any, plots_: any): any => {
        let chart: any, chartSpace: any, plotArea: any;
        chartSpace = element("c:chartSpace/c:chart/c:plotArea");
        plotArea = $get(chartSpace.xpath("./c:chart/c:plotArea"), 0);
        chart = $call(Chart, chartSpace, null);
        return [chart, plots_, _Plots_, plotArea];
      } },
    replace_fixture: { args: ["chart_data_", "SeriesXmlRewriterFactory_", "series_rewriter_", "workbook_", "workbook_prop_"], fn: (self: any, chart_data_: any, SeriesXmlRewriterFactory_: any, series_rewriter_: any, workbook_: any, workbook_prop_: any): any => {
        let chart: any, chartSpace: any, chart_type: any, xlsx_blob: any;
        chartSpace = element("c:chartSpace/c:chart/c:plotArea/c:pieChart");
        chart = $call(Chart, chartSpace, null);
        chart_type = XL_CHART_TYPE.PIE;
        xlsx_blob = "fooblob";
        chart_data_.xlsxBlob = xlsx_blob;
        return [chart, chart_data_, SeriesXmlRewriterFactory_, chart_type, series_rewriter_, chartSpace, workbook_, xlsx_blob];
      } },
    series_fixture: { args: ["SeriesCollection_", "series_collection_"], fn: (self: any, SeriesCollection_: any, series_collection_: any): any => {
        let chart: any, chartSpace: any, plotArea: any;
        chartSpace = element("c:chartSpace/c:chart/c:plotArea");
        plotArea = $get(chartSpace.xpath(".//c:plotArea"), 0);
        chart = $call(Chart, chartSpace, null);
        return [chart, SeriesCollection_, plotArea, series_collection_];
      } },
    style_get_fixture: { args: ["request"], params: [["c:chartSpace/c:style{val=42}", 42], ["c:chartSpace", null]], fn: (self: any, request: any): any => {
        let chart: any, chartSpace_cxml: any, expected_value: any;
        [chartSpace_cxml, expected_value] = $list(request.param);
        chart = $call(Chart, element(chartSpace_cxml), null);
        return [chart, expected_value];
      } },
    style_set_fixture: { args: ["request"], params: [["c:chartSpace", 4, "c:chartSpace/c:style{val=4}"], ["c:chartSpace", null, "c:chartSpace"], ["c:chartSpace/c:style{val=4}", 2, "c:chartSpace/c:style{val=2}"], ["c:chartSpace/c:style{val=4}", null, "c:chartSpace"]], fn: (self: any, request: any): any => {
        let chart: any, chartSpace_cxml: any, expected_chartSpace_cxml: any, expected_xml: any, new_value: any;
        [chartSpace_cxml, new_value, expected_chartSpace_cxml] = $list(request.param);
        chart = $call(Chart, element(chartSpace_cxml), null);
        expected_xml = xml(expected_chartSpace_cxml);
        return [chart, new_value, expected_xml];
      } },
    title_fixture: { args: ["request", "ChartTitle_", "chart_title_"], params: [["c:chartSpace/c:chart", "c:title/(c:layout,c:overlay{val=0})"], ["c:chartSpace/c:chart/c:title/c:layout", "c:title/c:layout"]], fn: (self: any, request: any, ChartTitle_: any, chart_title_: any): any => {
        let chart: any, chartSpace_cxml: any, expected_cxml: any, expected_xml: any;
        [chartSpace_cxml, expected_cxml] = $list(request.param);
        chart = $call(Chart, element(chartSpace_cxml), null);
        expected_xml = xml(expected_cxml);
        return [chart, expected_xml, ChartTitle_, chart_title_];
      } },
    val_ax_fixture: { args: ["request", "ValueAxis_", "value_axis_"], params: [["c:chartSpace/c:chart/c:plotArea/(c:catAx,c:valAx)", 0], ["c:chartSpace/c:chart/c:plotArea/(c:valAx,c:valAx)", 1]], fn: (self: any, request: any, ValueAxis_: any, value_axis_: any): any => {
        let chart: any, chartSpace: any, chartSpace_xml: any, idx: any, valAx: any;
        [chartSpace_xml, idx] = $list(request.param);
        chartSpace = element(chartSpace_xml);
        chart = $call(Chart, chartSpace, null);
        valAx = $get(chartSpace.xpath(".//c:valAx"), idx);
        return [chart, ValueAxis_, valAx, value_axis_];
      } },
    val_ax_raise_fixture: { args: [], fn: (self: any): any => {
        let chart: any;
        chart = $call(Chart, element("c:chartSpace/c:chart/c:plotArea"), null);
        return chart;
      } },
    CategoryAxis_: { args: ["request", "category_axis_"], fn: (self: any, request: any, category_axis_: any): any => {
        return $call(classMock, request, "pptx.chart.chart.CategoryAxis", $kw({return_value: category_axis_}));
      } },
    category_axis_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, CategoryAxis);
      } },
    chart_data_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, ChartData);
      } },
    ChartTitle_: { args: ["request", "chart_title_"], fn: (self: any, request: any, chart_title_: any): any => {
        return $call(classMock, request, "pptx.chart.chart.ChartTitle", $kw({return_value: chart_title_}));
      } },
    chart_title_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, ChartTitle);
      } },
    DateAxis_: { args: ["request", "date_axis_"], fn: (self: any, request: any, date_axis_: any): any => {
        return $call(classMock, request, "pptx.chart.chart.DateAxis", $kw({return_value: date_axis_}));
      } },
    date_axis_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, DateAxis);
      } },
    Font_: { args: ["request"], fn: (self: any, request: any): any => {
        return classMock(request, "pptx.chart.chart.Font");
      } },
    font_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Font);
      } },
    Legend_: { args: ["request", "legend_"], fn: (self: any, request: any, legend_: any): any => {
        return $call(classMock, request, "pptx.chart.chart.Legend", $kw({return_value: legend_}));
      } },
    legend_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Legend);
      } },
    PlotTypeInspector_: { args: ["request"], fn: (self: any, request: any): any => {
        return classMock(request, "pptx.chart.chart.PlotTypeInspector");
      } },
    _Plots_: { args: ["request", "plots_"], fn: (self: any, request: any, plots_: any): any => {
        return $call(classMock, request, "pptx.chart.chart._Plots", $kw({return_value: plots_}));
      } },
    plot_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _BasePlot);
      } },
    plots_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _Plots);
      } },
    SeriesCollection_: { args: ["request", "series_collection_"], fn: (self: any, request: any, series_collection_: any): any => {
        return $call(classMock, request, "pptx.chart.chart.SeriesCollection", $kw({return_value: series_collection_}));
      } },
    SeriesXmlRewriterFactory_: { args: ["request", "series_rewriter_"], fn: (self: any, request: any, series_rewriter_: any): any => {
        return $call(functionMock, request, "pptx.chart.chart.SeriesXmlRewriterFactory", $kw({return_value: series_rewriter_, autospec: true}));
      } },
    series_collection_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, SeriesCollection);
      } },
    series_rewriter_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _BaseSeriesXmlRewriter);
      } },
    ValueAxis_: { args: ["request", "value_axis_"], fn: (self: any, request: any, value_axis_: any): any => {
        return $call(classMock, request, "pptx.chart.chart.ValueAxis", $kw({return_value: value_axis_}));
      } },
    value_axis_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, ValueAxis);
      } },
    workbook_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, ChartWorkbook);
      } },
    workbook_prop_: { args: ["request", "workbook_"], fn: (self: any, request: any, workbook_: any): any => {
        return $call(propertyMock, request, Chart, "_workbook", $kw({return_value: workbook_}));
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_access_to_its_font", args: ["font_fixture", "Font_", "font_"], fn: (self: any, font_fixture: any, Font_: any, font_: any): any => {
        let chart: any, chartSpace: any, expected_xml: any, font: any;
        [chartSpace, expected_xml] = $list(font_fixture);
        Font_.returnValue = font_;
        chart = $call(Chart, chartSpace, null);
        font = chart.font;
        $assertEqual(chartSpace.xml, expected_xml);
        Font_.assertCalledOnceWith($get(chartSpace.xpath("./c:txPr/a:p/a:pPr/a:defRPr"), 0));
        $assert($is(font, font_));
      } },
    { name: "it_knows_whether_it_has_a_title", args: ["has_title_get_fixture"], fn: (self: any, has_title_get_fixture: any): any => {
        let chart: any, expected_value: any;
        [chart, expected_value] = $list(has_title_get_fixture);
        $assert($is(chart.hasTitle, expected_value));
      } },
    { name: "it_can_change_whether_it_has_a_title", args: ["has_title_set_fixture"], fn: (self: any, has_title_set_fixture: any): any => {
        let chart: any, expected_xml: any, new_value: any;
        [chart, new_value, expected_xml] = $list(has_title_set_fixture);
        chart.hasTitle = new_value;
        $assertEqual(chart._chartSpace.chart.xml, expected_xml);
      } },
    { name: "it_provides_access_to_the_chart_title", args: ["title_fixture"], fn: (self: any, title_fixture: any): any => {
        let ChartTitle_: any, chart: any, chart_title: any, chart_title_: any, expected_xml: any;
        [chart, expected_xml, ChartTitle_, chart_title_] = $list(title_fixture);
        chart_title = chart.chartTitle;
        $assertEqual($get(chart.element.xpath("c:chart/c:title"), 0).xml, expected_xml);
        ChartTitle_.assertCalledOnceWith(chart.element.chart.title);
        $assert($is(chart_title, chart_title_));
      } },
    { name: "it_provides_access_to_the_category_axis", args: ["category_axis_fixture"], fn: (self: any, category_axis_fixture: any): any => {
        let AxisCls_: any, category_axis: any, category_axis_: any, chart: any, xAx: any;
        [chart, category_axis_, AxisCls_, xAx] = $list(category_axis_fixture);
        category_axis = chart.categoryAxis;
        AxisCls_.assertCalledOnceWith(xAx);
        $assert($is(category_axis, category_axis_));
      } },
    { name: "it_raises_when_no_category_axis", args: ["cat_ax_raise_fixture"], fn: (self: any, cat_ax_raise_fixture: any): any => {
        let chart: any;
        chart = cat_ax_raise_fixture;
        $raises(ValueError, () => {
          chart.categoryAxis;
        });
      } },
    { name: "it_provides_access_to_the_value_axis", args: ["val_ax_fixture"], fn: (self: any, val_ax_fixture: any): any => {
        let ValueAxis_: any, chart: any, valAx: any, value_axis: any, value_axis_: any;
        [chart, ValueAxis_, valAx, value_axis_] = $list(val_ax_fixture);
        value_axis = chart.valueAxis;
        ValueAxis_.assertCalledOnceWith(valAx);
        $assert($is(value_axis, value_axis_));
      } },
    { name: "it_raises_when_no_value_axis", args: ["val_ax_raise_fixture"], fn: (self: any, val_ax_raise_fixture: any): any => {
        let chart: any;
        chart = val_ax_raise_fixture;
        $raises(ValueError, () => {
          chart.valueAxis;
        });
      } },
    { name: "it_provides_access_to_its_series", args: ["series_fixture"], fn: (self: any, series_fixture: any): any => {
        let SeriesCollection_: any, chart: any, plotArea: any, series: any, series_: any;
        [chart, SeriesCollection_, plotArea, series_] = $list(series_fixture);
        series = chart.series;
        SeriesCollection_.assertCalledOnceWith(plotArea);
        $assert($is(series, series_));
      } },
    { name: "it_provides_access_to_its_plots", args: ["plots_fixture"], fn: (self: any, plots_fixture: any): any => {
        let _Plots_: any, chart: any, plotArea: any, plots: any, plots_: any;
        [chart, plots_, _Plots_, plotArea] = $list(plots_fixture);
        plots = chart.plots;
        _Plots_.assertCalledOnceWith(plotArea, chart);
        $assert($is(plots, plots_));
      } },
    { name: "it_knows_whether_it_has_a_legend", args: ["has_legend_get_fixture"], fn: (self: any, has_legend_get_fixture: any): any => {
        let chart: any, expected_value: any;
        [chart, expected_value] = $list(has_legend_get_fixture);
        $assertEqual(chart.hasLegend, expected_value);
      } },
    { name: "it_can_change_whether_it_has_a_legend", args: ["has_legend_set_fixture"], fn: (self: any, has_legend_set_fixture: any): any => {
        let chart: any, expected_xml: any, new_value: any;
        [chart, new_value, expected_xml] = $list(has_legend_set_fixture);
        chart.hasLegend = new_value;
        $assertEqual(chart._chartSpace.xml, expected_xml);
      } },
    { name: "it_provides_access_to_its_legend", args: ["legend_fixture"], fn: (self: any, legend_fixture: any): any => {
        let Legend_: any, chart: any, expected_calls: any, expected_value: any, legend: any;
        [chart, Legend_, expected_calls, expected_value] = $list(legend_fixture);
        legend = chart.legend;
        $assertEqual(Legend_.callArgsList, expected_calls);
        $assert($is(legend, expected_value));
      } },
    { name: "it_knows_its_chart_type", args: ["request", "PlotTypeInspector_", "plot_"], fn: (self: any, request: any, PlotTypeInspector_: any, plot_: any): any => {
        let chart: any, chart_type: any;
        $call(propertyMock, request, Chart, "plots", $kw({return_value: [plot_]}));
        PlotTypeInspector_.chartType.returnValue = XL_CHART_TYPE.PIE;
        chart = $call(Chart, null, null);
        chart_type = chart.chartType;
        PlotTypeInspector_.chartType.assertCalledOnceWith(plot_);
        $assertEqual(chart_type, XL_CHART_TYPE.PIE);
      } },
    { name: "it_knows_its_style", args: ["style_get_fixture"], fn: (self: any, style_get_fixture: any): any => {
        let chart: any, expected_value: any;
        [chart, expected_value] = $list(style_get_fixture);
        $assertEqual(chart.chartStyle, expected_value);
      } },
    { name: "it_can_change_its_style", args: ["style_set_fixture"], fn: (self: any, style_set_fixture: any): any => {
        let chart: any, expected_xml: any, new_value: any;
        [chart, new_value, expected_xml] = $list(style_set_fixture);
        chart.chartStyle = new_value;
        $assertEqual(chart._chartSpace.xml, expected_xml);
      } },
    { name: "it_can_replace_the_chart_data", args: ["replace_fixture"], fn: (self: any, replace_fixture: any): any => {
        let SeriesXmlRewriterFactory_: any, chart: any, chartSpace: any, chart_data_: any, chart_type: any, rewriter_: any, workbook_: any, xlsx_blob: any;
        [chart, chart_data_, SeriesXmlRewriterFactory_, chart_type, rewriter_, chartSpace, workbook_, xlsx_blob] = $list(replace_fixture);
        chart.replaceData(chart_data_);
        SeriesXmlRewriterFactory_.assertCalledOnceWith(chart_type, chart_data_);
        rewriter_.replaceSeriesData.assertCalledOnceWith(chartSpace);
        workbook_.updateFromXlsxBlob.assertCalledOnceWith(xlsx_blob);
      } },
  ],
};
pytestClass("DescribeChart", DescribeChart);

const DescribeChartTitle: any = {
  fixtures: {
    format_fixture: { args: ["request", "ChartFormat_", "format_"], fn: (self: any, request: any, ChartFormat_: any, format_: any): any => {
        let chart_title: any;
        chart_title = $call(ChartTitle, element("c:title"));
        return [chart_title, ChartFormat_, format_];
      } },
    has_tf_get_fixture: { args: ["request"], params: [["c:title", false], ["c:title/c:tx", false], ["c:title/c:tx/c:strRef", false], ["c:title/c:tx/c:rich", true]], fn: (self: any, request: any): any => {
        let chart_title: any, expected_value: any, title_cxml: any;
        [title_cxml, expected_value] = $list(request.param);
        chart_title = $call(ChartTitle, element(title_cxml));
        return [chart_title, expected_value];
      } },
    has_tf_set_fixture: { args: ["request"], params: [["c:title{a:b=c}", true, "c:title{a:b=c}/c:tx/c:rich/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr)"], ["c:title{a:b=c}/c:tx", true, "c:title{a:b=c}/c:tx/c:rich/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr)"], ["c:title{a:b=c}/c:tx/c:strRef", true, "c:title{a:b=c}/c:tx/c:rich/(a:bodyPr,a:lstStyle,a:p/a:pPr/a:defRPr)"], ["c:title/c:tx/c:rich", true, "c:title/c:tx/c:rich"], ["c:title", false, "c:title"], ["c:title/c:tx", false, "c:title"], ["c:title/c:tx/c:rich", false, "c:title"], ["c:title/c:tx/c:strRef", false, "c:title"]], fn: (self: any, request: any): any => {
        let chart_title: any, expected_cxml: any, expected_xml: any, title_cxml: any, value: any;
        [title_cxml, value, expected_cxml] = $list(request.param);
        chart_title = $call(ChartTitle, element(title_cxml));
        expected_xml = xml(expected_cxml);
        return [chart_title, value, expected_xml];
      } },
    text_frame_fixture: { args: ["request", "TextFrame_"], fn: (self: any, request: any, TextFrame_: any): any => {
        let chart_title: any, text_frame_: any;
        chart_title = $call(ChartTitle, element("c:title"));
        text_frame_ = TextFrame_.returnValue;
        return [chart_title, TextFrame_, text_frame_];
      } },
    ChartFormat_: { args: ["request", "format_"], fn: (self: any, request: any, format_: any): any => {
        return $call(classMock, request, "pptx.chart.chart.ChartFormat", $kw({return_value: format_}));
      } },
    format_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, ChartFormat);
      } },
    TextFrame_: { args: ["request"], fn: (self: any, request: any): any => {
        return classMock(request, "pptx.chart.chart.TextFrame");
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_access_to_its_format", args: ["format_fixture"], fn: (self: any, format_fixture: any): any => {
        let ChartFormat_: any, chart_title: any, format: any, format_: any;
        [chart_title, ChartFormat_, format_] = $list(format_fixture);
        format = chart_title.format;
        ChartFormat_.assertCalledOnceWith(chart_title.element);
        $assert($is(format, format_));
      } },
    { name: "it_knows_whether_it_has_a_text_frame", args: ["has_tf_get_fixture"], fn: (self: any, has_tf_get_fixture: any): any => {
        let chart_title: any, expected_value: any, value: any;
        [chart_title, expected_value] = $list(has_tf_get_fixture);
        value = chart_title.hasTextFrame;
        $assert($is(value, expected_value));
      } },
    { name: "it_can_change_whether_it_has_a_text_frame", args: ["has_tf_set_fixture"], fn: (self: any, has_tf_set_fixture: any): any => {
        let chart_title: any, expected_xml: any, value: any;
        [chart_title, value, expected_xml] = $list(has_tf_set_fixture);
        chart_title.hasTextFrame = value;
        $assertEqual(chart_title._element.xml, expected_xml);
      } },
    { name: "it_provides_access_to_its_text_frame", args: ["text_frame_fixture"], fn: (self: any, text_frame_fixture: any): any => {
        let TextFrame_: any, chart_title: any, text_frame: any, text_frame_: any;
        [chart_title, TextFrame_, text_frame_] = $list(text_frame_fixture);
        text_frame = chart_title.textFrame;
        TextFrame_.assertCalledOnceWith(chart_title._element.tx.rich, chart_title);
        $assert($is(text_frame, text_frame_));
      } },
  ],
};
pytestClass("DescribeChartTitle", DescribeChartTitle);

const Describe_Plots: any = {
  fixtures: {
    getitem_fixture: { args: ["request", "PlotFactory_", "chart_", "plot_"], params: [["c:plotArea/c:barChart", 0], ["c:plotArea/(c:radarChart,c:barChart)", 1]], fn: (self: any, request: any, PlotFactory_: any, chart_: any, plot_: any): any => {
        let idx: any, plotArea: any, plotArea_cxml: any, plot_elm: any, plots: any;
        [plotArea_cxml, idx] = $list(request.param);
        plotArea = element(plotArea_cxml);
        plot_elm = $get(plotArea, idx);
        plots = $call(_Plots, plotArea, chart_);
        return [plots, idx, PlotFactory_, plot_elm, chart_, plot_];
      } },
    len_fixture: { args: ["request"], params: [["c:plotArea", 0], ["c:plotArea/c:barChart", 1], ["c:plotArea/(c:barChart,c:lineChart)", 2]], fn: (self: any, request: any): any => {
        let expected_len: any, plotArea_cxml: any, plots: any;
        [plotArea_cxml, expected_len] = $list(request.param);
        plots = $call(_Plots, element(plotArea_cxml), null);
        return [plots, expected_len];
      } },
    chart_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Chart);
      } },
    PlotFactory_: { args: ["request", "plot_"], fn: (self: any, request: any, plot_: any): any => {
        return $call(functionMock, request, "pptx.chart.chart.PlotFactory", $kw({return_value: plot_}));
      } },
    plot_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _BasePlot);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_supports_indexed_access", args: ["getitem_fixture"], fn: (self: any, getitem_fixture: any): any => {
        let PlotFactory_: any, chart_: any, idx: any, plot: any, plot_: any, plot_elm: any, plots: any;
        [plots, idx, PlotFactory_, plot_elm, chart_, plot_] = $list(getitem_fixture);
        plot = $get(plots, idx);
        PlotFactory_.assertCalledOnceWith(plot_elm, chart_);
        $assert($is(plot, plot_));
      } },
    { name: "it_supports_len", args: ["len_fixture"], fn: (self: any, len_fixture: any): any => {
        let expected_len: any, plots: any;
        [plots, expected_len] = $list(len_fixture);
        $assertEqual($len(plots), expected_len);
      } },
  ],
};
pytestClass("Describe_Plots", Describe_Plots);

