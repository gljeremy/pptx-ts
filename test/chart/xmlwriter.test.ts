// Machine-translated from python-pptx tests/chart/test_xmlwriter.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $add, $assert, $assertEqual, $call, $d, $get, $getattr, $is, $iter, $kw, $len, $list, $m, $mod, $mul, $range, $round, $slice, $truthy, call, classMock, instanceMock, methodMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { BubbleChartData as BubbleChartData$imported, CategoryChartData as CategoryChartData$imported, CategorySeriesData as CategorySeriesData$imported, XyChartData as XyChartData$imported, _BaseChartData as _BaseChartData$imported, _BaseSeriesData as _BaseSeriesData$imported } from "../../src/chart/data.js";
import { ChartXmlWriter as ChartXmlWriter$imported, SeriesXmlRewriterFactory as SeriesXmlRewriterFactory$imported, _AreaChartXmlWriter as _AreaChartXmlWriter$imported, _BarChartXmlWriter as _BarChartXmlWriter$imported, _BaseSeriesXmlRewriter as _BaseSeriesXmlRewriter$imported, _BubbleChartXmlWriter as _BubbleChartXmlWriter$imported, _BubbleSeriesXmlRewriter as _BubbleSeriesXmlRewriter$imported, _BubbleSeriesXmlWriter as _BubbleSeriesXmlWriter$imported, _CategorySeriesXmlRewriter as _CategorySeriesXmlRewriter$imported, _CategorySeriesXmlWriter as _CategorySeriesXmlWriter$imported, _DoughnutChartXmlWriter as _DoughnutChartXmlWriter$imported, _LineChartXmlWriter as _LineChartXmlWriter$imported, _PieChartXmlWriter as _PieChartXmlWriter$imported, _RadarChartXmlWriter as _RadarChartXmlWriter$imported, _XyChartXmlWriter as _XyChartXmlWriter$imported, _XySeriesXmlRewriter as _XySeriesXmlRewriter$imported, _XySeriesXmlWriter as _XySeriesXmlWriter$imported } from "../../src/chart/xmlwriter.js";
import { XL_CHART_TYPE as XL_CHART_TYPE$imported } from "../../src/enum/chart.js";
import { parseXml as parseXml$imported } from "../../src/oxml/index.js";
import { element, xml } from "../unitutil/cxml.js";
import { snippetSeq, snippetText } from "../unitutil/file.js";
import { count } from "../unitutil/index.js";
// python binds imported names at import time; patches made later don't change them
const BubbleChartData = BubbleChartData$imported;
const CategoryChartData = CategoryChartData$imported;
const CategorySeriesData = CategorySeriesData$imported;
const XyChartData = XyChartData$imported;
const _BaseChartData = _BaseChartData$imported;
const _BaseSeriesData = _BaseSeriesData$imported;
const ChartXmlWriter = ChartXmlWriter$imported;
const SeriesXmlRewriterFactory = SeriesXmlRewriterFactory$imported;
const _AreaChartXmlWriter = _AreaChartXmlWriter$imported;
const _BarChartXmlWriter = _BarChartXmlWriter$imported;
const _BaseSeriesXmlRewriter = _BaseSeriesXmlRewriter$imported;
const _BubbleChartXmlWriter = _BubbleChartXmlWriter$imported;
const _BubbleSeriesXmlRewriter = _BubbleSeriesXmlRewriter$imported;
const _BubbleSeriesXmlWriter = _BubbleSeriesXmlWriter$imported;
const _CategorySeriesXmlRewriter = _CategorySeriesXmlRewriter$imported;
const _CategorySeriesXmlWriter = _CategorySeriesXmlWriter$imported;
const _DoughnutChartXmlWriter = _DoughnutChartXmlWriter$imported;
const _LineChartXmlWriter = _LineChartXmlWriter$imported;
const _PieChartXmlWriter = _PieChartXmlWriter$imported;
const _RadarChartXmlWriter = _RadarChartXmlWriter$imported;
const _XyChartXmlWriter = _XyChartXmlWriter$imported;
const _XySeriesXmlRewriter = _XySeriesXmlRewriter$imported;
const _XySeriesXmlWriter = _XySeriesXmlWriter$imported;
const XL_CHART_TYPE = XL_CHART_TYPE$imported;
const parseXml = parseXml$imported;
const date = $py.$datetime.date;
const islice = $py.$itertools.islice;
void $py;


const DescribeChartXmlWriter: any = {
  fixtures: {
    call_fixture: { args: ["request", "series_seq_"], params: [["AREA", _AreaChartXmlWriter], ["AREA_STACKED", _AreaChartXmlWriter], ["AREA_STACKED_100", _AreaChartXmlWriter], ["BAR_CLUSTERED", _BarChartXmlWriter], ["BAR_STACKED", _BarChartXmlWriter], ["BAR_STACKED_100", _BarChartXmlWriter], ["BUBBLE", _BubbleChartXmlWriter], ["BUBBLE_THREE_D_EFFECT", _BubbleChartXmlWriter], ["COLUMN_CLUSTERED", _BarChartXmlWriter], ["COLUMN_STACKED", _BarChartXmlWriter], ["COLUMN_STACKED_100", _BarChartXmlWriter], ["DOUGHNUT", _DoughnutChartXmlWriter], ["DOUGHNUT_EXPLODED", _DoughnutChartXmlWriter], ["LINE", _LineChartXmlWriter], ["LINE_MARKERS", _LineChartXmlWriter], ["LINE_MARKERS_STACKED", _LineChartXmlWriter], ["LINE_MARKERS_STACKED_100", _LineChartXmlWriter], ["LINE_STACKED", _LineChartXmlWriter], ["LINE_STACKED_100", _LineChartXmlWriter], ["PIE", _PieChartXmlWriter], ["PIE_EXPLODED", _PieChartXmlWriter], ["RADAR", _RadarChartXmlWriter], ["RADAR_FILLED", _RadarChartXmlWriter], ["RADAR_MARKERS", _RadarChartXmlWriter], ["XY_SCATTER", _XyChartXmlWriter], ["XY_SCATTER_LINES", _XyChartXmlWriter], ["XY_SCATTER_LINES_NO_MARKERS", _XyChartXmlWriter], ["XY_SCATTER_SMOOTH", _XyChartXmlWriter], ["XY_SCATTER_SMOOTH_NO_MARKERS", _XyChartXmlWriter]], fn: (self: any, request: any, series_seq_: any): any => {
        let XmlWriterClass: any, XmlWriterClass_: any, chart_type: any, chart_type_member: any, class_spec: any, xml_writer_: any;
        [chart_type_member, XmlWriterClass] = $list(request.param);
        xml_writer_ = instanceMock(request, XmlWriterClass);
        class_spec = $mod("pptx.chart.xmlwriter.%s", XmlWriterClass.name);
        XmlWriterClass_ = $call(classMock, request, class_spec, $kw({return_value: xml_writer_}));
        chart_type = $getattr(XL_CHART_TYPE, chart_type_member);
        return [chart_type, series_seq_, XmlWriterClass_, xml_writer_];
      } },
    series_seq_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Array);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_contructs_an_xml_writer_for_a_chart_type", args: ["call_fixture"], fn: (self: any, call_fixture: any): any => {
        let XmlWriterClass_: any, chart_type: any, series_seq_: any, xml_writer: any, xml_writer_: any;
        [chart_type, series_seq_, XmlWriterClass_, xml_writer_] = $list(call_fixture);
        xml_writer = $call(ChartXmlWriter, chart_type, series_seq_);
        XmlWriterClass_.assertCalledOnceWith(chart_type, series_seq_);
        $assert($is(xml_writer, xml_writer_));
      } },
  ],
};
pytestClass("DescribeChartXmlWriter", DescribeChartXmlWriter);

const DescribeSeriesXmlRewriterFactory: any = {
  fixtures: {
    call_fixture: { args: ["request", "chart_data_"], params: [["BAR_CLUSTERED", _CategorySeriesXmlRewriter], ["BUBBLE", _BubbleSeriesXmlRewriter], ["BUBBLE_THREE_D_EFFECT", _BubbleSeriesXmlRewriter], ["XY_SCATTER", _XySeriesXmlRewriter], ["XY_SCATTER_LINES", _XySeriesXmlRewriter], ["XY_SCATTER_LINES_NO_MARKERS", _XySeriesXmlRewriter], ["XY_SCATTER_SMOOTH", _XySeriesXmlRewriter], ["XY_SCATTER_SMOOTH_NO_MARKERS", _XySeriesXmlRewriter]], fn: (self: any, request: any, chart_data_: any): any => {
        let XmlRewriterClass_: any, chart_type: any, chart_type_member: any, rewriter_cls: any, xml_rewriter_: any;
        [chart_type_member, rewriter_cls] = $list(request.param);
        chart_type = $getattr(XL_CHART_TYPE, chart_type_member);
        xml_rewriter_ = instanceMock(request, rewriter_cls);
        XmlRewriterClass_ = $call(classMock, request, $mod("pptx.chart.xmlwriter.%s", rewriter_cls.name), $kw({return_value: xml_rewriter_}));
        return [chart_type, chart_data_, XmlRewriterClass_, xml_rewriter_];
      } },
    chart_data_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _BaseChartData);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_contructs_an_xml_rewriter_for_a_chart_type", args: ["call_fixture"], fn: (self: any, call_fixture: any): any => {
        let XmlRewriterClass_: any, chart_data_: any, chart_type: any, xml_rewriter: any, xml_rewriter_: any;
        [chart_type, chart_data_, XmlRewriterClass_, xml_rewriter_] = $list(call_fixture);
        xml_rewriter = $call(SeriesXmlRewriterFactory, chart_type, chart_data_);
        XmlRewriterClass_.assertCalledOnceWith(chart_data_);
        $assert($is(xml_rewriter, xml_rewriter_));
      } },
  ],
};
pytestClass("DescribeSeriesXmlRewriterFactory", DescribeSeriesXmlRewriterFactory);

const Describe_AreaChartXmlWriter: any = {
  fixtures: {
    xml_fixture: { args: ["request"], params: [["AREA", 2, 2, String, "2x2-area"], ["AREA", 2, 2, date, "2x2-area-date"], ["AREA", 2, 2, Number, "2x2-area-float"], ["AREA_STACKED", 2, 2, String, "2x2-area-stacked"], ["AREA_STACKED_100", 2, 2, String, "2x2-area-stacked-100"]], fn: (self: any, request: any): any => {
        let cat_count: any, cat_type: any, chart_data: any, chart_type: any, expected_xml: any, member: any, ser_count: any, snippet_name: any, xml_writer: any;
        [member, cat_count, ser_count, cat_type, snippet_name] = $list(request.param);
        chart_type = $getattr(XL_CHART_TYPE, member);
        chart_data = make_category_chart_data(cat_count, cat_type, ser_count);
        xml_writer = $call(_AreaChartXmlWriter, chart_type, chart_data);
        expected_xml = snippetText(snippet_name);
        return [xml_writer, expected_xml];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_generate_xml_for_area_type_charts", args: ["xml_fixture"], fn: (self: any, xml_fixture: any): any => {
        let expected_xml: any, xml_writer: any;
        [xml_writer, expected_xml] = $list(xml_fixture);
        $assertEqual(xml_writer.xml, expected_xml);
      } },
  ],
};
pytestClass("Describe_AreaChartXmlWriter", Describe_AreaChartXmlWriter);

const Describe_BarChartXmlWriter: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_can_generate_xml_for_bar_type_charts", args: ["member", "cat_count", "ser_count", "cat_type", "snippet_name"], parametrize: [{ names: ["member", "cat_count", "ser_count", "cat_type", "snippet_name"], values: [["BAR_CLUSTERED", 2, 2, String, "2x2-bar-clustered"], ["BAR_CLUSTERED", 2, 2, date, "2x2-bar-clustered-date"], ["BAR_CLUSTERED", 2, 2, Number, "2x2-bar-clustered-float"], ["BAR_STACKED", 2, 2, String, "2x2-bar-stacked"], ["BAR_STACKED_100", 2, 2, String, "2x2-bar-stacked-100"], ["COLUMN_CLUSTERED", 2, 2, String, "2x2-column-clustered"], ["COLUMN_STACKED", 2, 2, String, "2x2-column-stacked"], ["COLUMN_STACKED_100", 2, 2, String, "2x2-column-stacked-100"]] }], fn: (self: any, member: any, cat_count: any, ser_count: any, cat_type: any, snippet_name: any): any => {
        let chart_data: any, chart_type: any, xml_writer: any;
        chart_type = $getattr(XL_CHART_TYPE, member);
        chart_data = make_category_chart_data(cat_count, cat_type, ser_count);
        xml_writer = $call(_BarChartXmlWriter, chart_type, chart_data);
        $assertEqual(xml_writer.xml, snippetText(snippet_name));
      } },
    { name: "it_can_generate_xml_for_multi_level_cat_charts", args: [], fn: (self: any): any => {
        let EAST: any, WEST: any, chart_data: any, xml_writer: any;
        chart_data = $call(CategoryChartData);
        WEST = chart_data.addCategory("WEST");
        WEST.addSubCategory("SF");
        WEST.addSubCategory("LA");
        EAST = chart_data.addCategory("EAST");
        EAST.addSubCategory("NY");
        EAST.addSubCategory("NJ");
        chart_data.addSeries("Series 1", [1, 2, null, 4]);
        chart_data.addSeries("Series 2", [5, null, 7, 8]);
        xml_writer = $call(_BarChartXmlWriter, XL_CHART_TYPE.BAR_CLUSTERED, chart_data);
        $assertEqual(xml_writer.xml, snippetText("4x2-multi-cat-bar"));
      } },
  ],
};
pytestClass("Describe_BarChartXmlWriter", Describe_BarChartXmlWriter);

const Describe_BubbleChartXmlWriter: any = {
  fixtures: {
    xml_fixture: { args: ["request"], params: [["BUBBLE", 2, 3, "2x3-bubble"], ["BUBBLE_THREE_D_EFFECT", 2, 3, "2x3-bubble-3d"]], fn: (self: any, request: any): any => {
        let chart_data: any, chart_type: any, enum_member: any, expected_xml: any, point_count: any, ser_count: any, snippet_name: any, xml_writer: any;
        [enum_member, ser_count, point_count, snippet_name] = $list(request.param);
        chart_type = $getattr(XL_CHART_TYPE, enum_member);
        chart_data = make_bubble_chart_data(ser_count, point_count);
        xml_writer = $call(_BubbleChartXmlWriter, chart_type, chart_data);
        expected_xml = snippetText(snippet_name);
        return [xml_writer, expected_xml];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_generate_xml_for_bubble_charts", args: ["xml_fixture"], fn: (self: any, xml_fixture: any): any => {
        let expected_xml: any, xml_writer: any;
        [xml_writer, expected_xml] = $list(xml_fixture);
        $assertEqual(xml_writer.xml, expected_xml.replace(/<c:v>(-?\d+)\.0<\/c:v>/g, "<c:v>$1</c:v>"));
      } },
  ],
};
pytestClass("Describe_BubbleChartXmlWriter", Describe_BubbleChartXmlWriter);

const Describe_DoughnutChartXmlWriter: any = {
  fixtures: {
    xml_fixture: { args: ["request"], params: [["DOUGHNUT", 3, 2, "3x2-doughnut"], ["DOUGHNUT_EXPLODED", 3, 2, "3x2-doughnut-exploded"]], fn: (self: any, request: any): any => {
        let cat_count: any, chart_data: any, chart_type: any, enum_member: any, expected_xml: any, ser_count: any, snippet_name: any, xml_writer: any;
        [enum_member, cat_count, ser_count, snippet_name] = $list(request.param);
        chart_type = $getattr(XL_CHART_TYPE, enum_member);
        chart_data = make_category_chart_data(cat_count, String, ser_count);
        xml_writer = $call(_DoughnutChartXmlWriter, chart_type, chart_data);
        expected_xml = snippetText(snippet_name);
        return [xml_writer, expected_xml];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_generate_xml_for_doughnut_type_charts", args: ["xml_fixture"], fn: (self: any, xml_fixture: any): any => {
        let expected_xml: any, xml_writer: any;
        [xml_writer, expected_xml] = $list(xml_fixture);
        $assertEqual(xml_writer.xml, expected_xml);
      } },
  ],
};
pytestClass("Describe_DoughnutChartXmlWriter", Describe_DoughnutChartXmlWriter);

const Describe_LineChartXmlWriter: any = {
  fixtures: {
    xml_fixture: { args: ["request"], params: [["LINE", 2, 2, String, "2x2-line"], ["LINE", 2, 2, date, "2x2-line-date"], ["LINE", 2, 2, Number, "2x2-line-float"], ["LINE_MARKERS", 2, 2, String, "2x2-line-markers"], ["LINE_MARKERS_STACKED", 2, 2, String, "2x2-line-markers-stacked"], ["LINE_MARKERS_STACKED_100", 2, 2, String, "2x2-line-markers-stacked-100"], ["LINE_STACKED", 2, 2, String, "2x2-line-stacked"], ["LINE_STACKED_100", 2, 2, String, "2x2-line-stacked-100"]], fn: (self: any, request: any): any => {
        let cat_count: any, cat_type: any, chart_data: any, chart_type: any, expected_xml: any, member: any, ser_count: any, snippet_name: any, xml_writer: any;
        [member, cat_count, ser_count, cat_type, snippet_name] = $list(request.param);
        chart_type = $getattr(XL_CHART_TYPE, member);
        chart_data = make_category_chart_data(cat_count, cat_type, ser_count);
        xml_writer = $call(_LineChartXmlWriter, chart_type, chart_data);
        expected_xml = snippetText(snippet_name);
        return [xml_writer, expected_xml];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_generate_xml_for_a_line_chart", args: ["xml_fixture"], fn: (self: any, xml_fixture: any): any => {
        let expected_xml: any, xml_writer: any;
        [xml_writer, expected_xml] = $list(xml_fixture);
        $assertEqual(xml_writer.xml, expected_xml);
      } },
  ],
};
pytestClass("Describe_LineChartXmlWriter", Describe_LineChartXmlWriter);

const Describe_PieChartXmlWriter: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_can_generate_xml_for_a_pie_chart", args: ["enum_member", "cat_count", "ser_count", "snippet_name"], parametrize: [{ names: ["enum_member", "cat_count", "ser_count", "snippet_name"], values: [["PIE", 3, 1, "3x1-pie"], ["PIE_EXPLODED", 3, 1, "3x1-pie-exploded"]] }], fn: (self: any, enum_member: any, cat_count: any, ser_count: any, snippet_name: any): any => {
        let chart_data: any, chart_type: any, xml_writer: any;
        chart_type = $getattr(XL_CHART_TYPE, enum_member);
        chart_data = make_category_chart_data(cat_count, String, ser_count);
        xml_writer = $call(_PieChartXmlWriter, chart_type, chart_data);
        $assertEqual(xml_writer.xml, snippetText(snippet_name));
      } },
  ],
};
pytestClass("Describe_PieChartXmlWriter", Describe_PieChartXmlWriter);

const Describe_RadarChartXmlWriter: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_can_generate_xml_for_a_radar_chart", args: [], fn: (self: any): any => {
        let series_data_seq: any, xml_writer: any;
        series_data_seq = $call(make_category_chart_data, $kw({cat_count: 5, cat_type: String, ser_count: 2}));
        xml_writer = $call(_RadarChartXmlWriter, XL_CHART_TYPE.RADAR, series_data_seq);
        $assertEqual(xml_writer.xml, snippetText("2x5-radar").replace(/<c:v>(-?\d+)\.0<\/c:v>/g, "<c:v>$1</c:v>"));
      } },
  ],
};
pytestClass("Describe_RadarChartXmlWriter", Describe_RadarChartXmlWriter);

const Describe_XyChartXmlWriter: any = {
  fixtures: {
    xml_fixture: { args: ["request"], params: [["XY_SCATTER", 2, 3, "2x3-xy"], ["XY_SCATTER_LINES", 2, 3, "2x3-xy-lines"], ["XY_SCATTER_LINES_NO_MARKERS", 2, 3, "2x3-xy-lines-no-markers"], ["XY_SCATTER_SMOOTH", 2, 3, "2x3-xy-smooth"], ["XY_SCATTER_SMOOTH_NO_MARKERS", 2, 3, "2x3-xy-smooth-no-markers"]], fn: (self: any, request: any): any => {
        let chart_data: any, chart_type: any, enum_member: any, expected_xml: any, point_count: any, ser_count: any, snippet_name: any, xml_writer: any;
        [enum_member, ser_count, point_count, snippet_name] = $list(request.param);
        chart_type = $getattr(XL_CHART_TYPE, enum_member);
        chart_data = make_xy_chart_data(ser_count, point_count);
        xml_writer = $call(_XyChartXmlWriter, chart_type, chart_data);
        expected_xml = snippetText(snippet_name);
        return [xml_writer, expected_xml];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_generate_xml_for_xy_charts", args: ["xml_fixture"], fn: (self: any, xml_fixture: any): any => {
        let expected_xml: any, xml_writer: any;
        [xml_writer, expected_xml] = $list(xml_fixture);
        $assertEqual(xml_writer.xml, expected_xml);
      } },
  ],
};
pytestClass("Describe_XyChartXmlWriter", Describe_XyChartXmlWriter);

const Describe_BubbleSeriesXmlWriter: any = {
  fixtures: {
    bubbleSize_fixture: { args: ["request"], params: [[null, null, "General"], [null, 42, "42"], [24, null, "24"], [24, 42, "42"]], fn: (self: any, request: any): any => {
        let chart_data: any, chart_number_format: any, cht_numfmt: any, expected_numfmt: any, expected_xml: any, ser_numfmt: any, series_data: any, series_number_format: any, xml_writer: any;
        [cht_numfmt, ser_numfmt, expected_numfmt] = $list(request.param);
        expected_xml = xml($mod("c:bubbleSize/c:numRef/(c:f\"Sheet1!$C$2:$C$1\",c:numCache/(c:formatCode\"%s\",c:ptCount{val=0}))", expected_numfmt));
        chart_number_format = ($truthy($is(cht_numfmt, null)) ? [] : [cht_numfmt]);
        series_number_format = ($truthy($is(ser_numfmt, null)) ? [] : [ser_numfmt]);
        chart_data = $call(BubbleChartData, ...$list(chart_number_format));
        series_data = $m(chart_data, "addSeries", null, ...$list(series_number_format));
        xml_writer = $call(_BubbleSeriesXmlWriter, series_data);
        return [xml_writer, expected_xml];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_bubbleSize_XML", args: ["bubbleSize_fixture"], fn: (self: any, bubbleSize_fixture: any): any => {
        let bubbleSize: any, expected_xml: any, xml_writer: any;
        [xml_writer, expected_xml] = $list(bubbleSize_fixture);
        bubbleSize = xml_writer.bubbleSize;
        $assertEqual(bubbleSize.xml, expected_xml);
      } },
  ],
};
pytestClass("Describe_BubbleSeriesXmlWriter", Describe_BubbleSeriesXmlWriter);

const Describe_CategorySeriesXmlWriter: any = {
  fixtures: {
    val_fixture: { args: ["series_data_"], fn: (self: any, series_data_: any): any => {
        let expected_xml: any, number_format: any, values_ref: any, xml_writer: any;
        [values_ref, number_format] = $list(["Sheet1!$B$2:$B$1", "Foobar"]);
        expected_xml = xml($mod("c:val/c:numRef/(c:f\"%s\",c:numCache/(c:formatCode\"%s\",c:ptCount{val=0}))", [values_ref, number_format]));
        xml_writer = $call(_CategorySeriesXmlWriter, series_data_);
        series_data_.valuesRef = values_ref;
        series_data_.numberFormat = number_format;
        return [xml_writer, expected_xml];
      } },
    series_data_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, CategorySeriesData);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_val_XML", args: ["val_fixture"], fn: (self: any, val_fixture: any): any => {
        let expected_xml: any, val: any, xml_writer: any;
        [xml_writer, expected_xml] = $list(val_fixture);
        val = xml_writer.val;
        $assertEqual(val.xml, expected_xml);
      } },
  ],
};
pytestClass("Describe_CategorySeriesXmlWriter", Describe_CategorySeriesXmlWriter);

const Describe_XySeriesXmlWriter: any = {
  fixtures: {
    xVal_fixture: { args: ["request"], params: [[null, null, "General"], [null, 42, "42"], [24, null, "24"], [24, 42, "42"]], fn: (self: any, request: any): any => {
        let chart_data: any, chart_number_format: any, cht_numfmt: any, expected_numfmt: any, expected_xml: any, ser_numfmt: any, series_data: any, series_number_format: any, xml_writer: any;
        [cht_numfmt, ser_numfmt, expected_numfmt] = $list(request.param);
        expected_xml = xml($mod("c:xVal/c:numRef/(c:f\"Sheet1!$A$2:$A$1\",c:numCache/(c:formatCode\"%s\",c:ptCount{val=0}))", expected_numfmt));
        chart_number_format = ($truthy($is(cht_numfmt, null)) ? [] : [cht_numfmt]);
        series_number_format = ($truthy($is(ser_numfmt, null)) ? [] : [ser_numfmt]);
        chart_data = $call(XyChartData, ...$list(chart_number_format));
        series_data = $m(chart_data, "addSeries", null, ...$list(series_number_format));
        xml_writer = $call(_XySeriesXmlWriter, series_data);
        return [xml_writer, expected_xml];
      } },
    yVal_fixture: { args: ["request"], params: [[null, null, "General"], [null, 42, "42"], [24, null, "24"], [24, 42, "42"]], fn: (self: any, request: any): any => {
        let chart_data: any, chart_number_format: any, cht_numfmt: any, expected_numfmt: any, expected_xml: any, ser_numfmt: any, series_data: any, series_number_format: any, xml_writer: any;
        [cht_numfmt, ser_numfmt, expected_numfmt] = $list(request.param);
        expected_xml = xml($mod("c:yVal/c:numRef/(c:f\"Sheet1!$B$2:$B$1\",c:numCache/(c:formatCode\"%s\",c:ptCount{val=0}))", expected_numfmt));
        chart_number_format = ($truthy($is(cht_numfmt, null)) ? [] : [cht_numfmt]);
        series_number_format = ($truthy($is(ser_numfmt, null)) ? [] : [ser_numfmt]);
        chart_data = $call(XyChartData, ...$list(chart_number_format));
        series_data = $m(chart_data, "addSeries", null, ...$list(series_number_format));
        xml_writer = $call(_XySeriesXmlWriter, series_data);
        return [xml_writer, expected_xml];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_knows_its_xVal_XML", args: ["xVal_fixture"], fn: (self: any, xVal_fixture: any): any => {
        let expected_xml: any, xVal: any, xml_writer: any;
        [xml_writer, expected_xml] = $list(xVal_fixture);
        xVal = xml_writer.xVal;
        $assertEqual(xVal.xml, expected_xml);
      } },
    { name: "it_knows_its_yVal_XML", args: ["yVal_fixture"], fn: (self: any, yVal_fixture: any): any => {
        let expected_xml: any, xml_writer: any, yVal: any;
        [xml_writer, expected_xml] = $list(yVal_fixture);
        yVal = xml_writer.yVal;
        $assertEqual(yVal.xml, expected_xml);
      } },
  ],
};
pytestClass("Describe_XySeriesXmlWriter", Describe_XySeriesXmlWriter);

const Describe_BaseSeriesXmlRewriter: any = {
  fixtures: {
    adjust_fixture: { args: ["request", "_add_cloned_sers_", "_trim_ser_count_by_"], params: [[3, true, false], [2, false, false], [1, false, true]], fn: (self: any, request: any, _add_cloned_sers_: any, _trim_ser_count_by_: any): any => {
        let add: any, add_calls: any, new_ser_count: any, plotArea: any, rewriter: any, sers: any, trim: any, trim_calls: any;
        [new_ser_count, add, trim] = $list(request.param);
        rewriter = $call(_BaseSeriesXmlRewriter, null);
        plotArea = element("c:plotArea/c:barChart/(c:ser/(c:idx{val=3},c:order{val=1}),c:ser/(c:idx{val=1},c:order{val=3}))");
        sers = plotArea.sers;
        add_calls = ($truthy(add) ? [call(rewriter, plotArea, 1)] : []);
        trim_calls = ($truthy(trim) ? [call(rewriter, plotArea, 1)] : []);
        return [rewriter, plotArea, new_ser_count, add_calls, trim_calls, sers];
      } },
    clone_fixture: { args: ["request"], params: [["c:plotArea/c:barChart/c:ser/(c:idx{val=3},c:order{val=1})", 1, "c:plotArea/c:barChart/(c:ser/(c:idx{val=3},c:order{val=1}),c:ser/(c:idx{val=4},c:order{val=2}))"], ["c:plotArea/(c:barChart/c:ser/(c:idx{val=3},c:order{val=1}),c:lineChart/(c:ser/(c:idx{val=4},c:order{val=7}),c:ser/(c:idx{val=1},c:order{val=2})))", 1, "c:plotArea/(c:barChart/c:ser/(c:idx{val=3},c:order{val=1}),c:lineChart/(c:ser/(c:idx{val=4},c:order{val=7}),c:ser/(c:idx{val=5},c:order{val=8}),c:ser/(c:idx{val=1},c:order{val=2})))"]], fn: (self: any, request: any): any => {
        let count: any, expected_cxml: any, expected_xml: any, plotArea: any, plotArea_cxml: any, rewriter: any;
        [plotArea_cxml, count, expected_cxml] = $list(request.param);
        rewriter = $call(_BaseSeriesXmlRewriter, null);
        plotArea = element(plotArea_cxml);
        expected_xml = xml(expected_cxml);
        return [rewriter, plotArea, count, expected_xml];
      } },
    replace_fixture: { args: ["request", "chart_data_", "_adjust_ser_count_", "_rewrite_ser_data_"], fn: (self: any, request: any, chart_data_: any, _adjust_ser_count_: any, _rewrite_ser_data_: any): any => {
        let calls: any, chartSpace: any, plotArea: any, rewriter: any, ser_count: any, series_datas: any, sers: any;
        rewriter = $call(_BaseSeriesXmlRewriter, chart_data_);
        chartSpace = element("c:chartSpace/c:chart/c:plotArea/c:barChart/(c:ser/c:order{val=0},c:ser/c:order{val=1})");
        plotArea = $get(chartSpace.xpath("c:chart/c:plotArea"), 0);
        sers = chartSpace.xpath(".//c:ser");
        ser_count = $len(sers);
        series_datas = [instanceMock(request, _BaseSeriesData), instanceMock(request, _BaseSeriesData)];
        calls = [call(rewriter, $get(sers, 0), $get(series_datas, 0), false), call(rewriter, $get(sers, 1), $get(series_datas, 1), false)];
        chart_data_.__len__.returnValue = $len(series_datas);
        chart_data_.__iter__.returnValue = $iter(series_datas);
        return [rewriter, chartSpace, plotArea, ser_count, calls];
      } },
    trim_fixture: { args: ["request"], params: [["c:plotArea/c:barChart/(c:ser/(c:idx{val=3},c:order{val=4}),c:ser/(c:idx{val=1},c:order{val=2}))", 1, "c:plotArea/c:barChart/(c:ser/(c:idx{val=1},c:order{val=2}))"], ["c:plotArea/(c:barChart/c:ser/(c:idx{val=3},c:order{val=1}),c:lineChart/(c:ser/(c:idx{val=4},c:order{val=7}),c:ser/(c:idx{val=1},c:order{val=2})))", 2, "c:plotArea/(c:barChart/c:ser/(c:idx{val=3},c:order{val=1}))"]], fn: (self: any, request: any): any => {
        let count: any, expected_cxml: any, expected_xml: any, plotArea: any, plotArea_cxml: any, rewriter: any;
        [plotArea_cxml, count, expected_cxml] = $list(request.param);
        rewriter = $call(_BaseSeriesXmlRewriter, null);
        plotArea = element(plotArea_cxml);
        expected_xml = xml(expected_cxml);
        return [rewriter, plotArea, count, expected_xml];
      } },
    _add_cloned_sers_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, _BaseSeriesXmlRewriter, "_add_cloned_sers", $kw({autospec: true}));
      } },
    _adjust_ser_count_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, _BaseSeriesXmlRewriter, "_adjust_ser_count", $kw({autospec: true}));
      } },
    chart_data_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, _BaseChartData);
      } },
    _rewrite_ser_data_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, _BaseSeriesXmlRewriter, "_rewrite_ser_data", $kw({autospec: true}));
      } },
    _trim_ser_count_by_: { args: ["request"], fn: (self: any, request: any): any => {
        return $call(methodMock, request, _BaseSeriesXmlRewriter, "_trim_ser_count_by", $kw({autospec: true}));
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_replace_series_data", args: ["replace_fixture"], fn: (self: any, replace_fixture: any): any => {
        let calls: any, chartSpace: any, plotArea: any, rewriter: any, ser_count: any;
        [rewriter, chartSpace, plotArea, ser_count, calls] = $list(replace_fixture);
        rewriter.replaceSeriesData(chartSpace);
        rewriter._adjustSerCount.assertCalledOnceWith(rewriter, plotArea, ser_count);
        $assertEqual(rewriter._rewriteSerData.callArgsList, calls);
      } },
    { name: "it_adjusts_the_ser_count_to_help", args: ["adjust_fixture"], fn: (self: any, adjust_fixture: any): any => {
        let add_calls: any, new_ser_count: any, plotArea: any, rewriter: any, sers: any, trim_calls: any;
        [rewriter, plotArea, new_ser_count, add_calls, trim_calls, sers] = $list(adjust_fixture);
        rewriter._adjustSerCount(plotArea, new_ser_count);
        $assertEqual(rewriter._addClonedSers.callArgsList, add_calls);
        $assertEqual(rewriter._trimSerCountBy.callArgsList, trim_calls);
      } },
    { name: "it_adds_cloned_sers_to_help", args: ["clone_fixture"], fn: (self: any, clone_fixture: any): any => {
        let count: any, expected_xml: any, plotArea: any, rewriter: any;
        [rewriter, plotArea, count, expected_xml] = $list(clone_fixture);
        rewriter._addClonedSers(plotArea, count);
        $assertEqual(plotArea.xml, expected_xml);
      } },
    { name: "it_trims_sers_to_help", args: ["trim_fixture"], fn: (self: any, trim_fixture: any): any => {
        let count: any, expected_xml: any, plotArea: any, rewriter: any;
        [rewriter, plotArea, count, expected_xml] = $list(trim_fixture);
        rewriter._trimSerCountBy(plotArea, count);
        $assertEqual(plotArea.xml, expected_xml);
      } },
  ],
};
pytestClass("Describe_BaseSeriesXmlRewriter", Describe_BaseSeriesXmlRewriter);

const Describe_BubbleSeriesXmlRewriter: any = {
  fixtures: {
    rewrite_fixture: { args: [], fn: (self: any): any => {
        let chart_data: any, expected_xml: any, rewriter: any, ser: any, ser_xml: any, series_data: any;
        [ser_xml, expected_xml] = $list($slice(snippetSeq("rewrite-ser"), 4, 6, null));
        chart_data = $call(BubbleChartData);
        series_data = chart_data.addSeries("Series 1");
        series_data.addDataPoint(1, 2, 10);
        series_data.addDataPoint(3, 4, 20);
        rewriter = $call(_BubbleSeriesXmlRewriter, chart_data);
        ser = parseXml(ser_xml);
        return [rewriter, ser, series_data, expected_xml];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_rewrite_a_ser_element", args: ["rewrite_fixture"], fn: (self: any, rewrite_fixture: any): any => {
        let expected_xml: any, rewriter: any, ser: any, series_data: any;
        [rewriter, ser, series_data, expected_xml] = $list(rewrite_fixture);
        rewriter._rewriteSerData(ser, series_data, null);
        $assertEqual(ser.xml, expected_xml);
      } },
  ],
};
pytestClass("Describe_BubbleSeriesXmlRewriter", Describe_BubbleSeriesXmlRewriter);

const Describe_CategorySeriesXmlRewriter: any = {
  fixtures: {
    rewrite_fixture: { args: ["request"], params: [[0, ["foo", "bar"]], [6, [42, 24]], [8, [date(2016, 12, 22), date(2016, 12, 23)]]], fn: (self: any, request: any): any => {
        let categories: any, chart_data: any, expected_xml: any, rewriter: any, ser: any, ser_xml: any, series_data: any, snippet_offset: any, snippets: any;
        [snippet_offset, categories] = $list(request.param);
        chart_data = $call(CategoryChartData);
        chart_data.categories = categories;
        rewriter = $call(_CategorySeriesXmlRewriter, chart_data);
        snippets = snippetSeq("rewrite-ser");
        ser_xml = $get(snippets, snippet_offset);
        ser = parseXml(ser_xml);
        series_data = chart_data.addSeries("Series 1", [1, 2]);
        expected_xml = $get(snippets, (snippet_offset + 1));
        return [rewriter, ser, series_data, expected_xml];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_rewrite_a_ser_element", args: ["rewrite_fixture"], fn: (self: any, rewrite_fixture: any): any => {
        let expected_xml: any, rewriter: any, ser: any, series_data: any;
        [rewriter, ser, series_data, expected_xml] = $list(rewrite_fixture);
        rewriter._rewriteSerData(ser, series_data, false);
        $assertEqual(ser.xml, expected_xml);
      } },
  ],
};
pytestClass("Describe_CategorySeriesXmlRewriter", Describe_CategorySeriesXmlRewriter);

const Describe_XySeriesXmlRewriter: any = {
  fixtures: {
    rewrite_fixture: { args: [], fn: (self: any): any => {
        let chart_data: any, expected_xml: any, rewriter: any, ser: any, ser_xml: any, series_data: any;
        [ser_xml, expected_xml] = $list($slice(snippetSeq("rewrite-ser"), 2, 4, null));
        chart_data = $call(XyChartData);
        series_data = chart_data.addSeries("Series 1");
        series_data.addDataPoint(1, 2);
        series_data.addDataPoint(3, 4);
        rewriter = $call(_XySeriesXmlRewriter, chart_data);
        ser = parseXml(ser_xml);
        return [rewriter, ser, series_data, expected_xml];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_rewrite_a_ser_element", args: ["rewrite_fixture"], fn: (self: any, rewrite_fixture: any): any => {
        let expected_xml: any, rewriter: any, ser: any, series_data: any;
        [rewriter, ser, series_data, expected_xml] = $list(rewrite_fixture);
        rewriter._rewriteSerData(ser, series_data, null);
        $assertEqual(ser.xml, expected_xml);
      } },
  ],
};
pytestClass("Describe_XySeriesXmlRewriter", Describe_XySeriesXmlRewriter);

export function make_bubble_chart_data(ser_count: any, point_count: any): any {
  let chart_data: any, point_idx: any, points: any, series: any, series_label: any, size: any, x: any, y: any;
  points = [[1.1, 11.1, 10.0], [2.1, 12.1, 20.0], [3.1, 13.1, 30.0], [1.2, 11.2, 40.0], [2.2, 12.2, 50.0], [3.2, 13.2, 60.0]];
  chart_data = $call(BubbleChartData);
  for (const i of $list($range(ser_count))) {
    series_label = $mod("Series %d", (i + 1));
    series = chart_data.addSeries(series_label);
    for (const j of $list($range(point_count))) {
      point_idx = $add($mul(i, point_count), j);
      [x, y, size] = $list($get(points, point_idx));
      series.addDataPoint(x, y, size);
    }
  }
  return chart_data;
}

export function make_category_chart_data(cat_count: any, cat_type: any, ser_count: any): any {
  let category_labels: any, chart_data: any, point_values: any, series_title: any, series_values: any;
  category_labels = $get($d({[date]: [date(2016, 12, 27), date(2016, 12, 28), date(2016, 12, 29)], [Number]: [1.1, 2.2, 3.3, 4.4, 5.5], [String]: ["Foo", "Bar", "Baz", "Boo", "Far", "Faz"]}), cat_type);
  point_values = count(1.1, 1.1);
  chart_data = $call(CategoryChartData);
  chart_data.categories = $slice(category_labels, null, cat_count, null);
  for (const idx of $list($range(ser_count))) {
    series_title = $mod("Series %d", (idx + 1));
    series_values = $list(islice(point_values, cat_count));
    series_values = $list(series_values).map((x: any) => ($round($mul(x, 10)) / 10.0));
    chart_data.addSeries(series_title, series_values);
  }
  return chart_data;
}

export function make_xy_chart_data(ser_count: any, point_count: any): any {
  let chart_data: any, point_idx: any, points: any, series: any, series_label: any, x: any, y: any;
  points = [[1.1, 11.1], [2.1, 12.1], [3.1, 13.1], [1.2, 11.2], [2.2, 12.2], [3.2, 13.2]];
  chart_data = $call(XyChartData);
  for (const i of $list($range(ser_count))) {
    series_label = $mod("Series %d", (i + 1));
    series = chart_data.addSeries(series_label);
    for (const j of $list($range(point_count))) {
      point_idx = $add($mul(i, point_count), j);
      [x, y] = $list($get(points, point_idx));
      series.addDataPoint(x, y);
    }
  }
  return chart_data;
}

