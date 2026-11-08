// Machine-translated from python-pptx tests/parts/test_chart.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $b, $call, $is, $kw, classMock, instanceMock, methodMock, propertyMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { Chart as Chart$imported } from "../../src/chart/chart.js";
import { ChartData as ChartData$imported } from "../../src/chart/data.js";
import { XL_CHART_TYPE as XCT$imported } from "../../src/enum/chart.js";
import { CONTENT_TYPE as CT$imported, RELATIONSHIP_TYPE as RT$imported } from "../../src/opc/constants.js";
import { OpcPackage as OpcPackage$imported } from "../../src/opc/package.js";
import { PackURI as PackURI$imported } from "../../src/opc/packuri.js";
import { CT_ChartSpace as CT_ChartSpace$imported } from "../../src/oxml/chart/chart.js";
import { ChartPart as ChartPart$imported, ChartWorkbook as ChartWorkbook$imported } from "../../src/parts/chart.js";
import { EmbeddedXlsxPart as EmbeddedXlsxPart$imported } from "../../src/parts/embeddedpackage.js";
import { element, xml } from "../unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const Chart = Chart$imported;
const ChartData = ChartData$imported;
const XCT = XCT$imported;
const CT = CT$imported;
const RT = RT$imported;
const OpcPackage = OpcPackage$imported;
const PackURI = PackURI$imported;
const CT_ChartSpace = CT_ChartSpace$imported;
const ChartPart = ChartPart$imported;
const ChartWorkbook = ChartWorkbook$imported;
const EmbeddedXlsxPart = EmbeddedXlsxPart$imported;
void $py;


const DescribeChartPart: any = {
  fixtures: {
    chartSpace_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, CT_ChartSpace);
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_construct_from_chart_type_and_data", args: ["request"], fn: (self: any, request: any): any => {
        let chart_data_: any, chart_part: any, chart_part_: any, load_: any, package_: any;
        chart_data_ = $call(instanceMock, request, ChartData, $kw({xlsx_blob: $b("xlsx-blob")}));
        chart_data_.xmlBytes.returnValue = $b("chart-blob");
        package_ = instanceMock(request, OpcPackage);
        package_.nextPartname.returnValue = $call(PackURI, "/ppt/charts/chart42.xml");
        chart_part_ = instanceMock(request, ChartPart);
        load_ = $call(methodMock, request, ChartPart, "load", $kw({autospec: false, return_value: chart_part_}));
        chart_part = ChartPart.new(XCT.RADAR, chart_data_, package_);
        package_.nextPartname.assertCalledOnceWith("/ppt/charts/chart%d.xml");
        chart_data_.xmlBytes.assertCalledOnceWith(XCT.RADAR);
        load_.assertCalledOnceWith("/ppt/charts/chart42.xml", CT.DML_CHART, package_, $b("chart-blob"));
        chart_part_.chartWorkbook.updateFromXlsxBlob.assertCalledOnceWith($b("xlsx-blob"));
        $assert($is(chart_part, chart_part_));
      } },
    { name: "it_provides_access_to_the_chart_object", args: ["request", "chartSpace_"], fn: (self: any, request: any, chartSpace_: any): any => {
        let Chart_: any, chart: any, chart_: any, chart_part: any;
        chart_ = instanceMock(request, Chart);
        Chart_ = $call(classMock, request, "pptx.parts.chart.Chart", $kw({return_value: chart_}));
        chart_part = $call(ChartPart, null, null, null, chartSpace_);
        chart = chart_part.chart;
        Chart_.assertCalledOnceWith(chart_part._element, chart_part);
        $assert($is(chart, chart_));
      } },
    { name: "it_provides_access_to_the_chart_workbook", args: ["request", "chartSpace_"], fn: (self: any, request: any, chartSpace_: any): any => {
        let ChartWorkbook_: any, chart_part: any, chart_workbook: any, chart_workbook_: any;
        chart_workbook_ = instanceMock(request, ChartWorkbook);
        ChartWorkbook_ = $call(classMock, request, "pptx.parts.chart.ChartWorkbook", $kw({return_value: chart_workbook_}));
        chart_part = $call(ChartPart, null, null, null, chartSpace_);
        chart_workbook = chart_part.chartWorkbook;
        ChartWorkbook_.assertCalledOnceWith(chartSpace_, chart_part);
        $assert($is(chart_workbook, chart_workbook_));
      } },
  ],
};
pytestClass("DescribeChartPart", DescribeChartPart);

const DescribeChartWorkbook: any = {
  fixtures: {
    chart_part_: { args: ["request", "package_", "xlsx_part_"], fn: (self: any, request: any, package_: any, xlsx_part_: any): any => {
        return instanceMock(request, ChartPart);
      } },
    package_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, OpcPackage);
      } },
    xlsx_part_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, EmbeddedXlsxPart);
      } },
    xlsx_part_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, ChartWorkbook, "xlsx_part");
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_get_the_chart_xlsx_part", args: ["chart_part_", "xlsx_part_"], fn: (self: any, chart_part_: any, xlsx_part_: any): any => {
        let chart_workbook: any, xlsx_part: any;
        chart_part_.relatedPart.returnValue = xlsx_part_;
        chart_workbook = $call(ChartWorkbook, element("c:chartSpace/c:externalData{r:id=rId42}"), chart_part_);
        xlsx_part = chart_workbook.xlsxPart;
        chart_part_.relatedPart.assertCalledOnceWith("rId42");
        $assert($is(xlsx_part, xlsx_part_));
      } },
    { name: "but_it_returns_None_when_the_chart_has_no_xlsx_part", args: [], fn: (self: any): any => {
        let chart_workbook: any;
        chart_workbook = $call(ChartWorkbook, element("c:chartSpace"), null);
        $assert($is(chart_workbook.xlsxPart, null));
      } },
    { name: "it_can_change_the_chart_xlsx_part", args: ["chart_part_", "xlsx_part_", "chartSpace_cxml", "expected_cxml"], parametrize: [{ names: ["chartSpace_cxml", "expected_cxml"], values: [["c:chartSpace{r:a=b}", "c:chartSpace{r:a=b}/c:externalData{r:id=rId42}/c:autoUpdate{val=0}"], ["c:chartSpace/c:externalData{r:id=rId66}", "c:chartSpace/c:externalData{r:id=rId42}"]] }], fn: (self: any, chart_part_: any, xlsx_part_: any, chartSpace_cxml: any, expected_cxml: any): any => {
        let chart_data: any;
        chart_part_.relateTo.returnValue = "rId42";
        chart_data = $call(ChartWorkbook, element(chartSpace_cxml), chart_part_);
        chart_data.xlsxPart = xlsx_part_;
        chart_part_.relateTo.assertCalledOnceWith(xlsx_part_, RT.PACKAGE);
        $assertEqual(chart_data._chartSpace.xml, xml(expected_cxml));
      } },
    { name: "it_adds_an_xlsx_part_on_update_if_needed", args: ["request", "chart_part_", "package_", "xlsx_part_", "xlsx_part_prop_"], fn: (self: any, request: any, chart_part_: any, package_: any, xlsx_part_: any, xlsx_part_prop_: any): any => {
        let EmbeddedXlsxPart_: any, chart_data: any;
        EmbeddedXlsxPart_ = classMock(request, "pptx.parts.chart.EmbeddedXlsxPart");
        EmbeddedXlsxPart_.new.returnValue = xlsx_part_;
        chart_part_.package = package_;
        xlsx_part_prop_.returnValue = null;
        chart_data = $call(ChartWorkbook, element("c:chartSpace"), chart_part_);
        chart_data.updateFromXlsxBlob($b("xlsx-blob"));
        EmbeddedXlsxPart_.new.assertCalledOnceWith($b("xlsx-blob"), package_);
        xlsx_part_prop_.assertCalledWith(xlsx_part_);
      } },
    { name: "but_it_replaces_the_xlsx_blob_when_the_part_exists", args: ["xlsx_part_prop_", "xlsx_part_"], fn: (self: any, xlsx_part_prop_: any, xlsx_part_: any): any => {
        let chart_data: any;
        xlsx_part_prop_.returnValue = xlsx_part_;
        chart_data = $call(ChartWorkbook, null, null);
        chart_data.updateFromXlsxBlob($b("xlsx-blob"));
        $assertEqual(chart_data.xlsxPart.blob, $b("xlsx-blob"));
      } },
  ],
};
pytestClass("DescribeChartWorkbook", DescribeChartWorkbook);

