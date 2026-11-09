// Machine-translated from python-pptx tests/shapes/test_graphfrm.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../pyshim.js";
import { $assert, $assertEqual, $b, $call, $is, $kw, $mod, $raises, $str, $truthy, classMock, instanceMock, propertyMock, pytestClass } from "../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../src/exc.js";
import { defineChildren } from "../../src/oxml/xmlchemy.js";
import { Chart as Chart$imported } from "../../src/chart/chart.js";
import { MSO_SHAPE_TYPE as MSO_SHAPE_TYPE$imported } from "../../src/enum/shapes.js";
import { ChartPart as ChartPart$imported } from "../../src/parts/chart.js";
import { EmbeddedPackagePart as EmbeddedPackagePart$imported } from "../../src/parts/embeddedpackage.js";
import { SlidePart as SlidePart$imported } from "../../src/parts/slide.js";
import { GraphicFrame as GraphicFrame$imported, _OleFormat as _OleFormat$imported } from "../../src/shapes/graphfrm.js";
import { SlideShapes as SlideShapes$imported } from "../../src/shapes/shapetree.js";
import { GRAPHIC_DATA_URI_CHART as GRAPHIC_DATA_URI_CHART$imported, GRAPHIC_DATA_URI_OLEOBJ as GRAPHIC_DATA_URI_OLEOBJ$imported, GRAPHIC_DATA_URI_TABLE as GRAPHIC_DATA_URI_TABLE$imported } from "../../src/spec.js";
import { element } from "../unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const Chart = Chart$imported;
const MSO_SHAPE_TYPE = MSO_SHAPE_TYPE$imported;
const ChartPart = ChartPart$imported;
const EmbeddedPackagePart = EmbeddedPackagePart$imported;
const SlidePart = SlidePart$imported;
const GraphicFrame = GraphicFrame$imported;
const _OleFormat = _OleFormat$imported;
const SlideShapes = SlideShapes$imported;
const GRAPHIC_DATA_URI_CHART = GRAPHIC_DATA_URI_CHART$imported;
const GRAPHIC_DATA_URI_OLEOBJ = GRAPHIC_DATA_URI_OLEOBJ$imported;
const GRAPHIC_DATA_URI_TABLE = GRAPHIC_DATA_URI_TABLE$imported;
void $py;


const DescribeGraphicFrame: any = {
  fixtures: {
    chart_: { args: ["request"], fn: (self: any, request: any): any => {
        return instanceMock(request, Chart);
      } },
    chart_part_: { args: ["request", "chart_"], fn: (self: any, request: any, chart_: any): any => {
        return $call(instanceMock, request, ChartPart, $kw({chart: chart_}));
      } },
    has_chart_prop_: { args: ["request"], fn: (self: any, request: any): any => {
        return propertyMock(request, GraphicFrame, "has_chart");
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_access_to_the_chart_it_contains", args: ["request", "has_chart_prop_", "chart_part_", "chart_"], fn: (self: any, request: any, has_chart_prop_: any, chart_part_: any, chart_: any): any => {
        has_chart_prop_.returnValue = true;
        $call(propertyMock, request, GraphicFrame, "chart_part", $kw({return_value: chart_part_}));
        chart_part_.chart = chart_;
        $assert($is($call(GraphicFrame, null, null).chart, chart_));
      } },
    { name: "but_it_raises_on_chart_if_there_isnt_one", args: ["has_chart_prop_"], fn: (self: any, has_chart_prop_: any): any => {
        let e: any;
        has_chart_prop_.returnValue = false;
        e = $raises(ValueError, () => {
          $call(GraphicFrame, null, null).chart;
        });
        $assertEqual($str(e.value), "shape does not contain a chart");
      } },
    { name: "it_provides_access_to_its_chart_part", args: ["request", "chart_part_"], fn: (self: any, request: any, chart_part_: any): any => {
        let chart_part: any, graphic_frame: any, slide_part_: any;
        slide_part_ = instanceMock(request, SlidePart);
        slide_part_.relatedPart.returnValue = chart_part_;
        $call(propertyMock, request, GraphicFrame, "part", $kw({return_value: slide_part_}));
        graphic_frame = $call(GraphicFrame, element("p:graphicFrame/a:graphic/a:graphicData/c:chart{r:id=rId42}"), null);
        chart_part = graphic_frame.chartPart;
        slide_part_.relatedPart.assertCalledOnceWith("rId42");
        $assert($is(chart_part, chart_part_));
      } },
    { name: "it_knows_whether_it_contains_a_chart", args: ["graphicData_uri", "expected_value"], parametrize: [{ names: ["graphicData_uri", "expected_value"], values: [[GRAPHIC_DATA_URI_CHART, true], [GRAPHIC_DATA_URI_OLEOBJ, false], [GRAPHIC_DATA_URI_TABLE, false]] }], fn: (self: any, graphicData_uri: any, expected_value: any): any => {
        let graphicFrame: any;
        graphicFrame = element($mod("p:graphicFrame/a:graphic/a:graphicData{uri=%s}", graphicData_uri));
        $assert($is($call(GraphicFrame, graphicFrame, null).hasChart, expected_value));
      } },
    { name: "it_knows_whether_it_contains_a_table", args: ["graphicData_uri", "expected_value"], parametrize: [{ names: ["graphicData_uri", "expected_value"], values: [[GRAPHIC_DATA_URI_CHART, false], [GRAPHIC_DATA_URI_OLEOBJ, false], [GRAPHIC_DATA_URI_TABLE, true]] }], fn: (self: any, graphicData_uri: any, expected_value: any): any => {
        let graphicFrame: any;
        graphicFrame = element($mod("p:graphicFrame/a:graphic/a:graphicData{uri=%s}", graphicData_uri));
        $assert($is($call(GraphicFrame, graphicFrame, null).hasTable, expected_value));
      } },
    { name: "it_provides_access_to_the_OleFormat_object", args: ["request"], fn: (self: any, request: any): any => {
        let _OleFormat_: any, graphicFrame: any, graphic_frame: any, ole_format: any, ole_format_: any, parent_: any;
        ole_format_ = instanceMock(request, _OleFormat);
        _OleFormat_ = $call(classMock, request, "pptx.shapes.graphfrm._OleFormat", $kw({return_value: ole_format_}));
        graphicFrame = element("p:graphicFrame/a:graphic/a:graphicData{uri=http://schemas.openxmlformats.org/presentationml/2006/ole}");
        parent_ = instanceMock(request, SlideShapes);
        graphic_frame = $call(GraphicFrame, graphicFrame, parent_);
        ole_format = graphic_frame.oleFormat;
        _OleFormat_.assertCalledOnceWith(graphicFrame.graphicData, parent_);
        $assert($is(ole_format, ole_format_));
      } },
    { name: "but_it_raises_on_ole_format_when_this_is_not_an_OLE_object", args: [], fn: (self: any): any => {
        let e: any, graphic_frame: any;
        graphic_frame = $call(GraphicFrame, element("p:graphicFrame/a:graphic/a:graphicData{uri=http://schemas.openxmlformats.org/drawingml/2006/table}"), null);
        e = $raises(ValueError, () => {
          graphic_frame.oleFormat;
        });
        $assertEqual($str(e.value), "not an OLE-object shape");
      } },
    { name: "it_raises_on_shadow", args: [], fn: (self: any): any => {
        let graphic_frame: any;
        graphic_frame = $call(GraphicFrame, null, null);
        $raises(NotImplementedError, () => {
          graphic_frame.shadow;
        });
      } },
    { name: "it_knows_its_shape_type", args: ["uri", "oleObj_child", "expected_value"], parametrize: [{ names: ["uri", "oleObj_child", "expected_value"], values: [[GRAPHIC_DATA_URI_CHART, null, MSO_SHAPE_TYPE.CHART], [GRAPHIC_DATA_URI_OLEOBJ, "embed", MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT], [GRAPHIC_DATA_URI_OLEOBJ, "link", MSO_SHAPE_TYPE.LINKED_OLE_OBJECT], [GRAPHIC_DATA_URI_TABLE, null, MSO_SHAPE_TYPE.TABLE], ["foobar", null, null]] }], fn: (self: any, uri: any, oleObj_child: any, expected_value: any): any => {
        let graphicFrame: any;
        graphicFrame = element(($truthy(oleObj_child) ? $mod("p:graphicFrame/a:graphic/a:graphicData{uri=%s}/p:oleObj/p:%s", [uri, oleObj_child]) : $mod("p:graphicFrame/a:graphic/a:graphicData{uri=%s}", uri)));
        $assert($is($call(GraphicFrame, graphicFrame, null).shapeType, expected_value));
      } },
  ],
};
pytestClass("DescribeGraphicFrame", DescribeGraphicFrame);

const Describe_OleFormat: any = {
  fixtures: {
  },
  helpers: {
  },
  tests: [
    { name: "it_provides_access_to_the_OLE_object_blob", args: ["request"], fn: (self: any, request: any): any => {
        let blob: any, ole_format: any, ole_obj_part_: any, slide_part_: any;
        ole_obj_part_ = $call(instanceMock, request, EmbeddedPackagePart, $kw({blob: $b("0123456789")}));
        slide_part_ = instanceMock(request, SlidePart);
        slide_part_.relatedPart.returnValue = ole_obj_part_;
        $call(propertyMock, request, _OleFormat, "part", $kw({return_value: slide_part_}));
        ole_format = $call(_OleFormat, element("a:graphicData/p:oleObj{r:id=rId7}"), null);
        blob = ole_format.blob;
        slide_part_.relatedPart.assertCalledOnceWith("rId7");
        $assertEqual(blob, $b("0123456789"));
      } },
    { name: "it_knows_the_OLE_object_prog_id", args: [], fn: (self: any): any => {
        let graphicData: any;
        graphicData = element("a:graphicData/p:oleObj{progId=Excel.Sheet.12}");
        $assertEqual($call(_OleFormat, graphicData, null).progId, "Excel.Sheet.12");
      } },
    { name: "it_knows_whether_to_show_the_OLE_object_as_an_icon", args: [], fn: (self: any): any => {
        let graphicData: any;
        graphicData = element("a:graphicData/p:oleObj{showAsIcon=1}");
        $assert($is($call(_OleFormat, graphicData, null).showAsIcon, true));
      } },
  ],
};
pytestClass("Describe_OleFormat", Describe_OleFormat);

