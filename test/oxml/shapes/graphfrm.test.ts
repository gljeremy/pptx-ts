// Machine-translated from python-pptx tests/oxml/shapes/test_graphfrm.py by
// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.
/* eslint-disable */
// @ts-nocheck
import * as $py from "../../pyshim.js";
import { $add, $assertEqual, $list, $mod, pytestClass } from "../../pyshim.js";
import { AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError } from "../../../src/errors.js";
import { InvalidXmlError, PackageNotFoundError, PythonPptxError } from "../../../src/exc.js";
import { defineChildren } from "../../../src/oxml/xmlchemy.js";
import { CT_GraphicalObjectFrame as CT_GraphicalObjectFrame$imported } from "../../../src/oxml/shapes/graphfrm.js";
import { xml } from "../../unitutil/cxml.js";
// python binds imported names at import time; patches made later don't change them
const CT_GraphicalObjectFrame = CT_GraphicalObjectFrame$imported;
void $py;


let CHART_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart";

let TABLE_URI = "http://schemas.openxmlformats.org/drawingml/2006/table";

const DescribeCT_GraphicalObjectFrame: any = {
  fixtures: {
    new_chart_graphicFrame_fixture: { args: [], fn: (self: any): any => {
        let cx: any, cy: any, expected_xml: any, id_: any, name: any, rId: any, x: any, xml_tmpl: any, y: any;
        [id_, name, rId, x, y, cx, cy] = $list([42, "foobar", "rId6", 1, 2, 3, 4]);
        xml_tmpl = xml($mod("p:graphicFrame/(p:nvGraphicFramePr/(p:cNvPr{id=42,name=foobar},p:cNvGraphicFramePr/a:graphicFrameLocks{noGrp=1},p:nvPr),p:xfrm/(a:off{x=1,y=2},a:ext{cx=3,cy=4}),a:graphic/a:graphicData{uri=%s}\"%%s\")", CHART_URI));
        expected_xml = $mod(xml_tmpl, $add($add("\n      ", xml("c:chart{r:id=rId6}")), "    "));
        return [id_, name, rId, x, y, cx, cy, expected_xml];
      } },
    new_graphicFrame_fixture: { args: [], fn: (self: any): any => {
        let cx: any, cy: any, expected_xml: any, id_: any, name: any, x: any, y: any;
        [id_, name, x, y, cx, cy] = $list([42, "foobar", 1, 2, 3, 4]);
        expected_xml = xml("p:graphicFrame/(p:nvGraphicFramePr/(p:cNvPr{id=42,name=foobar},p:cNvGraphicFramePr/a:graphicFrameLocks{noGrp=1},p:nvPr),p:xfrm/(a:off{x=1,y=2},a:ext{cx=3,cy=4}),a:graphic/a:graphicData)");
        return [id_, name, x, y, cx, cy, expected_xml];
      } },
    new_table_graphicFrame_fixture: { args: [], fn: (self: any): any => {
        let cols: any, cx: any, cy: any, expected_xml: any, id_: any, name: any, rows: any, x: any, y: any;
        [id_, name, rows, cols, x, y, cx, cy] = $list([42, "foobar", 1, 1, 1, 2, 3, 4]);
        expected_xml = xml($mod("p:graphicFrame/(p:nvGraphicFramePr/(p:cNvPr{id=42,name=foobar},p:cNvGraphicFramePr/a:graphicFrameLocks{noGrp=1},p:nvPr),p:xfrm/(a:off{x=1,y=2},a:ext{cx=3,cy=4}),a:graphic/a:graphicData{uri=%s}/a:tbl/(a:tblPr{firstRow=1,bandRow=1}/a:tableStyleId\"{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}\",a:tblGrid/a:gridCol{w=3},a:tr{h=4}/a:tc/(a:txBody/(a:bodyPr,a:lstStyle,a:p),a:tcPr)))", TABLE_URI));
        return [id_, name, rows, cols, x, y, cx, cy, expected_xml];
      } },
  },
  helpers: {
  },
  tests: [
    { name: "it_can_construct_a_new_graphicFrame", args: ["new_graphicFrame_fixture"], fn: (self: any, new_graphicFrame_fixture: any): any => {
        let cx: any, cy: any, expected_xml: any, graphicFrame: any, id_: any, name: any, x: any, y: any;
        [id_, name, x, y, cx, cy, expected_xml] = $list(new_graphicFrame_fixture);
        graphicFrame = CT_GraphicalObjectFrame.newGraphicFrame(id_, name, x, y, cx, cy);
        $assertEqual(graphicFrame.xml, expected_xml);
      } },
    { name: "it_can_construct_a_new_chart_graphicFrame", args: ["new_chart_graphicFrame_fixture"], fn: (self: any, new_chart_graphicFrame_fixture: any): any => {
        let cx: any, cy: any, expected_xml: any, graphicFrame: any, id_: any, name: any, rId: any, x: any, y: any;
        [id_, name, rId, x, y, cx, cy, expected_xml] = $list(new_chart_graphicFrame_fixture);
        graphicFrame = CT_GraphicalObjectFrame.newChartGraphicFrame(id_, name, rId, x, y, cx, cy);
        $assertEqual(graphicFrame.xml, expected_xml);
      } },
    { name: "it_can_construct_a_new_table_graphicFrame", args: ["new_table_graphicFrame_fixture"], fn: (self: any, new_table_graphicFrame_fixture: any): any => {
        let cols: any, cx: any, cy: any, expected_xml: any, graphicFrame: any, id_: any, name: any, rows: any, x: any, y: any;
        [id_, name, rows, cols, x, y, cx, cy, expected_xml] = $list(new_table_graphicFrame_fixture);
        graphicFrame = CT_GraphicalObjectFrame.newTableGraphicFrame(id_, name, rows, cols, x, y, cx, cy);
        $assertEqual(graphicFrame.xml, expected_xml);
      } },
  ],
};
pytestClass("DescribeCT_GraphicalObjectFrame", DescribeCT_GraphicalObjectFrame);

