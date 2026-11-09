import { describe, expect, it } from "vitest";
import { unzipSync } from "fflate";
import { BytesIO } from "../../src/fileio.js";
import { BubbleChartData, Category, CategoryChartData, XyChartData } from "../../src/chart/data.js";
import { Chart } from "../../src/chart/chart.js";
import { PySequence, ElementProxy } from "../../src/shared.js";
import { XL_AXIS_CROSSES, XL_CHART_TYPE, XL_TICK_MARK } from "../../src/enum/chart.js";
import { ValueAxis } from "../../src/chart/axis.js";
import { parseXml } from "../../src/oxml/parser.js";
import { _nsmap } from "../../src/oxml/ns.js";
import { xpath } from "../../src/oxml/xpath.js";
import { formatG16, Workbook } from "../../src/xlsx/workbook.js";

const decode = (bytes: Uint8Array) => new TextDecoder().decode(bytes);
const query = (node: any, expression: string): any => xpath(node, expression, {
  ..._nsmap,
  s: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  vt: "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
});

function workbookParts(values: unknown[]) {
  const file = new BytesIO();
  const workbook = new Workbook(file, { in_memory: true });
  const worksheet = workbook.addWorksheet();
  worksheet.writeColumn(0, 0, values);
  workbook.close();
  return unzipSync(file.getvalue());
}

describe("Python percent-G rounding contract", () => {
  it.each([
    [1000000000000000.5, "1000000000000000"],
    [1000000000000001.5, "1000000000000002"],
    [-1000000000000000.5, "-1000000000000000"],
    [-1000000000000001.5, "-1000000000000002"],
    [1.2345678901234567, "1.234567890123457"],
    [0.0001, "0.0001"],
    [0.00001, "1E-05"],
    [1e16, "1E+16"],
    [1e23, "9.999999999999999E+22"],
    [-1e23, "-9.999999999999999E+22"],
    [999999999999999.9, "999999999999999.9"],
    [Number.MIN_VALUE, "4.940656458412465E-324"],
    [Number.MAX_VALUE, "1.797693134862316E+308"],
    [0, "0"],
    [-0, "-0"],
  ] as const)("formats %.16G for %s", (value, expected) => {
    expect(formatG16(value)).toBe(expected);
  });

  it("writes ties-to-even values into the serialized worksheet", () => {
    const parts = workbookParts([1000000000000000.5, -1000000000000001.5, 1e23, -1e23]);
    const sheet = parseXml(decode(parts["xl/worksheets/sheet1.xml"]));
    expect(query(sheet, ".//s:v").map((node: any) => node.text)).toEqual([
      "1000000000000000", "-1000000000000002", "9.999999999999999E+22", "-9.999999999999999E+22",
    ]);
  });
});

describe("python-pptx integer chart-type equivalence", () => {
  const categoryTypes = [
    "AREA", "AREA_STACKED", "AREA_STACKED_100",
    "BAR_CLUSTERED", "BAR_STACKED", "BAR_STACKED_100",
    "COLUMN_CLUSTERED", "COLUMN_STACKED", "COLUMN_STACKED_100",
    "DOUGHNUT", "DOUGHNUT_EXPLODED", "LINE", "LINE_MARKERS",
    "LINE_MARKERS_STACKED", "LINE_MARKERS_STACKED_100", "LINE_STACKED", "LINE_STACKED_100",
    "PIE", "PIE_EXPLODED", "RADAR", "RADAR_FILLED", "RADAR_MARKERS",
  ] as const;

  it.each(categoryTypes)("preserves %s XML for an integer argument", (name) => {
    const data = new CategoryChartData();
    data.categories = ["A", "B"];
    data.addSeries("Values", [1, 2]);
    const type = XL_CHART_TYPE[name];
    expect(data.xmlBytes(type.value)).toEqual(data.xmlBytes(type));
  });

  it.each([
    "XY_SCATTER", "XY_SCATTER_LINES", "XY_SCATTER_LINES_NO_MARKERS",
    "XY_SCATTER_SMOOTH", "XY_SCATTER_SMOOTH_NO_MARKERS",
  ] as const)("preserves %s XML for an integer argument", (name) => {
    const data = new XyChartData();
    data.addSeries("Values").addDataPoint(1, 2);
    const type = XL_CHART_TYPE[name];
    expect(data.xmlBytes(type.value)).toEqual(data.xmlBytes(type));
  });

  it.each(["BUBBLE", "BUBBLE_THREE_D_EFFECT"] as const)("preserves %s XML for an integer argument", (name) => {
    const data = new BubbleChartData();
    data.addSeries("Values").addDataPoint(1, 2, 3);
    const type = XL_CHART_TYPE[name];
    expect(data.xmlBytes(type.value)).toEqual(data.xmlBytes(type));
  });

  it("emits the requested pie explosion and bubble 3D effect", () => {
    const pie = new CategoryChartData();
    pie.categories = ["A"];
    pie.addSeries("Values", [1]);
    expect(parseXml(decode(pie.xmlBytes(69))).xpath(".//c:explosion")[0].get("val")).toBe("25");
    const bubble = new BubbleChartData();
    bubble.addSeries("Values").addDataPoint(1, 2, 3);
    expect(parseXml(decode(bubble.xmlBytes(87))).xpath(".//c:bubble3D")[0].get("val")).toBe("1");
  });
});

describe("python-pptx integer axis-enum setter contract", () => {
  it.each(["column", "bar", "date", "scatter"] as const)("preserves tick and crossing transitions on %s axes", (kind) => {
    for (const integers of [false, true]) {
      let data: CategoryChartData | XyChartData;
      if (kind === "scatter") {
        data = new XyChartData();
        data.addSeries("Values").addDataPoint(1, 2);
      } else {
        data = new CategoryChartData();
        data.categories = kind === "date" ? [new Date("2024-01-01T00:00:00Z")] : ["East"];
        data.addSeries("Values", [1]);
      }
      const type = kind === "scatter" ? XL_CHART_TYPE.XY_SCATTER : kind === "bar" ? XL_CHART_TYPE.BAR_CLUSTERED : XL_CHART_TYPE.COLUMN_CLUSTERED;
      const chart = new Chart(parseXml(decode(data.xmlBytes(type))), null);
      for (const axis of [chart.categoryAxis, chart.valueAxis]) {
        for (const property of ["majorTickMark", "minorTickMark"] as const) {
          for (const member of [XL_TICK_MARK.OUTSIDE, XL_TICK_MARK.CROSS, XL_TICK_MARK.INSIDE, XL_TICK_MARK.NONE, XL_TICK_MARK.CROSS]) {
            axis[property] = integers ? member.value : member;
            expect(axis[property]).toBe(member);
            const elements = axis._element.xpath(`./c:${property}`);
            if (member === XL_TICK_MARK.CROSS) expect(elements).toHaveLength(0);
            else {
              expect(elements).toHaveLength(1);
              expect(elements[0].get("val")).toBe(member.xmlValue);
            }
          }
        }
        if (!(axis instanceof ValueAxis)) continue;
        axis.crosses = integers ? -4114 : XL_AXIS_CROSSES.CUSTOM;
        expect(axis.crosses).toBe(XL_AXIS_CROSSES.CUSTOM);
        expect(axis.crossesAt).toBe(0);
        axis.crossesAt = 2.4;
        const crossing = axis._crossXAx.crossesAt;
        axis.crosses = integers ? -4114 : XL_AXIS_CROSSES.CUSTOM;
        expect(axis.crossesAt).toBe(2.4);
        expect(axis._crossXAx.crossesAt).toBe(crossing);
        for (const member of [XL_AXIS_CROSSES.MINIMUM, XL_AXIS_CROSSES.MAXIMUM, XL_AXIS_CROSSES.AUTOMATIC]) {
          axis.crosses = integers ? member.value : member;
          expect(axis.crosses).toBe(member);
          expect(axis.crossesAt).toBeNull();
          expect(axis._crossXAx.crosses.get("val")).toBe(member.xmlValue);
        }
      }
    }
  });
});

describe("XlsxWriter Unicode code-point string limit", () => {
  const prefix = "a".repeat(32766);
  it.each([
    ["label", "label", 0],
    ["a".repeat(32767), "a".repeat(32767), 0],
    [prefix + "😀", prefix + "😀", 0],
    [prefix + "😀x", prefix + "😀", -2],
    ["😀".repeat(32767), "😀".repeat(32767), 0],
    ["😀".repeat(32768), "😀".repeat(32767), -2],
    ["a".repeat(32768), "a".repeat(32767), -2],
  ] as const)("preserves complete characters in string case %#", (input, expected, status) => {
    const file = new BytesIO();
    const workbook = new Workbook(file);
    expect(workbook.addWorksheet().write(0, 0, input)).toBe(status);
    workbook.close();
    const parts = unzipSync(file.getvalue());
    const strings = query(parseXml(decode(parts["xl/sharedStrings.xml"])), ".//s:t");
    expect(strings.map((node: any) => node.text)).toEqual([expected]);
  });

  it.each(["category", "xy", "bubble"] as const)("preserves a full-length supplementary character in %s chart labels", (kind) => {
    const label = prefix + "😀";
    let data: CategoryChartData | XyChartData | BubbleChartData;
    let type = XL_CHART_TYPE.COLUMN_CLUSTERED;
    if (kind === "category") {
      data = new CategoryChartData();
      data.categories = [label];
      data.addSeries(label, [1]);
    } else if (kind === "xy") {
      data = new XyChartData();
      data.addSeries(label).addDataPoint(1, 2);
      type = XL_CHART_TYPE.XY_SCATTER;
    } else {
      data = new BubbleChartData();
      data.addSeries(label).addDataPoint(1, 2, 3);
      type = XL_CHART_TYPE.BUBBLE;
    }
    const chart = parseXml(decode(data.xmlBytes(type)));
    expect(chart.xpath(".//c:tx/c:strRef/c:strCache/c:pt/c:v")[0].text).toBe(label);
    if (kind === "category") expect(chart.xpath(".//c:cat/c:strRef/c:strCache/c:pt/c:v")[0].text).toBe(label);
    const strings = query(parseXml(decode(unzipSync(data.xlsxBlob)["xl/sharedStrings.xml"])), ".//s:t");
    expect(strings.map((node: any) => node.text)).toEqual(kind === "bubble" ? [label, "Size"] : [label]);
  });
});

describe("Python string-subclass sequence membership", () => {
  it.each([false, true])("compares category labels by value, hierarchical=%s", (hierarchical) => {
    const data = new CategoryChartData();
    if (hierarchical) data.categories.addCategory("Region").addSubCategory("East");
    else data.categories = ["East"];
    data.addSeries("Values", [1]);
    const chart = new Chart(parseXml(decode(data.xmlBytes(XL_CHART_TYPE.COLUMN_CLUSTERED))), null);
    const categories = chart.plots.at(0).categories;
    for (const sequence of [categories, ...categories.levels]) {
      const category = sequence.at(0);
      expect(sequence.includes(category.label)).toBe(true);
      expect(sequence.includes(category)).toBe(true);
      expect(sequence.includes(new String(category.label))).toBe(true);
      expect(sequence.includes("Missing")).toBe(false);
      expect(sequence.includes({ toString: () => category.label })).toBe(false);
    }
  });

  it("unwraps either string operand without changing object equality", () => {
    const element = parseXml('<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>');
    const object = { label: "East" };
    const items: unknown[] = ["East", new String("West"), object, new ElementProxy(element)];
    const sequence = new class extends PySequence<unknown> {
      at(index: number) { return items[index]; }
      get length() { return items.length; }
    }();
    expect(sequence.includes(new String("East"))).toBe(true);
    expect(sequence.includes("West")).toBe(true);
    expect(sequence.includes(new String("West"))).toBe(true);
    expect(sequence.includes(object)).toBe(true);
    expect(sequence.includes({ label: "East" })).toBe(false);
    expect(sequence.includes(new ElementProxy(element))).toBe(true);
    expect(sequence.includes(new ElementProxy(parseXml('<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>')))).toBe(false);
    expect(sequence.includes(null)).toBe(false);
    expect(sequence.includes(undefined)).toBe(false);
  });
});

describe("Gregorian category dates in both Excel epochs", () => {
  it.each([
    ["0001-01-01", -693594, -695055],
    ["0050-01-01", -675697, -677158],
    ["0099-01-01", -657800, -659261],
    ["1900-01-01", 1, -1460],
    ["1900-03-01", 61, -1401],
    ["1904-01-01", 1462, 0],
  ] as const)("preserves %s when discarding time", (day, serial1900, serial1904) => {
    for (const time of ["00:00:00", "12:34:56"]) {
      const label = new Date(`${day}T${time}Z`);
      const category = new Category(label, null);
      expect(category.numericStrVal()).toBe(`${serial1900}.0`);
      expect(category.numericStrVal(true)).toBe(`${serial1904}.0`);
      expect(label.toISOString()).toBe(`${day}T${time}.000Z`);
    }
  });

  it("puts the year-50 serial in both the chart cache and embedded workbook", () => {
    const data = new CategoryChartData();
    data.categories = [new Date("0050-01-01T00:00:00Z")];
    data.addSeries("Values", [1]);
    const chart = parseXml(decode(data.xmlBytes(XL_CHART_TYPE.COLUMN_CLUSTERED)));
    expect(chart.xpath(".//c:cat/c:numRef/c:numCache/c:pt/c:v")[0].text).toBe("-675697.0");
    const sheet = parseXml(decode(unzipSync(data.xlsxBlob)["xl/worksheets/sheet1.xml"]));
    expect(query(sheet, './/s:c[@r="A2"]/s:v')[0].text).toBe("-675697");
  });
});

describe("single Sheet1 XLSX package contract", () => {
  it.each([[1], ["label"]])("keeps worksheet metadata and relationships consistent for %j", (value) => {
    const parts = workbookParts([value]);
    expect(Object.keys(parts).filter((name) => name.startsWith("xl/worksheets/"))).toEqual(["xl/worksheets/sheet1.xml"]);
    const workbook = parseXml(decode(parts["xl/workbook.xml"]));
    const sheets = query(workbook, ".//s:sheet");
    expect(sheets).toHaveLength(1);
    expect(sheets[0].get("name")).toBe("Sheet1");
    expect(sheets[0].get("sheetId")).toBe("1");
    expect(sheets[0].get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")).toBe("rId1");
    const rels = parseXml(decode(parts["xl/_rels/workbook.xml.rels"]));
    const relationships = rels.xpath("./pr:Relationship").map((node: any) => [node.get("Id"), node.get("Target")]);
    expect(relationships).toEqual([
      ["rId1", "worksheets/sheet1.xml"], ["rId2", "theme/theme1.xml"], ["rId3", "styles.xml"],
      ...(typeof value === "string" ? [["rId4", "sharedStrings.xml"]] : []),
    ]);
    const contentTypes = parseXml(decode(parts["[Content_Types].xml"]));
    expect(contentTypes.xpath('./ct:Override[@PartName="/xl/worksheets/sheet1.xml"]')[0].get("ContentType"))
      .toBe("application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
    const app = parseXml(decode(parts["docProps/app.xml"]));
    expect(query(app, ".//vt:i4")[0].text).toBe("1");
    expect(query(app, ".//vt:vector[@baseType='lpstr']/vt:lpstr").map((node: any) => node.text)).toEqual(["Sheet1"]);
    const sheet = parseXml(decode(parts["xl/worksheets/sheet1.xml"]));
    expect(query(sheet, ".//s:sheetView")[0].get("tabSelected")).toBe("1");
  });
});
