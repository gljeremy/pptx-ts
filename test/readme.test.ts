// The README's example, end to end; `scripts/crosscheck` and the README note that the same
// steps in python-pptx produce byte-identical parts.
import { expect, it } from "vitest";
import { unzipSync } from "fflate";
import { BytesIO, CategoryChartData, Inches, Presentation, XL_CHART_TYPE, XL_LEGEND_POSITION } from "../src/index.js";

it("runs the README example", () => {
  const prs = Presentation();
  const slide = prs.slides.addSlide(prs.slideLayouts.at(0));
  slide.shapes.title!.textFrame.text = "Hello, pptx-ts";
  slide.placeholders.get(1).textFrame.text = "built in TypeScript";
  prs.coreProperties.title = "Hello";
  expect([...prs.slides].map((s) => [s.slideId, s.slideLayout.name])).toEqual([[256, "Title Slide"]]);

  const out = new BytesIO();
  prs.save(out);
  const reopened = Presentation(out.getvalue());
  expect([...reopened.slides.at(0).shapes].map((sh: any) => sh.textFrame.text)).toEqual(["Hello, pptx-ts", "built in TypeScript"]);
  expect(reopened.coreProperties.title).toBe("Hello");
});

it("runs the README chart example", () => {
  const prs = Presentation();
  const data = new CategoryChartData();
  data.categories = ["East", "West", "Midwest"];
  data.addSeries("Q1 Sales", [19.2, 21.4, 16.7]);
  data.addSeries("Q2 Sales", [22.3, 28.6, 15.2]);

  const slide2 = prs.slides.addSlide(prs.slideLayouts.at(6));
  const chart = slide2.shapes.addChart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1), Inches(8), Inches(5), data).chart;
  chart.hasLegend = true;
  chart.legend.position = XL_LEGEND_POSITION.BOTTOM;
  chart.plots.at(0).hasDataLabels = true;

  const out = new BytesIO();
  prs.save(out);
  const shape: any = [...Presentation(out.getvalue()).slides.at(0).shapes][0];
  const c = shape.chart;
  expect(c.chartType).toBe(XL_CHART_TYPE.COLUMN_CLUSTERED);
  expect(c.legend.position).toBe(XL_LEGEND_POSITION.BOTTOM);
  expect(c.plots.at(0).hasDataLabels).toBe(true);
  expect([...c.plots.at(0).categories].map(String)).toEqual(["East", "West", "Midwest"]);
  expect([...c.series].map((s: any) => [s.name, s.values])).toEqual([
    ["Q1 Sales", [19.2, 21.4, 16.7]],
    ["Q2 Sales", [22.3, 28.6, 15.2]],
  ]);
  // the embedded workbook carries the same table
  const xlsx = unzipSync(c._workbook.xlsxPart.blob);
  const sst = new TextDecoder().decode(xlsx["xl/sharedStrings.xml"]);
  for (const s of ["East", "West", "Midwest", "Q1 Sales", "Q2 Sales"]) expect(sst).toContain(`<t>${s}</t>`);
  expect(new TextDecoder().decode(xlsx["xl/worksheets/sheet1.xml"])).toContain('<c r="C3" s="1"><v>28.6</v></c>') // python-pptx writes this cell exactly so;
});
