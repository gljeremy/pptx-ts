// pptx-ts: a TypeScript port of python-pptx.
//
//   import { Presentation, Inches } from "@gljeremy/pptx-ts";
//   const prs = Presentation();
//   const slide = prs.slides.addSlide(prs.slideLayouts.at(1));
//   slide.shapes.title!.textFrame.text = "Hello";
//   prs.save("hello.pptx");
import "./oxml/index.js";
import "./parts/index.js";

export { Presentation } from "./api.js";
export type { Presentation as PresentationObject } from "./presentation.js";

export * as util from "./util.js";
export { Length, Inches, Centipoints, Cm, Emu, Mm, Pt } from "./util.js";
export { BytesIO } from "./fileio.js";
export * from "./errors.js";
export * from "./exc.js";

// enumerations (pptx.enum.*)
export * from "./enum/action.js";
export * from "./enum/chart.js";
export * from "./enum/dml.js";
export * from "./enum/lang.js";
export * from "./enum/shapes.js";
export * from "./enum/text.js";
export { BaseEnum, BaseXmlEnum } from "./enum/base.js";

// presentation, slides and parts
export * as presentation from "./presentation.js";
export { Slide, Slides, SlideLayout, SlideLayouts, SlideMaster, SlideMasters, NotesMaster, NotesSlide } from "./slide.js";
export { Package } from "./package.js";
export { Image, ImagePart } from "./parts/image.js";
export { Video } from "./media.js";
export { CorePropertiesPart } from "./parts/coreprops.js";
export { RGBColor, ColorFormat } from "./dml/color.js";
export { FillFormat } from "./dml/fill.js";

// charts (pptx.chart.*)
export { BubbleChartData, CategoryChartData, ChartData, XyChartData } from "./chart/data.js";
export { Chart } from "./chart/chart.js";

// stage 1: the OPC package layer
export { OpcPackage, Part, XmlPart, PartFactory } from "./opc/package.js";
export { PackURI, PACKAGE_URI, CONTENT_TYPES_URI } from "./opc/packuri.js";
export { CONTENT_TYPE, NAMESPACE, RELATIONSHIP_TARGET_MODE, RELATIONSHIP_TYPE } from "./opc/constants.js";
export { parseXml, registerElementCls } from "./oxml/parser.js";
export { BaseOxmlElement, OxmlElement, defineChildren } from "./oxml/xmlchemy.js";
