// Registers the part classes with `PartFactory` by content type, as python-pptx's
// `pptx/__init__.py` does. Importing the package (src/index.ts) loads this.
import { CONTENT_TYPE as CT } from "../opc/constants.js";
import { PartFactory, type Part } from "../opc/package.js";
import { ChartPart } from "./chart.js";
import { CorePropertiesPart } from "./coreprops.js";
import { ImagePart } from "./image.js";
import { MediaPart } from "./media.js";
import { PresentationPart } from "./presentation.js";
import { NotesMasterPart, NotesSlidePart, SlideLayoutPart, SlideMasterPart, SlidePart } from "./slide.js";

export const contentTypeToPartClassMap: Record<string, { load(...args: any[]): Part }> = {
  [CT.PML_PRESENTATION_MAIN]: PresentationPart,
  [CT.PML_PRES_MACRO_MAIN]: PresentationPart,
  [CT.PML_TEMPLATE_MAIN]: PresentationPart,
  [CT.PML_SLIDESHOW_MAIN]: PresentationPart,
  [CT.OPC_CORE_PROPERTIES]: CorePropertiesPart,
  [CT.PML_NOTES_MASTER]: NotesMasterPart,
  [CT.PML_NOTES_SLIDE]: NotesSlidePart,
  [CT.PML_SLIDE]: SlidePart,
  [CT.PML_SLIDE_LAYOUT]: SlideLayoutPart,
  [CT.PML_SLIDE_MASTER]: SlideMasterPart,
  [CT.DML_CHART]: ChartPart,
  [CT.BMP]: ImagePart,
  [CT.GIF]: ImagePart,
  [CT.JPEG]: ImagePart,
  [CT.MS_PHOTO]: ImagePart,
  [CT.PNG]: ImagePart,
  [CT.TIFF]: ImagePart,
  [CT.X_EMF]: ImagePart,
  [CT.X_WMF]: ImagePart,
  [CT.ASF]: MediaPart,
  [CT.AVI]: MediaPart,
  [CT.MOV]: MediaPart,
  [CT.MP4]: MediaPart,
  [CT.MPG]: MediaPart,
  [CT.MS_VIDEO]: MediaPart,
  [CT.SWF]: MediaPart,
  [CT.VIDEO]: MediaPart,
  [CT.WMV]: MediaPart,
  [CT.X_MS_VIDEO]: MediaPart,
  // accommodate "image/jpg" as an alias for "image/jpeg"
  "image/jpg": ImagePart,
};

Object.assign(PartFactory.partTypeFor, contentTypeToPartClassMap);
