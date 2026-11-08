// Base shape-related objects such as BaseShape (`pptx.oxml.shapes`).
import type { CT_Shape } from "./autoshape.js";
import type { CT_Connector } from "./connector.js";
import type { CT_GraphicalObjectFrame } from "./graphfrm.js";
import type { CT_GroupShape } from "./groupshape.js";
import type { CT_Picture } from "./picture.js";

/** Any of the shape elements that can appear in a shape tree. */
export type ShapeElement = CT_Shape | CT_Connector | CT_GraphicalObjectFrame | CT_GroupShape | CT_Picture;
