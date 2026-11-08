// Abstract types used by pptx-ts (`pptx.types`).
import type { XmlPart } from "./opc/package.js";

/** An object that has width and height. */
export interface ProvidesExtents {
  /** Distance between top and bottom extents of shape in EMUs. */
  readonly height: number;
  /** Distance between left and right extents of shape in EMUs. */
  readonly width: number;
}

/**
 * An object that provides access to its XmlPart.
 *
 * This type is for objects that need access to their part, possibly because they need access to
 * the package or related parts.
 */
export interface ProvidesPart {
  readonly part: XmlPart;
}
