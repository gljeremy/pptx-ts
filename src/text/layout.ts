// Objects related to layout of rendered text, such as TextFitter (`pptx.text.layout`).
//
// python-pptx measures rendered text with Pillow's TrueType support to find the largest font
// size that fits a shape (`TextFrame.fitText()`). The measuring is ported in stage 4
// (gljeremy/pptx-ts#4).
import { NotImplementedError } from "../errors.js";

/** Size-to-fit calculation for a single line of text inside a box (stage 4). */
export class TextFitter {
  /**
   * The largest integer point size not exceeding `maxSize` that allows `text` to fit inside
   * `extents` when rendered using the font defined in `fontFile`.
   */
  static bestFitFontSize(_text: string, _extents: [number, number], _maxSize: number, _fontFile: string): number {
    throw new NotImplementedError("TextFrame.fitText() text measurement is ported in stage 4 (gljeremy/pptx-ts#4)");
  }
}
