// Objects related to system font file lookup (`pptx.text.fonts`).
//
// python-pptx finds installed .ttf files by walking the OS font directories and reading their
// `name` tables; that, with text measurement in `./layout.ts`, is only used by
// `TextFrame.fitText()`. It is ported in stage 4 (gljeremy/pptx-ts#4).
import { NotImplementedError } from "../errors.js";

/** A class-based singleton serving as a lazy cache for system font details. */
export class FontFiles {
  /** Path to the font file for `family` with the given `isBold` and `isItalic` styles. */
  static find(_family: string, _isBold: boolean, _isItalic: boolean): string {
    throw new NotImplementedError("font file lookup for TextFrame.fitText() is ported in stage 4 (gljeremy/pptx-ts#4); pass fontFile explicitly once available");
  }
}
