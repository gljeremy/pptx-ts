// Image format, pixel size and resolution read from file headers: the subset of Pillow
// (`PIL.Image.open(...)`: `.format`, `.size`, `.info["dpi"]`) that python-pptx uses.
//
// Values follow Pillow 11.0's plugins, including their quirks, so `Image.dpi` matches
// python-pptx: PNG dpi comes from a pHYs chunk in pixels per meter (floats), JPEG from a JFIF
// density in dots per inch (unit 1 only) or else EXIF X resolution, BMP from pixels per meter,
// TIFF from X/Y resolution and resolution unit, and WMF/EMF report a scalar "dpi" (which
// python-pptx then treats as missing).

/** Pillow's `UnidentifiedImageError`: the bytes aren't an image format it recognizes. */
export class UnidentifiedImageError extends Error {
  constructor(message = "cannot identify image file") {
    super(message);
    this.name = "UnidentifiedImageError";
  }
}

/** Pillow `info["dpi"]`: usually a pair, a single number for WMF/EMF, absent when unknown. */
export type RawDpi = readonly [number, number] | number | null;

/** [format, [widthPx, heightPx], dpi] as python-pptx's `Image._pil_props` returns. */
export type ImageProps = [format: string, size: [number, number], dpi: RawDpi];

/** Format, size and dpi of the image in `blob`; UnidentifiedImageError when it's not an image. */
export function readImageProps(blob: Uint8Array): ImageProps {
  const b = blob;
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  const startsWith = (sig: number[], at = 0) => b.length >= at + sig.length && sig.every((x, i) => b[at + i] === x);
  try {
    if (startsWith([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) return png(b, dv);
    if (startsWith([0xff, 0xd8, 0xff])) return jpeg(b, dv);
    if (startsWith([0x47, 0x49, 0x46, 0x38]) && (b[4] === 0x37 || b[4] === 0x39) && b[5] === 0x61) return ["GIF", [dv.getUint16(6, true), dv.getUint16(8, true)], null];
    if (startsWith([0x42, 0x4d])) return bmp(b, dv);
    if (startsWith([0x49, 0x49, 0x2a, 0x00]) || startsWith([0x4d, 0x4d, 0x00, 0x2a])) return tiff(dv, 0, "TIFF");
    if (startsWith([0xd7, 0xcd, 0xc6, 0x9a, 0x00, 0x00])) return placeableWmf(b, dv);
    if (startsWith([0x01, 0x00, 0x00, 0x00]) && startsWith([0x20, 0x45, 0x4d, 0x46], 40)) return emf(dv);
    const other = otherFormat(b, startsWith);
    if (other) return [other, [0, 0], null];
  } catch (e) {
    if (e instanceof RangeError) throw new UnidentifiedImageError();
    throw e;
  }
  throw new UnidentifiedImageError();
}

function png(b: Uint8Array, dv: DataView): ImageProps {
  // IHDR comes first; Pillow reads chunks up to the first IDAT on open
  let size: [number, number] = [0, 0];
  let dpi: RawDpi = null;
  for (let pos = 8; pos + 8 <= b.length; ) {
    const len = dv.getUint32(pos);
    const type = String.fromCharCode(b[pos + 4], b[pos + 5], b[pos + 6], b[pos + 7]);
    const data = pos + 8;
    if (type === "IHDR") size = [dv.getUint32(data), dv.getUint32(data + 4)];
    else if (type === "pHYs" && len >= 9 && b[data + 8] === 1) dpi = [dv.getUint32(data) * 0.0254, dv.getUint32(data + 4) * 0.0254];
    else if (type === "IDAT" || type === "IEND") break;
    pos = data + len + 4;
  }
  return ["PNG", size, dpi];
}

function jpeg(b: Uint8Array, dv: DataView): ImageProps {
  let size: [number, number] = [0, 0];
  let dpi: RawDpi = null;
  let exif: [number, number] | null = null;
  let pos = 2;
  while (pos + 4 <= b.length) {
    if (b[pos] !== 0xff) {
      pos++;
      continue;
    }
    const marker = b[pos + 1];
    if (marker === 0xff) {
      pos++;
      continue;
    }
    // markers without a length
    if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
      pos += 2;
      continue;
    }
    if (marker === 0xd9 || marker === 0xda) break;
    const len = dv.getUint16(pos + 2);
    const seg = pos + 4;
    if (marker === 0xe0 && len >= 16 && b[seg] === 0x4a && b[seg + 1] === 0x46 && b[seg + 2] === 0x49 && b[seg + 3] === 0x46 && b[seg + 4] === 0) {
      // JFIF: Pillow only takes the density as dpi when the unit is dots per inch
      if (b[seg + 7] === 1) dpi = [dv.getUint16(seg + 8), dv.getUint16(seg + 10)];
    } else if (marker === 0xe1 && exif === null && b[seg] === 0x45 && b[seg + 1] === 0x78 && b[seg + 2] === 0x69 && b[seg + 3] === 0x66 && b[seg + 4] === 0 && b[seg + 5] === 0) {
      exif = [seg + 6, len - 8];
    } else if ([0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf].includes(marker)) {
      size = [dv.getUint16(seg + 3), dv.getUint16(seg + 1)];
    }
    pos = seg + len - 2;
  }
  if (dpi === null && exif !== null) dpi = exifDpi(dv, exif[0]);
  return ["JPEG", size, dpi];
}

/** Pillow `JpegImageFile._read_dpi_from_exif`: X resolution for both axes, (72, 72) on any problem. */
function exifDpi(dv: DataView, tiffStart: number): RawDpi {
  try {
    const ifd = readIfd0(dv, tiffStart);
    const unit = ifd.get(0x0128);
    const xres = ifd.get(0x011a);
    if (unit === undefined || xres === undefined) return [72, 72];
    let dpi = xres;
    if (Number.isNaN(dpi) || !Number.isFinite(dpi)) return [72, 72];
    if (unit === 3) dpi *= 2.54;
    return [dpi, dpi];
  } catch {
    return [72, 72];
  }
}

function bmp(b: Uint8Array, dv: DataView): ImageProps {
  const headerSize = dv.getUint32(14, true);
  const h = 18;
  if (headerSize === 12) return ["BMP", [dv.getUint16(h, true), dv.getUint16(h + 2, true)], null];
  if ([40, 52, 56, 64, 108, 124].includes(headerSize)) {
    const yFlip = b[h + 7] === 0xff;
    const width = dv.getUint32(h, true);
    const height = yFlip ? 2 ** 32 - dv.getUint32(h + 4, true) : dv.getUint32(h + 4, true);
    const ppm: [number, number] = [dv.getUint32(h + 20, true), dv.getUint32(h + 24, true)];
    return ["BMP", [width, height], [ppm[0] / 39.3701, ppm[1] / 39.3701]];
  }
  throw new UnidentifiedImageError();
}

function tiff(dv: DataView, start: number, format: string): ImageProps {
  const ifd = readIfd0(dv, start);
  const width = ifd.get(0x0100) ?? 0;
  const height = ifd.get(0x0101) ?? 0;
  const xres = ifd.get(0x011a) ?? 1;
  const yres = ifd.get(0x011b) ?? 1;
  let dpi: RawDpi = null;
  if (xres && yres) {
    const resunit = ifd.get(0x0128);
    if (resunit === 2 || resunit === undefined) dpi = [xres, yres];
    else if (resunit === 3) dpi = [xres * 2.54, yres * 2.54];
  }
  return [format, [width, height], dpi];
}

/** Numeric values of the first IFD of the TIFF structure at `start` (first value of each tag). */
function readIfd0(dv: DataView, start: number): Map<number, number> {
  const le = dv.getUint8(start) === 0x49;
  const u16 = (o: number) => dv.getUint16(start + o, le);
  const u32 = (o: number) => dv.getUint32(start + o, le);
  const ifd = u32(4);
  const n = u16(ifd);
  const out = new Map<number, number>();
  const sizes: Record<number, number> = { 1: 1, 3: 2, 4: 4, 5: 8, 6: 1, 8: 2, 9: 4, 10: 8, 11: 4, 12: 8 };
  for (let i = 0; i < n; i++) {
    const e = ifd + 2 + i * 12;
    const tag = u16(e);
    const type = u16(e + 2);
    const count = u32(e + 4);
    const size = (sizes[type] ?? 0) * count;
    if (!size) continue;
    const off = size <= 4 ? e + 8 : u32(e + 8);
    let v: number;
    switch (type) {
      case 1:
        v = dv.getUint8(start + off);
        break;
      case 6:
        v = dv.getInt8(start + off);
        break;
      case 3:
        v = u16(off);
        break;
      case 8:
        v = dv.getInt16(start + off, le);
        break;
      case 4:
        v = u32(off);
        break;
      case 9:
        v = dv.getInt32(start + off, le);
        break;
      case 5:
        v = u32(off) / u32(off + 4);
        break;
      case 10:
        v = dv.getInt32(start + off, le) / dv.getInt32(start + off + 4, le);
        break;
      case 11:
        v = dv.getFloat32(start + off, le);
        break;
      case 12:
        v = dv.getFloat64(start + off, le);
        break;
      default:
        continue;
    }
    out.set(tag, v);
  }
  return out;
}

function placeableWmf(b: Uint8Array, dv: DataView): ImageProps {
  const inch = dv.getUint16(14, true);
  if (inch === 0) throw new UnidentifiedImageError();
  const [x0, y0, x1, y1] = [6, 8, 10, 12].map((o) => dv.getInt16(o, true));
  // sanity check (standard metafile header)
  if (!(b[22] === 1 && b[23] === 0 && b[24] === 9 && b[25] === 0)) throw new UnidentifiedImageError();
  const dpi = 72;
  return ["WMF", [Math.floor(((x1 - x0) * dpi) / inch), Math.floor(((y1 - y0) * dpi) / inch)], dpi];
}

function emf(dv: DataView): ImageProps {
  const [x0, y0, x1, y1] = [8, 12, 16, 20].map((o) => dv.getInt32(o, true));
  const frame = [24, 28, 32, 36].map((o) => dv.getInt32(o, true));
  // Pillow 11.0 computes xdpi from (x1 - y0); kept for identical results
  const xdpi = (2540.0 * (x1 - y0)) / (frame[2] - frame[0]);
  const ydpi = (2540.0 * (y1 - y0)) / (frame[3] - frame[1]);
  return ["WMF", [x1 - x0, y1 - y0], xdpi === ydpi ? xdpi : [xdpi, ydpi]];
}

/** Formats Pillow opens but python-pptx rejects (reported as "unsupported image format"). */
function otherFormat(b: Uint8Array, startsWith: (sig: number[], at?: number) => boolean): string | null {
  if (startsWith([0x52, 0x49, 0x46, 0x46]) && startsWith([0x57, 0x45, 0x42, 0x50], 8)) return "WEBP";
  if (startsWith([0x00, 0x00, 0x01, 0x00])) return "ICO";
  if (startsWith([0x38, 0x42, 0x50, 0x53])) return "PSD";
  if (startsWith([0x69, 0x63, 0x6e, 0x73])) return "ICNS";
  // "%!PS" text EPS, or the binary EPS header
  if (startsWith([0x25, 0x21, 0x50, 0x53]) || startsWith([0xc5, 0xd0, 0xd3, 0xc6])) return "EPS";
  return null;
}
