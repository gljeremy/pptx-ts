// Text measurement from a TrueType/OpenType font file: what python-pptx gets from Pillow's
// `ImageFont.truetype(path, size).getbbox(text)` for `TextFrame.fitText()`.
//
// Glyphs are found through the `cmap` table (formats 4 and 12), advanced with `hmtx`, adjusted
// by a legacy `kern` table when the font has one, and their ink extent comes from the `glyf`
// bounding boxes (or the font's `head` box for CFF-outline fonts). Pillow additionally applies
// FreeType's hinting at the pixel size, so sizes can differ from Pillow's by a pixel or so;
// `test/text/ttf.test.ts` records the difference on the test font.

/** A parsed font file, cached per path/bytes by `TrueTypeFont.load()`. */
export class TrueTypeFont {
  readonly unitsPerEm: number;
  readonly ascender: number;
  private readonly dv: DataView;
  private readonly tables = new Map<string, { offset: number; length: number }>();
  private readonly cmap: (cp: number) => number;
  private readonly numHMetrics: number;
  private readonly hmtx: number;
  private readonly loca: number[] | null;
  private readonly glyf: number;
  private readonly headBox: [number, number, number, number];
  private readonly kern = new Map<number, number>();

  constructor(bytes: Uint8Array) {
    this.dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const dv = this.dv;
    let base = 0;
    // a TrueType collection: use its first font
    if (dv.getUint32(0) === 0x74746366) base = dv.getUint32(12);
    const numTables = dv.getUint16(base + 4);
    for (let i = 0; i < numTables; i++) {
      const r = base + 12 + i * 16;
      const tag = String.fromCharCode(dv.getUint8(r), dv.getUint8(r + 1), dv.getUint8(r + 2), dv.getUint8(r + 3));
      this.tables.set(tag, { offset: dv.getUint32(r + 8), length: dv.getUint32(r + 12) });
    }
    const head = this.table("head");
    this.unitsPerEm = dv.getUint16(head + 18);
    this.headBox = [dv.getInt16(head + 36), dv.getInt16(head + 38), dv.getInt16(head + 40), dv.getInt16(head + 42)];
    const indexToLocFormat = dv.getInt16(head + 50);
    const hhea = this.table("hhea");
    this.ascender = dv.getInt16(hhea + 4);
    this.numHMetrics = dv.getUint16(hhea + 34);
    this.hmtx = this.table("hmtx");
    this.cmap = this.readCmap();
    const numGlyphs = dv.getUint16(this.table("maxp") + 4);
    if (this.tables.has("loca") && this.tables.has("glyf")) {
      const loca = this.table("loca");
      this.loca = [];
      for (let i = 0; i <= numGlyphs; i++) this.loca.push(indexToLocFormat === 0 ? dv.getUint16(loca + i * 2) * 2 : dv.getUint32(loca + i * 4));
      this.glyf = this.table("glyf");
    } else {
      this.loca = null;
      this.glyf = 0;
    }
    this.readKern();
  }

  private static cache = new Map<string | Uint8Array, TrueTypeFont>();

  /** The font in `source` (a path in Node, or the font file's bytes), parsed once. */
  static load(source: string | Uint8Array, read: (path: string) => Uint8Array): TrueTypeFont {
    let f = this.cache.get(source);
    if (f === undefined) {
      f = new TrueTypeFont(typeof source === "string" ? read(source) : source);
      this.cache.set(source, f);
    }
    return f;
  }

  private table(tag: string): number {
    const t = this.tables.get(tag);
    if (t === undefined) throw new Error(`font has no '${tag}' table`);
    return t.offset;
  }

  private readCmap(): (cp: number) => number {
    const dv = this.dv;
    const cmap = this.table("cmap");
    const n = dv.getUint16(cmap + 2);
    let best: { off: number; fmt: number; rank: number } | null = null;
    for (let i = 0; i < n; i++) {
      const r = cmap + 4 + i * 8;
      const platform = dv.getUint16(r);
      const encoding = dv.getUint16(r + 2);
      const off = cmap + dv.getUint32(r + 4);
      const fmt = dv.getUint16(off);
      // prefer full-Unicode (3,10 / 0,4+) format 12, then BMP format 4
      const rank = fmt === 12 ? 2 : fmt === 4 && (platform === 3 || platform === 0) && encoding !== 0 ? 1 : fmt === 4 ? 0.5 : -1;
      if (rank >= 0 && (best === null || rank > best.rank)) best = { off, fmt, rank };
    }
    if (best === null) return () => 0;
    const { off, fmt } = best;
    if (fmt === 12) {
      const nGroups = dv.getUint32(off + 12);
      return (cp) => {
        let lo = 0;
        let hi = nGroups - 1;
        while (lo <= hi) {
          const mid = (lo + hi) >> 1;
          const g = off + 16 + mid * 12;
          const start = dv.getUint32(g);
          const end = dv.getUint32(g + 4);
          if (cp < start) hi = mid - 1;
          else if (cp > end) lo = mid + 1;
          else return dv.getUint32(g + 8) + (cp - start);
        }
        return 0;
      };
    }
    const segX2 = dv.getUint16(off + 6);
    const ends = off + 14;
    const starts = ends + segX2 + 2;
    const deltas = starts + segX2;
    const rangeOffsets = deltas + segX2;
    return (cp) => {
      if (cp > 0xffff) return 0;
      for (let s = 0; s < segX2; s += 2) {
        if (cp > dv.getUint16(ends + s)) continue;
        const start = dv.getUint16(starts + s);
        if (cp < start) return 0;
        const ro = dv.getUint16(rangeOffsets + s);
        if (ro === 0) return (cp + dv.getInt16(deltas + s)) & 0xffff;
        const gi = dv.getUint16(rangeOffsets + s + ro + (cp - start) * 2);
        return gi === 0 ? 0 : (gi + dv.getInt16(deltas + s)) & 0xffff;
      }
      return 0;
    };
  }

  private readKern(): void {
    const t = this.tables.get("kern");
    if (!t) return;
    const dv = this.dv;
    const version = dv.getUint16(t.offset);
    if (version !== 0) return;
    const nTables = dv.getUint16(t.offset + 2);
    let p = t.offset + 4;
    for (let i = 0; i < nTables; i++) {
      const length = dv.getUint16(p + 2);
      const coverage = dv.getUint16(p + 4);
      // format 0, horizontal kerning values
      if (coverage >> 8 === 0 && (coverage & 1) === 1) {
        const nPairs = dv.getUint16(p + 6);
        for (let k = 0; k < nPairs; k++) {
          const e = p + 14 + k * 6;
          this.kern.set(dv.getUint16(e) * 65536 + dv.getUint16(e + 2), dv.getInt16(e + 4));
        }
      }
      p += length;
    }
  }

  /** Glyph index for code point `cp` (0, the .notdef glyph, when the font lacks it). */
  glyphIndex(cp: number): number {
    return this.cmap(cp);
  }

  /** Advance width of glyph `gid` in font units. */
  advance(gid: number): number {
    const i = Math.min(gid, this.numHMetrics - 1);
    return this.dv.getUint16(this.hmtx + i * 4);
  }

  /** [xMin, yMin, xMax, yMax] of glyph `gid` in font units; null for a glyph with no outline. */
  bbox(gid: number): [number, number, number, number] | null {
    if (this.loca === null) {
      const lsb = gid < this.numHMetrics ? this.dv.getInt16(this.hmtx + gid * 4 + 2) : 0;
      return [lsb, this.headBox[1], lsb + this.advance(gid), this.headBox[3]];
    }
    const start = this.loca[gid];
    if (this.loca[gid + 1] === start) return null;
    const g = this.glyf + start;
    return [this.dv.getInt16(g + 2), this.dv.getInt16(g + 4), this.dv.getInt16(g + 6), this.dv.getInt16(g + 8)];
  }

  /** Kerning adjustment between glyphs `left` and `right`, in font units. */
  kerning(left: number, right: number): number {
    return this.kern.get(left * 65536 + right) ?? 0;
  }

  /**
   * [left, top, right, bottom] pixel box of `text` rendered at `pixelSize` px, measured like
   * Pillow's `getbbox()` with its default "la" anchor (origin at the left of the ascender line).
   */
  getbbox(text: string, pixelSize: number): [number, number, number, number] {
    const scale = pixelSize / this.unitsPerEm;
    // positions in pixels; FreeType's hinted advances are whole pixels
    let pen = 0;
    let xMin = 0;
    let xMax = 0;
    let yMin = Infinity;
    let yMax = -Infinity;
    for (const ch of text) {
      // Pillow's basic layout applies no kerning, so neither does this
      const gid = this.glyphIndex(ch.codePointAt(0)!);
      const box = this.bbox(gid);
      if (box !== null) {
        xMin = Math.min(xMin, pen + Math.floor(box[0] * scale));
        xMax = Math.max(xMax, pen + Math.ceil(box[2] * scale));
        yMin = Math.min(yMin, box[1]);
        yMax = Math.max(yMax, box[3]);
      }
      pen += Math.round(this.advance(gid) * scale);
    }
    // the box covers the ink and at least the advance width, as FreeType's does
    xMax = Math.max(xMax, pen);
    if (yMax === -Infinity) return [xMin, 0, xMax, 0];
    const asc = Math.round(this.ascender * scale);
    return [xMin, asc - Math.round(yMax * scale), xMax, asc - Math.round(yMin * scale)];
  }
}
