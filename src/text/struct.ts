// The parts of python's `struct` module that font-file parsing uses: `calcsize()` and
// `unpackFrom()` for big- or little-endian formats with s, b, B, h, H, i, I, l, L, q, Q codes.
// 64-bit fields are read as numbers (exact below 2**53).

const SIZES: Record<string, number> = { x: 1, s: 1, b: 1, B: 1, h: 2, H: 2, i: 4, I: 4, l: 4, L: 4, q: 8, Q: 8 };

/** A struct format given as a string or as bytes (python accepts both). */
export type StructFormat = string | Uint8Array;

const fmtText = (fmt: StructFormat): string => (typeof fmt === "string" ? fmt : new TextDecoder().decode(fmt));

function* fields(fmt: string): Generator<[count: number, code: string]> {
  const re = /(\d*)([xsbBhHiIlLqQ])/g;
  const body = /^[<>!=@]/.test(fmt) ? fmt.slice(1) : fmt;
  let m: RegExpExecArray | null;
  while ((m = re.exec(body)) !== null) yield [m[1] ? parseInt(m[1], 10) : 1, m[2]];
}

/** python `struct.calcsize(fmt)` (standard sizes, no alignment padding). */
export function calcsize(format: StructFormat): number {
  const fmt = fmtText(format);
  let n = 0;
  for (const [count, code] of fields(fmt)) n += SIZES[code] * count;
  return n;
}

/** python `struct.unpack_from(fmt, buffer, offset)`; `s` fields are returned as bytes. */
export function unpackFrom(format: StructFormat, buffer: Uint8Array, offset = 0): any[] {
  const fmt = fmtText(format);
  const le = fmt.startsWith("<");
  if (buffer.length < offset + calcsize(fmt)) throw new RangeError(`unpack_from requires a buffer of at least ${offset + calcsize(fmt)} bytes`);
  const dv = new DataView(buffer.buffer, buffer.byteOffset, buffer.byteLength);
  const out: any[] = [];
  let o = offset;
  for (const [count, code] of fields(fmt)) {
    if (code === "s") {
      out.push(buffer.slice(o, o + count));
      o += count;
      continue;
    }
    for (let k = 0; k < count; k++) {
      switch (code) {
        case "x":
          break;
        case "b":
          out.push(dv.getInt8(o));
          break;
        case "B":
          out.push(dv.getUint8(o));
          break;
        case "h":
          out.push(dv.getInt16(o, le));
          break;
        case "H":
          out.push(dv.getUint16(o, le));
          break;
        case "i":
        case "l":
          out.push(dv.getInt32(o, le));
          break;
        case "I":
        case "L":
          out.push(dv.getUint32(o, le));
          break;
        case "q":
          out.push(Number(dv.getBigInt64(o, le)));
          break;
        case "Q":
          out.push(Number(dv.getBigUint64(o, le)));
          break;
      }
      o += SIZES[code];
    }
  }
  return out;
}
