// Synchronous SHA-1, used to identify duplicate image and media parts (python `hashlib.sha1`).
// WebCrypto's digest is async, and Node's crypto isn't available in browsers.

/** Lowercase hex SHA-1 digest of `data` (python `hashlib.sha1(data).hexdigest()`). */
export function sha1Hex(data: Uint8Array): string {
  const len = data.length;
  const nBlocks = ((len + 8) >>> 6) + 1;
  const w = new Int32Array(80);
  let h0 = 0x67452301,
    h1 = 0xefcdab89 | 0,
    h2 = 0x98badcfe | 0,
    h3 = 0x10325476,
    h4 = 0xc3d2e1f0 | 0;
  const tail = new Uint8Array(nBlocks * 64 - Math.floor(len / 64) * 64);
  const tailStart = Math.floor(len / 64) * 64;
  tail.set(data.subarray(tailStart));
  tail[len - tailStart] = 0x80;
  const bitLen = len * 8;
  const tv = new DataView(tail.buffer);
  tv.setUint32(tail.length - 4, bitLen >>> 0);
  tv.setUint32(tail.length - 8, Math.floor(bitLen / 0x100000000));
  const dv = new DataView(data.buffer, data.byteOffset, data.byteLength);
  for (let b = 0; b < nBlocks; b++) {
    const off = b * 64;
    const inTail = off >= tailStart;
    for (let i = 0; i < 16; i++) w[i] = inTail ? tv.getInt32(off - tailStart + i * 4) : dv.getInt32(off + i * 4);
    for (let i = 16; i < 80; i++) {
      const x = w[i - 3] ^ w[i - 8] ^ w[i - 14] ^ w[i - 16];
      w[i] = (x << 1) | (x >>> 31);
    }
    let a = h0,
      bb = h1,
      c = h2,
      d = h3,
      e = h4;
    for (let i = 0; i < 80; i++) {
      let f: number, k: number;
      if (i < 20) (f = (bb & c) | (~bb & d)), (k = 0x5a827999);
      else if (i < 40) (f = bb ^ c ^ d), (k = 0x6ed9eba1);
      else if (i < 60) (f = (bb & c) | (bb & d) | (c & d)), (k = 0x8f1bbcdc | 0);
      else (f = bb ^ c ^ d), (k = 0xca62c1d6 | 0);
      const t = (((a << 5) | (a >>> 27)) + f + e + k + w[i]) | 0;
      e = d;
      d = c;
      c = (bb << 30) | (bb >>> 2);
      bb = a;
      a = t;
    }
    h0 = (h0 + a) | 0;
    h1 = (h1 + bb) | 0;
    h2 = (h2 + c) | 0;
    h3 = (h3 + d) | 0;
    h4 = (h4 + e) | 0;
  }
  return [h0, h1, h2, h3, h4].map((h) => (h >>> 0).toString(16).padStart(8, "0")).join("");
}
