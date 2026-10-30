// SHA-1, SHA-256, SHA-384 and SHA-512 (FIPS 180-4), synchronous and dependency-free, for the
// password verifiers of sheet and workbook protection (docs/facts/sheet-rules.md SR-50,
// D-C08). WebCrypto's `crypto.subtle.digest` is asynchronous and needs one promise per hash,
// which made the 100,000 iterations Excel uses take about 5 s in Node.js 22; this takes a
// fraction of that. The round constants and initial values of SHA-256/384/512 are not written
// out: they are computed once from the square and cube roots of the first primes, as FIPS
// 180-4 §4.2.2, §4.2.3, §5.3.3–§5.3.5 define them. tests/unit/sha.test.ts compares every
// algorithm with node:crypto.

/**
 * The hash algorithms xlvane computes for password verifiers, named as in `algorithmName`
 * ([MS-OFFCRYPTO] §2.4.2.4).
 *
 * @example
 * ```ts
 * import { hashPassword, type ShaAlgorithm } from 'xlvane';
 *
 * const algorithm: ShaAlgorithm = 'SHA-256';
 * console.log(hashPassword('secret', { algorithm, spinCount: 1000 }).algorithmName);
 * ```
 */
export type ShaAlgorithm = 'SHA-1' | 'SHA-256' | 'SHA-384' | 'SHA-512';

/** @internal */
export const SHA_ALGORITHMS: readonly ShaAlgorithm[] = ['SHA-1', 'SHA-256', 'SHA-384', 'SHA-512'];

interface Tables {
  readonly k512: Uint32Array; // 80 pairs hi, lo
  readonly h512: Uint32Array; // 8 pairs
  readonly h384: Uint32Array; // 8 pairs
  readonly k256: Uint32Array; // 64
  readonly h256: Uint32Array; // 8
}

let tables: Tables | null = null;

function iroot(n: bigint, k: bigint): bigint {
  // Integer k-th root (floor) by Newton's method.
  let x = 1n << (BigInt(n.toString(2).length) / k + 1n);
  for (;;) {
    const y = ((k - 1n) * x + n / x ** (k - 1n)) / k;
    if (y >= x) return x;
    x = y;
  }
}

function getTables(): Tables {
  if (tables) return tables;
  const primes: number[] = [];
  for (let n = 2; primes.length < 80; n++) if (primes.every((p) => n % p !== 0)) primes.push(n);
  const mask = (1n << 64n) - 1n;
  const pair = (out: Uint32Array, i: number, v: bigint): void => {
    out[2 * i] = Number((v >> 32n) & 0xffffffffn);
    out[2 * i + 1] = Number(v & 0xffffffffn);
  };
  const k512 = new Uint32Array(160);
  const h512 = new Uint32Array(16);
  const h384 = new Uint32Array(16);
  const k256 = new Uint32Array(64);
  const h256 = new Uint32Array(8);
  // The first 64 bits of the fractional parts of the cube roots of the first 80 primes, and of
  // the square roots of the first 8 (SHA-512) and the 9th to 16th (SHA-384) primes.
  primes.forEach((p, i) => pair(k512, i, iroot(BigInt(p) << 192n, 3n) & mask));
  for (let i = 0; i < 8; i++) {
    pair(h512, i, iroot(BigInt(primes[i]!) << 128n, 2n) & mask);
    pair(h384, i, iroot(BigInt(primes[i + 8]!) << 128n, 2n) & mask);
  }
  // SHA-256 uses the first 32 of those bits.
  for (let i = 0; i < 64; i++) k256[i] = k512[2 * i]!;
  for (let i = 0; i < 8; i++) h256[i] = h512[2 * i]!;
  tables = { k512, h512, h384, k256, h256 };
  return tables;
}

function pad(data: Uint8Array, blockBytes: number, lengthBytes: number): Uint8Array {
  const total = Math.ceil((data.length + 1 + lengthBytes) / blockBytes) * blockBytes;
  const out = new Uint8Array(total);
  out.set(data);
  out[data.length] = 0x80;
  const bits = data.length * 8;
  const view = new DataView(out.buffer);
  view.setUint32(total - 4, bits >>> 0);
  view.setUint32(total - 8, Math.floor(bits / 0x100000000));
  return out;
}

function sha1(data: Uint8Array): Uint8Array {
  const m = pad(data, 64, 8);
  const view = new DataView(m.buffer);
  const h = new Uint32Array([0x67452301, 0xefcdab89, 0x98badcfe, 0x10325476, 0xc3d2e1f0]);
  const w = new Uint32Array(80);
  // floor(2^30 × √2, √3, √5, √10), FIPS 180-4 §4.2.1.
  const k = [0x5a827999, 0x6ed9eba1, 0x8f1bbcdc, 0xca62c1d6];
  for (let off = 0; off < m.length; off += 64) {
    for (let t = 0; t < 16; t++) w[t] = view.getUint32(off + 4 * t);
    for (let t = 16; t < 80; t++) {
      const x = w[t - 3]! ^ w[t - 8]! ^ w[t - 14]! ^ w[t - 16]!;
      w[t] = (x << 1) | (x >>> 31);
    }
    let a = h[0]!, b = h[1]!, c = h[2]!, d = h[3]!, e = h[4]!;
    for (let t = 0; t < 80; t++) {
      const f = t < 20 ? (b & c) | (~b & d) : t < 40 ? b ^ c ^ d : t < 60 ? (b & c) | (b & d) | (c & d) : b ^ c ^ d;
      const tmp = (((a << 5) | (a >>> 27)) + f + e + k[(t / 20) | 0]! + w[t]!) >>> 0;
      e = d;
      d = c;
      c = (b << 30) | (b >>> 2);
      b = a;
      a = tmp;
    }
    h[0] = h[0]! + a;
    h[1] = h[1]! + b;
    h[2] = h[2]! + c;
    h[3] = h[3]! + d;
    h[4] = h[4]! + e;
  }
  return toBytes(h, 5);
}

function sha256(data: Uint8Array): Uint8Array {
  const { k256, h256 } = getTables();
  const m = pad(data, 64, 8);
  const view = new DataView(m.buffer);
  const h = Uint32Array.from(h256);
  const w = new Uint32Array(64);
  const rotr = (x: number, n: number): number => (x >>> n) | (x << (32 - n));
  for (let off = 0; off < m.length; off += 64) {
    for (let t = 0; t < 16; t++) w[t] = view.getUint32(off + 4 * t);
    for (let t = 16; t < 64; t++) {
      const x = w[t - 15]!;
      const y = w[t - 2]!;
      const s0 = rotr(x, 7) ^ rotr(x, 18) ^ (x >>> 3);
      const s1 = rotr(y, 17) ^ rotr(y, 19) ^ (y >>> 10);
      w[t] = w[t - 16]! + s0 + w[t - 7]! + s1;
    }
    let a = h[0]!, b = h[1]!, c = h[2]!, d = h[3]!, e = h[4]!, f = h[5]!, g = h[6]!, hh = h[7]!;
    for (let t = 0; t < 64; t++) {
      const t1 = (hh + (rotr(e, 6) ^ rotr(e, 11) ^ rotr(e, 25)) + ((e & f) ^ (~e & g)) + k256[t]! + w[t]!) >>> 0;
      const t2 = ((rotr(a, 2) ^ rotr(a, 13) ^ rotr(a, 22)) + ((a & b) ^ (a & c) ^ (b & c))) >>> 0;
      hh = g;
      g = f;
      f = e;
      e = (d + t1) >>> 0;
      d = c;
      c = b;
      b = a;
      a = (t1 + t2) >>> 0;
    }
    h[0] = h[0]! + a;
    h[1] = h[1]! + b;
    h[2] = h[2]! + c;
    h[3] = h[3]! + d;
    h[4] = h[4]! + e;
    h[5] = h[5]! + f;
    h[6] = h[6]! + g;
    h[7] = h[7]! + hh;
  }
  return toBytes(h, 8);
}

function sha512(data: Uint8Array, init: Uint32Array, words: number): Uint8Array {
  const { k512 } = getTables();
  const m = pad(data, 128, 16);
  const view = new DataView(m.buffer);
  const h = Uint32Array.from(init);
  const w = new Uint32Array(160);
  const s = new Uint32Array(16);
  for (let off = 0; off < m.length; off += 128) {
    for (let t = 0; t < 32; t++) w[t] = view.getUint32(off + 4 * t);
    for (let t = 16; t < 80; t++) {
      // σ0 = ROTR1 ^ ROTR8 ^ SHR7 of w[t-15]; σ1 = ROTR19 ^ ROTR61 ^ SHR6 of w[t-2].
      const xh = w[2 * (t - 15)]!, xl = w[2 * (t - 15) + 1]!;
      const s0h = ((xh >>> 1) | (xl << 31)) ^ ((xh >>> 8) | (xl << 24)) ^ (xh >>> 7);
      const s0l = ((xl >>> 1) | (xh << 31)) ^ ((xl >>> 8) | (xh << 24)) ^ ((xl >>> 7) | (xh << 25));
      const yh = w[2 * (t - 2)]!, yl = w[2 * (t - 2) + 1]!;
      const s1h = ((yh >>> 19) | (yl << 13)) ^ ((yl >>> 29) | (yh << 3)) ^ (yh >>> 6);
      const s1l = ((yl >>> 19) | (yh << 13)) ^ ((yh >>> 29) | (yl << 3)) ^ ((yl >>> 6) | (yh << 26));
      const lo = (s0l >>> 0) + (s1l >>> 0) + w[2 * (t - 16) + 1]! + w[2 * (t - 7) + 1]!;
      w[2 * t + 1] = lo;
      w[2 * t] = (s0h >>> 0) + (s1h >>> 0) + w[2 * (t - 16)]! + w[2 * (t - 7)]! + Math.floor(lo / 0x100000000);
    }
    s.set(h);
    for (let t = 0; t < 80; t++) {
      const ah = s[0]!, al = s[1]!, eh = s[8]!, el = s[9]!;
      // Σ1(e) = ROTR14 ^ ROTR18 ^ ROTR41; Ch(e, f, g).
      const S1h = ((eh >>> 14) | (el << 18)) ^ ((eh >>> 18) | (el << 14)) ^ ((el >>> 9) | (eh << 23));
      const S1l = ((el >>> 14) | (eh << 18)) ^ ((el >>> 18) | (eh << 14)) ^ ((eh >>> 9) | (el << 23));
      const chh = (eh & s[10]!) ^ (~eh & s[12]!);
      const chl = (el & s[11]!) ^ (~el & s[13]!);
      const t1l = s[15]! + (S1l >>> 0) + (chl >>> 0) + k512[2 * t + 1]! + w[2 * t + 1]!;
      const t1h = s[14]! + (S1h >>> 0) + (chh >>> 0) + k512[2 * t]! + w[2 * t]! + Math.floor(t1l / 0x100000000);
      // Σ0(a) = ROTR28 ^ ROTR34 ^ ROTR39; Maj(a, b, c).
      const S0h = ((ah >>> 28) | (al << 4)) ^ ((al >>> 2) | (ah << 30)) ^ ((al >>> 7) | (ah << 25));
      const S0l = ((al >>> 28) | (ah << 4)) ^ ((ah >>> 2) | (al << 30)) ^ ((ah >>> 7) | (al << 25));
      const mjh = (ah & s[2]!) ^ (ah & s[4]!) ^ (s[2]! & s[4]!);
      const mjl = (al & s[3]!) ^ (al & s[5]!) ^ (s[3]! & s[5]!);
      const t2l = (S0l >>> 0) + (mjl >>> 0);
      const t2h = (S0h >>> 0) + (mjh >>> 0) + Math.floor(t2l / 0x100000000);
      s.copyWithin(2, 0, 14);
      const el2 = s[9]! + (t1l >>> 0);
      s[9] = el2;
      s[8] = s[8]! + t1h + Math.floor(el2 / 0x100000000);
      const al2 = (t1l >>> 0) + (t2l >>> 0);
      s[1] = al2;
      s[0] = t1h + t2h + Math.floor(al2 / 0x100000000);
    }
    for (let i = 0; i < 8; i++) {
      const lo = h[2 * i + 1]! + s[2 * i + 1]!;
      h[2 * i + 1] = lo;
      h[2 * i] = h[2 * i]! + s[2 * i]! + Math.floor(lo / 0x100000000);
    }
  }
  return toBytes(h, words);
}

function toBytes(h: Uint32Array, words: number): Uint8Array {
  const out = new Uint8Array(words * 4);
  const view = new DataView(out.buffer);
  for (let i = 0; i < words; i++) view.setUint32(4 * i, h[i]!);
  return out;
}

/** @internal The digest of `data` with `algorithm`. */
export function sha(algorithm: ShaAlgorithm, data: Uint8Array): Uint8Array {
  switch (algorithm) {
    case 'SHA-1':
      return sha1(data);
    case 'SHA-256':
      return sha256(data);
    case 'SHA-384':
      return sha512(data, getTables().h384, 12);
    case 'SHA-512':
      return sha512(data, getTables().h512, 16);
  }
}
