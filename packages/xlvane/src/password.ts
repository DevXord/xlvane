// Password verifiers of sheet and workbook protection (docs/facts/sheet-rules.md SR-50 … SR-56;
// docs/decisions.md D-C08). Protection passwords are not encryption: anyone can remove them
// from a file. The verifier only lets a spreadsheet application check a password typed by its
// user, and `verifyPassword` lets callers do the same check; xlvane never needs a password to
// read or edit a protected sheet.
import { XlvError } from './errors';
import type { SheetProtection, WorkbookProtection } from './printing';
import { sha, SHA_ALGORITHMS, type ShaAlgorithm } from './util/sha';

/**
 * A strong password verifier as stored in `sheetProtection`, `workbookProtection` or
 * `protectedRange` ([MS-OFFCRYPTO] §2.4.2.4 "ISO Write Protection Method", SR-50). `hashValue`
 * and `saltValue` are base64.
 *
 * @example
 * ```ts
 * import { hashPassword, type PasswordHash } from 'xlvane';
 *
 * const hash: PasswordHash = hashPassword('secret');
 * console.log(hash.algorithmName, hash.spinCount);
 * ```
 */
export interface PasswordHash {
  readonly algorithmName: string;
  readonly hashValue: string;
  readonly saltValue: string;
  readonly spinCount: number;
}

/**
 * Options of {@link hashPassword}. The defaults are what the Excel-saved files of the test
 * corpus carry (SR-53): SHA-512, 100,000 iterations and a random 16-byte salt.
 *
 * @example
 * ```ts
 * import { hashPassword, type HashPasswordOptions } from 'xlvane';
 *
 * // A fixed salt gives the same hash every time, for reproducible files.
 * const options: HashPasswordOptions = { salt: 'c2FsdHNhbHRzYWx0c2FsdA==', spinCount: 1000 };
 * console.log(hashPassword('secret', options).hashValue);
 * ```
 */
export interface HashPasswordOptions {
  /** `'SHA-512'` (default), `'SHA-384'`, `'SHA-256'` or `'SHA-1'`. */
  readonly algorithm?: ShaAlgorithm;
  /** Iterations after the first hash, 0 to 10,000,000. Default 100,000. */
  readonly spinCount?: number;
  /** The salt as bytes or base64, at most 65,536 bytes. Default: 16 random bytes. */
  readonly salt?: Uint8Array | string;
}

/**
 * Anything that carries a password verifier: a {@link PasswordHash}, a sheet's
 * {@link SheetProtection} or the workbook's {@link WorkbookProtection} as read. A legacy
 * `password` is the 4-hex-digit verifier of Transitional files.
 *
 * @example
 * ```ts
 * import type { PasswordVerifier } from 'xlvane';
 *
 * const legacy: PasswordVerifier = { password: 'DAA7' };
 * console.log(legacy.password);
 * ```
 */
export interface PasswordVerifier {
  readonly algorithmName?: string | null;
  readonly hashValue?: string | null;
  readonly saltValue?: string | null;
  readonly spinCount?: number | null;
  readonly password?: string | null;
}

const MAX_SPIN = 10_000_000;
const MAX_SALT = 65_536;
const MAX_PASSWORD = 255;

function invalid(message: string, hint: string, option: string): XlvError {
  return new XlvError('XLV_E_INVALID_OPTION', message, { hint, details: { option } });
}

const B64 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';

/** @internal Base64 (RFC 4648 §4) of `bytes`. */
export function toBase64(bytes: Uint8Array): string {
  let out = '';
  for (let i = 0; i < bytes.length; i += 3) {
    const n = (bytes[i]! << 16) | ((bytes[i + 1] ?? 0) << 8) | (bytes[i + 2] ?? 0);
    out += B64[(n >> 18) & 63]! + B64[(n >> 12) & 63]!;
    out += i + 1 < bytes.length ? B64[(n >> 6) & 63]! : '=';
    out += i + 2 < bytes.length ? B64[n & 63]! : '=';
  }
  return out;
}

/** @internal The bytes of base64 text (white space ignored), or `null` when it is not base64. */
export function fromBase64(text: string): Uint8Array | null {
  const s = text.replace(/[\t\n\r ]/g, '');
  if (!/^[A-Za-z0-9+/]*={0,2}$/.test(s) || s.length % 4 !== 0) return null;
  const pad = s.endsWith('==') ? 2 : s.endsWith('=') ? 1 : 0;
  const out = new Uint8Array((s.length / 4) * 3 - pad);
  let o = 0;
  for (let i = 0; i < s.length; i += 4) {
    let n = 0;
    for (let j = 0; j < 4; j++) n = (n << 6) | (s[i + j] === '=' ? 0 : B64.indexOf(s[i + j]!));
    if (o < out.length) out[o++] = (n >> 16) & 255;
    if (o < out.length) out[o++] = (n >> 8) & 255;
    if (o < out.length) out[o++] = n & 255;
  }
  return out;
}

/** The password as hashed: UTF-16LE without a leading byte order mark (ECMA-376 Part 1, SR-51). */
function passwordBytes(password: string): Uint8Array {
  const p = password.startsWith('\uFEFF') ? password.slice(1) : password;
  const out = new Uint8Array(p.length * 2);
  for (let i = 0; i < p.length; i++) {
    const c = p.charCodeAt(i);
    out[2 * i] = c & 255;
    out[2 * i + 1] = c >> 8;
  }
  return out;
}

/** [MS-OFFCRYPTO] §2.4.2.4: H0 = H(salt + password); Hn = H(Hn-1 + iterator), iterator 0 … spinCount − 1, 4 bytes little-endian. */
function isoHash(algorithm: ShaAlgorithm, salt: Uint8Array, password: Uint8Array, spinCount: number): Uint8Array {
  const first = new Uint8Array(salt.length + password.length);
  first.set(salt);
  first.set(password, salt.length);
  let h = sha(algorithm, first);
  const buf = new Uint8Array(h.length + 4);
  const view = new DataView(buf.buffer);
  for (let i = 0; i < spinCount; i++) {
    buf.set(h);
    view.setUint32(h.length, i, true);
    h = sha(algorithm, buf);
  }
  return h;
}

function checkPassword(password: unknown): string {
  if (typeof password !== 'string') throw invalid('The password must be a string.', "For example hashPassword('secret').", 'password');
  const p = password.startsWith('\uFEFF') ? password.slice(1) : password;
  if (p.length < 1 || p.length > MAX_PASSWORD) {
    throw invalid(`The password has ${String(p.length)} characters; it must have 1 to 255.`, 'Use a password of 1 to 255 characters ([MS-OFFCRYPTO] §2.4.2.4).', 'password');
  }
  return p;
}

function randomSalt(): Uint8Array {
  const c = (globalThis as { crypto?: { getRandomValues?: (a: Uint8Array) => Uint8Array } }).crypto;
  if (typeof c?.getRandomValues !== 'function') {
    throw new XlvError('XLV_E_NOT_SUPPORTED', 'This runtime has no crypto.getRandomValues, so a random salt cannot be made.', {
      hint: 'Pass your own random salt: hashPassword(password, { salt }), 16 bytes from a secure random source.',
      details: { option: 'salt' },
    });
  }
  return c.getRandomValues(new Uint8Array(16));
}

/**
 * Computes the verifier that protects a sheet or workbook with `password` ([MS-OFFCRYPTO]
 * §2.4.2.4, SR-50 … SR-53): the password as UTF-16LE, prefixed by the salt and hashed, then
 * hashed again `spinCount` times with a 4-byte little-endian counter. Pass the result to
 * `NewSheet.protect({ password })` or `NewWorkbook.protect({ password })`. Without a `salt` a
 * random one is made, so each call gives a different (equally valid) hash.
 *
 * A protection password only asks spreadsheet applications to refuse changes; it is not
 * encryption, and the file's content stays readable.
 *
 * @example
 * ```ts
 * import { createWorkbook, hashPassword, writeWorkbook } from 'xlvane';
 *
 * const wb = createWorkbook();
 * wb.addSheet('Prices').set('A1', 10).protect({ password: hashPassword('secret') });
 * const bytes = writeWorkbook(wb);
 * ```
 */
export function hashPassword(password: string, options: HashPasswordOptions = {}): PasswordHash {
  const p = checkPassword(password);
  if (typeof options !== 'object' || options === null) throw invalid('The hashPassword options must be an object.', "For example { algorithm: 'SHA-512' }.", 'options');
  for (const k of Object.keys(options)) {
    if (k !== 'algorithm' && k !== 'spinCount' && k !== 'salt') throw invalid(`hashPassword has no option "${k}".`, 'Use algorithm, spinCount or salt.', k);
  }
  const algorithm = options.algorithm ?? 'SHA-512';
  if (!SHA_ALGORITHMS.includes(algorithm)) {
    throw new XlvError('XLV_E_NOT_SUPPORTED', `xlvane cannot compute the hash algorithm ${JSON.stringify(String(algorithm).slice(0, 40))}.`, {
      hint: `Use one of ${SHA_ALGORITHMS.map((a) => `'${a}'`).join(', ')}; 'SHA-512' is what Excel writes.`,
      details: { option: 'algorithm' },
    });
  }
  const spinCount = options.spinCount ?? 100_000;
  if (typeof spinCount !== 'number' || !Number.isInteger(spinCount) || spinCount < 0 || spinCount > MAX_SPIN) {
    throw invalid('The option "spinCount" must be a whole number from 0 to 10000000.', 'For example { spinCount: 100000 }, the value Excel writes.', 'spinCount');
  }
  let salt: Uint8Array;
  if (options.salt === undefined) salt = randomSalt();
  else if (options.salt instanceof Uint8Array) salt = options.salt;
  else if (typeof options.salt === 'string' && fromBase64(options.salt) !== null) salt = fromBase64(options.salt)!;
  else throw invalid('The option "salt" must be a Uint8Array or base64 text.', "For example { salt: 'c2FsdHNhbHRzYWx0c2FsdA==' }.", 'salt');
  if (salt.length > MAX_SALT) throw invalid(`The salt has ${String(salt.length)} bytes; at most 65536 are allowed.`, 'Use a salt of 16 bytes, as Excel does.', 'salt');
  return Object.freeze({
    algorithmName: algorithm,
    hashValue: toBase64(isoHash(algorithm, salt, passwordBytes(p), spinCount)),
    saltValue: toBase64(salt),
    spinCount,
  });
}

/**
 * The legacy 16-bit verifier of Transitional files (`password` attributes; ECMA-376 Part 4
 * §15.3.3, [MS-OFFCRYPTO] §2.3.7.1; SR-54), as 4 uppercase hexadecimal digits, high byte first
 * as LibreOffice writes it (SR-55). Characters above U+00FF have no single-byte form here and
 * count as `?` (0x3F), the replacement [MS-XLS] §4 describes; the real conversion uses the
 * system's ANSI code page, which a file does not record. Weak by design: many passwords share a
 * verifier. Use {@link hashPassword} for new files.
 *
 * @example
 * ```ts
 * import { legacyPasswordHash } from 'xlvane';
 *
 * console.log(legacyPasswordHash('secret')); // "DAA7"
 * ```
 */
export function legacyPasswordHash(password: string): string {
  return legacyVerifier(password).toString(16).toUpperCase().padStart(4, '0');
}

function legacyVerifier(password: string): number {
  if (typeof password !== 'string') throw invalid('The password must be a string.', "For example legacyPasswordHash('secret').", 'password');
  const bytes = [...password].map((ch) => {
    const c = ch.codePointAt(0)!;
    return c <= 0xff ? c : 0x3f;
  });
  if (bytes.length === 0) return 0;
  let h = 0;
  for (let i = bytes.length - 1; i >= 0; i--) {
    h = ((h >> 14) & 1) | ((h << 1) & 0x7fff);
    h ^= bytes[i]!;
  }
  h = ((h >> 14) & 1) | ((h << 1) & 0x7fff);
  h ^= bytes.length;
  return h ^ 0xce4b;
}

function verifierOf(v: PasswordVerifier | SheetProtection | WorkbookProtection): PasswordVerifier {
  if ('lockStructure' in v) {
    const w = v;
    return { algorithmName: w.workbookAlgorithmName, hashValue: w.workbookHashValue, saltValue: w.workbookSaltValue, spinCount: w.workbookSpinCount, password: w.workbookPassword };
  }
  return v;
}

/**
 * Checks `password` against a stored verifier: a sheet's `protection`, the workbook's
 * `protection` (its workbook password) or a {@link PasswordHash}. Strong verifiers are
 * recomputed per [MS-OFFCRYPTO] §2.4.2.4; as [MS-XLSX] §2.6.56 and [MS-XLSB] §2.2.10 require,
 * the variant that hashes the 16-bit legacy verifier as a single character is tried too
 * (SR-56). A legacy `password` is compared as stored, in either byte order. Returns `false` when
 * there is no verifier. A check helper only: xlvane does not need passwords to read or edit
 * protected sheets.
 *
 * @example
 * ```ts
 * import { hashPassword, verifyPassword } from 'xlvane';
 *
 * const hash = hashPassword('secret', { spinCount: 1000 });
 * console.log(verifyPassword(hash, 'secret'), verifyPassword(hash, 'Secret')); // true false
 * ```
 */
export function verifyPassword(verifier: PasswordVerifier | SheetProtection | WorkbookProtection, password: string): boolean {
  if (typeof verifier !== 'object' || verifier === null) throw invalid('verifyPassword needs a protection or a password hash.', 'Pass sheet.protection, workbook.protection or the result of hashPassword.', 'verifier');
  if (typeof password !== 'string') throw invalid('The password must be a string.', "For example verifyPassword(sheet.protection, 'secret').", 'password');
  const v = verifierOf(verifier);
  if (typeof v.hashValue === 'string' && typeof v.algorithmName === 'string') {
    const algorithm = SHA_ALGORITHMS.find((a) => a.toLowerCase() === v.algorithmName!.toLowerCase().replace(/^sha(\d)/, 'sha-$1'));
    if (algorithm === undefined) {
      throw new XlvError('XLV_E_NOT_SUPPORTED', `xlvane cannot compute the hash algorithm ${JSON.stringify(v.algorithmName.slice(0, 40))}.`, {
        hint: `Only ${SHA_ALGORITHMS.join(', ')} are supported.`,
        details: { algorithmName: v.algorithmName.slice(0, 40) },
      });
    }
    const salt = typeof v.saltValue === 'string' ? fromBase64(v.saltValue) : new Uint8Array(0);
    const expected = fromBase64(v.hashValue);
    const spin = v.spinCount ?? 0;
    if (salt === null || expected === null || !Number.isInteger(spin) || spin < 0 || spin > MAX_SPIN || salt.length > MAX_SALT) return false;
    const p = password.startsWith('\uFEFF') ? password.slice(1) : password;
    const same = (h: Uint8Array): boolean => h.length === expected.length && h.every((b, i) => b === expected[i]);
    if (same(isoHash(algorithm, salt, passwordBytes(p), spin))) return true;
    // The verifier of the 16-bit legacy hash, read as one UTF-16 character ([MS-XLSX] §2.6.56).
    const legacy = legacyVerifier(p);
    return same(isoHash(algorithm, salt, new Uint8Array([legacy & 255, legacy >> 8]), spin));
  }
  if (typeof v.password === 'string' && /^[0-9a-f]{4}$/i.test(v.password)) {
    if (password.length === 0) return false;
    const stored = parseInt(v.password, 16);
    const computed = legacyVerifier(password);
    return stored === computed || stored === (((computed & 255) << 8) | (computed >> 8));
  }
  return false;
}
