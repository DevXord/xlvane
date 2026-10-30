// Password verifiers of sheet and workbook protection (docs/facts/sheet-rules.md SR-50 … SR-56,
// docs/decisions.md D-C08). Test vectors:
// - exceljs 4.4.0 `Encryptor.convertPasswordToHash('secret', 'SHA512', salt, n)` (an independent
//   implementation of [MS-OFFCRYPTO] §2.4.2.4, bench/node_modules/exceljs/lib/utils/encryptor.js),
//   run once on 2026-10-30 (SR-57);
// - LibreOffice 24.2.7.2 protecting a sheet with "secret" writes password="daa7"
//   (tests/libreoffice/password.test.ts, SR-55);
// - ECMA-376 Part 4 §14.8.1 worked value for "Example": 0xED7E ([MS-OFFCRYPTO] report R4 Item 3).
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { hashPassword, legacyPasswordHash, verifyPassword } from '../../packages/xlvane/src/password';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { writeWorkbookStream } from '../../packages/xlvane/src/stream-write';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { XlvError } from '../../packages/xlvane/src/errors';

const SALT = 'c2FsdHNhbHRzYWx0c2FsdA==';

async function collect(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const parts: Uint8Array[] = [];
  for await (const p of stream as unknown as AsyncIterable<Uint8Array>) parts.push(p);
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
}

function code(fn: () => unknown): string {
  try {
    fn();
  } catch (error) {
    return error instanceof XlvError ? error.code : 'OTHER';
  }
  return 'NONE';
}

describe('hashPassword', () => {
  it('reproduces the exceljs hashes ([MS-OFFCRYPTO] §2.4.2.4, SR-57)', () => {
    expect(hashPassword('secret', { salt: SALT })).toEqual({
      algorithmName: 'SHA-512',
      hashValue: 'CuF5axNVXUBkO2qloE31bcdYHWdOluK7fApMpNlHDYFGW9qj/IRvc1titSnw4LZSDIf/LiR6uQJUpw/fHH0oWA==',
      saltValue: SALT,
      spinCount: 100_000,
    });
    expect(hashPassword('secret', { salt: 'XBC4HCRwrjnsaOYawLGxWw==', spinCount: 1000 }).hashValue).toBe(
      'njCaQTXGfZF0ezFkDDjiJF6T+o5AIlsR/CeqxuVSYlvMqI3zncQyWs4ozi0aEoyAGgxd1jCSK1Zki/rFlgsCpQ==',
    );
  });

  it('defaults to SHA-512, 100000 iterations and a random 16-byte salt', () => {
    const a = hashPassword('secret', { spinCount: 10 });
    const b = hashPassword('secret', { spinCount: 10 });
    expect(a.algorithmName).toBe('SHA-512');
    expect(atob(a.saltValue)).toHaveLength(16);
    expect(a.saltValue).not.toBe(b.saltValue);
    expect(hashPassword('x').spinCount).toBe(100_000);
  });

  it('removes a leading byte order mark and supports SHA-1, SHA-256 and SHA-384', () => {
    expect(hashPassword('\uFEFFsecret', { salt: SALT, spinCount: 5 })).toEqual(hashPassword('secret', { salt: SALT, spinCount: 5 }));
    for (const algorithm of ['SHA-1', 'SHA-256', 'SHA-384'] as const) {
      const h = hashPassword('secret', { salt: SALT, spinCount: 3, algorithm });
      expect(h.algorithmName).toBe(algorithm);
      expect(verifyPassword(h, 'secret')).toBe(true);
    }
  });

  it('checks the limits of [MS-OFFCRYPTO] §2.4.2.4', () => {
    expect(code(() => hashPassword(''))).toBe('XLV_E_INVALID_OPTION');
    expect(code(() => hashPassword('x'.repeat(256), { spinCount: 0 }))).toBe('XLV_E_INVALID_OPTION');
    expect(code(() => hashPassword('x'.repeat(255), { spinCount: 0 }))).toBe('NONE');
    expect(code(() => hashPassword('x', { spinCount: 10_000_001 }))).toBe('XLV_E_INVALID_OPTION');
    expect(code(() => hashPassword('x', { salt: new Uint8Array(65_537) }))).toBe('XLV_E_INVALID_OPTION');
    expect(code(() => hashPassword('x', { salt: 'not base64!' }))).toBe('XLV_E_INVALID_OPTION');
    expect(code(() => hashPassword('x', { algorithm: 'MD5' as never }))).toBe('XLV_E_NOT_SUPPORTED');
    expect(code(() => hashPassword('x', { unknown: 1 } as never))).toBe('XLV_E_INVALID_OPTION');
  });
});

describe('legacy verifier', () => {
  it('gives the documented and observed values, high byte first (SR-54, SR-55)', () => {
    expect(legacyPasswordHash('Example')).toBe('ED7E');
    expect(legacyPasswordHash('secret')).toBe('DAA7');
    expect(legacyPasswordHash('')).toBe('0000');
    expect(legacyPasswordHash('żółw')).toBe(legacyPasswordHash('?ó?w'));
  });

  it('verifies legacy passwords in either byte order, case-insensitively', () => {
    expect(verifyPassword({ password: 'daa7' }, 'secret')).toBe(true);
    expect(verifyPassword({ password: 'A7DA' }, 'secret')).toBe(true);
    expect(verifyPassword({ password: 'DAA7' }, 'Secret')).toBe(false);
    expect(verifyPassword({ password: 'DAA7' }, '')).toBe(false);
    expect(verifyPassword({}, 'secret')).toBe(false);
  });
});

describe('verifyPassword', () => {
  it('accepts the password, rejects others, and tries the 16-bit-verifier variant (SR-56)', () => {
    const h = hashPassword('secret', { salt: SALT, spinCount: 50 });
    expect(verifyPassword(h, 'secret')).toBe(true);
    expect(verifyPassword(h, 'secret ')).toBe(false);
    // A hash of the legacy verifier 0xDAA7 read as one UTF-16 character.
    const variant = hashPassword('\uDAA7', { salt: SALT, spinCount: 50 });
    expect(verifyPassword(variant, 'secret')).toBe(true);
    expect(verifyPassword({ ...h, saltValue: '', hashValue: hashPassword('secret', { salt: new Uint8Array(0), spinCount: 50 }).hashValue }, 'secret')).toBe(true);
    expect(verifyPassword({ ...h, algorithmName: 'sha512' }, 'secret')).toBe(true);
    expect(code(() => verifyPassword({ ...h, algorithmName: 'WHIRLPOOL' }, 'secret'))).toBe('XLV_E_NOT_SUPPORTED');
    expect(verifyPassword({ ...h, hashValue: '***' }, 'secret')).toBe(false);
  });
});

describe('protecting with a password', () => {
  it('writes the verifier on sheetProtection and workbookProtection, and reads it back', async () => {
    const hash = hashPassword('secret', { salt: SALT, spinCount: 100 });
    const wb = createWorkbook().protect({ password: hash });
    wb.addSheet('S').set('A1', 1).protect({ password: hash, sort: false });
    wb.addSheet('T').set('A1', 1).protect({ password: 'other' });
    const bytes = writeWorkbook(wb);
    const files = unzipSync(bytes);
    const sheetXml = strFromU8(files['xl/worksheets/sheet1.xml']!);
    expect(sheetXml).toContain(`<sheetProtection algorithmName="SHA-512" hashValue="${hash.hashValue}" saltValue="${SALT}" spinCount="100" sheet="1" sort="0"/>`);
    const workbookXml = strFromU8(files['xl/workbook.xml']!);
    expect(workbookXml).toMatch(/<workbookProtection lockStructure="1" workbookAlgorithmName="SHA-512" workbookHashValue="[^"]+" workbookSaltValue="c2FsdHNhbHRzYWx0c2FsdA==" workbookSpinCount="100"\/><bookViews>/);
    const read = readWorkbook(bytes);
    expect(verifyPassword(read.getSheet('S').protection!, 'secret')).toBe(true);
    expect(verifyPassword(read.getSheet('T').protection!, 'other')).toBe(true);
    expect(verifyPassword(read.getSheet('T').protection!, 'secret')).toBe(false);
    expect(read.protection?.lockStructure).toBe(true);
    expect(verifyPassword(read.protection!, 'secret')).toBe(true);
    // The streaming writer writes the same.
    const streamed = unzipSync(
      await collect(writeWorkbookStream({ protection: { password: hash }, sheets: [{ name: 'S', protect: { password: hash, sort: false }, rows: [[1]] }] })),
    );
    expect(strFromU8(streamed['xl/worksheets/sheet1.xml']!)).toContain(`hashValue="${hash.hashValue}"`);
    expect(strFromU8(streamed['xl/workbook.xml']!)).toContain('<workbookProtection lockStructure="1" workbookAlgorithmName="SHA-512"');
  });

  it('refuses malformed verifiers and removes workbook protection with null', () => {
    const sheet = createWorkbook().addSheet('S');
    expect(code(() => sheet.protect({ password: { algorithmName: 'SHA-512', hashValue: 'x', saltValue: '', spinCount: 1 } as never }))).toBe('XLV_E_INVALID_OPTION');
    expect(code(() => sheet.protect({ password: 42 as never }))).toBe('XLV_E_INVALID_OPTION');
    const wb = createWorkbook().protect({ lockWindows: true }).protect(null);
    wb.addSheet('S');
    expect(strFromU8(unzipSync(writeWorkbook(wb))['xl/workbook.xml']!)).not.toContain('workbookProtection');
  });
});
