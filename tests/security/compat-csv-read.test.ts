// @xlvane/compat csv.read after the 1.2.2 penetration tests (docs/decisions.md D-J01): the input is
// limited while it is collected, and the dayjs "word" expression of the date formats MMM, MMMM, A,
// a and Do no longer retries every position of a long run of digits (quadratic before).
import { describe, expect, it } from 'vitest';
import ExcelJS from '../../packages/compat/src/index';
import { csvValue } from '../../packages/compat/src/csv';

describe('csv.read', () => {
  it('stops an endless input at the input limit instead of collecting it', async () => {
    const chunk = new Uint8Array(1 << 20).fill(0x61);
    let pulled = 0;
    async function* endless(): AsyncGenerator<Uint8Array> {
      for (;;) {
        pulled++;
        yield chunk;
      }
    }
    await expect(new ExcelJS.Workbook().csv.read(endless())).rejects.toMatchObject({ code: 'XLV_E_LIMIT_EXCEEDED', details: { limit: 'maxInputBytes' } });
    expect(pulled).toBeLessThanOrEqual(257);
  }, 60_000);

  it('a field of 200 000 digits with word date formats takes linear time', () => {
    const digits = '9'.repeat(200_000);
    const start = performance.now();
    expect(csvValue(digits, ['MMM D YYYY', 'Do MMMM', 'h A'])).toBe(digits);
    expect(performance.now() - start).toBeLessThan(2000);
  });

  it('the word expression finds the same matches as the one of dayjs', () => {
    const dayjs = /\d*[^-_:/,()\s\d]+/;
    const ours = /(?<!\d)\d*[^-_:/,()\s\d]+/;
    const alphabet = ['1', '2', 'a', 'Z', 'é', '-', ':', ' ', '/', ',', '(', '.'];
    let seed = 1;
    const next = (): number => (seed = (seed * 1103515245 + 12345) % 2147483648);
    for (let i = 0; i < 20_000; i++) {
      let s = '';
      for (let n = next() % 12; n > 0; n--) s += alphabet[next() % alphabet.length];
      const a = dayjs.exec(s);
      const b = ours.exec(s);
      expect([b?.index, b?.[0]], s).toEqual([a?.index, a?.[0]]);
    }
    expect(csvValue('Jan 5 2024', ['MMM D YYYY'])).toBeInstanceOf(Date);
  });
});
