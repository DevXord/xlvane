// Type inference of readTable with Standard Schema libraries (phase 3 acceptance criterion).
// `npm run typecheck` checks this file; Vitest's expectTypeOf assertions are compile-time.
import { type } from 'arktype';
import * as v from 'valibot';
import { describe, expectTypeOf, it } from 'vitest';
import { z } from 'zod';
import { readTable, readTableAsync, type TableResult, type TableRow } from '../../packages/xlvane/src/index';

declare const bytes: Uint8Array;

/** The assertions are checked by the compiler only; the bodies never run (`bytes` does not exist). */
function typeOnly(_check: () => void): void {}

describe('readTable row types', () => {
  it('infers the output type of a Zod schema', () => {
    typeOnly(() => {
    const schema = z.object({ name: z.string(), total: z.number().optional(), when: z.date() });
    const t = readTable(bytes, { schema });
    expectTypeOf(t.rows).toEqualTypeOf<z.infer<typeof schema>[]>();
    expectTypeOf(t.rows[0]!.name).toEqualTypeOf<string>();
    // Transforms change the output type.
    const parsed = readTable(bytes, { schema: z.object({ n: z.string().transform((s) => s.length) }) });
    expectTypeOf(parsed.rows).toEqualTypeOf<{ n: number }[]>();
    });
  });

  it('infers the output type of a Valibot schema', () => {
    typeOnly(() => {
    const schema = v.object({ name: v.string(), total: v.optional(v.number()) });
    const t = readTable(bytes, { schema });
    expectTypeOf(t.rows).toEqualTypeOf<v.InferOutput<typeof schema>[]>();
    });
  });

  it('infers the output type of an ArkType schema', () => {
    typeOnly(() => {
    const schema = type({ name: 'string', 'total?': 'number' });
    const t = readTable(bytes, { schema });
    expectTypeOf(t.rows).toEqualTypeOf<(typeof schema.infer)[]>();
    });
  });

  it('returns generic rows without a schema', () => {
    typeOnly(() => {
    expectTypeOf(readTable(bytes)).toEqualTypeOf<TableResult<TableRow>>();
    expectTypeOf(readTable(bytes, { sheet: 'S' }).rows[0]!['x']).toEqualTypeOf<string | number | boolean | Date | null | undefined>();
    });
  });

  it('works the same for readTableAsync', () => {
    typeOnly(() => {
    const schema = z.object({ id: z.number() });
    expectTypeOf(readTableAsync(bytes, { schema })).resolves.toEqualTypeOf<TableResult<{ id: number }>>();
    });
  });

  it('rejects things that are not schemas at compile time', () => {
    typeOnly(() => {
    // @ts-expect-error a plain object is not a Standard Schema
    readTable(bytes, { schema: { name: 'string' } });
    });
  });
});
