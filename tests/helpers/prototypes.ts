import { expect } from 'vitest';

const SNAPSHOT = new Map<object, string[]>(
  [Object.prototype, Array.prototype, Function.prototype, String.prototype, Map.prototype].map((proto) => [
    proto,
    Reflect.ownKeys(proto).map(String).sort(),
  ]),
);

/** Fails if any well-known prototype gained or lost a property since this module was loaded. */
export function expectPrototypesUntouched(): void {
  for (const [proto, keys] of SNAPSHOT) {
    expect(Reflect.ownKeys(proto).map(String).sort()).toEqual(keys);
  }
  const probe: Record<string, unknown> = {};
  for (const key of ['polluted', 'x', 'yes', 'isAdmin']) expect(probe[key]).toBeUndefined();
}
