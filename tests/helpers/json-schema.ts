// A small JSON Schema (draft 2020-12) validator for the subset used by docs/cli-json/*.schema.json
// (docs/decisions.md D-605). Unknown keywords throw, so the schemas cannot silently use a keyword
// this validator ignores.
import { readFileSync } from 'node:fs';
import { join } from 'node:path';

export type Schema = boolean | { readonly [keyword: string]: unknown };

const ANNOTATIONS = new Set(['$schema', '$id', '$defs', 'title', 'description', '$comment', 'examples']);
const KEYWORDS = new Set(['type', 'properties', 'required', 'additionalProperties', 'items', 'enum', 'const', 'oneOf', 'anyOf', '$ref', 'minimum', 'pattern']);

/** Loads schemas by file name from a directory; `$ref` may point into another file there. */
export class SchemaSet {
  private readonly cache = new Map<string, Schema>();
  constructor(private readonly dir: string) {}

  load(file: string): Schema {
    let s = this.cache.get(file);
    if (s === undefined) {
      s = JSON.parse(readFileSync(join(this.dir, file), 'utf8')) as Schema;
      this.cache.set(file, s);
    }
    return s;
  }

  /** Returns the validation errors of `value` against the schema in `file` (empty when valid). */
  validate(file: string, value: unknown): string[] {
    const errors: string[] = [];
    this.check(this.load(file), file, value, '$', errors);
    return errors;
  }

  private resolve(ref: string, file: string): { schema: Schema; file: string } {
    const [target, pointer = ''] = ref.split('#');
    const inFile = target === '' || target === undefined ? file : target;
    let node: unknown = this.load(inFile);
    for (const part of pointer.split('/').filter((p) => p !== '')) {
      node = (node as Record<string, unknown>)[part.replace(/~1/g, '/').replace(/~0/g, '~')];
      if (node === undefined) throw new Error(`Unresolved $ref ${ref} in ${file}`);
    }
    return { schema: node as Schema, file: inFile };
  }

  private check(schema: Schema, file: string, value: unknown, path: string, errors: string[]): void {
    if (schema === true) return;
    if (schema === false) {
      errors.push(`${path}: not allowed`);
      return;
    }
    for (const key of Object.keys(schema)) {
      if (!KEYWORDS.has(key) && !ANNOTATIONS.has(key)) throw new Error(`Unsupported JSON Schema keyword "${key}" in ${file}`);
    }
    if (typeof schema['$ref'] === 'string') {
      const r = this.resolve(schema['$ref'], file);
      this.check(r.schema, r.file, value, path, errors);
    }
    if (schema['type'] !== undefined) {
      const types = Array.isArray(schema['type']) ? (schema['type'] as string[]) : [schema['type'] as string];
      if (!types.some((t) => typeMatches(t, value))) {
        errors.push(`${path}: expected ${types.join(' or ')}, got ${describe(value)}`);
        return;
      }
    }
    if (schema['const'] !== undefined && JSON.stringify(schema['const']) !== JSON.stringify(value)) errors.push(`${path}: expected ${JSON.stringify(schema['const'])}`);
    if (Array.isArray(schema['enum']) && !schema['enum'].some((e) => JSON.stringify(e) === JSON.stringify(value))) {
      errors.push(`${path}: ${JSON.stringify(value)} is not one of ${JSON.stringify(schema['enum'])}`);
    }
    if (typeof schema['minimum'] === 'number' && typeof value === 'number' && value < schema['minimum']) errors.push(`${path}: less than ${String(schema['minimum'])}`);
    if (typeof schema['pattern'] === 'string' && typeof value === 'string' && !new RegExp(schema['pattern'], 'u').test(value)) errors.push(`${path}: does not match ${schema['pattern']}`);
    for (const [key, combine] of [['oneOf', 1], ['anyOf', 0]] as const) {
      const options = schema[key];
      if (!Array.isArray(options)) continue;
      const passing = options.filter((o) => {
        const e: string[] = [];
        this.check(o as Schema, file, value, path, e);
        return e.length === 0;
      }).length;
      if (combine === 1 ? passing !== 1 : passing === 0) {
        // Report the errors of the closest option to make failures readable.
        const details = options.map((o) => {
          const e: string[] = [];
          this.check(o as Schema, file, value, path, e);
          return e;
        });
        const best = details.sort((a, b) => a.length - b.length)[0] ?? [];
        errors.push(`${path}: ${key} matched ${String(passing)} schemas`, ...best.map((e) => `  ${e}`));
      }
    }
    if (isObject(value)) {
      const props = (schema['properties'] ?? {}) as Record<string, Schema>;
      for (const req of (schema['required'] ?? []) as string[]) if (!(req in value)) errors.push(`${path}: missing property "${req}"`);
      for (const [k, v] of Object.entries(value)) {
        if (k in props) this.check(props[k]!, file, v, `${path}.${k}`, errors);
        else if (schema['additionalProperties'] !== undefined) this.check(schema['additionalProperties'] as Schema, file, v, `${path}.${k}`, errors);
      }
    }
    if (Array.isArray(value) && schema['items'] !== undefined) value.forEach((v, i) => this.check(schema['items'] as Schema, file, v, `${path}[${String(i)}]`, errors));
  }
}

function isObject(v: unknown): v is Record<string, unknown> {
  return typeof v === 'object' && v !== null && !Array.isArray(v);
}

function typeMatches(type: string, v: unknown): boolean {
  switch (type) {
    case 'object':
      return isObject(v);
    case 'array':
      return Array.isArray(v);
    case 'string':
      return typeof v === 'string';
    case 'number':
      return typeof v === 'number' && Number.isFinite(v);
    case 'integer':
      return typeof v === 'number' && Number.isInteger(v);
    case 'boolean':
      return typeof v === 'boolean';
    case 'null':
      return v === null;
    default:
      throw new Error(`Unknown type ${type}`);
  }
}

function describe(v: unknown): string {
  return v === null ? 'null' : Array.isArray(v) ? 'array' : typeof v;
}
