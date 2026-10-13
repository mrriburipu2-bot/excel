// Minimal JSON Schema checker for tool inputs. Tool inputs stream eagerly, so
// the API does not validate them; a truncated or malformed input must be
// caught here before it touches the workbook.

export interface JsonSchema {
  type?: string | string[];
  properties?: Record<string, JsonSchema>;
  required?: string[];
  items?: JsonSchema;
  enum?: readonly unknown[];
  description?: string;
  minItems?: number;
  additionalProperties?: boolean | JsonSchema;
}

function typeOf(value: unknown): string {
  if (value === null) return "null";
  if (Array.isArray(value)) return "array";
  if (typeof value === "number") return Number.isInteger(value) ? "integer" : "number";
  return typeof value;
}

function matchesType(value: unknown, type: string): boolean {
  const actual = typeOf(value);
  if (type === "number") return actual === "number" || actual === "integer";
  return actual === type;
}

/** Returns a list of problems; an empty list means the value is valid. */
export function validate(value: unknown, schema: JsonSchema, path = "input"): string[] {
  const errors: string[] = [];
  if (schema.type) {
    const types = Array.isArray(schema.type) ? schema.type : [schema.type];
    if (!types.some((t) => matchesType(value, t))) {
      return [`${path} should be ${types.join(" or ")} but is ${typeOf(value)}`];
    }
  }
  if (schema.enum && !schema.enum.includes(value)) {
    errors.push(`${path} must be one of ${schema.enum.map((e) => JSON.stringify(e)).join(", ")}`);
  }
  if (Array.isArray(value)) {
    if (schema.minItems !== undefined && value.length < schema.minItems) {
      errors.push(`${path} needs at least ${schema.minItems} item(s)`);
    }
    if (schema.items) value.forEach((v, i) => errors.push(...validate(v, schema.items!, `${path}[${i}]`)));
  } else if (value !== null && typeof value === "object") {
    const obj = value as Record<string, unknown>;
    for (const key of schema.required ?? []) {
      if (!(key in obj) || obj[key] === undefined) errors.push(`${path}.${key} is required`);
    }
    for (const [key, sub] of Object.entries(schema.properties ?? {})) {
      if (key in obj && obj[key] !== undefined) errors.push(...validate(obj[key], sub, `${path}.${key}`));
    }
  }
  return errors;
}
