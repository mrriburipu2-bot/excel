import { test } from "node:test";
import assert from "node:assert/strict";
import { validate, type JsonSchema } from "../src/shared/schema.ts";

const schema: JsonSchema = {
  type: "object",
  properties: {
    address: { type: "string" },
    values: { type: "array", minItems: 1, items: { type: "array", items: { type: ["string", "number", "boolean", "null"] } } },
    mode: { type: "string", enum: ["a", "b"] },
    count: { type: "integer" },
  },
  required: ["address", "values"],
};

test("accepts valid input", () => {
  assert.deepEqual(validate({ address: "A1", values: [["x", 1, null, true]], mode: "a", count: 3 }, schema), []);
});

test("reports missing, mistyped and out-of-enum fields", () => {
  const errs = validate({ values: [[{}]], mode: "z", count: 1.5 }, schema);
  assert.ok(errs.some((e) => e.includes("input.address is required")));
  assert.ok(errs.some((e) => e.includes("input.values[0][0]")));
  assert.ok(errs.some((e) => e.includes("input.mode must be one of")));
  assert.ok(errs.some((e) => e.includes("input.count should be integer")));
});

test("catches a truncated (empty) matrix", () => {
  assert.ok(validate({ address: "A1", values: [] }, schema).some((e) => e.includes("at least 1")));
});
