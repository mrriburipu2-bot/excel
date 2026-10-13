import { test } from "node:test";
import assert from "node:assert/strict";
import {
  columnLetter,
  columnIndex,
  cellAddress,
  splitAddress,
  padMatrix,
  toFormulaInput,
  columnStats,
  resolveColumnOffset,
} from "../src/shared/excel-utils.ts";

test("column letters round-trip", () => {
  for (const [i, l] of [[0, "A"], [25, "Z"], [26, "AA"], [51, "AZ"], [701, "ZZ"], [702, "AAA"]] as const) {
    assert.equal(columnLetter(i), l);
    assert.equal(columnIndex(l), i);
  }
  assert.equal(columnIndex("A1"), -1);
  assert.equal(cellAddress(0, 2), "C1");
});

test("splitAddress handles quoted and plain sheet names", () => {
  assert.deepEqual(splitAddress("Sheet1!A1:B2"), { sheet: "Sheet1", ref: "A1:B2" });
  assert.deepEqual(splitAddress("'Q1 Sales'!B:B"), { sheet: "Q1 Sales", ref: "B:B" });
  assert.deepEqual(splitAddress("'Bob''s'!A1"), { sheet: "Bob's", ref: "A1" });
  assert.deepEqual(splitAddress(" C3 "), { sheet: null, ref: "C3" });
});

test("padMatrix and toFormulaInput square ragged rows", () => {
  assert.deepEqual(padMatrix([[1, 2, 3], [4]]), [[1, 2, 3], [4, null, null]]);
  assert.deepEqual(toFormulaInput([["a", null], ["=A1"]]), [["a", ""], ["=A1", ""]]);
});

test("columnStats summarises numeric and text columns", () => {
  const stats = columnStats(
    [
      ["Region", "Sales"],
      ["North", 10],
      ["South", 20],
      ["North", null],
      ["North", 30],
    ],
    true,
    1,
  );
  assert.equal(stats[0].column, "B");
  assert.equal(stats[0].header, "Region");
  assert.deepEqual(stats[0].topValues?.[0], { value: "North", count: 3 });
  assert.equal(stats[1].sum, 60);
  assert.equal(stats[1].mean, 20);
  assert.equal(stats[1].median, 20);
  assert.equal(stats[1].blanks, 1);
  assert.equal(stats[1].stdev, 10);
  assert.equal(stats[1].topValues, undefined);
});

test("resolveColumnOffset accepts header, letter or offset", () => {
  const headers = ["Name", "Revenue", "Cost"];
  assert.equal(resolveColumnOffset("revenue", 2, headers), 1);
  assert.equal(resolveColumnOffset("E", 2, headers), 2);
  assert.equal(resolveColumnOffset(0, 2, headers), 0);
  assert.equal(resolveColumnOffset("1", 2, null), 1);
  assert.throws(() => resolveColumnOffset("Profit!", 0, headers));
});
