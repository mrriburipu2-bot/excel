// Pure helpers used by the workbook tools. No Office.js here so they can be unit tested in Node.

export type Cell = string | number | boolean | null;

/** 0 -> "A", 25 -> "Z", 26 -> "AA". */
export function columnLetter(index: number): string {
  let n = index + 1;
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** "A" -> 0, "AA" -> 26. Returns -1 when the text is not a column letter. */
export function columnIndex(letters: string): number {
  if (!/^[A-Za-z]{1,3}$/.test(letters)) return -1;
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

export function cellAddress(row: number, col: number): string {
  return `${columnLetter(col)}${row + 1}`;
}

/** Splits "'My Sheet'!A1:B2" into { sheet: "My Sheet", ref: "A1:B2" }. */
export function splitAddress(address: string): { sheet: string | null; ref: string } {
  const trimmed = address.trim();
  const bang = trimmed.lastIndexOf("!");
  if (bang < 0) return { sheet: null, ref: trimmed };
  let sheet = trimmed.slice(0, bang);
  if (sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
  return { sheet, ref: trimmed.slice(bang + 1) };
}

/** Pads ragged rows with nulls so every row has the same width. */
export function padMatrix(values: Cell[][]): Cell[][] {
  const width = Math.max(1, ...values.map((r) => r.length));
  return values.map((r) => {
    const row = r.map((v) => (v === undefined ? null : v));
    while (row.length < width) row.push(null);
    return row;
  });
}

/** Excel's formulas array wants "" for an empty cell, not null. */
export function toFormulaInput(values: Cell[][]): (string | number | boolean)[][] {
  return padMatrix(values).map((row) => row.map((v) => (v === null ? "" : v)));
}

export interface ColumnStats {
  column: string;
  header: string | null;
  filled: number;
  blanks: number;
  numeric: number;
  distinct: number;
  sum?: number;
  min?: number;
  max?: number;
  mean?: number;
  median?: number;
  stdev?: number;
  topValues?: { value: string; count: number }[];
}

export function columnStats(values: Cell[][], hasHeaders: boolean, firstColumn = 0): ColumnStats[] {
  if (values.length === 0) return [];
  const width = values[0].length;
  const body = hasHeaders ? values.slice(1) : values;
  const stats: ColumnStats[] = [];
  for (let c = 0; c < width; c++) {
    const nums: number[] = [];
    const counts = new Map<string, number>();
    let blanks = 0;
    for (const row of body) {
      const v = row[c];
      if (v === null || v === "") {
        blanks++;
        continue;
      }
      if (typeof v === "number") nums.push(v);
      const key = String(v);
      counts.set(key, (counts.get(key) ?? 0) + 1);
    }
    const s: ColumnStats = {
      column: columnLetter(firstColumn + c),
      header: hasHeaders ? String(values[0][c] ?? "") : null,
      filled: body.length - blanks,
      blanks,
      numeric: nums.length,
      distinct: counts.size,
    };
    if (nums.length > 0) {
      const sorted = [...nums].sort((a, b) => a - b);
      const sum = nums.reduce((a, b) => a + b, 0);
      const mean = sum / nums.length;
      const mid = Math.floor(sorted.length / 2);
      s.sum = round(sum);
      s.min = sorted[0];
      s.max = sorted[sorted.length - 1];
      s.mean = round(mean);
      s.median = round(sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2);
      s.stdev =
        nums.length > 1 ? round(Math.sqrt(nums.reduce((a, b) => a + (b - mean) ** 2, 0) / (nums.length - 1))) : 0;
    }
    if (nums.length < s.filled) {
      s.topValues = [...counts.entries()]
        .sort((a, b) => b[1] - a[1])
        .slice(0, 5)
        .map(([value, count]) => ({ value, count }));
    }
    stats.push(s);
  }
  return stats;
}

function round(n: number): number {
  return Math.round(n * 1e6) / 1e6;
}

/** Resolves a sort/filter column given as a letter ("C"), a header name, or a 0-based offset. */
export function resolveColumnOffset(
  column: string | number,
  rangeFirstColumn: number,
  headers: Cell[] | null,
): number {
  if (typeof column === "number") return column;
  if (headers) {
    const byHeader = headers.findIndex((h) => String(h ?? "").trim().toLowerCase() === column.trim().toLowerCase());
    if (byHeader >= 0) return byHeader;
  }
  const letter = columnIndex(column.trim());
  if (letter >= 0) return letter - rangeFirstColumn;
  if (/^\d+$/.test(column.trim())) return Number(column.trim());
  throw new Error(`Column "${column}" not found. Use a column letter, a header name, or a 0-based offset.`);
}
