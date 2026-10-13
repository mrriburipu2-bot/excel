// Workbook tools the assistant can call. Each tool has a JSON schema (sent to
// Claude) and a run() that performs the operation with Office.js.

import {
  type Cell,
  cellAddress,
  columnStats,
  resolveColumnOffset,
  splitAddress,
  toFormulaInput,
} from "../shared/excel-utils";
import type { JsonSchema } from "../shared/schema";
import type { Settings } from "../shared/settings";

export interface ToolContext {
  settings: Settings;
  /** Asks the user to approve an action in the task pane. */
  confirm(title: string, detail: string): Promise<boolean>;
}

export interface ToolDef {
  name: string;
  description: string;
  input_schema: JsonSchema & { type: "object" };
  run(input: any, ctx: ToolContext): Promise<unknown>;
}

// ---------- shared helpers ----------

function getSheet(context: Excel.RequestContext, name: string | null | undefined): Excel.Worksheet {
  return name ? context.workbook.worksheets.getItem(name) : context.workbook.worksheets.getActiveWorksheet();
}

/** Resolves "Sheet1!A1:B2", "A1:B2" (active sheet) or a defined name. */
function getRange(context: Excel.RequestContext, address: string): Excel.Range {
  const { sheet, ref } = splitAddress(address);
  return getSheet(context, sheet).getRange(ref);
}

const MAX_READ_CELLS = 4000;
const CELL_SCHEMA: JsonSchema = { type: ["string", "number", "boolean", "null"] };
const MATRIX_SCHEMA: JsonSchema = { type: "array", minItems: 1, items: { type: "array", items: CELL_SCHEMA } };
const ADDRESS = (what = "Range"): JsonSchema => ({
  type: "string",
  description: `${what} in A1 notation, optionally sheet-qualified, e.g. "Sheet1!A1:D20" or "'Q1 Sales'!B:B". Without a sheet name the active sheet is used.`,
});
const COLOR: JsonSchema = { type: "string", description: 'Color as hex "#RRGGBB" or a name like "red".' };

const OPERATOR_ENUM = ["between", "not_between", "equal", "not_equal", "greater", "less", "greater_or_equal", "less_or_equal"];
const CF_OPERATORS: Record<string, string> = {
  between: "Between",
  not_between: "NotBetween",
  equal: "EqualTo",
  not_equal: "NotEqualTo",
  greater: "GreaterThan",
  less: "LessThan",
  greater_or_equal: "GreaterThanOrEqual",
  less_or_equal: "LessThanOrEqual",
};
const DV_OPERATORS: Record<string, string> = {
  ...CF_OPERATORS,
  greater_or_equal: "GreaterThanOrEqualTo",
  less_or_equal: "LessThanOrEqualTo",
};

/** Reads a range, capping very large reads to the used area and MAX_READ_CELLS. */
async function readRange(context: Excel.RequestContext, address: string, maxCells: number, includeFormats: boolean) {
  let range = getRange(context, address);
  const used = range.worksheet.getUsedRangeOrNullObject();
  range = range.getIntersectionOrNullObject(used);
  range.load("isNullObject");
  await context.sync();
  if (range.isNullObject) return { address, empty: true, note: "Range is empty." };

  range.load("address,rowCount,columnCount,rowIndex,columnIndex");
  await context.sync();
  const cols = Math.min(range.columnCount, 200);
  const rows = Math.max(1, Math.min(range.rowCount, Math.floor(maxCells / cols)));
  const slice = range.getCell(0, 0).getResizedRange(rows - 1, cols - 1);
  slice.load(["address", "values", "formulas", ...(includeFormats ? ["numberFormat"] : [])]);
  await context.sync();

  const formulas: [string, string][] = [];
  slice.formulas.forEach((row, r) =>
    row.forEach((f, c) => {
      if (typeof f === "string" && f.startsWith("=")) {
        formulas.push([cellAddress(range.rowIndex + r, range.columnIndex + c), f]);
      }
    }),
  );
  return {
    address: slice.address,
    totalRows: range.rowCount,
    totalColumns: range.columnCount,
    truncated: rows < range.rowCount || cols < range.columnCount,
    values: slice.values,
    ...(formulas.length ? { formulas: formulas.slice(0, 500) } : {}),
    ...(includeFormats ? { numberFormat: slice.numberFormat } : {}),
  };
}

// ---------- tools ----------

export const TOOLS: ToolDef[] = [
  {
    name: "get_workbook_info",
    description:
      "Overview of the workbook: sheets (with used ranges), active sheet, current selection, tables, named ranges, charts and pivot tables. Call this first when you don't know the workbook layout.",
    input_schema: { type: "object", properties: {} },
    async run() {
      return Excel.run(async (context) => {
        const wb = context.workbook;
        const sheets = wb.worksheets;
        sheets.load("items/name,items/visibility,items/position");
        const active = sheets.getActiveWorksheet();
        active.load("name");
        const selection = wb.getSelectedRange();
        selection.load("address");
        const tables = wb.tables;
        tables.load("items/name,items/style,items/showTotals");
        const names = wb.names;
        names.load("items/name,items/formula,items/type");
        await context.sync();

        const perSheet = sheets.items.map((s) => {
          const used = s.getUsedRangeOrNullObject(true);
          used.load("address,rowCount,columnCount");
          s.charts.load("items/name,items/chartType");
          s.pivotTables.load("items/name");
          return { s, used };
        });
        const tableRanges = tables.items.map((t) => {
          const r = t.getRange();
          r.load("address");
          const h = t.getHeaderRowRange();
          h.load("values");
          return { t, r, h };
        });
        await context.sync();

        return {
          activeSheet: active.name,
          selection: selection.address,
          sheets: perSheet.map(({ s, used }) => ({
            name: s.name,
            visibility: s.visibility,
            usedRange: used.isNullObject ? null : used.address,
            rows: used.isNullObject ? 0 : used.rowCount,
            columns: used.isNullObject ? 0 : used.columnCount,
            charts: s.charts.items.map((c) => ({ name: c.name, type: c.chartType })),
            pivotTables: s.pivotTables.items.map((p) => p.name),
          })),
          tables: tableRanges.map(({ t, r, h }) => ({
            name: t.name,
            range: r.address,
            headers: h.values[0],
            style: t.style,
            showTotals: t.showTotals,
          })),
          namedRanges: names.items.map((n) => ({ name: n.name, refersTo: n.formula, type: n.type })),
        };
      });
    },
  },
  {
    name: "read_range",
    description:
      "Read cell values (and any formulas) from a range. Large ranges are clipped to the sheet's used area and to about 4000 cells; the result says when it was truncated. Use get_range_statistics to summarise big data instead of reading it all.",
    input_schema: {
      type: "object",
      properties: {
        address: ADDRESS(),
        include_number_formats: { type: "boolean", description: "Also return number formats." },
        max_cells: { type: "integer", description: "Cap on cells returned (default 4000)." },
      },
      required: ["address"],
    },
    run: (input) =>
      Excel.run((context) =>
        readRange(context, input.address, Math.min(input.max_cells ?? MAX_READ_CELLS, 20000), !!input.include_number_formats),
      ),
  },
  {
    name: "get_selection",
    description: "Return the address and contents of the cells the user currently has selected.",
    input_schema: { type: "object", properties: {} },
    async run() {
      return Excel.run(async (context) => {
        const sel = context.workbook.getSelectedRange();
        sel.load("address");
        await context.sync();
        return readRange(context, sel.address, 2000, true);
      });
    },
  },
  {
    name: "write_range",
    description:
      'Write a 2D array of values into the sheet starting at the top-left cell of `address`. Strings starting with "=" are entered as formulas (use English function names and commas). Rows may be ragged; missing cells are left blank. Use null to clear a cell.',
    input_schema: {
      type: "object",
      properties: {
        address: ADDRESS("Top-left cell (or the full target range)"),
        values: { ...MATRIX_SCHEMA, description: "Rows of cell values, e.g. [[\"Name\",\"Total\"],[\"A\",\"=SUM(B2:B9)\"]]." },
        number_format: { type: "string", description: 'Optional number format for the whole block, e.g. "#,##0.00", "0%", "yyyy-mm-dd", "$#,##0".' },
      },
      required: ["address", "values"],
    },
    async run(input) {
      const values = toFormulaInput(input.values as Cell[][]);
      return Excel.run(async (context) => {
        const target = getRange(context, input.address)
          .getCell(0, 0)
          .getResizedRange(values.length - 1, values[0].length - 1);
        target.formulas = values;
        if (input.number_format) target.numberFormat = values.map((r) => r.map(() => input.number_format));
        target.load("address");
        await context.sync();
        return { written: target.address, rows: values.length, columns: values[0].length };
      });
    },
  },
  {
    name: "fill_formula",
    description:
      "Put one formula in the first cell of a range and fill it down/across the whole range with relative references adjusted (like dragging the fill handle). Ideal for calculated columns.",
    input_schema: {
      type: "object",
      properties: {
        address: ADDRESS("Range to fill, e.g. \"Sheet1!E2:E500\""),
        formula: { type: "string", description: 'Formula for the first cell, e.g. "=C2*D2".' },
        number_format: { type: "string" },
      },
      required: ["address", "formula"],
    },
    async run(input) {
      return Excel.run(async (context) => {
        const range = getRange(context, input.address);
        const first = range.getCell(0, 0);
        first.formulas = [[input.formula]];
        range.load("address,cellCount");
        await context.sync();
        if (range.cellCount > 1) first.autoFill(range, Excel.AutoFillType.fillDefault);
        if (input.number_format) range.numberFormat = [[input.number_format]] as any;
        await context.sync();
        return { filled: range.address };
      });
    },
  },
  {
    name: "format_range",
    description:
      "Format cells: font, fill, number format, alignment, wrap, borders, column width/row height, autofit and merge. Only the properties you pass are changed.",
    input_schema: {
      type: "object",
      properties: {
        address: ADDRESS(),
        bold: { type: "boolean" },
        italic: { type: "boolean" },
        underline: { type: "boolean" },
        strikethrough: { type: "boolean" },
        font_color: COLOR,
        font_size: { type: "number" },
        font_name: { type: "string" },
        fill_color: { type: "string", description: 'Background color, or "none" to clear.' },
        number_format: { type: "string", description: 'e.g. "#,##0", "0.0%", "$#,##0.00", "dd-mmm-yyyy", "@" for text.' },
        horizontal_alignment: { type: "string", enum: ["General", "Left", "Center", "Right", "Fill", "Justify", "CenterAcrossSelection"] },
        vertical_alignment: { type: "string", enum: ["Top", "Center", "Bottom", "Justify"] },
        wrap_text: { type: "boolean" },
        indent: { type: "integer" },
        text_orientation: { type: "integer", description: "Rotation in degrees, -90 to 90." },
        borders: {
          type: "object",
          properties: {
            edges: { type: "string", enum: ["all", "outline", "inside", "top", "bottom", "left", "right", "none"] },
            style: { type: "string", enum: ["Continuous", "Dash", "DashDot", "Dot", "Double"] },
            weight: { type: "string", enum: ["Hairline", "Thin", "Medium", "Thick"] },
            color: COLOR,
          },
          required: ["edges"],
        },
        column_width: { type: "number", description: "Width in points." },
        row_height: { type: "number", description: "Height in points." },
        autofit: { type: "string", enum: ["columns", "rows", "both"] },
        merge: { type: "string", enum: ["merge", "merge_across", "unmerge"] },
      },
      required: ["address"],
    },
    async run(input) {
      return Excel.run(async (context) => {
        const r = getRange(context, input.address);
        const f = r.format;
        if (input.bold !== undefined) f.font.bold = input.bold;
        if (input.italic !== undefined) f.font.italic = input.italic;
        if (input.underline !== undefined) f.font.underline = input.underline ? "Single" : "None";
        if (input.strikethrough !== undefined) f.font.strikethrough = input.strikethrough;
        if (input.font_color) f.font.color = input.font_color;
        if (input.font_size) f.font.size = input.font_size;
        if (input.font_name) f.font.name = input.font_name;
        if (input.fill_color) {
          if (input.fill_color.toLowerCase() === "none") f.fill.clear();
          else f.fill.color = input.fill_color;
        }
        if (input.number_format) {
          r.load("rowCount,columnCount");
          await context.sync();
          r.numberFormat = Array.from({ length: r.rowCount }, () => Array(r.columnCount).fill(input.number_format));
        }
        if (input.horizontal_alignment) f.horizontalAlignment = input.horizontal_alignment;
        if (input.vertical_alignment) f.verticalAlignment = input.vertical_alignment;
        if (input.wrap_text !== undefined) f.wrapText = input.wrap_text;
        if (input.indent !== undefined) f.indentLevel = input.indent;
        if (input.text_orientation !== undefined) f.textOrientation = input.text_orientation;
        if (input.borders) {
          const edgeMap: Record<string, string[]> = {
            all: ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"],
            outline: ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"],
            inside: ["InsideHorizontal", "InsideVertical"],
            top: ["EdgeTop"],
            bottom: ["EdgeBottom"],
            left: ["EdgeLeft"],
            right: ["EdgeRight"],
            none: ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"],
          };
          for (const edge of edgeMap[input.borders.edges] ?? []) {
            const b = f.borders.getItem(edge as Excel.BorderIndex);
            if (input.borders.edges === "none") {
              b.style = "None";
              continue;
            }
            b.style = input.borders.style ?? "Continuous";
            if (input.borders.weight) b.weight = input.borders.weight;
            if (input.borders.color) b.color = input.borders.color;
          }
        }
        if (input.column_width) f.columnWidth = input.column_width;
        if (input.row_height) f.rowHeight = input.row_height;
        if (input.autofit === "columns" || input.autofit === "both") f.autofitColumns();
        if (input.autofit === "rows" || input.autofit === "both") f.autofitRows();
        if (input.merge === "merge") r.merge(false);
        if (input.merge === "merge_across") r.merge(true);
        if (input.merge === "unmerge") r.unmerge();
        r.load("address");
        await context.sync();
        return { formatted: r.address };
      });
    },
  },
  {
    name: "clear_range",
    description: "Clear contents, formats, or both from a range (cells stay in place).",
    input_schema: {
      type: "object",
      properties: {
        address: ADDRESS(),
        what: { type: "string", enum: ["all", "contents", "formats", "hyperlinks"] },
      },
      required: ["address"],
    },
    async run(input) {
      const applyTo = { all: "All", contents: "Contents", formats: "Formats", hyperlinks: "RemoveHyperlinks" }[
        (input.what ?? "contents") as string
      ] as Excel.ClearApplyTo;
      return Excel.run(async (context) => {
        const r = getRange(context, input.address);
        r.clear(applyTo);
        r.load("address");
        await context.sync();
        return { cleared: r.address, what: input.what ?? "contents" };
      });
    },
  },
  {
    name: "modify_rows_columns",
    description:
      "Insert, delete, hide, unhide, group, ungroup or autofit whole rows or columns. Address examples: \"5:7\" (rows 5-7), \"C:D\" (columns C-D), or any cell range whose rows/columns should be affected.",
    input_schema: {
      type: "object",
      properties: {
        address: ADDRESS(),
        dimension: { type: "string", enum: ["rows", "columns"] },
        action: { type: "string", enum: ["insert", "delete", "hide", "unhide", "group", "ungroup", "autofit"] },
      },
      required: ["address", "dimension", "action"],
    },
    async run(input) {
      return Excel.run(async (context) => {
        const base = getRange(context, input.address);
        const rows = input.dimension === "rows";
        const r = rows ? base.getEntireRow() : base.getEntireColumn();
        switch (input.action) {
          case "insert":
            r.insert(rows ? "Down" : "Right");
            break;
          case "delete":
            r.delete(rows ? "Up" : "Left");
            break;
          case "hide":
          case "unhide":
            if (rows) r.rowHidden = input.action === "hide";
            else r.columnHidden = input.action === "hide";
            break;
          case "group":
            r.group(rows ? "ByRows" : "ByColumns");
            break;
          case "ungroup":
            r.ungroup(rows ? "ByRows" : "ByColumns");
            break;
          case "autofit":
            if (rows) r.format.autofitRows();
            else r.format.autofitColumns();
            break;
        }
        await context.sync();
        return { done: `${input.action} ${input.dimension} ${input.address}` };
      });
    },
  },
  {
    name: "manage_worksheet",
    description:
      "Worksheet operations: add, delete, rename, activate, copy, move, hide, unhide, set tab color, freeze panes, unfreeze, protect, unprotect.",
    input_schema: {
      type: "object",
      properties: {
        action: {
          type: "string",
          enum: ["add", "delete", "rename", "activate", "copy", "move", "hide", "unhide", "tab_color", "freeze", "unfreeze", "protect", "unprotect"],
        },
        name: { type: "string", description: "Target sheet name (for add: the new sheet's name). Defaults to the active sheet." },
        new_name: { type: "string", description: "For rename and copy." },
        position: { type: "integer", description: "0-based position for move." },
        color: COLOR,
        freeze_rows: { type: "integer", description: "Number of top rows to freeze." },
        freeze_columns: { type: "integer", description: "Number of left columns to freeze." },
        password: { type: "string" },
      },
      required: ["action"],
    },
    async run(input, ctx) {
      if (input.action === "delete") {
        const ok = await ctx.confirm("Delete worksheet?", `Claude wants to delete sheet "${input.name ?? "(active sheet)"}". This cannot be undone.`);
        if (!ok) return { declined: "The user declined deleting the sheet." };
      }
      return Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        if (input.action === "add") {
          const s = sheets.add(input.name);
          s.activate();
          s.load("name");
          await context.sync();
          return { added: s.name };
        }
        const s = getSheet(context, input.name);
        switch (input.action) {
          case "delete":
            s.delete();
            break;
          case "rename":
            s.name = input.new_name;
            break;
          case "activate":
            s.activate();
            break;
          case "copy": {
            const copy = s.copy("End");
            if (input.new_name) copy.name = input.new_name;
            break;
          }
          case "move":
            s.position = input.position ?? 0;
            break;
          case "hide":
            s.visibility = "Hidden";
            break;
          case "unhide":
            s.visibility = "Visible";
            break;
          case "tab_color":
            s.tabColor = input.color ?? "";
            break;
          case "freeze": {
            const rows = input.freeze_rows ?? 0;
            const cols = input.freeze_columns ?? 0;
            s.freezePanes.unfreeze();
            if (rows && cols) s.freezePanes.freezeAt(s.getRangeByIndexes(0, 0, rows, cols));
            else if (rows) s.freezePanes.freezeRows(rows);
            else if (cols) s.freezePanes.freezeColumns(cols);
            break;
          }
          case "unfreeze":
            s.freezePanes.unfreeze();
            break;
          case "protect":
            s.protection.protect(undefined, input.password);
            break;
          case "unprotect":
            s.protection.unprotect(input.password);
            break;
        }
        await context.sync();
        return { done: input.action, sheet: input.name ?? "(active)" };
      });
    },
  },
  {
    name: "create_table",
    description: "Turn a range into an Excel table (ListObject) with filter buttons, banded rows and a style.",
    input_schema: {
      type: "object",
      properties: {
        address: ADDRESS("Data range including the header row"),
        has_headers: { type: "boolean" },
        name: { type: "string", description: "Table name (letters, digits, underscores; no spaces)." },
        style: { type: "string", description: 'e.g. "TableStyleMedium2", "TableStyleLight9", "TableStyleDark1".' },
        show_totals: { type: "boolean" },
      },
      required: ["address"],
    },
    async run(input) {
      return Excel.run(async (context) => {
        const { sheet, ref } = splitAddress(input.address);
        const ws = getSheet(context, sheet);
        const t = ws.tables.add(ws.getRange(ref), input.has_headers ?? true);
        if (input.name) t.name = input.name;
        t.style = input.style ?? "TableStyleMedium2";
        if (input.show_totals !== undefined) t.showTotals = input.show_totals;
        ws.getRange(ref).format.autofitColumns();
        t.load("name");
        await context.sync();
        return { table: t.name };
      });
    },
  },
  {
    name: "modify_table",
    description:
      "Change an existing table: set style, toggle totals row or banded rows, rename, append rows, set a column's totals function, convert to a plain range, or delete it.",
    input_schema: {
      type: "object",
      properties: {
        table: { type: "string", description: "Table name." },
        action: {
          type: "string",
          enum: ["style", "show_totals", "hide_totals", "banded_rows", "rename", "add_rows", "column_total", "convert_to_range", "delete"],
        },
        style: { type: "string" },
        banded: { type: "boolean" },
        new_name: { type: "string" },
        rows: { ...MATRIX_SCHEMA, description: "Rows to append for add_rows." },
        column: { type: "string", description: "Column header for column_total." },
        total_function: {
          type: "string",
          enum: ["SUM", "AVERAGE", "COUNT", "COUNTA", "MAX", "MIN", "STDEV", "VAR", "NONE"],
        },
      },
      required: ["table", "action"],
    },
    async run(input) {
      return Excel.run(async (context) => {
        const t = context.workbook.tables.getItem(input.table);
        switch (input.action) {
          case "style":
            t.style = input.style;
            break;
          case "show_totals":
          case "hide_totals":
            t.showTotals = input.action === "show_totals";
            break;
          case "banded_rows":
            t.showBandedRows = input.banded ?? true;
            break;
          case "rename":
            t.name = input.new_name;
            break;
          case "add_rows":
            t.rows.add(undefined, toFormulaInput(input.rows));
            break;
          case "column_total": {
            t.showTotals = true;
            const codes: Record<string, number> = { AVERAGE: 101, COUNT: 102, COUNTA: 103, MAX: 104, MIN: 105, STDEV: 107, SUM: 109, VAR: 110 };
            const cell = t.columns.getItem(input.column).getTotalRowRange();
            const fn = input.total_function ?? "SUM";
            cell.formulas = [[fn === "NONE" ? "" : `=SUBTOTAL(${codes[fn]},[${input.column}])`]];
            break;
          }
          case "convert_to_range":
            t.convertToRange();
            break;
          case "delete":
            t.delete();
            break;
        }
        await context.sync();
        return { done: input.action, table: input.table };
      });
    },
  },
  {
    name: "sort_range",
    description: "Sort a range (or table range) by one or more columns.",
    input_schema: {
      type: "object",
      properties: {
        address: ADDRESS("Range to sort, including headers if has_headers"),
        has_headers: { type: "boolean" },
        keys: {
          type: "array",
          minItems: 1,
          items: {
            type: "object",
            properties: {
              column: { type: ["string", "integer"], description: 'Column letter ("C"), header name ("Revenue"), or 0-based offset in the range.' },
              ascending: { type: "boolean" },
            },
            required: ["column"],
          },
        },
        match_case: { type: "boolean" },
      },
      required: ["address", "keys"],
    },
    async run(input) {
      return Excel.run(async (context) => {
        const r = getRange(context, input.address);
        r.load("columnIndex,address");
        const header = r.getRow(0);
        header.load("values");
        await context.sync();
        const hasHeaders = input.has_headers ?? true;
        const fields = (input.keys as { column: string | number; ascending?: boolean }[]).map((k) => ({
          key: resolveColumnOffset(k.column, r.columnIndex, hasHeaders ? (header.values[0] as Cell[]) : null),
          ascending: k.ascending ?? true,
        }));
        r.sort.apply(fields, input.match_case ?? false, hasHeaders, "Rows");
        await context.sync();
        return { sorted: r.address, by: fields };
      });
    },
  },
  {
    name: "filter_data",
    description:
      "Apply or clear an AutoFilter on a sheet range. Filter one column by a list of values, a custom condition (e.g. \">=100\", \"=*north*\"), or top/bottom N items.",
    input_schema: {
      type: "object",
      properties: {
        address: ADDRESS("Data range including headers"),
        column: { type: ["string", "integer"], description: "Column letter, header name, or 0-based offset." },
        values: { type: "array", items: { type: "string" }, description: "Keep rows whose value is one of these." },
        condition: { type: "string", description: 'Custom criterion, e.g. ">100", "<>0", "=*abc*".' },
        condition2: { type: "string" },
        operator: { type: "string", enum: ["and", "or"] },
        top_items: { type: "integer" },
        bottom_items: { type: "integer" },
        clear: { type: "boolean", description: "Remove the sheet's AutoFilter instead." },
      },
      required: ["address"],
    },
    async run(input) {
      return Excel.run(async (context) => {
        const r = getRange(context, input.address);
        const ws = r.worksheet;
        if (input.clear) {
          ws.autoFilter.remove();
          await context.sync();
          return { cleared: true };
        }
        r.load("columnIndex,address");
        const header = r.getRow(0);
        header.load("values");
        await context.sync();
        if (input.column === undefined) {
          ws.autoFilter.apply(r);
          await context.sync();
          return { autofilter: r.address };
        }
        const col = resolveColumnOffset(input.column, r.columnIndex, header.values[0] as Cell[]);
        let criteria: Excel.FilterCriteria;
        if (input.values) criteria = { filterOn: "Values", values: input.values };
        else if (input.top_items) criteria = { filterOn: "TopItems", criterion1: String(input.top_items) };
        else if (input.bottom_items) criteria = { filterOn: "BottomItems", criterion1: String(input.bottom_items) };
        else if (input.condition)
          criteria = {
            filterOn: "Custom",
            criterion1: input.condition,
            ...(input.condition2 ? { criterion2: input.condition2, operator: input.operator === "or" ? "Or" : "And" } : {}),
          };
        else throw new Error("Give values, condition, top_items or bottom_items (or clear: true).");
        ws.autoFilter.apply(r, col, criteria);
        await context.sync();
        return { filtered: r.address, columnOffset: col };
      });
    },
  },
  {
    name: "create_chart",
    description: "Create a chart from a data range (first row/column used as labels) and place it on a sheet.",
    input_schema: {
      type: "object",
      properties: {
        data_address: ADDRESS("Source data including headers"),
        chart_type: {
          type: "string",
          enum: [
            "ColumnClustered", "ColumnStacked", "ColumnStacked100", "BarClustered", "BarStacked", "Line", "LineMarkers",
            "LineStacked", "Pie", "Doughnut", "Area", "AreaStacked", "XYScatter", "XYScatterLines", "XYScatterSmooth",
            "Bubble", "Radar", "Waterfall", "Histogram", "Pareto", "BoxWhisker", "Funnel", "Treemap", "Sunburst", "3DColumn", "3DPie",
          ],
        },
        title: { type: "string" },
        series_by: { type: "string", enum: ["Auto", "Columns", "Rows"] },
        position: { type: "string", description: 'Cell range the chart should cover, e.g. "H2:P20". Can be on another sheet: "Dashboard!B2:J18".' },
        x_axis_title: { type: "string" },
        y_axis_title: { type: "string" },
        legend: { type: "string", enum: ["Right", "Bottom", "Top", "Left", "None"] },
        data_labels: { type: "boolean" },
        name: { type: "string" },
      },
      required: ["data_address", "chart_type"],
    },
    async run(input) {
      return Excel.run(async (context) => {
        const data = getRange(context, input.data_address);
        const pos = input.position ? splitAddress(input.position) : null;
        const host = pos?.sheet ? getSheet(context, pos.sheet) : data.worksheet;
        const chart = host.charts.add(input.chart_type, data, input.series_by ?? "Auto");
        if (input.title) chart.title.text = input.title;
        if (input.name) chart.name = input.name;
        if (pos) {
          const [start, end] = pos.ref.split(":");
          chart.setPosition(start, end);
        }
        if (input.legend === "None") chart.legend.visible = false;
        else if (input.legend) chart.legend.position = input.legend;
        if (input.data_labels) chart.dataLabels.showValue = true;
        if (input.x_axis_title) {
          chart.axes.categoryAxis.title.text = input.x_axis_title;
          chart.axes.categoryAxis.title.visible = true;
        }
        if (input.y_axis_title) {
          chart.axes.valueAxis.title.text = input.y_axis_title;
          chart.axes.valueAxis.title.visible = true;
        }
        chart.load("name");
        await context.sync();
        return { chart: chart.name };
      });
    },
  },
  {
    name: "add_conditional_format",
    description:
      "Add conditional formatting: color scale, data bar, icon set, cell-value rule, text-contains rule, top/bottom N, duplicate/unique values, or a custom formula rule. Use type \"clear\" to remove all conditional formats from the range.",
    input_schema: {
      type: "object",
      properties: {
        address: ADDRESS(),
        type: {
          type: "string",
          enum: ["color_scale", "data_bar", "icon_set", "cell_value", "text_contains", "top_bottom", "duplicates", "unique", "above_average", "below_average", "blanks", "errors", "formula", "clear"],
        },
        operator: { type: "string", enum: OPERATOR_ENUM },
        value1: { type: "string", description: 'First value or formula, e.g. "100" or "=$B$1".' },
        value2: { type: "string", description: "Second value for between/not_between." },
        text: { type: "string", description: "Text for text_contains." },
        formula: { type: "string", description: 'Formula for type formula, relative to the first cell, e.g. "=$D2>1000".' },
        rank: { type: "integer", description: "N for top_bottom." },
        top_bottom: { type: "string", enum: ["TopItems", "TopPercent", "BottomItems", "BottomPercent"] },
        icon_style: { type: "string", description: 'e.g. "ThreeTrafficLights1", "ThreeArrows", "FiveArrows", "ThreeSymbols", "FourRating", "ThreeStars".' },
        colors: { type: "array", items: { type: "string" }, description: "Color scale colors low→(mid)→high, 2 or 3 entries." },
        fill_color: COLOR,
        font_color: COLOR,
        bold: { type: "boolean" },
      },
      required: ["address", "type"],
    },
    async run(input) {
      return Excel.run(async (context) => {
        const r = getRange(context, input.address);
        const cfs = r.conditionalFormats;
        if (input.type === "clear") {
          cfs.clearAll();
          await context.sync();
          return { cleared: input.address };
        }
        const applyFormat = (fmt: Excel.ConditionalRangeFormat) => {
          if (input.fill_color) fmt.fill.color = input.fill_color;
          if (input.font_color) fmt.font.color = input.font_color;
          if (input.bold !== undefined) fmt.font.bold = input.bold;
          if (!input.fill_color && !input.font_color && input.bold === undefined) {
            fmt.fill.color = "#FFC7CE";
            fmt.font.color = "#9C0006";
          }
        };
        switch (input.type) {
          case "color_scale": {
            const cf = cfs.add("ColorScale");
            const c: string[] = input.colors ?? ["#F8696B", "#FFEB84", "#63BE7B"];
            cf.colorScale.criteria =
              c.length >= 3
                ? {
                    minimum: { formula: null as any, type: "LowestValue", color: c[0] },
                    midpoint: { formula: "50", type: "Percentile", color: c[1] },
                    maximum: { formula: null as any, type: "HighestValue", color: c[2] },
                  }
                : {
                    minimum: { formula: null as any, type: "LowestValue", color: c[0] },
                    maximum: { formula: null as any, type: "HighestValue", color: c[1] ?? "#63BE7B" },
                  };
            break;
          }
          case "data_bar": {
            const cf = cfs.add("DataBar");
            if (input.fill_color) cf.dataBar.positiveFormat.fillColor = input.fill_color;
            break;
          }
          case "icon_set": {
            const cf = cfs.add("IconSet");
            cf.iconSet.style = (input.icon_style ?? "ThreeTrafficLights1") as Excel.IconSet;
            break;
          }
          case "cell_value": {
            const cf = cfs.add("CellValue");
            cf.cellValue.rule = {
              formula1: input.value1 ?? "0",
              ...(input.value2 ? { formula2: input.value2 } : {}),
              operator: CF_OPERATORS[input.operator ?? "greater"] as Excel.ConditionalCellValueOperator,
            };
            applyFormat(cf.cellValue.format);
            break;
          }
          case "text_contains": {
            const cf = cfs.add("ContainsText");
            cf.textComparison.rule = { operator: "Contains", text: input.text ?? "" };
            applyFormat(cf.textComparison.format);
            break;
          }
          case "top_bottom": {
            const cf = cfs.add("TopBottom");
            cf.topBottom.rule = { rank: input.rank ?? 10, type: input.top_bottom ?? "TopItems" };
            applyFormat(cf.topBottom.format);
            break;
          }
          case "duplicates":
          case "unique":
          case "above_average":
          case "below_average":
          case "blanks":
          case "errors": {
            const criterion = {
              duplicates: "DuplicateValues",
              unique: "UniqueValues",
              above_average: "AboveAverage",
              below_average: "BelowAverage",
              blanks: "Blanks",
              errors: "Errors",
            }[input.type as string] as Excel.ConditionalFormatPresetCriterion;
            const cf = cfs.add("PresetCriteria");
            cf.preset.rule = { criterion };
            applyFormat(cf.preset.format);
            break;
          }
          case "formula": {
            const cf = cfs.add("Custom");
            cf.custom.rule.formula = input.formula;
            applyFormat(cf.custom.format);
            break;
          }
        }
        await context.sync();
        return { added: input.type, address: input.address };
      });
    },
  },
  {
    name: "set_data_validation",
    description:
      "Restrict what can be typed into cells: dropdown list, whole number, decimal, date, text length, or custom formula; optional input prompt and error alert. Use type \"clear\" to remove validation.",
    input_schema: {
      type: "object",
      properties: {
        address: ADDRESS(),
        type: { type: "string", enum: ["list", "whole_number", "decimal", "date", "text_length", "custom", "clear"] },
        list_source: { type: "string", description: 'Comma-separated items ("Yes,No,Maybe") or a range reference ("=Lists!$A$2:$A$20").' },
        operator: { type: "string", enum: OPERATOR_ENUM },
        value1: { type: "string" },
        value2: { type: "string" },
        formula: { type: "string", description: "Custom validation formula." },
        prompt_message: { type: "string" },
        error_message: { type: "string" },
      },
      required: ["address", "type"],
    },
    async run(input) {
      return Excel.run(async (context) => {
        const dv = getRange(context, input.address).dataValidation;
        dv.clear();
        if (input.type !== "clear") {
          const basic = {
            formula1: input.value1 ?? "0",
            ...(input.value2 ? { formula2: input.value2 } : {}),
            operator: DV_OPERATORS[input.operator ?? "between"] as Excel.DataValidationOperator,
          };
          const rule: Excel.DataValidationRule =
            input.type === "list"
              ? { list: { inCellDropDown: true, source: input.list_source ?? "" } }
              : input.type === "whole_number"
                ? { wholeNumber: basic }
                : input.type === "decimal"
                  ? { decimal: basic }
                  : input.type === "date"
                    ? { date: basic }
                    : input.type === "text_length"
                      ? { textLength: basic }
                      : { custom: { formula: input.formula ?? "" } };
          dv.rule = rule;
          if (input.prompt_message) dv.prompt = { message: input.prompt_message, showPrompt: true, title: "Input" };
          if (input.error_message)
            dv.errorAlert = { message: input.error_message, showAlert: true, style: "Stop", title: "Invalid entry" };
        }
        await context.sync();
        return { validation: input.type, address: input.address };
      });
    },
  },
  {
    name: "manage_named_range",
    description: "Add or delete a workbook-level defined name (named range or named formula).",
    input_schema: {
      type: "object",
      properties: {
        action: { type: "string", enum: ["add", "delete"] },
        name: { type: "string" },
        refers_to: { type: "string", description: 'e.g. "=Sheet1!$A$2:$A$100" or "=0.08".' },
        comment: { type: "string" },
      },
      required: ["action", "name"],
    },
    async run(input) {
      return Excel.run(async (context) => {
        if (input.action === "add") context.workbook.names.add(input.name, input.refers_to, input.comment);
        else context.workbook.names.getItem(input.name).delete();
        await context.sync();
        return { done: input.action, name: input.name };
      });
    },
  },
  {
    name: "create_pivot_table",
    description:
      "Create a PivotTable from a source range or table. Fields are header names from the source. The destination sheet is created if it does not exist.",
    input_schema: {
      type: "object",
      properties: {
        source: { type: "string", description: 'Source range ("Data!A1:F500") or table name.' },
        destination: { type: "string", description: 'Top-left cell, e.g. "Pivot!A3".' },
        name: { type: "string" },
        rows: { type: "array", items: { type: "string" } },
        columns: { type: "array", items: { type: "string" } },
        values: {
          type: "array",
          items: {
            type: "object",
            properties: {
              field: { type: "string" },
              summarize_by: {
                type: "string",
                enum: ["Sum", "Count", "Average", "Max", "Min", "Product", "CountNumbers", "StandardDeviation", "Variance"],
              },
              number_format: { type: "string" },
            },
            required: ["field"],
          },
        },
        filters: { type: "array", items: { type: "string" } },
      },
      required: ["source", "destination", "values"],
    },
    async run(input) {
      return Excel.run(async (context) => {
        const dest = splitAddress(input.destination);
        let destSheet: Excel.Worksheet;
        if (dest.sheet) {
          const existing = context.workbook.worksheets.getItemOrNullObject(dest.sheet);
          existing.load("isNullObject");
          await context.sync();
          destSheet = existing.isNullObject ? context.workbook.worksheets.add(dest.sheet) : existing;
        } else {
          destSheet = context.workbook.worksheets.getActiveWorksheet();
        }
        const source: Excel.Range | Excel.Table = input.source.includes("!") || /[A-Za-z]+\d+/.test(input.source)
          ? getRange(context, input.source)
          : context.workbook.tables.getItem(input.source);
        const name = input.name ?? `PivotTable${Date.now() % 100000}`;
        const pt = destSheet.pivotTables.add(name, source, destSheet.getRange(dest.ref));
        for (const f of input.rows ?? []) pt.rowHierarchies.add(pt.hierarchies.getItem(f));
        for (const f of input.columns ?? []) pt.columnHierarchies.add(pt.hierarchies.getItem(f));
        for (const f of input.filters ?? []) pt.filterHierarchies.add(pt.hierarchies.getItem(f));
        for (const v of input.values as { field: string; summarize_by?: string; number_format?: string }[]) {
          const dh = pt.dataHierarchies.add(pt.hierarchies.getItem(v.field));
          dh.summarizeBy = (v.summarize_by ?? "Sum") as Excel.AggregationFunction;
          if (v.number_format) dh.numberFormat = v.number_format;
        }
        destSheet.activate();
        await context.sync();
        return { pivotTable: name, at: input.destination };
      });
    },
  },
  {
    name: "find_replace",
    description: "Find cells containing text (returns their addresses), or replace text throughout a range or sheet.",
    input_schema: {
      type: "object",
      properties: {
        search: { type: "string" },
        replace: { type: "string", description: "Omit to only find." },
        address: ADDRESS("Where to search (defaults to the active sheet's used range)"),
        match_case: { type: "boolean" },
        whole_cell: { type: "boolean" },
      },
      required: ["search"],
    },
    async run(input) {
      return Excel.run(async (context) => {
        const r = input.address
          ? getRange(context, input.address)
          : context.workbook.worksheets.getActiveWorksheet().getUsedRange();
        const opts = { completeMatch: !!input.whole_cell, matchCase: !!input.match_case };
        if (input.replace !== undefined) {
          const count = r.replaceAll(input.search, input.replace, opts);
          await context.sync();
          return { replaced: count.value };
        }
        // findAll lives on the worksheet; narrow the hits to the requested range.
        const all = r.worksheet.findAllOrNullObject(input.search, opts);
        all.load("isNullObject");
        await context.sync();
        if (all.isNullObject) return { found: 0 };
        const found = all.getIntersectionOrNullObject(r);
        found.load("address,cellCount");
        await context.sync();
        return found.isNullObject ? { found: 0 } : { found: found.cellCount, addresses: found.address };
      });
    },
  },
  {
    name: "remove_duplicates",
    description: "Remove duplicate rows from a range, comparing the given columns.",
    input_schema: {
      type: "object",
      properties: {
        address: ADDRESS(),
        columns: { type: "array", items: { type: "integer" }, description: "0-based column offsets to compare; default all columns." },
        has_headers: { type: "boolean" },
      },
      required: ["address"],
    },
    async run(input) {
      return Excel.run(async (context) => {
        const r = getRange(context, input.address);
        r.load("columnCount");
        await context.sync();
        const cols = input.columns ?? Array.from({ length: r.columnCount }, (_, i) => i);
        const result = r.removeDuplicates(cols, input.has_headers ?? true);
        result.load("removed,uniqueRemaining");
        await context.sync();
        return { removed: result.removed, remaining: result.uniqueRemaining };
      });
    },
  },
  {
    name: "copy_range",
    description: "Copy (or move) a range to another location, optionally pasting only values, formats or formulas, or transposing.",
    input_schema: {
      type: "object",
      properties: {
        source: ADDRESS("Source range"),
        destination: ADDRESS("Destination top-left cell"),
        paste: { type: "string", enum: ["all", "values", "formats", "formulas"] },
        transpose: { type: "boolean" },
        skip_blanks: { type: "boolean" },
        move: { type: "boolean", description: "Cut instead of copy." },
      },
      required: ["source", "destination"],
    },
    async run(input) {
      return Excel.run(async (context) => {
        const src = getRange(context, input.source);
        const dest = getRange(context, input.destination);
        if (input.move) {
          src.moveTo(dest);
        } else {
          const type = { all: "All", values: "Values", formats: "Formats", formulas: "Formulas" }[
            (input.paste ?? "all") as string
          ] as Excel.RangeCopyType;
          dest.copyFrom(src, type, !!input.skip_blanks, !!input.transpose);
        }
        await context.sync();
        return { done: input.move ? "moved" : "copied", from: input.source, to: input.destination };
      });
    },
  },
  {
    name: "annotate_cell",
    description: "Add a comment and/or a hyperlink to a cell.",
    input_schema: {
      type: "object",
      properties: {
        cell: ADDRESS("Single cell"),
        comment: { type: "string" },
        url: { type: "string", description: "Web URL, or an internal reference like \"#Sheet2!A1\"." },
        display_text: { type: "string" },
      },
      required: ["cell"],
    },
    async run(input) {
      return Excel.run(async (context) => {
        const cell = getRange(context, input.cell);
        if (input.url) {
          const internal = input.url.startsWith("#");
          cell.hyperlink = internal
            ? { documentReference: input.url.slice(1), textToDisplay: input.display_text ?? input.url.slice(1) }
            : { address: input.url, textToDisplay: input.display_text ?? input.url };
        }
        if (input.comment) {
          cell.load("address");
          await context.sync();
          context.workbook.comments.add(cell.address, input.comment);
        }
        await context.sync();
        return { annotated: input.cell };
      });
    },
  },
  {
    name: "get_range_statistics",
    description:
      "Summarise each column of a data range: filled/blank counts, distinct values, numeric sum/min/max/mean/median/stdev, and top text values. Use this to understand large datasets cheaply.",
    input_schema: {
      type: "object",
      properties: { address: ADDRESS(), has_headers: { type: "boolean" } },
      required: ["address"],
    },
    async run(input) {
      return Excel.run(async (context) => {
        let r = getRange(context, input.address);
        r = r.getIntersectionOrNullObject(r.worksheet.getUsedRange());
        r.load("isNullObject");
        await context.sync();
        if (r.isNullObject) return { empty: true };
        r.load("values,address,rowCount,columnIndex");
        await context.sync();
        return {
          address: r.address,
          rows: r.rowCount,
          columns: columnStats(r.values as Cell[][], input.has_headers ?? true, r.columnIndex),
        };
      });
    },
  },
  {
    name: "select_range",
    description: "Select a range and scroll it into view so the user can see the result.",
    input_schema: { type: "object", properties: { address: ADDRESS() }, required: ["address"] },
    async run(input) {
      return Excel.run(async (context) => {
        const r = getRange(context, input.address);
        r.worksheet.activate();
        r.select();
        await context.sync();
        return { selected: input.address };
      });
    },
  },
  {
    name: "recalculate",
    description: "Force a full recalculation of the workbook.",
    input_schema: { type: "object", properties: {} },
    async run() {
      return Excel.run(async (context) => {
        context.workbook.application.calculate("Full");
        await context.sync();
        return { recalculated: true };
      });
    },
  },
];

/** Escape hatch for anything the other tools can't do. Only offered when enabled in Settings. */
export const RUN_CODE_TOOL: ToolDef = {
  name: "run_office_js",
  description:
    "Run Office.js code against the workbook for tasks no other tool covers (e.g. slicers, sparklines, shapes, page layout, chart series tweaks). The code is the body of an async function with `context` (Excel.RequestContext) and `Excel` in scope; call `await context.sync()` and `return` a JSON-serialisable result. The user must approve every run.",
  input_schema: {
    type: "object",
    properties: {
      explanation: { type: "string", description: "One sentence the user sees before approving." },
      code: { type: "string" },
    },
    required: ["explanation", "code"],
  },
  async run(input, ctx) {
    const ok = await ctx.confirm("Run code in your workbook?", `${input.explanation}\n\n${input.code}`);
    if (!ok) return { declined: "The user declined running this code." };
    const AsyncFunction = Object.getPrototypeOf(async () => {}).constructor;
    const fn = new AsyncFunction("context", "Excel", input.code);
    return Excel.run(async (context) => {
      const result = await fn(context, Excel);
      await context.sync();
      return result ?? { ok: true };
    });
  },
};

export function activeTools(settings: Settings): ToolDef[] {
  return settings.allowCode ? [...TOOLS, RUN_CODE_TOOL] : TOOLS;
}
