import { z } from "zod";
import { CLEAN_OPERATIONS } from "../excel/cleaning";

/** A1-style address, optionally sheet-qualified: A1, $A$1:B9, A:A, Sheet1!A1, 'My Sheet'!A1:C3 */
const ADDRESS_PATTERN =
  /^(?:(?:'[^']{1,120}'|[A-Za-z0-9_.]{1,120})!)?\$?[A-Za-z]{1,3}\$?\d{0,7}(?::\$?[A-Za-z]{1,3}\$?\d{0,7})?$/;

const address = z.string().regex(ADDRESS_PATTERN, "Invalid address. Use A1 notation like A1, B2:D9 or Sheet1!A1:C5.");

const hexColor = z.string().regex(/^#[0-9A-Fa-f]{6}$/, "Colour must be hex like #FFEE00.");

const cell = z.union([z.string().max(32000), z.number(), z.boolean(), z.null()]);

const MAX_ROWS = 2000;
const MAX_CELLS = 20000;

const writeData = z
  .array(z.array(cell).min(1).max(200))
  .min(1)
  .max(MAX_ROWS)
  .refine((rows) => rows.reduce((n, r) => n + r.length, 0) <= MAX_CELLS, `At most ${MAX_CELLS} cells per write.`)
  .refine((rows) => rows.every((r) => r.length === rows[0]?.length), "All rows must have the same length.");

const chartTypes = [
  "ColumnClustered",
  "ColumnStacked",
  "BarClustered",
  "BarStacked",
  "Line",
  "LineMarkers",
  "Pie",
  "Area",
  "XYScatter",
  "Doughnut",
] as const;

const CELL_OPERATORS = [
  "GreaterThan",
  "LessThan",
  "GreaterThanOrEqual",
  "LessThanOrEqual",
  "EqualTo",
  "NotEqualTo",
  "Between",
  "NotBetween",
] as const;

/** Excel sheet names: 1-31 chars, none of \ / ? * [ ] :, and no leading/trailing apostrophe. */
const sheetName = z
  .string()
  .min(1)
  .max(31)
  .regex(/^[^\\/?*[\]:]+$/, "Sheet name cannot contain \\ / ? * [ ] :")
  .refine((s) => !s.startsWith("'") && !s.endsWith("'"), "Sheet name cannot start or end with an apostrophe.");

const VALIDATION_OPERATORS = [
  "Between",
  "NotBetween",
  "EqualTo",
  "NotEqualTo",
  "GreaterThan",
  "LessThan",
  "GreaterThanOrEqualTo",
  "LessThanOrEqualTo",
] as const;

const schemas = {
  get_workbook_info: z.object({}),
  read_range: z.object({
    address,
    maxRows: z.number().int().min(1).max(500).optional(),
  }),
  write_cells: z.object({
    address,
    data: writeData,
  }),
  format_range: z.object({
    address,
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    fontColor: hexColor.optional(),
    fillColor: hexColor.optional(),
    numberFormat: z.string().max(100).optional(),
    horizontalAlignment: z.enum(["Left", "Center", "Right"]).optional(),
    wrapText: z.boolean().optional(),
  }),
  sort_range: z.object({
    address,
    columnIndex: z.number().int().min(0).max(200),
    ascending: z.boolean(),
    hasHeaders: z.boolean(),
  }),
  clear_range: z.object({
    address,
    what: z.enum(["contents", "formats", "all"]),
  }),
  create_table: z.object({
    address,
    hasHeaders: z.boolean(),
    name: z.string().regex(/^[A-Za-z_][A-Za-z0-9_]{0,60}$/, "Table name: letters, digits, underscore.").optional(),
  }),
  create_chart: z.object({
    sourceAddress: address,
    chartType: z.enum(chartTypes),
    title: z.string().max(200).optional(),
  }),
  autofit_columns: z.object({ address }),
  audit_formulas: z.object({ address: address.optional() }),
  add_sheet: z.object({ name: sheetName }),
  rename_sheet: z.object({ currentName: sheetName, newName: sheetName }),
  freeze_panes: z.object({
    rows: z.number().int().min(0).max(100).default(0),
    columns: z.number().int().min(0).max(50).default(0),
  }),
  data_validation: z
    .object({
      address,
      rule: z.enum(["list", "wholeNumber", "decimal", "textLength", "clear"]),
      listValues: z.array(z.string().min(1).max(100)).min(1).max(100).optional(),
      operator: z.enum(VALIDATION_OPERATORS).optional(),
      value1: z.number().optional(),
      value2: z.number().optional(),
      errorMessage: z.string().max(255).optional(),
    })
    .superRefine((a, ctx) => {
      const fail = (message: string) => ctx.addIssue({ code: "custom", message });
      if (a.rule === "list") {
        if (!a.listValues) fail("list rule needs listValues.");
        else if (a.listValues.some((v) => v.includes(","))) fail("List items cannot contain commas.");
      } else if (a.rule !== "clear") {
        if (!a.operator || a.value1 === undefined) fail(`${a.rule} rule needs operator and value1.`);
        else if ((a.operator === "Between" || a.operator === "NotBetween") && a.value2 === undefined) fail(`${a.operator} needs value2.`);
      }
    }),
  remove_duplicates: z.object({
    address,
    columnIndexes: z.array(z.number().int().min(0).max(200)).min(1).max(200),
    hasHeaders: z.boolean(),
  }),
  filter_range: z
    .object({
      address,
      action: z.enum(["apply", "remove"]),
      columnIndex: z.number().int().min(0).max(200).optional(),
      values: z.array(z.string().max(255)).min(1).max(100).optional(),
    })
    .refine((a) => a.values === undefined || a.columnIndex !== undefined, "values needs columnIndex."),
  find_replace: z.object({
    address: address.optional(),
    find: z.string().min(1).max(255),
    replace: z.string().max(32000),
    matchCase: z.boolean().default(false),
    matchEntireCell: z.boolean().default(false),
  }),
  conditional_format: z
    .object({
      address,
      rule: z.enum(["cellValue", "text", "colorScale"]),
      operator: z.enum(CELL_OPERATORS).optional(),
      value1: z.union([z.string().max(255), z.number()]).optional(),
      value2: z.union([z.string().max(255), z.number()]).optional(),
      text: z.string().min(1).max(255).optional(),
      fillColor: hexColor.optional(),
      fontColor: hexColor.optional(),
      minColor: hexColor.optional(),
      midColor: hexColor.optional(),
      maxColor: hexColor.optional(),
    })
    .superRefine((a, ctx) => {
      const fail = (message: string) => ctx.addIssue({ code: "custom", message });
      const hasStyle = a.fillColor !== undefined || a.fontColor !== undefined;
      if (a.rule === "cellValue") {
        if (!a.operator || a.value1 === undefined) fail("cellValue rule needs operator and value1.");
        else if ((a.operator === "Between" || a.operator === "NotBetween") && a.value2 === undefined) fail(`${a.operator} needs value2.`);
        if (!hasStyle) fail("Give fillColor and/or fontColor.");
      } else if (a.rule === "text") {
        if (!a.text) fail("text rule needs text.");
        if (!hasStyle) fail("Give fillColor and/or fontColor.");
      } else if (!a.minColor || !a.maxColor) {
        fail("colorScale needs minColor and maxColor (midColor optional).");
      }
    }),
  clean_data: z
    .object({
      address,
      operations: z.array(z.enum(CLEAN_OPERATIONS)).min(1).max(CLEAN_OPERATIONS.length),
    })
    .refine(
      (a) => a.operations.filter((o) => o === "upper" || o === "lower" || o === "proper").length <= 1,
      "Pick at most one of upper, lower, proper.",
    ),
  create_pivot_table: z.object({
    sourceAddress: address,
    destinationAddress: address.optional(),
    visibleRowsOnly: z.boolean().default(true),
    name: z.string().regex(/^[A-Za-z_][A-Za-z0-9_ ]{0,60}$/, "Pivot name: letters, digits, underscore, space.").optional(),
    rows: z.array(z.string().min(1).max(255)).max(10),
    columns: z.array(z.string().min(1).max(255)).max(10).default([]),
    filters: z.array(z.string().min(1).max(255)).max(10).default([]),
    values: z
      .array(
        z.object({
          field: z.string().min(1).max(255),
          summarizeBy: z.enum(["sum", "count", "average", "max", "min"]),
        }),
      )
      .min(1)
      .max(10),
  }),
} as const;

type Schemas = typeof schemas;
export type ToolName = keyof Schemas;
export type ToolArgs<N extends ToolName> = z.infer<Schemas[N]>;
export type ToolCall = { [N in ToolName]: { readonly name: N; readonly args: ToolArgs<N> } }[ToolName];

export type ParseResult = { readonly ok: true; readonly call: ToolCall } | { readonly ok: false; readonly error: string };

const isToolName = (name: string): name is ToolName => Object.hasOwn(schemas, name);

/** Validates untrusted model output into a typed tool call. Never throws. */
export const parseToolCall = (name: string, args: unknown): ParseResult => {
  if (!isToolName(name)) return { ok: false, error: `Unknown tool "${name}".` };
  const parsed = schemas[name].safeParse(args ?? {});
  if (!parsed.success) {
    const issue = parsed.error.issues[0];
    const where = issue && issue.path.length > 0 ? `${issue.path.join(".")}: ` : "";
    return { ok: false, error: `Invalid arguments for ${name}. ${where}${issue?.message ?? "invalid"}` };
  }
  return { ok: true, call: { name, args: parsed.data } as ToolCall };
};

const descriptions: Record<ToolName, string> = {
  get_workbook_info: "List all worksheets with their used range addresses. Call this first if you do not know the layout.",
  read_range:
    "Read cell values and formulas from a range. Use to inspect data before changing it. Returns up to maxRows rows (default 100). Rows hidden by a filter are skipped; the result reports how many.",
  write_cells:
    "Write a 2D array of values starting at the top-left cell of address. A string starting with '=' is a formula. Prefer formulas over hard-coded results so values stay live.",
  format_range:
    "Apply formatting to a range: bold, italic, font colour, fill colour (hex like #FFEE00), number format (e.g. '0.00%', '$#,##0.00', 'yyyy-mm-dd'), alignment, wrap.",
  sort_range: "Sort a range by one column. columnIndex is 0-based within the range. Set hasHeaders true to keep the header row in place.",
  clear_range: "Clear contents, formats, or both from a range. Destructive; use only when the user asked.",
  create_table: "Convert a range into an Excel table with filters and banded rows.",
  create_chart: "Create a chart from a source range (include header row/column for labels) on the same sheet.",
  autofit_columns: "Auto-fit column widths for the columns covering the range.",
  audit_formulas:
    "Read-only check for spreadsheet problems in a range (default: the whole active sheet's used range, max 50,000 cells): cells showing errors (#REF!, #DIV/0!, #N/A...), formulas that break their column's pattern, and typed numbers inside a column of formulas. Returns A1 addresses with a short reason each. Findings are hints, not proof: totals rows often differ on purpose. Use read_range on flagged cells to explain or fix them.",
  add_sheet: "Add a new worksheet with the given name (max 31 chars, none of \\ / ? * [ ] :). Fails if the name is taken.",
  rename_sheet: "Rename an existing worksheet.",
  freeze_panes:
    "Freeze the top `rows` rows and/or the first `columns` columns of the active sheet so they stay visible while scrolling. rows=0 and columns=0 removes any freeze.",
  data_validation:
    "Restrict what can be typed into a range. rule 'list': dropdown from listValues (no commas inside items). rule 'wholeNumber' | 'decimal' | 'textLength': needs operator (Between, NotBetween, EqualTo, NotEqualTo, GreaterThan, LessThan, GreaterThanOrEqualTo, LessThanOrEqualTo) and value1 (value2 for Between/NotBetween). rule 'clear' removes validation. errorMessage is shown on bad input.",
  remove_duplicates:
    "Delete duplicate rows in a range, comparing the columns in columnIndexes (0-based within the range). The first occurrence is kept. Set hasHeaders true if the first row is a header. Destructive but undoable.",
  filter_range:
    "action 'apply': turn on filter dropdowns for the range's sheet, optionally keeping only rows where column columnIndex (0-based within the range) equals one of values. action 'remove': clear the sheet's filter and show all rows.",
  find_replace:
    "Find and replace text in cell contents. address limits the search; omit it to search the whole active sheet. replace may be '' to delete the text. matchCase and matchEntireCell default to false. Returns how many cells changed. Undoable.",
  conditional_format:
    "Add a conditional formatting rule that highlights cells automatically. rule 'cellValue': operator (GreaterThan, LessThan, GreaterThanOrEqual, LessThanOrEqual, EqualTo, NotEqualTo, Between, NotBetween) with value1 (and value2 for Between/NotBetween) plus fillColor and/or fontColor. rule 'text': cells containing text, plus fillColor and/or fontColor. rule 'colorScale': minColor and maxColor (optional midColor) gradient. Colours are hex like #FFC7CE (light red fill) with #9C0006 (dark red font). Numbers or text only for values; start a value with '=' for a cell reference like =$F$1.",
  clean_data:
    "Tidy messy data in a range in place (formulas are never touched). operations: 'trim' (cut leading/trailing/double spaces and non-breaking spaces), 'strip_nonprinting' (remove control and zero-width characters), 'fill_down' (fill blank cells from the cell above in the same column), 'text_to_number' (turn text like '1,200', '$5', '(45)', '12%' into real numbers), 'upper', 'lower', 'proper' (case; pick at most one). Operations run in a fixed order whatever order you list them. Returns how many cells each operation changed. Undoable. Use remove_duplicates for duplicate rows and find_replace for custom text fixes.",
  create_pivot_table:
    "Create a PivotTable from a source range whose first row is headers (every header non-blank and unique). rows/columns/filters/values use header names exactly. If destinationAddress is omitted, a new worksheet is created. summarizeBy: sum|count|average|max|min. Read the source headers first if unsure of names. If the source has rows hidden by a filter, only the visible rows are used (via a copied snapshot sheet) unless visibleRowsOnly is false. The source may be a whole sheet or whole columns; only the used part is read.",
};

export interface ToolDeclaration {
  readonly name: ToolName;
  readonly description: string;
  readonly parametersJsonSchema: Record<string, unknown>;
}

const toJsonSchema = (schema: z.ZodType): Record<string, unknown> => {
  const { $schema: _ignored, ...rest } = z.toJSONSchema(schema) as Record<string, unknown>;
  return rest;
};

export const toolDeclarations: readonly ToolDeclaration[] = (Object.keys(schemas) as ToolName[]).map((name) => ({
  name,
  description: descriptions[name],
  parametersJsonSchema: toJsonSchema(schemas[name]),
}));
