import { z } from "zod";

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
  create_pivot_table: z.object({
    sourceAddress: address,
    destinationAddress: address.optional(),
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
    "Read cell values and formulas from a range. Use to inspect data before changing it. Returns up to maxRows rows (default 100).",
  write_cells:
    "Write a 2D array of values starting at the top-left cell of address. A string starting with '=' is a formula. Prefer formulas over hard-coded results so values stay live.",
  format_range:
    "Apply formatting to a range: bold, italic, font colour, fill colour (hex like #FFEE00), number format (e.g. '0.00%', '$#,##0.00', 'yyyy-mm-dd'), alignment, wrap.",
  sort_range: "Sort a range by one column. columnIndex is 0-based within the range. Set hasHeaders true to keep the header row in place.",
  clear_range: "Clear contents, formats, or both from a range. Destructive; use only when the user asked.",
  create_table: "Convert a range into an Excel table with filters and banded rows.",
  create_chart: "Create a chart from a source range (include header row/column for labels) on the same sheet.",
  autofit_columns: "Auto-fit column widths for the columns covering the range.",
  create_pivot_table:
    "Create a PivotTable from a source range whose first row is headers (every header non-blank and unique). rows/columns/filters/values use header names exactly. If destinationAddress is omitted, a new worksheet is created. summarizeBy: sum|count|average|max|min. Read the source headers first if unsure of names.",
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
