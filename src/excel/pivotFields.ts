export type PivotAggregation = "sum" | "count" | "average" | "max" | "min";

export interface PivotValueSpec {
  readonly field: string;
  readonly summarizeBy: PivotAggregation;
}

export interface PivotFieldRequest {
  readonly rows: readonly string[];
  readonly columns: readonly string[];
  readonly filters: readonly string[];
  readonly values: readonly PivotValueSpec[];
}

export type PivotFieldResolution =
  | ({ readonly ok: true } & PivotFieldRequest)
  | { readonly ok: false; readonly error: string };

const fail = (error: string): PivotFieldResolution => ({ ok: false, error });

/**
 * Match requested pivot fields against the source header row (case-insensitive, trimmed)
 * and return the exact header text Excel expects. Pure; explains failures to the model.
 */
export const resolvePivotFields = (headers: readonly string[], request: PivotFieldRequest): PivotFieldResolution => {
  if (headers.some((h) => h.trim() === "")) {
    return fail("The pivot source has a blank header cell. Every column in the source range needs a header.");
  }
  const lowered = headers.map((h) => h.trim().toLowerCase());
  if (new Set(lowered).size !== lowered.length) {
    return fail("The pivot source has duplicate header names. Rename them so each column header is unique.");
  }
  if (request.values.length === 0) return fail("A pivot table needs at least one value field.");
  if (request.rows.length + request.columns.length === 0) {
    return fail("A pivot table needs at least one row or column field.");
  }

  const available = headers.join(", ");
  const find = (name: string): string | null => {
    const i = lowered.indexOf(name.trim().toLowerCase());
    return i >= 0 ? (headers[i] ?? null) : null;
  };
  const missing = (name: string) => fail(`Field "${name}" not found in the source. Available fields: ${available}`);

  const mapAll = (names: readonly string[]): readonly string[] | { readonly missing: string } => {
    const out: string[] = [];
    for (const n of names) {
      const hit = find(n);
      if (hit === null) return { missing: n };
      out.push(hit);
    }
    return out;
  };

  const rows = mapAll(request.rows);
  if ("missing" in rows) return missing(rows.missing);
  const columns = mapAll(request.columns);
  if ("missing" in columns) return missing(columns.missing);
  const filters = mapAll(request.filters);
  if ("missing" in filters) return missing(filters.missing);

  const values: PivotValueSpec[] = [];
  for (const v of request.values) {
    const hit = find(v.field);
    if (hit === null) return missing(v.field);
    values.push({ field: hit, summarizeBy: v.summarizeBy });
  }

  return { ok: true, rows, columns, filters, values };
};
