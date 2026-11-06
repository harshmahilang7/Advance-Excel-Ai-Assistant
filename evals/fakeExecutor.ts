import type { ToolExecutor } from "../src/agent/loop";
import { describeCall } from "../src/agent/preview";
import { parseToolCall } from "../src/agent/tools";
import type { CellValue } from "../src/excel/context";
import type { RecordedCall } from "./score";

/** The fake workbook the model sees. Formulas, when present, line up cell for cell with values. */
export interface FakeSheet {
  readonly name: string;
  readonly address: string;
  readonly values: readonly (readonly CellValue[])[];
  readonly formulas?: readonly (readonly (string | null)[])[];
}

export interface FakeExecutor {
  readonly execute: ToolExecutor;
  /** Every call the model made, in order, with whether the tool accepted it. */
  readonly calls: () => readonly RecordedCall[];
}

/**
 * In-memory stand-in for the Excel executor. It validates arguments with the real schemas (so a model
 * that sends bad arguments is caught), answers reads from the fake sheet, and accepts every valid write
 * without changing anything.
 */
export const createFakeExecutor = (sheet: FakeSheet): FakeExecutor => {
  let recorded: readonly RecordedCall[] = [];
  const record = (name: string, args: unknown, ok: boolean) => {
    recorded = [...recorded, { name, args, ok }];
  };

  const execute: ToolExecutor = async (name, args) => {
    const parsed = parseToolCall(name, args);
    if (!parsed.ok) {
      record(name, args, false);
      return { ok: false, error: parsed.error };
    }
    record(name, parsed.call.args, true);
    switch (parsed.call.name) {
      case "get_workbook_info":
        return { ok: true, summary: "1 sheet(s)", data: [{ name: sheet.name, usedRange: sheet.address }] };
      case "read_range":
        return {
          ok: true,
          summary: `Read ${sheet.address}`,
          data: {
            address: sheet.address,
            totalRows: sheet.values.length,
            hiddenRowsSkipped: 0,
            columns: sheet.values[0]?.length ?? 0,
            values: sheet.values,
            formulas: sheet.formulas ?? sheet.values.map((row) => row.map(() => null)),
          },
        };
      case "audit_formulas":
        return { ok: true, summary: "Audited: 0 error value, 0 inconsistent formula, 0 hardcoded in formulas", data: { findings: [] } };
      default:
        return { ok: true, summary: `Simulated ${describeCall(parsed.call.name, parsed.call.args)}`, data: null };
    }
  };

  return { execute, calls: () => recorded };
};
