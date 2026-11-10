/** The planning tool: the model lists its steps up front for multi-part requests. It changes nothing in the workbook. */

export const PLAN_TOOL = "propose_plan";

export const MAX_PLAN_STEPS = 8;
export const MAX_STEP_CHARS = 160;

export const isPlanTool = (name: string): boolean => name === PLAN_TOOL;

/** Human-readable numbered plan. Blank steps are dropped. */
export const formatPlan = (steps: readonly string[]): string => {
  const lines = steps
    .map((s) => s.trim())
    .filter((s) => s.length > 0)
    .map((s, i) => `${i + 1}. ${s}`);
  return ["Plan:", ...lines].join("\n");
};
