export interface ModelInfo {
  readonly id: string;
  readonly label: string;
}

export const DEFAULT_MODEL_ID = "gemini-3.8-flash";

/**
 * The only models offered in the picker, best first. Chosen for fast, reliable tool calling on
 * spreadsheet tasks. Image, speech, live, embedding, lite and deprecated models are deliberately
 * excluded, as is gemini-4-argon (it needs the Interactions API, not generateContent).
 */
const SUPPORTED: readonly ModelInfo[] = [
  { id: DEFAULT_MODEL_ID, label: "Gemini 3.8 Flash (recommended)" },
  { id: "gemini-3.6-flash", label: "Gemini 3.6 Flash (stable)" },
  { id: "gemini-3.1-pro-preview", label: "Gemini 3.1 Pro (preview, slower, harder tasks)" },
];

const stripPrefix = (id: string): string => id.replace(/^models\//, "");

/** Intersect the models a key can see with the supported list. Preserves preferred order. */
export const selectExcelModels = (available: readonly ModelInfo[]): readonly ModelInfo[] => {
  const ids = new Set(available.map((m) => stripPrefix(m.id)));
  return SUPPORTED.filter((m) => ids.has(m.id));
};

export const pickDefaultModel = (models: readonly ModelInfo[]): string =>
  models.find((m) => m.id === DEFAULT_MODEL_ID)?.id ?? models[0]?.id ?? "";
