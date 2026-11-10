import { useState, type CSSProperties } from "react";
import { Button, Field, Switch, Text, Textarea } from "@fluentui/react-components";
import { validateKeys, type ModelInfo } from "../agent/gemini";
import { parseKeys } from "../lib/keyStore";

interface Props {
  readonly initialKeys?: readonly string[];
  readonly onReady: (apiKeys: readonly string[], models: readonly ModelInfo[]) => void;
  readonly saveKeys: (keys: readonly string[]) => void;
}

const mask = (key: string): string => (key.length > 10 ? `${key.slice(0, 6)}...${key.slice(-4)}` : "****");

const MASKED: CSSProperties = { WebkitTextSecurity: "disc" } as CSSProperties;

export const ApiKeyGate = ({ initialKeys = [], onReady, saveKeys }: Props) => {
  const [value, setValue] = useState(initialKeys.join("\n"));
  const [reveal, setReveal] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const keys = parseKeys(value);

  const submit = async () => {
    setBusy(true);
    setError(null);
    try {
      const result = await validateKeys(keys);
      if (result.rejected.length > 0) {
        throw new Error(`Not valid: ${result.rejected.map(mask).join(", ")}. Remove or fix them and try again.`);
      }
      saveKeys(result.keep);
      onReady(result.keep, result.models);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Could not validate keys.");
    } finally {
      setBusy(false);
    }
  };

  return (
    <div style={{ padding: 16, display: "grid", gap: 12 }}>
      <Text size={500} weight="semibold">Enter your Gemini API keys</Text>
      <Text size={200}>
        One key per line. When a key hits its limit, the next one is used automatically. Stored only on this
        computer and sent only to Google. Get free keys at{" "}
        <a href="https://aistudio.google.com/apikey" target="_blank" rel="noreferrer">
          aistudio.google.com/apikey
        </a>
        .
      </Text>
      <Field
        label={`API keys (${keys.length})`}
        validationMessage={error ?? undefined}
        validationState={error ? "error" : "none"}
      >
        <Textarea
          value={value}
          onChange={(_, d) => setValue(d.value)}
          placeholder={"AIza...\nAIza..."}
          rows={5}
          resize="vertical"
          autoComplete="off"
          spellCheck={false}
          textarea={{ style: reveal ? undefined : MASKED }}
        />
      </Field>
      <Switch label="Show keys" checked={reveal} onChange={(_, d) => setReveal(d.checked)} />
      <Button appearance="primary" disabled={keys.length === 0 || busy} onClick={() => void submit()}>
        {busy ? "Checking keys..." : "Save and continue"}
      </Button>
    </div>
  );
};
