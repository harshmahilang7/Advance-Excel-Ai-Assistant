import { useState } from "react";
import { Button, Field, Input, Text } from "@fluentui/react-components";
import { listChatModels, type ModelInfo } from "../agent/gemini";

interface Props {
  readonly onReady: (apiKey: string, models: readonly ModelInfo[]) => void;
  readonly saveKey: (key: string) => void;
}

export const ApiKeyGate = ({ onReady, saveKey }: Props) => {
  const [value, setValue] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const submit = async () => {
    setBusy(true);
    setError(null);
    try {
      const key = value.trim();
      const models = await listChatModels(key);
      saveKey(key);
      onReady(key, models);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Could not validate key.");
    } finally {
      setBusy(false);
    }
  };

  return (
    <div style={{ padding: 16, display: "grid", gap: 12 }}>
      <Text size={500} weight="semibold">Enter your Gemini API key</Text>
      <Text size={200}>
        Stored only on this computer and sent only to Google. Get a free key at{" "}
        <a href="https://aistudio.google.com/apikey" target="_blank" rel="noreferrer">
          aistudio.google.com/apikey
        </a>
        .
      </Text>
      <Field
        label="API key"
        validationMessage={error ?? undefined}
        validationState={error ? "error" : "none"}
      >
        <Input
          type="password"
          value={value}
          onChange={(_, d) => setValue(d.value)}
          onKeyDown={(e) => e.key === "Enter" && value.trim() && !busy && void submit()}
          placeholder="AIza..."
          autoComplete="off"
        />
      </Field>
      <Button appearance="primary" disabled={!value.trim() || busy} onClick={() => void submit()}>
        {busy ? "Checking key..." : "Save and continue"}
      </Button>
    </div>
  );
};
