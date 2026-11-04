import { useEffect, useMemo, useState } from "react";
import { Button, Dropdown, Option, Text } from "@fluentui/react-components";
import { browserKeyStore } from "../lib/keyStore";
import { listChatModels, type ModelInfo } from "../agent/gemini";
import { ApiKeyGate } from "./ApiKeyGate";
import { Chat } from "./Chat";
import { SelectionPanel } from "./SelectionPanel";

interface Session {
  readonly apiKey: string;
  readonly models: readonly ModelInfo[];
}

/** Prefer a fast "flash" model by default; the user can switch in the dropdown. */
const pickDefaultModel = (models: readonly ModelInfo[]): string =>
  models.find((m) => /flash/.test(m.id) && !/(lite|image|live|tts|preview|exp)/.test(m.id))?.id ?? models[0]?.id ?? "";

export const App = () => {
  const store = useMemo(() => browserKeyStore(), []);
  const [session, setSession] = useState<Session | null>(null);
  const [modelId, setModelId] = useState<string>("");
  const [restoring, setRestoring] = useState(() => store.get() !== null);

  const start = (apiKey: string, models: readonly ModelInfo[]) => {
    setSession({ apiKey, models });
    setModelId(pickDefaultModel(models));
  };

  // Restore a previously saved key, re-validating it once on startup.
  useEffect(() => {
    const saved = store.get();
    if (!saved) return;
    let cancelled = false;
    listChatModels(saved)
      .then((models) => !cancelled && start(saved, models))
      .catch(() => store.clear())
      .finally(() => !cancelled && setRestoring(false));
    return () => {
      cancelled = true;
    };
  }, [store]);

  if (restoring) {
    return (
      <div style={{ padding: 16 }}>
        <Text>Loading...</Text>
      </div>
    );
  }

  if (!session) return <ApiKeyGate saveKey={store.set} onReady={start} />;

  return (
    <div style={{ padding: 12, display: "flex", flexDirection: "column", gap: 8, height: "100vh", boxSizing: "border-box" }}>
      <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
        <Text weight="semibold">Excel-Ai-Assistant</Text>
        <Dropdown
          size="small"
          style={{ flex: 1, minWidth: 0 }}
          value={session.models.find((m) => m.id === modelId)?.label ?? modelId}
          selectedOptions={[modelId]}
          onOptionSelect={(_, d) => setModelId(d.optionValue ?? modelId)}
        >
          {session.models.map((m) => (
            <Option key={m.id} value={m.id} text={m.label}>
              {m.label}
            </Option>
          ))}
        </Dropdown>
      </div>

      <Chat key={`${session.apiKey}:${modelId}`} apiKey={session.apiKey} modelId={modelId} />

      <details>
        <summary style={{ cursor: "pointer", fontSize: 12 }}>Debug: selection summary</summary>
        <SelectionPanel />
      </details>
      <Button
        size="small"
        appearance="subtle"
        onClick={() => {
          store.clear();
          setSession(null);
        }}
      >
        Remove API key
      </Button>
    </div>
  );
};
