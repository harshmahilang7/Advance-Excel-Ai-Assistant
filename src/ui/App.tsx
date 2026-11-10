import { useEffect, useMemo, useState } from "react";
import { Button, Dropdown, Option, Text } from "@fluentui/react-components";
import { browserKeyStore } from "../lib/keyStore";
import { validateKeys, type ModelInfo } from "../agent/gemini";
import { pickDefaultModel } from "../agent/models";
import { ApiKeyGate } from "./ApiKeyGate";
import { Chat } from "./Chat";
import { SelectionPanel } from "./SelectionPanel";

interface Session {
  readonly apiKeys: readonly string[];
  readonly models: readonly ModelInfo[];
}

export const App = () => {
  const store = useMemo(() => browserKeyStore(), []);
  const [session, setSession] = useState<Session | null>(null);
  const [modelId, setModelId] = useState<string>("");
  const [editing, setEditing] = useState(false);
  const [restoring, setRestoring] = useState(() => store.list().length > 0);

  const start = (apiKeys: readonly string[], models: readonly ModelInfo[]) => {
    setSession({ apiKeys, models });
    setModelId(pickDefaultModel(models));
    setEditing(false);
  };

  // Restore previously saved keys, re-validating them once on startup.
  useEffect(() => {
    const saved = store.list();
    if (saved.length === 0) return;
    let cancelled = false;
    validateKeys(saved)
      .then(({ keep, rejected, models }) => {
        // Forget only keys Google says are bad, not ones that hit a network blip or rate limit.
        if (rejected.length > 0) store.save(keep);
        if (!cancelled) start(keep, models);
      })
      .catch(() => undefined)
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

  if (!session || editing) {
    return (
      <ApiKeyGate initialKeys={session?.apiKeys ?? store.list()} saveKeys={store.save} onReady={start} />
    );
  }

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

      <Chat key={`${session.apiKeys.join("|")}:${modelId}`} apiKeys={session.apiKeys} modelId={modelId} />

      <details>
        <summary style={{ cursor: "pointer", fontSize: 12 }}>Debug: selection summary</summary>
        <SelectionPanel />
      </details>
      <Button
        size="small"
        appearance="subtle"
        onClick={() => setEditing(true)}
      >
        Manage API keys ({session.apiKeys.length})
      </Button>
      <Button
        size="small"
        appearance="subtle"
        onClick={() => {
          store.clear();
          setSession(null);
        }}
      >
        Remove all API keys
      </Button>
    </div>
  );
};
