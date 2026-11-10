import { useEffect, useMemo, useState } from "react";
import { Button, Dropdown, Menu, MenuItem, MenuList, MenuPopover, MenuTrigger, Option, Text } from "@fluentui/react-components";
import { browserKeyStore } from "../lib/keyStore";
import { copyText } from "../lib/clipboard";
import { buildDiagnostics } from "../lib/diagnostics";
import { validateKeys, type ModelInfo } from "../agent/gemini";
import { pickDefaultModel } from "../agent/models";
import { ApiKeyGate } from "./ApiKeyGate";
import { Chat } from "./Chat";
import { SelectionPanel } from "./SelectionPanel";

interface Session {
  readonly apiKeys: readonly string[];
  readonly models: readonly ModelInfo[];
}

const APP_VERSION = "0.1.0";
const API_SETS = ["1.9", "1.18"] as const;

const collectDiagnostics = (modelId: string, keyCount: number): string => {
  const office = typeof Office !== "undefined" ? Office : undefined;
  const apiSets: Record<string, boolean> = {};
  for (const v of API_SETS) {
    try {
      apiSets[v] = office?.context?.requirements?.isSetSupported("ExcelApi", v) ?? false;
    } catch {
      apiSets[v] = false;
    }
  }
  return buildDiagnostics({
    appVersion: APP_VERSION,
    modelId,
    keyCount,
    host: office?.context?.host?.toString(),
    platform: office?.context?.platform?.toString(),
    apiSets,
    userAgent: navigator.userAgent,
    now: new Date(),
  });
};

export const App = () => {
  const store = useMemo(() => browserKeyStore(), []);
  const [session, setSession] = useState<Session | null>(null);
  const [modelId, setModelId] = useState<string>("");
  const [editing, setEditing] = useState(false);
  const [showSelection, setShowSelection] = useState(false);
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
    <div
      style={{
        padding: 8,
        display: "flex",
        flexDirection: "column",
        gap: 6,
        height: "100vh",
        boxSizing: "border-box",
        overflow: "hidden",
      }}
    >
      <div style={{ display: "flex", gap: 4, alignItems: "center", flexShrink: 0 }}>
        <Dropdown
          size="small"
          aria-label="Model"
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
        <Menu>
          <MenuTrigger disableButtonEnhancement>
            <Button size="small" appearance="subtle" aria-label="Settings" title="Settings">
              ⋯
            </Button>
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              <MenuItem onClick={() => setEditing(true)}>Manage API keys ({session.apiKeys.length})</MenuItem>
              <MenuItem onClick={() => setShowSelection((v) => !v)}>
                {showSelection ? "Hide selection summary" : "Show selection summary"}
              </MenuItem>
              <MenuItem onClick={() => void copyText(collectDiagnostics(modelId, session.apiKeys.length))}>Copy diagnostics</MenuItem>
              <MenuItem
                onClick={() => {
                  store.clear();
                  setSession(null);
                }}
              >
                Remove all API keys
              </MenuItem>
            </MenuList>
          </MenuPopover>
        </Menu>
      </div>

      {showSelection && (
        <div style={{ flexShrink: 0, maxHeight: 140, overflow: "auto" }}>
          <SelectionPanel />
        </div>
      )}

      <Chat key={`${session.apiKeys.join("|")}:${modelId}`} apiKeys={session.apiKeys} modelId={modelId} />
    </div>
  );
};
